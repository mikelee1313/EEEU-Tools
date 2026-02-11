<#
.SYNOPSIS
    Script to remove Everyone Except External Users (EEEU) permissions from SharePoint objects listed in a CSV.

.DESCRIPTION
    This script processes a list of SharePoint objects (files, folders, lists, webs) from a CSV file and removes 
    the "Everyone except external users" permissions from each object. It uses PnP PowerShell with certificate-based 
    authentication to connect to SharePoint Online sites and modify permissions.

.PARAMETER None
    This script uses hardcoded values for authentication and file paths.

.NOTES
    File Name      : Remove-EEEUFromFileList.ps1
    Author         : Mike Lee
    Date           : 9/22/25

.DISCLAIMER
    Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
    
.INPUTS
    CSV file with columns for URL (site URL), ItemURL (object path within the site), and Type (File, Folder, List, Web)
    Located at: C:\temp\EEEUFileList.csv

.OUTPUTS
    Log file created in the user's temp directory with timestamp in the filename
    Console output with color-coded status messages

.EXAMPLE
    .\Remove-EEEUFromFileList.ps1

.FUNCTIONALITY
    1. Connects to SharePoint Online sites using app-based authentication
    2. Reads a CSV file containing site URLs, object paths, and object types
    3. For each object, retrieves its permissions
    4. Removes any permissions assigned to the "Everyone except external users" group
    5. Handles different types of SharePoint objects: Files, Folders, Lists, and Webs
    6. Logs all operations to a log file
#>
# Tenant Level Information
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID

# Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"    # User principal name for "Everyone except external users"
$csvFilePath = "C:\Temp\Find_EEEU_In_Sites_20250625_152020.csv" # Path to the CSV file containing file list
$startime = Get-Date -Format "yyyyMMdd_HHmmss"                  # Timestamp for log file
$logFilePath = "$env:TEMP\Remove_EEEU_From_File_List_$startime.txt" # Path to the log file

# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

# Connect to SharePoint Online
function Connect-ToSharePoint {
    param (
        [string]$SiteURL
    )
    try {
        Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        Write-Log "Connected to SharePoint Online: $SiteURL"
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online: $_" "ERROR"
        throw $_
    }
}

# Get the file
function Get-File {
    param (
        [string]$FilePath
    )
    try {
        $file = Get-PnPFile -Url $FilePath -AsListItem
        Write-Log "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL"
        return $file
    }
    catch {
        Write-Log "Initial Get-PnPFile failed for: $FilePath - $_" "WARNING"
        Write-Log "Attempting fallback: resolving file via parent folder and list item query..." "INFO"
        
        try {
            # Extract the folder path and file name from the full path
            $folderPath = $FilePath.Substring(0, $FilePath.LastIndexOf("/"))
            $fileName = $FilePath.Substring($FilePath.LastIndexOf("/") + 1)
            
            # Get the folder's parent list, then query for the file by name within that folder
            $folder = Get-PnPFolder -Url $folderPath -Includes ListItemAllFields
            $parentList = Get-PnPProperty -ClientObject $folder -Property ListItemAllFields
            $listId = Get-PnPProperty -ClientObject $folder.ListItemAllFields -Property ParentList
            $list = Get-PnPList -Identity $listId.Id
            
            # Use a CAML query to find the file by its name and folder path
            $camlQuery = "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>$fileName</Value></Eq><Eq><FieldRef Name='FileDirRef'/><Value Type='Text'>$folderPath</Value></Eq></And></Where></Query></View>"
            $items = Get-PnPListItem -List $list -Query $camlQuery
            
            if ($items -and $items.Count -gt 0) {
                $file = $items[0]
                Write-Log "Retrieved file via fallback: $($file.FieldValues.FileLeafRef) on $SiteURL"
                return $file
            }
            else {
                Write-Log "Fallback: File not found via CAML query: $FilePath" "ERROR"
                throw "File not found: $FilePath"
            }
        }
        catch {
            Write-Log "Fallback also failed for file: $FilePath - $_" "ERROR"
            throw $_
        }
    }
}

# Get the folder
function Get-Folder {
    param (
        [string]$FolderPath
    )
    try {
        $folder = Get-PnPFolder -Url $FolderPath -Includes ListItemAllFields
        Write-Log "Retrieved folder: $($folder.Name) on $SiteURL"
        return $folder.ListItemAllFields
    }
    catch {
        Write-Log "Failed to retrieve folder: $_" "ERROR"
        throw $_
    }
}

# Get the list
function Get-SPList {
    param (
        [string]$ListPath
    )
    try {
        # Check if this is a document library view path and fix it
        if ($ListPath -match "(.+)/Forms/AllItems\.aspx$") {
            $ListPath = $Matches[1]
            Write-Log "Modified document library path to: $ListPath"
        }
        
        # Extract list name from path
        $listName = $ListPath
        if ($ListPath -match "/Lists/([^/]+)") {
            $listName = $Matches[1]
        }
        elseif ($ListPath -match "/([^/]+)$") {
            $listName = $Matches[1]
        }
        
        $list = Get-PnPList -Identity $listName -Includes RoleAssignments
        Write-Log "Retrieved list: $($list.Title) on $SiteURL"
        return $list
    }
    catch {
        Write-Log "Failed to retrieve list: $_" "ERROR"
        throw $_
    }
}

# Get the web
function Get-SPWeb {
    param (
        [string]$WebPath
    )
    try {
        # For web, we might already be connected to it via the SiteURL
        # or it could be a subweb
        if ([string]::IsNullOrEmpty($WebPath) -or $WebPath -eq "/") {
            # It's the root web
            $web = Get-PnPWeb -Includes RoleAssignments
        }
        else {
            # It's a subweb - using -Identity instead of -Url
            $web = Get-PnPWeb -Identity $WebPath -Includes RoleAssignments
        }
        
        Write-Log "Retrieved web: $($web.Title) on $SiteURL"
        return $web
    }
    catch {
        Write-Log "Failed to retrieve web: $_" "ERROR"
        throw $_
    }
}

# Remove EEEU from file
function Remove-EEEUfromFile {
    param (
        [Microsoft.SharePoint.Client.ListItem]$file
    )
    try {
        $RoleAssignment = @()
        $Permissions = Get-PnPProperty -ClientObject $file -Property RoleAssignments
        foreach ($RoleAssignment in $Permissions) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            if ($RoleAssignment.Member.LoginName -eq $LoginName) {
                $roleuser = $RoleAssignment.Member.LoginName
                $rolelevel = $RoleAssignment.RoleDefinitionBindings
                foreach ($role in $rolelevel) {
                    Write-Host "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Yellow
                    Write-Log "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL"
                    Set-PnPListItemPermission -List $file.ParentList.Title -Identity $file.Id -RemoveRole $role.Name -User $roleuser
                    Write-Host "Removed Role: $($role.Name) from User: $($roleuser) in File: $($file.FieldValues.FileLeafRef)" -ForegroundColor cyan
                    Write-Log "Removed Role: $($role.Name) from User: $($roleuser) in File: $($file.FieldValues.FileLeafRef)"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to remove EEEU from file: $_" "ERROR"
        throw $_
    }
}

# Remove EEEU from folder
function Remove-EEEUfromFolder {
    param (
        [Microsoft.SharePoint.Client.ListItem]$folder
    )
    try {
        $RoleAssignment = @()
        $Permissions = Get-PnPProperty -ClientObject $folder -Property RoleAssignments
        foreach ($RoleAssignment in $Permissions) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            if ($RoleAssignment.Member.LoginName -eq $LoginName) {
                $roleuser = $RoleAssignment.Member.LoginName
                $rolelevel = $RoleAssignment.RoleDefinitionBindings
                foreach ($role in $rolelevel) {
                    Write-Host "Retrieved folder: $($folder.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Yellow
                    Write-Log "Retrieved folder: $($folder.FieldValues.FileLeafRef) on $SiteURL"
                    Set-PnPListItemPermission -List $folder.ParentList.Title -Identity $folder.Id -RemoveRole $role.Name -User $roleuser
                    Write-Host "Removed Role: $($role.Name) from User: $($roleuser) in Folder: $($folder.FieldValues.FileLeafRef)" -ForegroundColor cyan
                    Write-Log "Removed Role: $($role.Name) from User: $($roleuser) in Folder: $($folder.FieldValues.FileLeafRef)"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to remove EEEU from folder: $_" "ERROR"
        throw $_
    }
}

# Remove EEEU from list
function Remove-EEEUfromList {
    param (
        [Microsoft.SharePoint.Client.List]$list
    )
    try {
        $RoleAssignment = @()
        $Permissions = Get-PnPProperty -ClientObject $list -Property RoleAssignments
        foreach ($RoleAssignment in $Permissions) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            if ($RoleAssignment.Member.LoginName -eq $LoginName) {
                $roleuser = $RoleAssignment.Member.LoginName
                $rolelevel = $RoleAssignment.RoleDefinitionBindings
                foreach ($role in $rolelevel) {
                    Write-Host "Retrieved list: $($list.Title) on $SiteURL" -ForegroundColor Yellow
                    Write-Log "Retrieved list: $($list.Title) on $SiteURL"
                    Set-PnPListPermission -Identity $list.Title -RemoveRole $role.Name -User $roleuser
                    Write-Host "Removed Role: $($role.Name) from User: $($roleuser) in List: $($list.Title)" -ForegroundColor cyan
                    Write-Log "Removed Role: $($role.Name) from User: $($roleuser) in List: $($list.Title)"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to remove EEEU from list: $_" "ERROR"
        throw $_
    }
}

# Remove EEEU from web
function Remove-EEEUfromWeb {
    param (
        [Microsoft.SharePoint.Client.Web]$web
    )
    try {
        $RoleAssignment = @()
        $Permissions = Get-PnPProperty -ClientObject $web -Property RoleAssignments
        foreach ($RoleAssignment in $Permissions) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            if ($RoleAssignment.Member.LoginName -eq $LoginName) {
                $roleuser = $RoleAssignment.Member.LoginName
                $rolelevel = $RoleAssignment.RoleDefinitionBindings
                foreach ($role in $rolelevel) {
                    Write-Host "Retrieved web: $($web.Title) on $SiteURL" -ForegroundColor Yellow
                    Write-Log "Retrieved web: $($web.Title) on $SiteURL"
                    Set-PnPWebPermission -Identity $web -RemoveRole $role.Name -User $roleuser
                    Write-Host "Removed Role: $($role.Name) from User: $($roleuser) in Web: $($web.Title)" -ForegroundColor cyan
                    Write-Log "Removed Role: $($role.Name) from User: $($roleuser) in Web: $($web.Title)"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to remove EEEU from web: $_" "ERROR"
        throw $_
    }
}

# Resolve the correct web URL for an item path (handles classic subsites)
function Resolve-SubsiteURL {
    param (
        [string]$SiteURL,
        [string]$ObjectPath
    )
    try {
        # Connect to the provided site URL first
        Connect-ToSharePoint -SiteURL $SiteURL

        # Use GetSubwebs to get all direct child webs
        $web = Get-PnPWeb -Includes Webs
        $subwebs = Get-PnPProperty -ClientObject $web -Property Webs

        # Check if any subweb's ServerRelativeUrl is a prefix of the ObjectPath
        foreach ($subweb in $subwebs) {
            $subwebUrl = $subweb.ServerRelativeUrl
            if ($ObjectPath.StartsWith($subwebUrl + "/") -or $ObjectPath -eq $subwebUrl) {
                # Found a matching subsite - resolve its full URL and recurse deeper
                $subSiteFullURL = $SiteURL.TrimEnd('/').Replace($web.ServerRelativeUrl.TrimEnd('/'), '') + $subwebUrl
                Write-Log "Detected subsite: $subSiteFullURL for path: $ObjectPath"
                Write-Host "Detected subsite: $subSiteFullURL" -ForegroundColor Magenta
                
                # Recurse to check for deeper subsites
                return Resolve-SubsiteURL -SiteURL $subSiteFullURL -ObjectPath $ObjectPath
            }
        }

        # No matching subweb found - the current SiteURL is the correct web
        Write-Log "Resolved web URL for ObjectPath '$ObjectPath': $SiteURL"
        return $SiteURL
    }
    catch {
        Write-Log "Error resolving subsite for path $ObjectPath from $SiteURL - using original URL: $_" "WARNING"
        return $SiteURL
    }
}

# Process each object based on its type
function Invoke-SharePointObjectProcessing {
    param (
        [string]$SiteURL,
        [string]$ObjectPath,
        [string]$ObjectType
    )
    
    try {
        # Resolve the correct web URL (handles classic subsites)
        $resolvedSiteURL = Resolve-SubsiteURL -SiteURL $SiteURL -ObjectPath $ObjectPath
        if ($resolvedSiteURL -ne $SiteURL) {
            Write-Host "Resolved subsite URL: $resolvedSiteURL (original: $SiteURL)" -ForegroundColor Magenta
            Write-Log "Resolved subsite URL from $SiteURL to $resolvedSiteURL for path $ObjectPath"
            $SiteURL = $resolvedSiteURL
        }

        Connect-ToSharePoint -SiteURL $SiteURL
        
        Write-Host "Processing $ObjectType object: $ObjectPath" -ForegroundColor Yellow
        Write-Log "Processing $ObjectType object: $ObjectPath"
        
        # Clean up document library paths if needed
        if ($ObjectType -eq "List" -or $ObjectType -eq "DocLib") {
            if ($ObjectPath -match "(.+)/Forms/AllItems\.aspx$") {
                $cleanPath = $Matches[1]
                Write-Log "Cleaning document library path from: $ObjectPath to: $cleanPath"
                $ObjectPath = $cleanPath
            }
        }
        
        switch ($ObjectType) {
            "File" {
                $object = Get-File -FilePath $ObjectPath
                if ($object) {
                    # Expand the ParentList property if needed
                    Get-PnPProperty -ClientObject $object -Property ParentList | Out-Null
                    Remove-EEEUfromFile -file $object | Out-Null
                }
            }
            "Folder" {
                $object = Get-Folder -FolderPath $ObjectPath
                if ($object) {
                    # Expand the ParentList property if needed
                    Get-PnPProperty -ClientObject $object -Property ParentList | Out-Null
                    Remove-EEEUfromFolder -folder $object | Out-Null
                }
            }
            "List" {
                $object = Get-SPList -ListPath $ObjectPath
                if ($object) {
                    Remove-EEEUfromList -list $object | Out-Null
                }
            }
            "DocLib" {
                # Handle Document Libraries specifically
                $object = Get-SPList -ListPath $ObjectPath
                if ($object) {
                    Remove-EEEUfromList -list $object | Out-Null
                }
            }
            "Web" {
                # For Web objects, if the ObjectPath is a full URL, we need to use "/"
                # to indicate we want the root web of the site we just connected to
                if ($ObjectPath -eq $SiteURL) {
                    Write-Log "Web object is the root web of the site, using '/' as WebPath"
                    $object = Get-SPWeb -WebPath "/"
                }
                else {
                    # For subwebs, extract the relative path
                    $relativeWebPath = $ObjectPath
                    if ($ObjectPath.StartsWith($SiteURL)) {
                        $relativeWebPath = $ObjectPath.Substring($SiteURL.Length)
                        if ($relativeWebPath -eq "") { $relativeWebPath = "/" }
                    }
                    Write-Log "Using relative web path: $relativeWebPath"
                    $object = Get-SPWeb -WebPath $relativeWebPath
                }
                
                if ($object) {
                    Remove-EEEUfromWeb -web $object | Out-Null
                }
            }
            default {
                Write-Log "Unknown object type: $ObjectType. Skipping this item." "WARNING"
                Write-Host "Unknown object type: $ObjectType. Skipping this item." -ForegroundColor Yellow
            }
        }
        
        if (-not $object) {
            Write-Log "Object is null for $ObjectPath with type $ObjectType. Skipping further processing for this item." "WARNING"
            Write-Host "Object is null for $ObjectPath with type $ObjectType. Skipping." -ForegroundColor Red
        }
    }
    catch {
        Write-Log "An error occurred while processing $ObjectPath in $SiteURL : $_" "ERROR"
        Write-Host "Error processing $ObjectPath : $_" -ForegroundColor Red
    }
}

# Read the CSV file and process each row
$csvData = Import-Csv -Path $csvFilePath
foreach ($row in $csvData) {
    # The headings can be different in the CSV file, so we need to use the correct column names
    $SiteURL = $row.URL
    $ObjectPath = $row.ItemURL
    
    # Get the object type from the "ItemType" column in the CSV
    $ObjectType = "File" # Default to File
    if ($row.PSObject.Properties.Name -contains "ItemType") {
        $ObjectType = $row.ItemType
        Write-Log "Using ItemType from CSV: $ObjectType"
    }
    elseif ($row.PSObject.Properties.Name -contains "Type") {
        $ObjectType = $row.Type
        Write-Log "Using Type from CSV: $ObjectType"
    }
    else {
        # Fallback: Try to determine type based on path
        Write-Log "No Type column found in CSV, determining type based on path pattern"
        
        if ([string]::IsNullOrEmpty($ObjectPath) -or $ObjectPath -eq "/") {
            $ObjectType = "Web"
        }
        elseif ($ObjectPath -match "/Lists/[^/]+$") {
            $ObjectType = "List"
        }
        elseif ($ObjectPath -match "/([^/]+)$" -and -not ($ObjectPath -match "\.[^/]+$")) {
            # Check if this might be a document library
            try {
                Connect-ToSharePoint -SiteURL $SiteURL
                $listName = $Matches[1]
                $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
                if ($list -and $list.BaseTemplate -eq 101) {
                    # 101 is the template ID for document libraries
                    $ObjectType = "DocLib"
                }
                else {
                    $ObjectType = "Folder"
                }
            }
            catch {
                $ObjectType = "Folder" # Default to folder if we can't check
                Write-Log "Error checking if path is a document library: $_" "WARNING"
            }
        }
        elseif ($ObjectPath -match "\.[^/]+$") {
            $ObjectType = "File"
        }
        else {
            $ObjectType = "Folder"
        }
    }

    Write-Host "Working on $ObjectType : $ObjectPath in $siteURL to remove EEEU" -ForegroundColor Green
    Write-Log "Attempting to process $ObjectType : $ObjectPath in site: $SiteURL"

    Invoke-SharePointObjectProcessing -SiteURL $SiteURL -ObjectPath $ObjectPath -ObjectType $ObjectType
}

Write-Log "Operations completed successfully"
Write-Host "Operations completed successfully"
