<#
.SYNOPSIS
    Scans SharePoint Online sites to identify occurrences of the "Everyone Except External Users" (EEEU) group in file permissions.

.DESCRIPTION
    This script connects to SharePoint Online using provided tenant-level credentials and iterates through a list of
    site URLs specified in an input file. It recursively scans document libraries and lists (excluding specified folders)
    to locate files where the "Everyone Except External Users" group has permissions assigned (excluding "Limited Access").
    The script logs its operations and outputs the results to a CSV file, detailing the site URL, file URL, and assigned roles.

    The script works with both standard SharePoint domains (contoso.sharepoint.com) and vanity domains (contoso.myspace.com).

.PARAMETER inputFilePath
    Text file containing a list of sites to scan

.PARAMETER appID
    Client ID of the App Registration to use in the Connect-PnpOnline cmdlet.
    Requires at minimum read access to all sites

.PARAMETER thumbprint
    Thumbprint of the certificate to use for authentication with the app registration

.PARAMETER tenant
    Tenant ID

.PARAMETER logFilePath
    Output log file path. Optional
    Specify either the full path to a log file. If omitted then a log file will be created called 'Find_EEEU_In_Sites_yyyyMMdd_HHmmss.txt'

.PARAMETER outputFilePath
    Output results CSV path. Optional
    Specify either the full path to a csv file. If omitted then a csv file will be created called 'Find_EEEU_In_Sites_yyyyMMdd_HHmmss.csv'

.PARAMETER permissionLevels
    The permission levels to scan. Can be customised if you are only looking for permissions at a specific level.
    Options are @('Web', 'List', 'Folder', 'File', 'Group')
    By default all levels will be scanned

.PARAMETER debugLogging
    Boolean. Enable debug logging.
    Default is false.

.PARAMETER includeLimitedAccessPermissions
    Boolean. When set to true, permnissions with an access level of 'Limited Access' are included in the output report.
    Default is false

.INPUTS
    A text file containing SharePoint site URLs to scan (path specified in $inputFilePath variable).

.OUTPUTS
    - A CSV file containing all found EEEU occurrences (path: $env:TEMP\Find_EEEU_In_Sites_[timestamp].csv)
    - A log file documenting the script's execution (path: $env:TEMP\Find_EEEU_In_Sites_[timestamp].txt)

.NOTES
    File Name      : Find-EEEUInSites.ps1
    Author         : Mike Lee
    Date Created   : 6/26/2025
    Update History :
        6/26/2025 - Initial script creation
        7/02/2025 - Improved error handling and logging
        7/10/2025 - Added folder-level permission checks
        7/15/2025 - Enhanced URL encoding for file access
        7/20/2025 - Added throttling handling with exponential backoff
        7/25/2025 - Improved list URL retrieval logic
        8/01/2025 - Added debug logging option
        10/15/2025 - Updated to use latest PnP PowerShell module cmdlets
        10/16/2025 - Add support for vanity domains
        12/23/2025 - Performance improvements in get items from lists (Craig Tolley)
        12/23/2025 - Use ArrayLists for improvements (Craig Tolley)
        12/23/2025 - Add discovery for Everyone permissions (Craig Tolley)
        12/23/2025 - Support searching for Everyone/EEEU in Site Groups (Craig Tolley)
        12/23/2025 - Move configuration to parameters (Craig Tolley)
        12/23/2025 - Implement option to only scan specific objects for permissions (Craig Tolley)

    The script uses app-only authentication with a certificate thumbprint. Make sure the app has
    proper permissions in your tenant (SharePoint: Sites.FullControl.All is recommended).

    The script ignores several system folders and lists to improve performance and avoid errors.

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

.EXAMPLE
    Executes the script using the specific values for AppId, Thumbprint and Tenant.
    Sites in the Sites.txt file will be scanned

    $appId = "5baa1427-1e90-4501-831d-a8e67465f0d9"
    $thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"
    $tenantId = "85612ccb-4c28-4a34-88df-a538cc139a51"

    .\Find-EEEUInSites.ps1 -inputFilePath Sites.txt -appID $appId -thumbprint $thumbprint -tenant $tenantId

.EXAMPLE
    As above, but only scans for Group and Web level permissions
    $appId = "5baa1427-1e90-4501-831d-a8e67465f0d9"
    $thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"
    $tenantId = "85612ccb-4c28-4a34-88df-a538cc139a51"
    $permissionLevels = @("Group", "Web")

    .\Find-EEEUInSites.ps1 -inputFilePath Sites.txt -appID $appId -thumbprint $thumbprint -tenant $tenantId -permissionLevels $permissionLevels

#>
param (
    # Path to the input file containing site URLs to scan
    [Parameter(Mandatory = $true)]
    $inputFilePath,

    # Entra App ID for authentication
    [Parameter(Mandatory = $true)]
    $appID,

    # Certificate thumbprint for authentication
    [Parameter(Mandatory = $true)]
    $thumbprint,

    # Tenant ID for your tenant
    [Parameter(Mandatory = $true)]
    $tenant,

    # Path to save the output log file
    # If not specified then it will saved as 'Find_EEEU_In_Sites_yyyyMMdd_HHmmss.txt'
    $logFilePath,

    # Path to save the output CSV file
    # If not specified then it will saved as 'Find_EEEU_In_Sites_yyyyMMdd_HHmmss.txt'
    $outputFilePath,

    # Types of object to scan. Options are Web, List, Folder, File and Group.
    # Default is all
    $permissionLevels = @('Web', 'List', 'Folder', 'File', 'Group'),

    # Set to $true for verbose logging, $false for essential logging only
    # Default is false
    [bool]$debugLogging = $false,

    # When set to true, 'Limited Access' permissions will be included in the report.
    # Default is false
    [bool]$includeLimitedAccessPermissions = $false
)

# Script Parameters
Add-Type -AssemblyName System.Web
$EEEU = '*spo-grid-all-users*'
$Everyone = 'c:0(.s|true'
$AllUsers = 'c:0!.s|windows' # Shows as 'All Users (Windows)' - a legacy claim that is similar in scope to EEEU

$startime = Get-Date -Format 'yyyyMMdd_HHmmss'
if ([String]::IsNullOrEmpty($logFilePath)) {
    $logFilePath = ".\Find_EEEU_In_Sites_$($startime).txt"
}

if ([String]::IsNullOrEmpty($outputFilePath)) {
    $outputFilePath = ".\Find_EEEU_In_Sites_$($startime).csv"
}

# List of folder patterns to ignore (uses wildcard matching for better tenant compatibility)
$ignoreFolderPatterns = @(
    '*VivaEngage*',    #Viva Engage folder for Storyline attachments EEEU is read by default
    '*Style Library*',
    '*_catalogs*',
    '*_cts*',
    '*_private*',
    '*_vti_pvt*',
    '*Reference*',  # Matches any folder with "Reference" and a GUID
    '*Sharing Links*',
    '*Social*',
    '*FavoriteLists*',  # Matches FavoriteLists with any GUID
    '*User Information List*',
    '*Web Template Extensions*',
    '*SmartCache*',  # Matches SmartCache with any GUID
    '*SharePointHomeCacheList*',
    '*RecentLists*',  # Matches RecentLists with any GUID
    '*PersonalCacheLibrary*',
    '*microsoft.ListSync.Endpoints*',
    '*Maintenance Log Library*',
    '*DO_NOT_DELETE_ENTERPRISE_USER_CONTAINER_ENUM_LIST*',  # Matches with any GUID
    '*appfiles*'
)

# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = 'INFO'
    )

    # Only log essential messages when debug is false
    $essentialLevels = @('ERROR', 'WARNING')
    $isEssential = $level -in $essentialLevels -or
    $message -like '*Located EEEU/Everyone*' -or
    $message -like '*Connected to SharePoint*' -or
    $message -like '*Failed to connect*' -or
    $message -like '*Processing site:*' -or
    $message -like '*Completed processing*' -or
    $message -like '*scan completed*'

    if ($debugLogging -or $isEssential) {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logMessage = "$timestamp - $level - $message"
        Add-Content -Path $logFilePath -Value $logMessage -ErrorAction Stop
    }
}

# Handle SharePoint Online throttling with exponential backoff
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$InitialDelaySeconds = 5
    )

    $retryCount = 0
    $delay = $InitialDelaySeconds
    $success = $false
    $result = $null

    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $exception = $_.Exception

            # Check if this is a throttling error (look for specific status codes or messages)
            $isThrottlingError = $false
            $retryAfterSeconds = $delay

            if ($exception.Response) {
                # Check for Retry-After header
                $retryAfterHeader = $exception.Response.Headers['Retry-After']
                if ($retryAfterHeader) {
                    $isThrottlingError = $true
                    $retryAfterSeconds = [int]$retryAfterHeader
                    Write-Log "Received Retry-After header: $retryAfterSeconds seconds" 'WARNING'
                }

                # Check for 429 (Too Many Requests) or 503 (Service Unavailable)
                $statusCode = [int]$exception.Response.StatusCode
                if ($statusCode -eq 429 -or $statusCode -eq 503) {
                    $isThrottlingError = $true
                    Write-Log "Detected throttling response (Status code: $statusCode)" 'WARNING'
                }
            }

            # Also check for specific throttling error messages
            if ($exception.Message -match 'throttl' -or
                $exception.Message -match 'too many requests' -or
                $exception.Message -match 'temporarily unavailable') {
                $isThrottlingError = $true
                Write-Log "Detected throttling error in message: $($exception.Message)" 'WARNING'
            }

            if ($isThrottlingError) {
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Log "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." 'WARNING'
                    Write-Host "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfterSeconds

                    # Implement exponential backoff if no Retry-After header was provided
                    if ($retryAfterSeconds -eq $delay) {
                        $delay = $delay * 2 # Exponential backoff
                    }
                }
                else {
                    Write-Log 'Maximum retry attempts reached. Giving up on operation.' 'ERROR'
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow
                # Check if it's an expected error that we can handle gracefully
                if ($_.Exception.Message -like '*Object reference not set to an instance of an object*' -or
                    $_.Exception.Message -like '*ListItemAllFields*' -or
                    $_.Exception.Message -like '*object is associated with property*') {
                    Write-Log "Expected retrieval error (likely null object reference): $($_.Exception.Message)" 'DEBUG'
                }
                elseif ($_.Exception.Message -like '*does not exist at site*') {
                    Write-Log "Resource not found (likely folder/list doesn't exist): $($_.Exception.Message)" 'DEBUG'
                }
                else {
                    Write-Log "General Error occurred During retrieval : $($_.Exception.Message)" 'WARNING'
                }
                throw $_
            }
        }
    }

    return $result
}

# Read site URLs from input file
function Read-SiteURLs {
    param (
        [string]$filePath
    )
    $urls = Get-Content -Path $filePath
    return $urls
}

# Connect to SharePoint Online
function Connect-SharePoint {
    param (
        [string]$siteURL
    )
    try {
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $siteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        Write-Log "Connected to SharePoint Online at $siteURL"
        return $true # Connection successful
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online at $siteURL : $($_.Exception.Message)" 'ERROR'
        return $false # Connection failed
    }
}

# Get all items in a folder and subfolders (using absolute URLs)
function Get-AllItemsInList {
    param (
        [string]$listTitleOrUrl,
        [string]$listId
    )

    # Get list GUID to use in REST
    Write-Host "  Getting items in '$listTitleOrUrl'" -ForegroundColor Yellow
    # Base REST URL
    $baseUrl = "/_api/web/lists(guid'$($listId)')/items"
    $select = 'Id,UniqueId,FileLeafRef,FileRef,HasUniqueRoleAssignments,FileSystemObjectType'

    $top = 5000
    $uri = "$baseUrl`?`$select=$select&`$top=$top"

    $results = [System.Collections.ArrayList]::new()
    while ($uri) {
        Write-Log -message "Query URI: $uri" -level DEBUG
        $response = Invoke-PnPSPRestMethod -Url $uri -Method Get
        $responseCountPreFilter = $response.value.count
        $filteredresponses = $response.value | Where-Object { $_ -ne $null -and $null -ne $_.FileLeafRef -and $_.HasUniqueRoleAssignments -eq 'True' }
        Write-Log -message "Retrieved $responseCountPreFilter items, of which $($filteredresponses.count) have unique permissions" -level INFO
        Write-Host "  Retrieved $responseCountPreFilter items, of which $($filteredresponses.count) have unique permissions" -ForegroundColor Yellow
        foreach ( $r in $filteredresponses) {
            $results.Add($r) | Out-Null
        }

        if ($response.'odata.nextLink') {
            $uri = $response.'odata.nextLink'
        }
        else {
            $uri = $null
        }
    }
    Write-Host "  Completed. Retrieved $($results.count) items, which have unique permissions" -ForegroundColor Yellow
    $results
}

# Function to check for EEEU in web-level permissions
function Find-EEEUinWeb {
    param (
        [string]$siteURL,
        [ref]$EEEUOccurrences
    )
    try {
        Write-Host "Checking web-level permissions for $siteURL..." -ForegroundColor Yellow
        Write-Log "Checking web-level permissions for $siteURL"

        # Get web with throttling protection
        $web = Invoke-WithRetry -ScriptBlock {
            Get-PnPWeb
        }

        # Check if web has unique permissions
        $hasUniquePermissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $web -Property HasUniqueRoleAssignments
        }

        if (-not $hasUniquePermissions) {
            Write-Log 'Web does not have unique permissions. Skipping.' 'DEBUG'
            Write-Host 'Web does not have unique permissions. Skipping.' -ForegroundColor Yellow
            return
        }

        # Get web permissions with throttling protection
        $Permissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $web -Property RoleAssignments
        }

        if ($Permissions) {
            $roles = [System.Collections.ArrayList]::new()
            foreach ($RoleAssignment in $Permissions) {
                # Get role assignments with throttling protection
                Invoke-WithRetry -ScriptBlock {
                    Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                }


                if ($RoleAssignment.Member.LoginName -like $EEEU -or $RoleAssignment.Member.LoginName -eq $Everyone -or $RoleAssignment.Member.LoginName -eq $AllUsers) {
                    $rolelevel = $RoleAssignment.RoleDefinitionBindings
                    foreach ($role in $rolelevel) {
                        # Only add roles that are not 'Limited Access'
                        if ($role.Name -ne 'Limited Access' -or $includeLimitedAccessPermissions) {
                            $roles.Add([PsCustomObject]@{
                                    Member   = $RoleAssignment.Member.Title + ' (' + $RoleAssignment.Member.LoginName + ')'
                                    RoleName = $role.Name
                                }) | Out-Null
                        }
                    }
                }
            }

            if ($roles.Count -gt 0) {
                # Store "/" as the relative path for the root web
                $relativeUrl = '/'
                $roles | Group-Object member | ForEach-Object {
                    $newOccurrence = [PSCustomObject]@{
                        Url         = $SiteURL
                        ItemURL     = $relativeUrl
                        ItemType    = 'Web'
                        Member      = $_.Name
                        RoleNames   = ($_.Group.RoleName -join ', ')
                        OwnerName   = 'N/A'
                        OwnerEmail  = 'N/A'
                        CreatedDate = 'N/A'
                    }

                    $EEEUOccurrences.Value.Add($newOccurrence) | Out-Null
                    Write-Host "Located '$($_.Name)' at Web level on $SiteURL" -ForegroundColor Red
                    Write-Log "Located '$($_.Name)' at Web level on $SiteURL - Added to collection (Count: $($EEEUOccurrences.Value.Count))"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to process web-level permissions: $_" 'ERROR'
    }
}

# Function to check for EEEU in list-level permissions
function Find-EEEUinLists {
    param (
        [string]$siteURL,
        [ref]$EEEUOccurrences
    )
    try {
        Write-Host "Checking list-level permissions for $siteURL..." -ForegroundColor Yellow
        Write-Log "Checking list-level permissions for $siteURL"

        # Get all lists and libraries with throttling protection
        $lists = Invoke-WithRetry -ScriptBlock {
            Get-PnPList | Where-Object {
                $listTitle = $_.Title
                $shouldIgnore = $false
                foreach ($pattern in $ignoreFolderPatterns) {
                    if ($listTitle -like $pattern) {
                        $shouldIgnore = $true
                        break
                    }
                }
                -not $shouldIgnore
            }
        }

        foreach ($list in $lists) {
            # Skip processing hidden lists
            if ($list.Hidden) {
                continue
            }

            # Check if list has unique permissions
            $hasUniquePermissions = Invoke-WithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments
            }

            if (-not $hasUniquePermissions) {
                Write-Log "List '$($list.Title)' does not have unique permissions. Skipping." 'DEBUG'
                continue
            }

            # Get list permissions with throttling protection
            $Permissions = Invoke-WithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $list -Property RoleAssignments
            }

            if ($Permissions) {
                $roles = [System.Collections.ArrayList]::new()
                foreach ($RoleAssignment in $Permissions) {
                    # Get role assignments with throttling protection
                    Invoke-WithRetry -ScriptBlock {
                        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                    }

                    if ($RoleAssignment.Member.LoginName -like $EEEU -or $RoleAssignment.Member.LoginName -eq $Everyone -or $RoleAssignment.Member.LoginName -eq $AllUsers) {
                        $rolelevel = $RoleAssignment.RoleDefinitionBindings
                        foreach ($role in $rolelevel) {
                            # Only add roles that are not 'Limited Access'
                            if ($role.Name -ne 'Limited Access' -or $includeLimitedAccessPermissions) {
                                $roles.Add([PsCustomObject]@{
                                        Member   = $RoleAssignment.Member.Title + ' (' + $RoleAssignment.Member.LoginName + ')'
                                        RoleName = $role.Name
                                    }) | Out-Null
                            }
                        }
                    }
                }
                if ($roles.Count -gt 0) {
                    # Get the list's root folder server relative URL instead of the DefaultViewUrl
                    # This will give us the clean list URL without Forms/Views
                    $relativeUrl = ''

                    try {
                        # Try to get the root folder URL with throttling protection
                        $rootFolder = Invoke-WithRetry -ScriptBlock {
                            Get-PnPProperty -ClientObject $list -Property RootFolder
                        }
                        if ($rootFolder -and $rootFolder.ServerRelativeUrl) {
                            $relativeUrl = $rootFolder.ServerRelativeUrl
                            Write-Log "Retrieved RootFolder URL for $($list.Title): $relativeUrl" 'DEBUG'
                        }
                    }
                    catch {
                        Write-Log "Could not retrieve root folder for list $($list.Title): $_" 'DEBUG'
                    }

                    # If we can't get the root folder URL, fall back to constructing it from the list title
                    if (-not $relativeUrl -or $relativeUrl -eq '') {
                        # Parse the site URL to get the site path and construct the list path
                        $uri = New-Object System.Uri($siteURL)
                        $sitePath = $uri.AbsolutePath.TrimEnd('/')

                        # Handle special case for Site Pages (should be SitePages, not "Site Pages")
                        $listTitle = $list.Title
                        if ($listTitle -eq 'Site Pages') {
                            $listTitle = 'SitePages'
                        }
                        elseif ($listTitle -eq 'Shared Documents') {
                            # Documents library is often called "Shared Documents" in title but URL is just the folder name
                            $listTitle = $listTitle.Replace(' ', '')
                        }
                        else {
                            # For other lists, replace spaces with empty string (common SharePoint URL pattern)
                            $listTitle = $listTitle.Replace(' ', '')
                        }

                        $relativeUrl = "$sitePath/$listTitle"
                        Write-Log "Constructed fallback URL for $($list.Title): $relativeUrl" 'DEBUG'
                    }

                    # Additional check: if the URL still contains Forms/ or other view paths, clean it up
                    if ($relativeUrl -like '*/Forms/*' -or $relativeUrl -like '*/Forms' -or $relativeUrl -like '*/AllItems.aspx' -or $relativeUrl -like '*ByAuthor.aspx') {
                        Write-Log "Cleaning up view path from URL: $relativeUrl" 'DEBUG'
                        # Remove /Forms and everything after it, or specific view files
                        $relativeUrl = $relativeUrl -replace '/Forms.*$', ''
                        $relativeUrl = $relativeUrl -replace '/AllItems\.aspx$', ''
                        $relativeUrl = $relativeUrl -replace '/.*ByAuthor\.aspx$', ''
                        Write-Log "Cleaned URL result: $relativeUrl" 'DEBUG'
                    }

                    $roles | Group-Object member | ForEach-Object {
                        $newOccurrence = [PSCustomObject]@{
                            Url         = $SiteURL
                            ItemURL     = $relativeUrl
                            ItemType    = 'List'
                            Member      = $_.Name
                            RoleNames   = ($_.Group.RoleName -join ', ')
                            OwnerName   = 'N/A'
                            OwnerEmail  = 'N/A'
                            CreatedDate = 'N/A'
                        }
                        $EEEUOccurrences.Value.Add($newOccurrence) | Out-Null
                        Write-Host "Located '$($_.Name)' at List level: $($list.Title) on $SiteURL" -ForegroundColor Red
                        Write-Log "Located '$($_.Name)' at List level: $($list.Title) on $SiteURL - Added to collection (Count: $($EEEUOccurrences.Value.Count))"
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Failed to process list-level permissions: $_" 'ERROR'
    }
}

# Function to check for EEEU in folder-level permissions
function Find-EEEUinFolders {
    param (
        $item,
        [string]$siteURL,
        [string]$listTitle,
        [ref]$EEEUOccurrences
    )
    try {
        if ($null -eq $item -or $null -eq $item.FileRef) {
            Write-Log 'Item or FileRef is null, skipping folder processing' 'DEBUG'
            return
        }

        Write-Host "Checking folder-level permissions in list '$listTitle' for folder '$($item.FileLeafRef)'..." -ForegroundColor Yellow
        Write-Log "Checking folder-level permissions in list '$listTitle' for folder '$($item.FileLeafRef)..."

        # Get the list object first
        $list = Invoke-WithRetry -ScriptBlock {
            Get-PnPList -Identity $listTitle -ErrorAction Stop
        }

        if ($null -eq $list) {
            Write-Log "List '$listTitle' not found" 'WARNING'
            return
        }

        $folderName = $item['FileLeafRef']
        $folderUrl = $item['FileRef']

        # Skip ignored folders using wildcard patterns
        $shouldIgnoreFolder = $false
        foreach ($pattern in $ignoreFolderPatterns) {
            if ($folderName -like $pattern) {
                $shouldIgnoreFolder = $true
                break
            }
        }
        if ($shouldIgnoreFolder) {
            continue
        }

        # Check if folder has unique permissions
        $hasUniquePermissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $folderItem -Property HasUniqueRoleAssignments
        }

        if (-not $hasUniquePermissions) {
            Write-Log "Folder '$folderName' does not have unique permissions. Skipping." 'DEBUG'
            continue
        }

        # Get folder permissions
        $Permissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $folderItem -Property RoleAssignments
        }

        if ($Permissions) {
            $roles = [System.Collections.ArrayList]::new()
            foreach ($RoleAssignment in $Permissions) {
                # Get role assignments with throttling protection
                Invoke-WithRetry -ScriptBlock {
                    Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                }

                if ($RoleAssignment.Member.LoginName -like $EEEU -or $RoleAssignment.Member.LoginName -eq $Everyone -or $RoleAssignment.Member.LoginName -eq $AllUsers) {
                    $rolelevel = $RoleAssignment.RoleDefinitionBindings
                    foreach ($role in $rolelevel) {
                        # Only add roles that are not 'Limited Access'
                        if ($role.Name -ne 'Limited Access' -or $includeLimitedAccessPermissions) {
                            $roles.Add([PsCustomObject]@{
                                    Member   = $RoleAssignment.Member.Title + ' (' + $RoleAssignment.Member.LoginName + ')'
                                    RoleName = $role.Name
                                }) | Out-Null
                        }
                    }
                }
            }
            if ($roles.Count -gt 0) {
                # Get folder owner information if available
                $owner = 'N/A'
                $ownerEmail = 'N/A'
                $createdDate = 'N/A'

                try {
                    # Try to get folder author/owner information
                    if ($null -ne $folderItem['Author']) {
                        $authorId = $folderItem['Author'].LookupId

                        if ($authorId) {
                            $ownerInfo = Invoke-WithRetry -ScriptBlock {
                                Get-PnPUser -Identity $authorId
                            }

                            if ($ownerInfo) {
                                $owner = $ownerInfo.Title
                                $ownerEmail = $ownerInfo.Email
                            }
                        }
                    }

                    # Get created date
                    if ($null -ne $folderItem['Created']) {
                        $createdDate = $folderItem['Created'].ToString('yyyy-MM-dd HH:mm:ss')
                    }
                }
                catch {
                    Write-Log "Error retrieving folder owner information: $_" 'WARNING'
                }

                $roles | Group-Object member | ForEach-Object {
                    $newOccurrence = [PSCustomObject]@{
                        Url         = $SiteURL
                        ItemURL     = $folderUrl
                        ItemType    = 'Folder'
                        Member      = $_.Name
                        RoleNames   = ($_.Group.RoleName -join ', ')
                        OwnerName   = $owner
                        OwnerEmail  = $ownerEmail
                        CreatedDate = $createdDate
                    }
                    $EEEUOccurrences.Value.Add($newOccurrence) | Out-Null
                    Write-Host "Located '$($_.Name)' at Folder level: $folderName on $SiteURL" -ForegroundColor Red
                    Write-Log "Located '$($_.Name)' at Folder level: $folderName on $SiteURL - Added to collection (Count: $($EEEUOccurrences.Value.Count))"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to process folder-level permissions in list '$listTitle': $_" 'ERROR'
    }
}

function Find-EEEUinListRootFolder {
    param (
        [string]$siteURL,
        [string]$listTitle,
        [ref]$EEEUOccurrences
    )
    # Also check root folder of the list if it has unique permissions
    try {
        Write-Host "Checking root-level permissions in list '$listTitle'..." -ForegroundColor Yellow
        Write-Log "Checking root-level permissions in list '$listTitle'..."

        # Get the list object first
        $list = Invoke-WithRetry -ScriptBlock {
            Get-PnPList -Identity $listTitle -ErrorAction Stop
        }

        if ($null -eq $list) {
            Write-Log "List '$listTitle' not found" 'WARNING'
            return
        }

        $rootFolder = Invoke-WithRetry -ScriptBlock {
            Get-PnPFolder -Url $list.RootFolder.ServerRelativeUrl -Includes ListItemAllFields
        }

        if ($rootFolder -and $rootFolder.ListItemAllFields) {
            $rootFolderItem = $rootFolder.ListItemAllFields

            # Check if root folder has unique permissions
            $hasUniquePermissions = Invoke-WithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $rootFolderItem -Property HasUniqueRoleAssignments
            }

            if (-not $hasUniquePermissions) {
                Write-Log "Root folder of list '$listTitle' does not have unique permissions. Skipping." 'DEBUG'
                return
            }

            # Get folder permissions
            $Permissions = Invoke-WithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $rootFolder -Property RoleAssignments
            }

            if ($Permissions) {
                $roles = [System.Collections.ArrayList]::new()
                foreach ($RoleAssignment in $Permissions) {
                    # Get role assignments with throttling protection
                    Invoke-WithRetry -ScriptBlock {
                        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                    }

                    if ($RoleAssignment.Member.LoginName -like $EEEU -or $RoleAssignment.Member.LoginName -eq $Everyone -or $RoleAssignment.Member.LoginName -eq $AllUsers) {
                        $rolelevel = $RoleAssignment.RoleDefinitionBindings
                        foreach ($role in $rolelevel) {
                            # Only add roles that are not 'Limited Access'
                            if ($role.Name -ne 'Limited Access' -or $includeLimitedAccessPermissions) {
                                $roles.Add([PsCustomObject]@{
                                        Member   = $RoleAssignment.Member.Title + ' (' + $RoleAssignment.Member.LoginName + ')'
                                        RoleName = $role.Name
                                    }) | Out-Null
                            }
                        }
                    }
                }
                if ($roles.Count -gt 0) {
                    $roles | Group-Object member | ForEach-Object {
                        $newOccurrence = [PSCustomObject]@{
                            Url         = $SiteURL
                            ItemURL     = $rootFolder.ServerRelativeUrl
                            ItemType    = 'Folder'
                            Member      = $_.Name
                            RoleNames   = ($_.Group.RoleName -join ', ')
                            OwnerName   = 'N/A'
                            OwnerEmail  = 'N/A'
                            CreatedDate = 'N/A'
                        }
                        $EEEUOccurrences.Value.Add($newOccurrence) | Out-Null
                        Write-Host "Located '$($_.Name)' at Root Folder level: $($list.Title) on $SiteURL" -ForegroundColor Red
                        Write-Log "Located '$($_.Name)' at Root Folder level: $($list.Title) on $SiteURL - Added to collection (Count: $($EEEUOccurrences.Value.Count))"
                    }
                }
            }
        }
    }
    catch {
        # Check if it's the expected ListItemAllFields error
        if ($_.Exception.Message -like '*Object reference not set to an instance of an object*' -or
            $_.Exception.Message -like '*ListItemAllFields*' -or
            $_.Exception.Message -like '*object is associated with property*') {
            Write-Log "Expected root folder error (likely null ListItemAllFields): $($_.Exception.Message)" 'DEBUG'
        }
        else {
            Write-Log "Failed to process root folder permissions: $_" 'WARNING'
        }
    }
}
# Update the existing Find-EEEUinFiles function to include ItemType
function Find-EEEUinFiles {
    param (
        $listTitle,
        $item,
        [string]$siteURL,
        [ref]$EEEUOccurrences
    )
    try {
        if ($null -eq $item -or $null -eq $item.FileRef) {
            Write-Log 'Item or FileRef is null, skipping file processing' 'DEBUG'
            return
        }

        Write-Host "Checking file-level permissions in list '$listTitle' for file '$($item.FileLeafRef)'..." -ForegroundColor Yellow
        Write-Log "Checking file-level permissions in list '$listTitle' for file '$($item.FileLeafRef)..."

        $file = [System.Collections.ArrayList]::new()
        $fileUrl = $item.FileRef

        # Check if the file URL contains any of the ignore folder patterns
        foreach ($ignorePattern in $ignoreFolderPatterns) {
            # Remove wildcards from pattern for URL matching
            $cleanPattern = $ignorePattern.Replace('*', '')
            if ($fileUrl -like "*/$cleanPattern/*" -or $fileUrl -like "*/$cleanPattern") {
                return # Skip processing the ignored file
            }
        }

        try {
            # Try direct approach first with throttling protection
            $file = Invoke-WithRetry -ScriptBlock {
                Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop
            }
        }
        catch {
            # If direct approach fails, try with URL encoding
            try {
                Write-Log "Initial file access failed, trying with URL encoding: $fileUrl" 'WARNING'

                # Parse the URL into parts
                $urlParts = $fileUrl.Split('/')

                # Encode each part of the URL separately (except the protocol and domain)
                $encodedParts = [System.Collections.ArrayList]::new()
                $skipEncoding = $true
                foreach ($part in $urlParts) {
                    # Skip encoding for the protocol and domain parts
                    if ($skipEncoding -and ($part -eq 'https:' -or $part -eq '' -or $part -match '^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')) {
                        $encodedParts.Add($part) | Out-Null
                    }
                    else {
                        $skipEncoding = $false
                        $encodedParts.Add([System.Web.HttpUtility]::UrlEncode($part)) | Out-Null
                    }
                }

                # Rebuild the URL with encoded parts
                $encodedFileUrl = $encodedParts -join '/'

                # Try with encoded URL and throttling protection
                $file = Invoke-WithRetry -ScriptBlock {
                    Get-PnPFile -Url $encodedFileUrl -AsListItem
                }
                Write-Log "Successfully accessed file with encoded URL: $encodedFileUrl" 'DEBUG'
            }
            catch {
                Write-Log "Failed to access file even with URL encoding: $fileUrl - $_" 'ERROR'
                return
            }
        }

        # Check if file has unique permissions
        $hasUniquePermissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $file -Property HasUniqueRoleAssignments
        }

        if (-not $hasUniquePermissions) {
            Write-Log "File '$($file.FieldValues.FileLeafRef)' does not have unique permissions. Skipping." 'DEBUG'
            return
        }

        # Get permissions with throttling protection
        $Permissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $file -Property RoleAssignments
        }

        if ($Permissions) {
            $roles = [System.Collections.ArrayList]::new()
            foreach ($RoleAssignment in $Permissions) {
                # Get role assignments with throttling protection
                Invoke-WithRetry -ScriptBlock {
                    Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                }

                if ($RoleAssignment.Member.LoginName -like $EEEU -or $RoleAssignment.Member.LoginName -eq $Everyone -or $RoleAssignment.Member.LoginName -eq $AllUsers) {
                    $rolelevel = $RoleAssignment.RoleDefinitionBindings
                    foreach ($role in $rolelevel) {
                        # Only add roles that are not 'Limited Access'
                        if ($role.Name -ne 'Limited Access' -or $includeLimitedAccessPermissions) {
                            $roles.Add([PsCustomObject]@{
                                    Member   = $RoleAssignment.Member.Title + ' (' + $RoleAssignment.Member.LoginName + ')'
                                    RoleName = $role.Name
                                }) | Out-Null
                        }
                    }
                }
            }
            if ($roles.Count -gt 0) {
                # Get file owner information
                $owner = 'Unknown'
                $ownerEmail = 'Unknown'
                $createdDate = 'Unknown'

                try {
                    # Try to get file author/owner information using PnP methods
                    if ($file.FieldValues.ContainsKey('Author')) {
                        $authorId = $file.FieldValues.Author.LookupId

                        if ($authorId) {
                            $ownerInfo = Invoke-WithRetry -ScriptBlock {
                                Get-PnPUser -Identity $authorId
                            }

                            if ($ownerInfo) {
                                $owner = $ownerInfo.Title
                                $ownerEmail = $ownerInfo.Email
                            }
                        }
                    }

                    # Get created date
                    if ($file.FieldValues.ContainsKey('Created')) {
                        $createdDate = $file.FieldValues.Created.ToString('yyyy-MM-dd HH:mm:ss')
                    }
                }
                catch {
                    Write-Log "Error retrieving file owner information: $_" 'WARNING'
                }

                $roles | Group-Object member | ForEach-Object {
                    $newOccurrence = [PSCustomObject]@{
                        Url         = $SiteURL
                        ItemURL     = $file.FieldValues.FileRef
                        ItemType    = 'File'
                        Member      = $_.Name
                        RoleNames   = ($_.Group.RoleName -join ', ')
                        OwnerName   = $owner
                        OwnerEmail  = $ownerEmail
                        CreatedDate = $createdDate
                    }
                    $EEEUOccurrences.Value.Add($newOccurrence) | Out-Null
                    Write-Host "Located '$($_.Name)' in file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Red
                    Write-Log "Located '$($_.Name)' in file: $($file.FieldValues.FileLeafRef) on $SiteURL - Added to collection (Count: $($EEEUOccurrences.Value.Count))"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to process file: $_" 'ERROR'
    }
}

# Find EEEU in Site Level group memberships
function Find-EEEUinSiteGroups {
    param (
        [string]$siteURL,
        [ref]$EEEUOccurrences
    )

    Write-Host "Checking site group level permissions for $siteURL..." -ForegroundColor Yellow
    Write-Log "Checking site group level permissions for $siteURL"

    $Groups = Invoke-WithRetry -ScriptBlock {
        $web = Get-PnPWeb -Includes ParentWeb

        if (-not $web.ParentWeb.Id) {
            Get-PnPProperty -ClientObject $web -Property SiteGroups -ErrorAction SilentlyContinue
        }
        else {
            Write-Host "Site $siteURL is a subweb. Skipping group check as they inherit from the parent web"
            Write-Log -level 'INFO' -message "Site $siteURL is a subweb. Skipping group check as they inherit from the parent web"
        }
    }

    foreach ($group in $groups) {
        try {
            # Get-PnPGroupMembers returns principals in the SharePoint group
            $members = Get-PnPGroupMember -Identity $group.Title -ErrorAction Stop
        }
        catch {
            Write-Log -level 'WARNING' -message "Could not enumerate members of group '$($Group.Title)': $($_.Exception.Message)"
        }

        foreach ($m in $members | Where-Object { $_.LoginName -like $EEEU -or $_.LoginName -eq $Everyone -or $_.LoginName -eq $AllUsers }) {
            $newOccurrence = [PSCustomObject]@{
                Url         = $SiteURL
                ItemURL     = ''
                ItemType    = 'Group'
                Member      = $m.Title + ' (' + $m.LoginName + ')'
                RoleNames   = ($Group.Title)
                OwnerName   = ''
                OwnerEmail  = ''
                CreatedDate = ''
            }
            $EEEUOccurrences.Value.Add($newOccurrence) | Out-Null
            Write-Host "Located '$($m.Title)' in Site Group '$($group.Title)' on $SiteURL" -ForegroundColor Red
            Write-Log "Located '$($m.Title)' in Site Group '$($group.Title)' on $SiteURL - Added to collection (Count: $($EEEUOccurrences.Value.Count))"
        }
    }
}

# Update the CSV output function to include ItemType
function Write-EEEUOccurrencesToCSV {
    param (
        [string]$filePath,
        [switch]$Append = $false,
        [array]$OccurrencesData = $global:EEEUOccurrences
    )
    try {
        # Create the file with headers if it doesn't exist or if we're not appending
        if (-not (Test-Path $filePath) -or -not $Append) {
            # Create empty file with headers - adding ItemType column
            Write-Host "Creating output CSV file at $filePath"
            'Url,ItemURL,ItemType,Member,RoleNames,OwnerName,OwnerEmail,CreatedDate' | Out-File -FilePath $filePath
        }

        # Group by URL, Item URL, ItemType and Roles to remove duplicates
        # Also handle cases where we might have both Forms paths and clean paths for the same list
        $uniqueOccurrences = $OccurrencesData |
        ForEach-Object {
            # Clean up any remaining Forms paths in the ItemURL before deduplication
            if ($_.ItemURL -like '*/Forms/*' -or $_.ItemURL -like '*/Forms' -or $_.ItemURL -like '*/AllItems.aspx' -or $_.ItemURL -like '*ByAuthor.aspx') {
                # Remove /Forms and everything after it, or specific view files
                $_.ItemURL = $_.ItemURL -replace '/Forms.*$', ''
                $_.ItemURL = $_.ItemURL -replace '/AllItems\.aspx$', ''
                $_.ItemURL = $_.ItemURL -replace '/.*ByAuthor\.aspx$', ''
            }
            $_
        } |
        Group-Object -Property Url, ItemURL, ItemType, RoleNames, Member |
        ForEach-Object { $_.Group[0] }

        # Append data to CSV
        foreach ($occurrence in $uniqueOccurrences) {
            # Manual CSV creation to handle special characters correctly
            $csvLine = "`"$($occurrence.Url)`",`"$($occurrence.ItemURL)`",`"$($occurrence.ItemType)`",`"$($occurrence.Member)`",`"$($occurrence.RoleNames)`",`"$($occurrence.OwnerName)`",`"$($occurrence.OwnerEmail)`",`"$($occurrence.CreatedDate)`""
            Add-Content -Path $filePath -Value $csvLine
        }

        Write-Log "EEEU occurrences have been written to $filePath" 'DEBUG'
    }
    catch {
        Write-Log "Failed to write EEEU occurrences to CSV file: $_" 'ERROR'
        Write-Error "Failed to write EEEU occurrences to CSV file: $_" -ErrorAction Stop
    }
}

# Add a helper function to convert server relative paths to site relative paths
function Convert-ToRelativePath {
    param (
        [string]$serverRelativePath,
        [string]$siteUrl
    )

    try {
        # If it's already a relative path (not starting with /)
        if (-not $serverRelativePath.StartsWith('/')) {
            return $serverRelativePath
        }

        # Parse the site URL to get the site path
        $uri = New-Object System.Uri($siteUrl)
        $sitePath = $uri.AbsolutePath

        # If the server relative path starts with the site path, remove it
        if ($serverRelativePath.StartsWith($sitePath)) {
            $relativePath = $serverRelativePath.Substring($sitePath.Length)
            # Ensure it starts with /
            if (-not $relativePath.StartsWith('/')) {
                $relativePath = '/' + $relativePath
            }
            return $relativePath
        }

        # Return the original path if we couldn't convert it
        return $serverRelativePath
    }
    catch {
        Write-Log "Error converting path to relative: $_" 'WARNING'
        return $serverRelativePath
    }
}

# Add a function to recursively process subsites
function Process-SiteAndSubsites {
    param (
        [string]$siteURL
    )

    Write-Host "Processing site: $siteURL" -ForegroundColor Green
    Write-Log "Processing site: $siteURL"

    # Initialize local collection for this site (don't clear global yet)
    $siteEEEUOccurrences = [System.Collections.ArrayList]::new()

    if (Connect-SharePoint -siteURL $siteURL) {
        # Check web-level permissions
        if ($permissionLevels -contains 'Web') {
            Find-EEEUinWeb -siteURL $siteURL -EEEUOccurrences ([ref]$siteEEEUOccurrences)
        }

        # Check site groups
        if ($permissionLevels -contains 'Group') {
            Find-EEEUinSiteGroups -siteURL $siteURL -EEEUOccurrences ([ref]$siteEEEUOccurrences)
        }

        # Check list-level permissions
        if ($permissionLevels -contains 'List') {
            Find-EEEUinLists -siteURL $siteURL -EEEUOccurrences ([ref]$siteEEEUOccurrences)
        }

        # Get all lists and libraries with throttling protection
        if ($permissionLevels -contains 'Folder' -or $permissionLevels -contains 'File') {
            $lists = Invoke-WithRetry -ScriptBlock {
                Get-PnPList | Where-Object { -not $_.Hidden -and -not ($ignoreFolderPatterns | Where-Object { $_.Title -like $_ }) }
            }
        }
        else {
            $Lists = [System.Collections.ArrayList]::new()
        }

        foreach ($list in $lists) {
            $ListItems = Get-AllItemsInList -listTitleOrUrl $list.Title -listId $list.Id

            # Check folder-level permissions
            if ($permissionLevels -contains 'Folder') {
                Find-EEEUinListRootFolder -siteURL $siteURL -listTitle $list.Title -EEEUOccurrences ([ref]$siteEEEUOccurrences)

                foreach ($item in $ListItems | Where-Object { $_.FileSystemObjectType -eq 1 }) {
                    Find-EEEUinFolders -siteURL $siteURL -item $item -listTitle $list.Title -EEEUOccurrences ([ref]$siteEEEUOccurrences)
                }
            }

            # Check file-level permissions
            if ($permissionLevels -contains 'File') {
                foreach ($item in $ListItems | Where-Object { $_.FileSystemObjectType -eq 0 }) {
                    Find-EEEUinFiles -siteURL $siteURL -item $item -listTitle $list.Title -EEEUOccurrences ([ref]$siteEEEUOccurrences)
                }
            }
        }

        # Write the results for this site collection to the CSV
        if ($siteEEEUOccurrences.Count -gt 0) {
            Write-Host "Writing $($siteEEEUOccurrences.Count) EEEU/Everyone occurrences from $siteURL to CSV..." -ForegroundColor Cyan
            Write-Log "About to write $($siteEEEUOccurrences.Count) EEEU/Everyone occurrences from $siteURL to CSV"

            # Debug: Log each occurrence before writing
            foreach ($occurrence in $siteEEEUOccurrences) {
                Write-Log "DEBUG - Occurrence: URL=$($occurrence.Url), ItemURL=$($occurrence.ItemURL), Type=$($occurrence.ItemType), Roles=$($occurrence.RoleNames)" 'DEBUG'
            }

            Write-EEEUOccurrencesToCSV -filePath $outputFilePath -Append -OccurrencesData $siteEEEUOccurrences
            Write-Log 'Finished writing occurrences to CSV'
        }
        else {
            Write-Host "No EEEU/Everyone occurrences found in $siteURL" -ForegroundColor Green
            Write-Log "No EEEU/Everyone occurrences found in $siteURL"
        }

        # Now process all subsites recursively
        $subsites = Invoke-WithRetry -ScriptBlock {
            Get-PnPSubWeb -Recurse:$false
        }

        if ($subsites -and $subsites.Count -gt 0) {
            Write-Host "Found $($subsites.Count) subsites to process" -ForegroundColor Yellow
            Write-Log "Found $($subsites.Count) subsites to process" 'DEBUG'

            foreach ($subsite in $subsites) {
                Write-Host "Processing subsite: $($subsite.Url)" -ForegroundColor Yellow
                Write-Log "Processing subsite: $($subsite.Url)" 'DEBUG'
                Process-SiteAndSubsites -siteURL $subsite.Url
            }
        }
    }

    Write-Host "Completed processing for $siteURL" -ForegroundColor Green
    Write-Log "Completed processing for $siteURL"
}

Write-Host ''
Write-Host '======================================================' -ForegroundColor 'Cyan'
Write-Host 'SCRIPT MODE: DETECTION'
Write-Host "  - Scanning for permissions at the $($permissionLevels -join ', ') level(s)"
Write-Host '  - Included identities are:'
Write-Host "    - Everyone ($Everyone)"
Write-Host "    - Everyone Except External Users ($EEEU)"
Write-Host "    - All Users (Windows) ($AllUsers) - a legacy claim similar in scope to EEEU"
Write-Host "  - Results will be saved to: $outputFilePath" -ForegroundColor Cyan
Write-Host "  - Log File will be saved to: $logFilePath" -ForegroundColor Cyan
if ($includeLimitedAccessPermissions) {
    Write-Host '  - Limited Access permissions will be included' -ForegroundColor Cyan
}
if ($debugLogging) {
    Write-Host '  - Debug logging is enabled' -ForegroundColor Cyan
}
Write-Host '  - NO modifications will be made to permissions' -ForegroundColor Cyan
Write-Host '======================================================' -ForegroundColor 'Cyan'
Write-Host ''

# Main script execution
$siteURLs = Read-SiteURLs -filePath $inputFilePath

# Create an empty output file with headers
Write-EEEUOccurrencesToCSV -filePath $outputFilePath

foreach ($siteURL in $siteURLs) {
    # Process the site and all its subsites recursively
    Process-SiteAndSubsites -siteURL $siteURL
}

# Final message, don't need another CSV write since we've been writing after each site
Write-Host "EEEU occurrences scan completed. Results available in $outputFilePath" -ForegroundColor Green
Write-Log "EEEU occurrences scan completed. Results available in $outputFilePath"
