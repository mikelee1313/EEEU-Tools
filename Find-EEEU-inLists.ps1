<#
.SYNOPSIS
    Scans SharePoint Online sites to identify occurrences of the "Everyone Except External Users" (EEEU) group
    in PowerApps-enabled list permissions.

.DESCRIPTION
    This script connects to SharePoint Online using provided tenant-level credentials and iterates through a list of
    site URLs specified in an input file. It scans document libraries and lists (excluding specified folders)
    to locate PowerApps-enabled lists where the "Everyone Except External Users" group has permissions assigned
    (excluding "Limited Access"). The script logs its operations and outputs the results to a CSV file, detailing
    the site URL, list name, PowerApps App ID, and assigned roles.

.NOTES
    Configure the $appsonly variable in the Configuration Section to control scan scope:
    - $true (default): Only scan PowerApps-enabled lists for EEEU
    - $false: Scan ALL lists for EEEU, regardless of PowerApps integration

    This script scans only Lists/Libraries level for EEEU permissions.

.INPUTS
    A text file containing SharePoint site URLs to scan (path specified in $inputFilePath variable).

.OUTPUTS
    - A CSV file containing all found EEEU occurrences (path: $env:TEMP\Find_EEEU_In_Sites_[timestamp].csv)
    - A log file documenting the script's execution (path: $env:TEMP\Find_EEEU_In_Sites_[timestamp].txt)

.NOTES
    File Name      : Find-EEEU-inLists.ps1
    Author         : Mike Lee
    Date Created   : 10/7/25

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
    .\Find-EEEUinLists.ps1
    Executes the script with the configured settings.
    Check the $appsonly variable in the Configuration Section to control scan scope.
#>

#==============================================================================
#                        CONFIGURATION SECTION
#==============================================================================
# Modify the settings below to match your environment and requirements
#==============================================================================

#------------------------------------------------------------------------------
# Scan Mode Settings
#------------------------------------------------------------------------------
$appsonly = $false           # $true = Only PowerApps-enabled lists | $false = ALL lists with EEEU

#------------------------------------------------------------------------------
# Authentication Settings
#------------------------------------------------------------------------------
# Entra App Registration details for app-only authentication
# Required permissions: SharePoint Sites.FullControl.All
$appID = '1e488dc4-1977-48ef-8d4d-9856f4e04536'                 # Entra App ID (Client ID)
$thumbprint = '5EAD7303A5C7E27DB4245878AD554642940BA082'        # Certificate thumbprint
$tenant = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'                # Tenant ID (GUID)

#------------------------------------------------------------------------------
# Input/Output File Paths
#------------------------------------------------------------------------------
$inputFilePath = 'C:\temp\sitelist.txt'  # Input file with site URLs (one per line)
$startime = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFilePath = "$env:TEMP\Find-EEEU-inApps_$startime.txt"       # Log file output path
$outputFilePath = "$env:TEMP\Find-EEEU-inApps_$startime.csv"    # CSV results output path

#------------------------------------------------------------------------------
# Scan Behavior Settings
#------------------------------------------------------------------------------
$debugLogging = $false          # Set to $true for verbose logging, $false for essential logging only

#------------------------------------------------------------------------------
# Retry and Throttling Settings
#------------------------------------------------------------------------------
$maxRetries = 5                 # Maximum number of retry attempts for throttled requests
$initialDelaySeconds = 5        # Initial delay in seconds before first retry (uses exponential backoff)

#------------------------------------------------------------------------------
# Exclusion Patterns
#------------------------------------------------------------------------------
# List of folder/list patterns to ignore during scanning (uses wildcard matching)
# These are typically system lists that should not be scanned
$ignoreFolderPatterns = @(
    '*VivaEngage*',                                             # Viva Engage folder for Storyline attachments
    '*DO_NOT_DELETE_REVIEW_INSTANCE*',                          # Review Instances (handles both singular and plural)
    '*Style Library*',                                          # Style Library
    '*_catalogs*',                                              # System catalogs
    '*_cts*',                                                   # Content Type Syndication
    '*_private*',                                               # Private folders
    '*_vti_pvt*',                                               # FrontPage folders
    '*Reference*',                                              # Reference folders
    '*Sharing Links*',                                          # Sharing Links
    '*Social*',                                                 # Social features
    '*FavoriteLists*',                                          # Favorite Lists
    '*User Information List*',                                  # User Information List
    '*Web Template Extensions*',                                # Web Template Extensions
    '*SmartCache*',                                             # SmartCache
    '*SharePointHomeCacheList*',                                # SharePoint Home Cache
    '*RecentLists*',                                            # Recent Lists
    '*PersonalCacheLibrary*',                                   # Personal Cache Library
    '*microsoft.ListSync.Endpoints*',                           # List Sync Endpoints
    '*Maintenance Log Library*',                                # Maintenance Logs
    '*DO_NOT_DELETE_ENTERPRISE_USER_CONTAINER_ENUM_LIST*',      # Enterprise User Container
    '*appfiles*'                                                # App files
)

#==============================================================================
#                   END OF CONFIGURATION SECTION
#==============================================================================
# Do not modify below this line unless you know what you are doing
#==============================================================================

# Internal constants (do not modify)
Add-Type -AssemblyName System.Web
$EEEU = '*spo-grid-all-users*'  # EEEU group login name pattern

# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = 'INFO'
    )

    # Only log essential messages when debug is false
    $essentialLevels = @('ERROR', 'WARNING')
    $isEssential = $level -in $essentialLevels -or
    $message -like '*Located EEEU*' -or
    $message -like '*Connected to SharePoint*' -or
    $message -like '*Failed to connect*' -or
    $message -like '*Processing site:*' -or
    $message -like '*Completed processing*' -or
    $message -like '*scan completed*'

    if ($debugLogging -or $isEssential) {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logMessage = "$timestamp - $level - $message"
        Add-Content -Path $logFilePath -Value $logMessage
    }
}

# Handle SharePoint Online throttling with exponential backoff
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = $maxRetries,
        [int]$InitialDelaySeconds = $initialDelaySeconds
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

            if ($null -ne $exception.Response) {
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
                # Check if it's an expected object reference error and log as DEBUG
                if ($_.Exception.Message -like '*Object reference not set to an instance of an object*' -or
                    $_.Exception.Message -like '*ListItemAllFields*' -or
                    $_.Exception.Message -like '*object is associated with property*') {
                    Write-Log "Expected retrieval error (likely null object reference): $($_.Exception.Message)" 'DEBUG'
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

# Helper function to check if a list has PowerApps enabled
function Test-PowerAppsEnabled {
    param (
        [Microsoft.SharePoint.Client.List]$list
    )

    try {
        # Get the list's properties
        $listProperties = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $list -Property RootFolder, Id, Title
        }

        $listServerRelativeUrl = $list.RootFolder.ServerRelativeUrl

        Write-Host "🔍 Checking PowerApps for list: $($list.Title)" -ForegroundColor Cyan
        Write-Log "Checking PowerApps via RenderListDataAsStream for list '$($list.Title)'" 'DEBUG'

        # Use RenderListDataAsStream API with AppAdditionalData as the deciding factor
        try {
            # Get the current site URL
            $web = Get-PnPWeb
            $siteUrl = $web.Url

            # Build the encoded list URL parameter
            $encodedListUrl = [System.Web.HttpUtility]::UrlEncode("'$listServerRelativeUrl'")
            $apiUrl = "$siteUrl/_api/web/GetList(@listUrl)/RenderListDataAsStream?@listUrl=$encodedListUrl"

            # Build the request body matching the working example
            $requestBody = @{
                parameters = @{
                    RenderOptions               = 64
                    ViewXml                     = "<View><ViewFields><FieldRef Name=`"ID`"/></ViewFields></View>"
                    AddRequiredFields           = $true
                    RequireFolderColoringFields = $true
                }
            }

            Write-Log "API URL: $apiUrl" 'DEBUG'
            Write-Host '  🔄 Calling RenderListDataAsStream API...' -ForegroundColor Gray

            # Use Invoke-PnPSPRestMethod which handles authentication automatically
            $response = Invoke-PnPSPRestMethod -Url $apiUrl -Method Post -Content $requestBody

            Write-Host '  ✅ API call succeeded!' -ForegroundColor Green
            Write-Log "RenderListDataAsStream response received for list '$($list.Title)'" 'DEBUG'

            # Check for AppId in the response - this is the definitive indicator of PowerApps
            $appId = $null
            if ($response.PSObject.Properties.Name -contains 'AppId') {
                $appId = $response.AppId
            }

            if ($appId) {
                Write-Host "  🎯 PowerApps AppId found: $appId" -ForegroundColor Magenta
                Write-Log "PowerApps AppId found: $appId" 'DEBUG'
                Write-Host '  ✅✅ PowerApps CONFIRMED!' -ForegroundColor Green
                Write-Log "✅ PowerApps CONFIRMED for list '$($list.Title)' - AppId: $appId" 'DEBUG'
                return $appId  # Return the AppId value
            }
            else {
                Write-Host '  ❌ No AppId in response - NOT PowerApps enabled' -ForegroundColor Yellow
                Write-Log "No AppId in response for list '$($list.Title)'" 'DEBUG'
                return $null  # Return null if no PowerApps
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Host "  ⚠️ API call failed: $errorMsg" -ForegroundColor Yellow
            Write-Log "RenderListDataAsStream API failed for list '$($list.Title)': $errorMsg" 'WARNING'
            return $false
        }
    }
    catch {
        Write-Log "Error checking PowerApps status for list '$($list.Title)': $_" 'WARNING'
        return $false
    }
}

# Function to check for EEEU in list-level permissions
function Find-EEEUinLists {
    param (
        [string]$siteURL
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

            # Get relative path by checking if DefaultViewUrl has a full URL or is relative
            $relativeUrl = $list.DefaultViewUrl
            # If it's already a relative URL, use it as is
            # If it contains the site URL, convert to relative
            if ($relativeUrl.StartsWith('https://')) {
                # Remove the site URL part to get the relative path
                $uri = New-Object System.Uri($siteURL)
                $siteRoot = $uri.AbsolutePath
                if ($relativeUrl.Contains($siteURL)) {
                    $relativeUrl = $relativeUrl.Replace($siteURL, '')
                }
            }

            # Check PowerApps status based on appsonly setting
            $appId = $null
            $isPowerAppEnabled = $false
            $powerAppStatus = 'N/A'

            if ($appsonly) {
                # Only check PowerApps when appsonly is true
                $appId = Test-PowerAppsEnabled -list $list
                $isPowerAppEnabled = $appId -ne $null
                $powerAppStatus = if ($isPowerAppEnabled) {
                    'True' 
                }
                else {
                    'False' 
                }

                # Skip EEEU check if PowerApps is not enabled
                if (-not $isPowerAppEnabled) {
                    Write-Host "Skipping list: $($list.Title) (PowerApps not enabled, appsonly=true)" -ForegroundColor Gray
                    Write-Log "Skipping list '$($list.Title)' - PowerApps not enabled, appsonly=true" 'DEBUG'
                    continue
                }

                Write-Host "PowerApps enabled - checking EEEU for list: $($list.Title)" -ForegroundColor Cyan
                Write-Log "PowerApps enabled - checking EEEU for list '$($list.Title)'" 'DEBUG'
            }
            else {
                # When appsonly is false, skip PowerApps detection and check all lists
                Write-Host "Checking EEEU for list: $($list.Title)" -ForegroundColor Yellow
                Write-Log "Checking EEEU for list '$($list.Title)' (appsonly=false, skipping PowerApps detection)" 'DEBUG'
            }

            # Initialize variables
            $roles = @()
            $hasEEEU = $false

            # Get list creator email and creation date information
            $listCreatorEmail = 'Unknown'
            $listCreatedDate = 'Unknown'

            try {
                # Get list properties including Author and Created
                Invoke-WithRetry -ScriptBlock {
                    Get-PnPProperty -ClientObject $list -Property Author, Created
                }

                # Get creator information
                if ($null -ne $list.Author) {
                    # Load the Author's properties (Title and Email)
                    Invoke-WithRetry -ScriptBlock {
                        Get-PnPProperty -ClientObject $list.Author -Property Title, Email, LoginName
                    }

                    # Now access the loaded properties
                    if ($null -ne $list.Author.Email -and $list.Author.Email -ne '') {
                        $listCreatorEmail = $list.Author.Email
                    }
                    else {
                        # If email is not available, try to use LoginName as fallback
                        if ($null -ne $list.Author.LoginName -and $list.Author.LoginName -ne '') {
                            $listCreatorEmail = $list.Author.LoginName
                        }
                    }
                }

                # Get creation date
                if ($list.Created) {
                    $listCreatedDate = $list.Created.ToString('yyyy-MM-dd HH:mm:ss')
                }
            }
            catch {
                Write-Log "Error retrieving list creator information for '$($list.Title)': $_" 'WARNING'
            }

            # Check if list has unique permissions
            $hasUniquePermissions = Invoke-WithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments
            }

            # Check for EEEU only if list has unique permissions
            if ($hasUniquePermissions) {
                # Get list permissions with throttling protection
                $Permissions = Invoke-WithRetry -ScriptBlock {
                    Get-PnPProperty -ClientObject $list -Property RoleAssignments
                }

                if ($Permissions) {
                    foreach ($RoleAssignment in $Permissions) {
                        # Get role assignments with throttling protection
                        Invoke-WithRetry -ScriptBlock {
                            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                        }

                        if ($RoleAssignment.Member.LoginName -like $EEEU -and $RoleAssignment.RoleDefinitionBindings.name -ne 'Limited Access') {
                            $hasEEEU = $true
                            $rolelevel = $RoleAssignment.RoleDefinitionBindings
                            foreach ($role in $rolelevel) {
                                $roles += $role.Name
                            }
                        }
                    }
                }
            }

            # Add list to output regardless of EEEU presence
            $global:EEEUOccurrences += [PSCustomObject]@{
                Url             = $SiteURL
                ListName        = $list.Title
                ItemURL         = $relativeUrl
                ItemType        = 'List'
                'EEEU Role'     = if ($roles.Count -gt 0) {
                    ($roles -join ', ') 
                }
                else {
                    'No EEEU' 
                }
                PowerAppEnabled = $powerAppStatus
                AppID           = if ($appId) {
                    $appId 
                }
                else {
                    'N/A' 
                }
                CreatorEmail    = $listCreatorEmail
                CreatedDate     = $listCreatedDate
            }

            if ($hasEEEU) {
                Write-Host "Located EEEU at List level: $($list.Title) on $SiteURL (PowerApps: $powerAppStatus)" -ForegroundColor Red
                Write-Log "Located EEEU at List level: $($list.Title) on $SiteURL (PowerApps: $powerAppStatus)"
            }
            else {
                Write-Host "List found (No EEEU): $($list.Title) on $SiteURL (PowerApps: $powerAppStatus)" -ForegroundColor Green
                Write-Log "List found (No EEEU): $($list.Title) on $SiteURL (PowerApps: $powerAppStatus)"
            }
        }
    }
    catch {
        Write-Log "Failed to process list-level permissions: $_" 'ERROR'
    }
}

# CSV output function to write EEEU occurrences to file
function Write-EEEUOccurrencesToCSV {
    param (
        [string]$filePath,
        [switch]$Append = $false,
        [array]$OccurrencesData = $global:EEEUOccurrences
    )
    try {
        # Create the file with headers if it doesn't exist or if we're not appending
        if (-not (Test-Path $filePath) -or -not $Append) {
            # Create empty file with headers
            'Url,ListName,ItemURL,ItemType,EEEU Role,PowerAppEnabled,AppID,CreatorEmail,CreatedDate' | Out-File -FilePath $filePath
        }

        # Group by URL, ListName, Item URL, ItemType and Roles to remove duplicates
        $uniqueOccurrences = $OccurrencesData |
        Group-Object -Property Url, ListName, ItemURL, ItemType, 'EEEU Role' |
        ForEach-Object { $_.Group[0] }

        # Append data to CSV
        foreach ($occurrence in $uniqueOccurrences) {
            # Manual CSV creation to handle special characters correctly
            $csvLine = "`"$($occurrence.Url)`",`"$($occurrence.ListName)`",`"$($occurrence.ItemURL)`",`"$($occurrence.ItemType)`",`"$($occurrence.'EEEU Role')`",`"$($occurrence.PowerAppEnabled)`",`"$($occurrence.AppID)`",`"$($occurrence.CreatorEmail)`",`"$($occurrence.CreatedDate)`""
            Add-Content -Path $filePath -Value $csvLine
        }

        Write-Log "EEEU occurrences have been written to $filePath" 'DEBUG'
    }
    catch {
        Write-Log "Failed to write EEEU occurrences to CSV file: $_" 'ERROR'
    }
}

# Function to recursively process subsites
function Invoke-SiteAndSubsites {
    param (
        [string]$siteURL
    )

    Write-Host "Processing site: $siteURL (List-Level Scan)" -ForegroundColor Green
    Write-Log "Processing site: $siteURL (List-Level Scan)"

    # Clear the global collection for this site
    $global:EEEUOccurrences = @()

    if (Connect-SharePoint -siteURL $siteURL) {
        # Check list-level permissions for PowerApps-enabled lists
        Find-EEEUinLists -siteURL $siteURL

        # Write the results for this site collection to the CSV
        if ($global:EEEUOccurrences.Count -gt 0) {
            Write-Host "Writing $($global:EEEUOccurrences.Count) EEEU occurrences from $siteURL to CSV..." -ForegroundColor Cyan
            Write-EEEUOccurrencesToCSV -filePath $outputFilePath -Append -OccurrencesData $global:EEEUOccurrences
        }
        else {
            Write-Host "No EEEU occurrences found in $siteURL" -ForegroundColor Green
            Write-Log "No EEEU occurrences found in $siteURL"
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
                Invoke-SiteAndSubsites -siteURL $subsite.Url
            }
        }
    }

    Write-Host "Completed processing for $siteURL" -ForegroundColor Green
    Write-Log "Completed processing for $siteURL"
}

# Main script execution
$global:EEEUOccurrences = @()

Write-Host '===== EEEU Scan Configuration =====' -ForegroundColor Cyan
Write-Host 'Scan Type: List-Level Only' -ForegroundColor Cyan
if ($appsonly) {
    Write-Host '  - PowerApps Detection: ENABLED' -ForegroundColor Green
    Write-Host '  - EEEU Check: Only for PowerApps-enabled lists' -ForegroundColor Green
    Write-Host '  - Mode: Apps Only (appsonly = true)' -ForegroundColor Green
}
else {
    Write-Host '  - PowerApps Detection: ENABLED (for reporting)' -ForegroundColor Green
    Write-Host '  - EEEU Check: ALL lists (regardless of PowerApps)' -ForegroundColor Yellow
    Write-Host '  - Mode: All Lists (appsonly = false)' -ForegroundColor Yellow
}
Write-Host '===================================' -ForegroundColor Cyan
Write-Log "Starting EEEU List-Level scan with appsonly=$appsonly"

$siteURLs = Read-SiteURLs -filePath $inputFilePath

# Create an empty output file with headers
Write-EEEUOccurrencesToCSV -filePath $outputFilePath

foreach ($siteURL in $siteURLs) {
    # Process the site and all its subsites recursively
    Invoke-SiteAndSubsites -siteURL $siteURL
}

# Final message, don't need another CSV write since we've been writing after each site
Write-Host "EEEU occurrences scan completed. Results available in $outputFilePath" -ForegroundColor Green
Write-Log "EEEU occurrences scan completed. Results available in $outputFilePath"
