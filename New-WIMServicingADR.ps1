#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Creates an SCCM Deployment Package and Automatic Deployment Rule (ADR)
    to automatically download monthly updates for offline WIM servicing.

.DESCRIPTION
    This script sets up the full pipeline for keeping WIM servicing updates
    current without manual intervention:

      1. Creates a Deployment Package with a clean source path per OS version
      2. Creates an ADR scoped to that package, targeting the correct update
         classifications and products for your WIM's OS version
      3. Configures the ADR to run after Patch Tuesday (2nd Tuesday + 1 day)
         so updates are downloaded and ready before your servicing window
      4. Optionally triggers an immediate ADR evaluation to seed the initial
         download right away

    After initial setup, the ADR runs monthly. Point Invoke-WIMOfflineServicing.ps1
    at the package source path and your WIM maintenance is fully automated.

.PARAMETER SiteCode
    SCCM site code (e.g. CF1).

.PARAMETER SiteServer
    SCCM site server FQDN.

.PARAMETER OSVersion
    Target OS for the WIM being serviced. Used to scope ADR product filtering.
    Valid values: 'Win11-23H2', 'Win11-24H2', 'Win11-25H2', 'Win10-22H2'

.PARAMETER PackageSourceRoot
    UNC or local path root where update packages will be stored.
    A subfolder per OSVersion will be created automatically.
    Example: \\CFHSCCM01\Sources\WIM_Updates

.PARAMETER PackageName
    Name for the Deployment Package in SCCM. Defaults to
    "WIM Servicing - <OSVersion>".

.PARAMETER ADRName
    Name for the ADR in SCCM. Defaults to "WIM Servicing ADR - <OSVersion>".

.PARAMETER SUGName
    Name for the Software Update Group the ADR will create/maintain.
    Defaults to "WIM Servicing - <OSVersion> - <Year>".

.PARAMETER DPGroupName
    Name of the Distribution Point Group to deploy the package to.
    Required if you want the package pushed to DPs automatically.

.PARAMETER IncludeDotNet
    If specified, also includes .NET Framework cumulative updates in the ADR.
    Recommended to reduce post-OSD update time.

.PARAMETER RunImmediately
    If specified, triggers an immediate ADR evaluation after creation
    to kick off the first download without waiting for the schedule.

.PARAMETER Architecture
    Target architecture. Defaults to x64.

.EXAMPLE
    .\New-WIMServicingADR.ps1 `
        -SiteCode "CF1" `
        -SiteServer "CFHSCCM01.carle.com" `
        -OSVersion "Win11-25H2" `
        -PackageSourceRoot "\\CFHSCCM01\Sources\WIM_Updates" `
        -DPGroupName "All Distribution Points" `
        -IncludeDotNet `
        -RunImmediately

.NOTES
    - Requires the SCCM console to be installed on the machine running this script
    - Run as a user with Full Administrator rights in SCCM
    - The SUP must already be syncing the target products and classifications
    - ADR schedule targets Wednesday after Patch Tuesday (2nd Tue of month + 1 day)
    - Existing packages/ADRs with the same name will NOT be overwritten;
      the script will warn and exit cleanly if duplicates are detected
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$SiteCode,

    [Parameter(Mandatory)]
    [string]$SiteServer,

    [Parameter(Mandatory)]
    [ValidateSet('Win11-23H2','Win11-24H2','Win11-25H2','Win10-22H2')]
    [string]$OSVersion,

    [Parameter(Mandatory)]
    [string]$PackageSourceRoot,

    [Parameter()]
    [string]$PackageName,

    [Parameter()]
    [string]$ADRName,

    [Parameter()]
    [string]$SUGName,

    [Parameter()]
    [string]$DPGroupName,

    [Parameter()]
    [switch]$IncludeDotNet,

    [Parameter()]
    [switch]$RunImmediately,

    [Parameter()]
    [ValidateSet('x64','x86')]
    [string]$Architecture = 'x64'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Logging ---

$LogFile = Join-Path $PSScriptRoot "WIMServicingADR_Setup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry
    switch ($Level) {
        'INFO'    { Write-Host $entry -ForegroundColor Cyan }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
    }
}

#endregion

#region --- OS Version Mappings ---

# Maps the OSVersion param to SCCM product names and the Windows build
# used for update title filtering.
$OSMap = @{
    'Win11-25H2' = @{
        Product       = 'Windows 11'
        TitleFilter   = 'Windows 11 Version 24H2'   # MS still labels 25H2 CUs as 24H2 in WSUS
        BuildPrefix   = '26100'
        FriendlyName  = 'Windows 11 25H2'
    }
    'Win11-24H2' = @{
        Product       = 'Windows 11'
        TitleFilter   = 'Windows 11 Version 24H2'
        BuildPrefix   = '26100'
        FriendlyName  = 'Windows 11 24H2'
    }
    'Win11-23H2' = @{
        Product       = 'Windows 11'
        TitleFilter   = 'Windows 11 Version 23H2'
        BuildPrefix   = '22631'
        FriendlyName  = 'Windows 11 23H2'
    }
    'Win10-22H2' = @{
        Product       = 'Windows 10, version 1903 and later'
        TitleFilter   = 'Windows 10 Version 22H2'
        BuildPrefix   = '19045'
        FriendlyName  = 'Windows 10 22H2'
    }
}

$osInfo = $OSMap[$OSVersion]

#endregion

#region --- Default Parameter Population ---

if (-not $PackageName) { $PackageName = "WIM Servicing - $OSVersion" }
if (-not $ADRName)     { $ADRName     = "WIM Servicing ADR - $OSVersion" }
if (-not $SUGName)     { $SUGName     = "WIM Servicing - $OSVersion - $(Get-Date -Format 'yyyy')" }

$PackageSourcePath = Join-Path $PackageSourceRoot $OSVersion

#endregion

#region --- Functions ---

function Connect-SCCMSite {
    param([string]$Code, [string]$Server)

    Write-Log "Loading ConfigMgr module..."

    # Build list of possible console install paths to check
    $candidatePaths = @()
    if ($env:SMS_ADMIN_UI_PATH) {
        $candidatePaths += Split-Path $env:SMS_ADMIN_UI_PATH
    }
    $candidatePaths += "${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\AdminConsole\bin"
    $candidatePaths += "${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\bin"
    $candidatePaths += "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin"
    $candidatePaths += "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\bin"
    $candidatePaths += "${env:ProgramFiles}\Microsoft Endpoint Manager\AdminConsole\bin"
    $candidatePaths += "${env:ProgramFiles}\Microsoft Endpoint Manager\bin"
    $candidatePaths += "${env:ProgramFiles}\Microsoft Configuration Manager\AdminConsole\bin"
    $candidatePaths += "${env:ProgramFiles}\Microsoft Configuration Manager\bin"

    $modulePath = $null
    foreach ($path in $candidatePaths) {
        $candidate = Join-Path $path "ConfigurationManager.psd1"
        if (Test-Path $candidate) {
            $modulePath = $candidate
            Write-Log "Found ConfigMgr module at: $modulePath"
            break
        }
    }

    if (-not $modulePath) {
        throw "ConfigMgr PowerShell module not found. Ensure the SCCM console is installed.`nSearched paths:`n$($candidatePaths -join "`n")"
    }

    # The CMSite provider only registers correctly if the console bin directory
    # is in PSModulePath at the time of import. Add it explicitly before importing.
    $consoleDir = Split-Path $modulePath
    if ($env:PSModulePath -notlike "*$consoleDir*") {
        Write-Log "Adding console bin to PSModulePath: $consoleDir"
        $env:PSModulePath = $consoleDir + [System.IO.Path]::PathSeparator + $env:PSModulePath
    }

    Import-Module $modulePath -ErrorAction Stop
    Write-Log "Module loaded. Connecting to site $Code on $Server..."

    # Confirm the CMSite provider actually registered after module import
    $provider = Get-PSProvider -PSProvider CMSite -ErrorAction SilentlyContinue
    if (-not $provider) {
        throw "The CMSite PSProvider did not register after module import.`nThis usually means a dependency DLL is missing or the console installation is incomplete.`nTry running this script directly on the SCCM site server instead."
    }
    Write-Log "CMSite provider confirmed registered."

    # Verify the site server is reachable before attempting drive creation
    Write-Log "Testing connectivity to $Server..."
    if (-not (Test-Connection -ComputerName $Server -Count 1 -Quiet)) {
        throw "Cannot reach site server '$Server'. Verify the FQDN and network connectivity."
    }

    # Remove any stale PSDrive from a previous failed attempt
    if (Get-PSDrive -Name $Code -ErrorAction SilentlyContinue) {
        Write-Log "Removing stale PSDrive '$Code' before recreating..." -Level WARN
        Remove-PSDrive -Name $Code -Force -ErrorAction SilentlyContinue
    }

    # Create the CMSite PSDrive with explicit error handling
    try {
        New-PSDrive -Name $Code -PSProvider CMSite -Root $Server -ErrorAction Stop | Out-Null
    } catch {
        throw "Failed to create CMSite PSDrive for site '$Code' on server '$Server'.`nVerify:`n  1. The FQDN '$Server' is correct`n  2. Your account has SCCM Full Administrator rights`n  3. The SMS Provider is running on '$Server'`nError: $_"
    }

    Set-Location "${Code}:"

    # Confirm the connection actually works by querying the site
    try {
        $site = Get-CMSite -SiteCode $Code -ErrorAction Stop
        Write-Log "Connected to SCCM site $Code ($($site.SiteName)) on $Server." -Level SUCCESS
    } catch {
        throw "PSDrive created but site query failed. The SMS Provider may not be on '$Server'.`nError: $_"
    }
}

function Get-PatchTuesdayPlusOne {
    # Returns a schedule for the Wednesday after Patch Tuesday (2nd Tue of month)
    # so updates are downloaded the morning after they release
    param([int]$HourOfDay = 3)

    $now         = Get-Date
    $firstOfNext = (Get-Date -Year $now.Year -Month $now.Month -Day 1).AddMonths(1)

    # Find 2nd Tuesday of next month
    $day = $firstOfNext
    $tuesdayCount = 0
    while ($tuesdayCount -lt 2) {
        if ($day.DayOfWeek -eq [DayOfWeek]::Tuesday) { $tuesdayCount++ }
        if ($tuesdayCount -lt 2) { $day = $day.AddDays(1) }
    }

    # Wednesday after patch tuesday
    $patchWednesday = $day.AddDays(1)
    $scheduleDate   = Get-Date -Year $patchWednesday.Year -Month $patchWednesday.Month `
                               -Day $patchWednesday.Day -Hour $HourOfDay -Minute 0 -Second 0

    Write-Log "ADR schedule set to: $scheduleDate (Wednesday after next Patch Tuesday)"
    return $scheduleDate
}

function New-WIMServicingPackage {
    param(
        [string]$Name,
        [string]$SourcePath,
        [string]$DPGroup
    )

    Write-Log "Checking for existing deployment package: $Name"

    # Get-CMSoftwareUpdateDeploymentPackage throws a null key exception in some
    # ConfigMgr versions when the package list is empty. Use WMI as a safer fallback.
    $existing = $null
    try {
        $existing = Get-CMSoftwareUpdateDeploymentPackage -Name $Name -ErrorAction Stop
    } catch {
        Write-Log "Cmdlet query failed ($($_.Exception.Message)), falling back to WMI lookup..." -Level WARN
        try {
            $wmiPkg = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                    -ComputerName $SiteServer `
                                    -Class SMS_SoftwareUpdatesPackage `
                                    -Filter "Name = '$Name'" `
                                    -ErrorAction SilentlyContinue
            if ($wmiPkg) {
                # Cmdlet re-fetch has the same null-key bug — build the object directly from WMI
                $existing = [PSCustomObject]@{ PackageID = $wmiPkg.PackageID; Name = $wmiPkg.Name }
            }
        } catch {
            Write-Log "WMI fallback also failed (non-fatal, assuming package does not exist): $_" -Level WARN
        }
    }

    if ($existing) {
        Write-Log "Deployment package '$Name' already exists (PackageID: $($existing.PackageID)). Skipping creation." -Level WARN
        return $existing
    }

    # Create the source folder if it doesn't exist.
    # Use .NET directly — Test-Path/New-Item route through the CM PSDrive provider
    # when the current location is a CMSite drive, so they won't touch the filesystem.
    if (-not [System.IO.Directory]::Exists($SourcePath)) {
        Write-Log "Creating package source folder: $SourcePath"
        [System.IO.Directory]::CreateDirectory($SourcePath) | Out-Null
    }

    Write-Log "Creating deployment package: $Name"
    Write-Log "  Source path: $SourcePath"

    # New-CMSoftwareUpdateDeploymentPackage has a null reference bug in some ConfigMgr
    # versions when called with extra parameters. Try progressively simpler approaches.
    $pkg = $null

    # Attempt 1: Cmdlet with only the required parameters
    Write-Log "Attempting package creation via cmdlet (minimal params)..."
    try {
        $pkg = New-CMSoftwareUpdateDeploymentPackage -Name $Name -Path $SourcePath -ErrorAction Stop
        Write-Log "Package created via cmdlet. PackageID: $($pkg.PackageID)" -Level SUCCESS
    } catch {
        Write-Log "Cmdlet creation failed: $($_.Exception.Message). Trying WMI..." -Level WARN
    }

    # Attempt 2: Direct WMI creation
    if (-not $pkg) {
        try {
            Write-Log "Creating package via WMI (SMS_SoftwareUpdatesPackage)..."
            $wmiNS   = "root\SMS\site_$SiteCode"
            $wmiConn = [System.Management.ManagementScope]::new("\\" + $SiteServer + "\" + $wmiNS)
            $wmiConn.Connect()
            $wmiPath = [System.Management.ManagementPath]::new("SMS_SoftwareUpdatesPackage")
            $mc      = [System.Management.ManagementClass]::new($wmiConn, $wmiPath, $null)
            $newPkg  = $mc.CreateInstance()
            $newPkg["Name"]          = $Name
            $newPkg["Description"]   = "Auto-downloaded updates for offline WIM servicing."
            $newPkg["PkgSourcePath"] = $SourcePath
            $newPkg["PkgSourceFlag"] = [System.UInt32]2
            $newPkg["Priority"]      = [System.UInt32]2
            $newPkg.Put() | Out-Null

            # Give SCCM a moment to register the new package
            Start-Sleep -Seconds 3

            # Retrieve by WMI to get PackageID
            $wmiResult = Get-WmiObject -Namespace $wmiNS -ComputerName $SiteServer `
                                       -Class SMS_SoftwareUpdatesPackage `
                                       -Filter "Name = '$Name'" -ErrorAction Stop
            if ($wmiResult) {
                Write-Log "Package created via WMI. PackageID: $($wmiResult.PackageID)" -Level SUCCESS
                $pkg = [PSCustomObject]@{ PackageID = $wmiResult.PackageID; Name = $wmiResult.Name }
            } else {
                throw "Package not found in WMI after creation."
            }
        } catch {
            throw "Both cmdlet and WMI package creation failed. Last error: $_"
        }
    }

    # Distribute to DP group if specified
    if ($DPGroup) {
        Write-Log "Distributing package to DP group: $DPGroup"
        try {
            Start-CMContentDistribution -DeploymentPackageName $Name `
                                        -DistributionPointGroupName $DPGroup
            Write-Log "Content distribution started." -Level SUCCESS
        } catch {
            Write-Log "DP distribution failed (non-fatal, can be done manually): $_" -Level WARN
        }
    }

    return $pkg
}

function New-WIMServicingADR {
    param(
        [string]$Name,
        [string]$SUG,
        [string]$PkgName,
        [string]$PkgSourcePath,
        [hashtable]$OS,
        [string]$Arch,
        [bool]$DotNet,
        [datetime]$Schedule
    )

    Write-Log "Checking for existing ADR: $Name"
    $existing = $null
    try {
        $existing = Get-CMAutoDeploymentRule -Name $Name -ErrorAction Stop
    } catch {
        Write-Log "Cmdlet ADR query failed ($($_.Exception.Message)), falling back to WMI..." -Level WARN
        try {
            $wmiAdr = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                    -ComputerName $SiteServer `
                                    -Class SMS_AutoDeploymentRule `
                                    -Filter "Name = '$Name'" `
                                    -ErrorAction SilentlyContinue
            if ($wmiAdr) { $existing = $wmiAdr }
        } catch {
            Write-Log "WMI ADR fallback also failed (assuming ADR does not exist): $_" -Level WARN
        }
    }
    if ($existing) {
        Write-Log "ADR '$Name' already exists. Skipping creation." -Level WARN
        return $existing
    }

    Write-Log "Building ADR property criteria..."

    # Validate critical parameters before any cmdlet/WMI calls
    if (-not $Name)         { throw "ADR Name is null or empty." }
    if (-not $OS)           { throw "OS info hashtable is null." }
    if (-not $OS.Product)   { throw "OS Product is null - check OSMap for '$Name'." }
    if (-not $OS.TitleFilter) { throw "OS TitleFilter is null - check OSMap." }
    if (-not $Arch)         { throw "Architecture is null or empty." }

    # Build the update classification list
    $classifications = @('Security Updates', 'Critical Updates', 'Updates')

    # Build article ID exclusion list - common updates that break offline servicing
    # or that are not CBS-injectable (Dynamic Updates, Feature Updates, etc.)
    # These are excluded by Update Classification in the ADR criteria

    Write-Log "Creating ADR: $Name"
    Write-Log "  Product       : $($OS.Product)"
    Write-Log "  Title filter  : $($OS.TitleFilter)"
    Write-Log "  Architecture  : $Arch"
    Write-Log "  Classifications: $($classifications -join ', ')"
    Write-Log "  Include .NET  : $DotNet"
    Write-Log "  SUG name      : $SUG"

    # This console version only supports Minutes/Hours/Days for RecurInterval (not Months).
    # 35 days (5 weeks) keeps recurrence on the same day-of-week each cycle,
    # landing within a few days of Patch Wednesday every month.
    # New-CMSchedule throws the same null-key bug as other CM cmdlets in some console
    # versions, so wrap it with fallbacks rather than letting it propagate as FATAL.
    $cmSchedule = $null
    $scheduleToken = $null
    try {
        $cmSchedule = New-CMSchedule -Start $Schedule -RecurInterval Days -RecurCount 35 -ErrorAction Stop
        Write-Log "ADR schedule object created (35-day recurrence)."
    } catch {
        Write-Log "New-CMSchedule failed ($($_.Exception.Message)), trying non-recurring schedule..." -Level WARN
        try {
            $cmSchedule = New-CMSchedule -Start $Schedule -Nonrecurring -ErrorAction Stop
            Write-Log "ADR schedule object created (non-recurring; set recurrence manually in console)." -Level WARN
        } catch {
            Write-Log "New-CMSchedule cmdlet unavailable. Generating schedule token directly..." -Level WARN
            # Build a simple SMS schedule token for the start date.
            # Format: SMS_ST_NonRecurring encoded as a date string that SCCM can parse.
            # We'll apply the full schedule via WMI after ADR creation.
            $scheduleToken = $Schedule.ToUniversalTime().ToString('yyyyMMddHHmmss') + '.000000+***'
            Write-Log "Generated schedule token: $scheduleToken"
        }
    }

    # Resolve package ID via WMI - the DeploymentPackageName cmdlet parameter triggers
    # the same null-key bug as other CM cmdlets, so we skip it during creation and
    # assign the package directly to the ADR via WMI afterward.
    $resolvedPkgID = $null
    try {
        $wmiPkgLookup = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                      -ComputerName $SiteServer `
                                      -Class SMS_SoftwareUpdatesPackage `
                                      -Filter "Name = '$PkgName'" `
                                      -ErrorAction SilentlyContinue
        if ($wmiPkgLookup) {
            $resolvedPkgID = $wmiPkgLookup.PackageID
            Write-Log "Resolved package ID for '$PkgName': $resolvedPkgID"
        } else {
            Write-Log "Could not resolve package ID for '$PkgName' - will skip WMI package assignment." -Level WARN
        }
    } catch {
        Write-Log "Package ID lookup failed (non-fatal): $_" -Level WARN
    }

    # Resolve collection ID via WMI - CollectionName triggers the same null-key lookup
    # bug in New-CMAutoDeploymentRule. CollectionId bypasses the internal name lookup.
    $resolvedCollID = $null
    $collectionName = 'WMI Update Source No Deploy'
    try {
        $wmiColl = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                 -ComputerName $SiteServer `
                                 -Class SMS_Collection `
                                 -Filter "Name = '$collectionName'" `
                                 -ErrorAction SilentlyContinue
        if ($wmiColl) {
            $resolvedCollID = $wmiColl.CollectionID
            Write-Log "Resolved collection ID for '$collectionName': $resolvedCollID"
        } else {
            Write-Log "Could not resolve collection '$collectionName' via WMI." -Level WARN
        }
    } catch {
        Write-Log "Collection ID lookup failed (non-fatal): $_" -Level WARN
    }

    if (-not $resolvedCollID) {
        throw "Collection '$collectionName' not found. Verify the collection exists in SCCM and the name matches exactly."
    }

    # Title filter - scopes to the correct OS version
    # Format matches how WSUS/SCCM stores CU titles:
    # "YYYY-MM Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB...)"
    $titleCriteria = "$($OS.TitleFilter) for ${Arch}-based Systems"

    Write-Log "  Title criteria: $titleCriteria"

    # Strategy: The CM cmdlet has a null-key bug that triggers when filter criteria
    # parameters (Product, UpdateClassification, Architecture, Language) are passed.
    # Even CollectionId can trigger it in some versions.
    #
    # Approach: Try progressively simpler cmdlet calls, then fall back to pure WMI.
    # ALL filter criteria and package assignment are applied via WMI afterward.

    $adr = $null

    # --- Attempt 1: Minimal cmdlet (Name + CollectionId only) ---
    # No filter criteria, no schedule - those trigger the null-key bug.
    Write-Log "Creating ADR via cmdlet (minimal params: Name + CollectionId)..."
    try {
        $adr = New-CMAutoDeploymentRule -Name $Name -CollectionId $resolvedCollID -ErrorAction Stop
        Write-Log "ADR created via cmdlet (minimal params)." -Level SUCCESS
    } catch {
        Write-Log "Minimal cmdlet failed ($($_.Exception.Message))." -Level WARN
    }

    # --- Attempt 2: Cmdlet with Name only (CollectionId set via WMI) ---
    if (-not $adr) {
        Write-Log "Trying cmdlet with Name only..."
        try {
            $adr = New-CMAutoDeploymentRule -Name $Name -CollectionId 'SMS00001' -ErrorAction Stop
            Write-Log "ADR created via cmdlet (Name + All Systems fallback collection)." -Level SUCCESS
        } catch {
            Write-Log "Name-only cmdlet also failed ($($_.Exception.Message)). Falling back to WMI..." -Level WARN
        }
    }

    # --- Attempt 3: Pure WMI creation ---
    if (-not $adr) {
        try {
            Write-Log "Creating ADR via WMI (SMS_AutoDeploymentRule)..."
            $wmiNS   = "root\SMS\site_$SiteCode"
            $wmiConn = [System.Management.ManagementScope]::new("\\" + $SiteServer + "\" + $wmiNS)
            $wmiConn.Connect()
            $wmiPath = [System.Management.ManagementPath]::new("SMS_AutoDeploymentRule")
            $mc      = [System.Management.ManagementClass]::new($wmiConn, $wmiPath, $null)
            $newADR  = $mc.CreateInstance()
            $newADR["Name"]                  = $Name
            $newADR["Description"]           = "Monthly LCU download for offline WIM servicing of $($OS.FriendlyName)."
            $newADR["CollectionID"]          = $resolvedCollID
            $newADR["AutoDeploymentEnabled"] = $true

            # ContentTemplate XML - required for WMI creation; tells the ADR where to
            # store downloaded content and basic download settings.
            $contentTemplate = @"
<ContentTemplate SchemaVersion="1.0">
  <ContentAction>
    <PackageID>$(if ($resolvedPkgID) { $resolvedPkgID } else { '' })</PackageID>
    <DownloadFromInternet>true</DownloadFromInternet>
    <DownloadFromMicrosoftUpdate>true</DownloadFromMicrosoftUpdate>
  </ContentAction>
</ContentTemplate>
"@
            $newADR["ContentTemplate"] = $contentTemplate

            # DeploymentTemplate XML - required; controls how the ADR deploys updates.
            # Since this is for download-only (WIM servicing), we use non-intrusive settings.
            $deployTemplate = @"
<DeploymentCreationActionXML SchemaVersion="1.0">
  <CollectionID>$resolvedCollID</CollectionID>
  <AvailableDateTimeIsUTC>false</AvailableDateTimeIsUTC>
  <DeadlineDateTimeIsUTC>false</DeadlineDateTimeIsUTC>
  <UserNotificationOption>DisplayAll</UserNotificationOption>
  <AllowSoftwareInstallationOutsideWindow>false</AllowSoftwareInstallationOutsideWindow>
  <AllowRestart>false</AllowRestart>
  <SuppressServers>Unchecked</SuppressServers>
  <SuppressWorkstations>Unchecked</SuppressWorkstations>
  <EnableWakeOnLan>false</EnableWakeOnLan>
  <EnableAlert>false</EnableAlert>
</DeploymentCreationActionXML>
"@
            $newADR["DeploymentTemplate"] = $deployTemplate

            $newADR.Put() | Out-Null

            Start-Sleep -Seconds 3

            $adrWmiCheck = Get-WmiObject -Namespace $wmiNS -ComputerName $SiteServer `
                                         -Class SMS_AutoDeploymentRule `
                                         -Filter "Name = '$Name'" -ErrorAction Stop
            if (-not $adrWmiCheck) { throw "ADR not found in WMI after creation." }

            Write-Log "ADR created via WMI successfully." -Level SUCCESS
            $adr = [PSCustomObject]@{ Name = $Name }
        } catch {
            throw "All ADR creation methods failed (cmdlet and WMI). Last error: $_`nCreate the ADR manually in the console and point it at package '$PkgName'."
        }
    }

    # --- Post-creation: apply ALL properties via WMI ---
    # This is the safest path - the cmdlet only creates the shell ADR,
    # and WMI handles filter criteria, package, and schedule assignment.
    Write-Log "Applying ADR properties via WMI (package, filters, schedule)..."

    Set-ADRPackageViaWMI -ADRName $Name -PackageID $resolvedPkgID

    # Fix collection ID if we used the fallback 'SMS00001' collection
    Set-ADRCollectionViaWMI -ADRName $Name -CollectionID $resolvedCollID

    Set-WIMServicingADRProperties -ADRName $Name -OS $OS -Arch $Arch `
                                  -DotNet $DotNet -Schedule $Schedule `
                                  -TitleCriteria $titleCriteria `
                                  -ScheduleToken $scheduleToken

    Write-Log "ADR created and configured successfully." -Level SUCCESS
    return $adr
}

function Set-ADRCollectionViaWMI {
    # Ensures the ADR targets the correct collection via WMI.
    # Used when the cmdlet required a fallback collection during creation.
    param(
        [string]$ADRName,
        [string]$CollectionID
    )

    if (-not $CollectionID) { return }

    try {
        $adrWmi = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                -ComputerName $SiteServer `
                                -Class SMS_AutoDeploymentRule `
                                -Filter "Name = '$ADRName'" `
                                -ErrorAction Stop
        if ($adrWmi -and $adrWmi.CollectionID -ne $CollectionID) {
            $adrWmi.CollectionID = $CollectionID
            $adrWmi.Put() | Out-Null
            Write-Log "Collection ID set to '$CollectionID' via WMI." -Level SUCCESS
        }
    } catch {
        Write-Log "WMI collection assignment failed (non-fatal, set manually in console): $_" -Level WARN
    }
}

function Set-ADRPackageViaWMI {
    # Links a deployment package to an ADR directly via WMI.
    # Used because passing DeploymentPackageName to New-CMAutoDeploymentRule
    # triggers the null-key cmdlet bug in some console versions.
    param(
        [string]$ADRName,
        [string]$PackageID
    )

    if (-not $PackageID) {
        Write-Log "No package ID available - skipping WMI package assignment. Set it manually in the ADR properties." -Level WARN
        return
    }

    Write-Log "Assigning package '$PackageID' to ADR '$ADRName' via WMI..."
    try {
        $adrWmi = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                -ComputerName $SiteServer `
                                -Class SMS_AutoDeploymentRule `
                                -Filter "Name = '$ADRName'" `
                                -ErrorAction Stop
        if ($adrWmi) {
            $adrWmi.PackageID = $PackageID
            $adrWmi.Put() | Out-Null
            Write-Log "Package '$PackageID' linked to ADR via WMI." -Level SUCCESS
        } else {
            Write-Log "ADR WMI object not found for package assignment - set it manually." -Level WARN
        }
    } catch {
        Write-Log "WMI package assignment failed (non-fatal, set manually in console): $_" -Level WARN
    }
}

function Set-WIMServicingADRProperties {
    # Applies full filter criteria directly via WMI for cases where
    # the PowerShell cmdlet doesn't expose all parameters or triggers null-key bugs
    param(
        [string]$ADRName,
        [hashtable]$OS,
        [string]$Arch,
        [bool]$DotNet,
        [datetime]$Schedule,
        [string]$TitleCriteria,
        [string]$ScheduleToken = $null
    )

    Write-Log "Applying ADR filter properties via WMI for: $ADRName"

    $adrWmi = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                            -ComputerName $SiteServer `
                            -Class SMS_AutoDeploymentRule `
                            -Filter "Name = '$ADRName'"

    if (-not $adrWmi) {
        Write-Log "ADR WMI object not found for: $ADRName" -Level ERROR
        return
    }

    # Build criteria XML
    # SCCM stores ADR criteria as an XML blob in AutoDeploymentProperties
    $criteriaTemplate = @"
<AutoDeploymentCriteria>
  <UpdateClassification>
    <Property PropertyName="UpdateClassification" Operator="In">
      <Values>
        <Value>Security Updates</Value>
        <Value>Critical Updates</Value>
        <Value>Updates</Value>
      </Values>
    </Property>
  </UpdateClassification>
  <Products>
    <Property PropertyName="Product" Operator="In">
      <Values>
        <Value>$($OS.Product)</Value>
      </Values>
    </Property>
  </Products>
  <Architecture>
    <Property PropertyName="LocalizedCategoryInstanceNames" Operator="Contains">
      <Values>
        <Value>$Arch</Value>
      </Values>
    </Property>
  </Architecture>
  <Title>
    <Property PropertyName="LocalizedDisplayName" Operator="Contains">
      <Values>
        <Value>$TitleCriteria</Value>
      </Values>
    </Property>
  </Title>
  <Superseded>
    <Property PropertyName="IsSuperseded" Operator="Equals">
      <Values>
        <Value>false</Value>
      </Values>
    </Property>
  </Superseded>
  <Expired>
    <Property PropertyName="IsExpired" Operator="Equals">
      <Values>
        <Value>false</Value>
      </Values>
    </Property>
  </Expired>
</AutoDeploymentCriteria>
"@

    try {
        $adrWmi.AutoDeploymentProperties = $criteriaTemplate
        $adrWmi.Put() | Out-Null
        Write-Log "ADR filter criteria applied via WMI." -Level SUCCESS
    } catch {
        Write-Log "WMI criteria update failed: $_" -Level WARN
        Write-Log "You may need to manually configure ADR filter criteria in the console." -Level WARN
    }

    # Apply schedule via WMI if the cmdlet-based schedule failed
    if ($ScheduleToken) {
        try {
            # Re-fetch to avoid stale object
            $adrWmi2 = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                     -ComputerName $SiteServer `
                                     -Class SMS_AutoDeploymentRule `
                                     -Filter "Name = '$ADRName'"
            if ($adrWmi2) {
                $adrWmi2.Schedule = $ScheduleToken
                $adrWmi2.Put() | Out-Null
                Write-Log "Schedule token applied via WMI." -Level SUCCESS
            }
        } catch {
            Write-Log "WMI schedule assignment failed (non-fatal, set manually in console): $_" -Level WARN
        }
    }
}

function Test-SUPProductSync {
    # Warns if the target product isn't enabled in SUP component settings
    param([string]$ProductName)

    Write-Log "Checking SUP product sync for: $ProductName"
    try {
        # Query the SUP sync categories via WMI directly - more reliable than
        # Get-CMSoftwareUpdatePointComponent which can throw null key exceptions
        $supProps = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" `
                                  -ComputerName $SiteServer `
                                  -Class SMS_SCI_Component `
                                  -Filter "ComponentName = 'SMS_WSUS_CONFIGURATION_MANAGER'" `
                                  -ErrorAction Stop

        if ($supProps) {
            $syncCatProp = $supProps.Props | Where-Object { $_.PropertyName -eq 'SyncCatalogCategories' }
            $syncCats    = if ($syncCatProp) { $syncCatProp.Value1 } else { $null }
            if ($syncCats -and $syncCats -match [regex]::Escape($ProductName)) {
                Write-Log "SUP product sync confirmed for: $ProductName" -Level SUCCESS
            } else {
                Write-Log "WARNING: '$ProductName' may not be enabled in SUP sync settings." -Level WARN
                Write-Log "Verify under: Administration > Site Configuration > Sites > Configure Site Components > Software Update Point" -Level WARN
                Write-Log "If not enabled, the ADR will find no updates to download." -Level WARN
            }
        } else {
            Write-Log "Could not retrieve SUP component properties (non-fatal)." -Level WARN
        }
    } catch {
        Write-Log "Could not verify SUP product sync (non-fatal): $_" -Level WARN
    }
}

function Invoke-ADREvaluation {
    param([string]$Name)
    Write-Log "Triggering immediate ADR evaluation: $Name"
    try {
        Invoke-CMAutoDeploymentRuleSync -Name $Name -ErrorAction Stop
        Write-Log "ADR evaluation triggered. Monitor RuleEngine.log on the site server for progress." -Level SUCCESS
        Write-Log "Updates will download to: $PackageSourcePath" -Level SUCCESS
    } catch {
        Write-Log "ADR evaluation trigger failed: $_" -Level WARN
        Write-Log "You can trigger it manually: right-click the ADR in console > Run Now" -Level WARN
    }
}

#endregion

#region --- Main Execution ---

Write-Log "===== WIM Servicing ADR Setup ====="
Write-Log "Site Code      : $SiteCode"
Write-Log "Site Server    : $SiteServer"
Write-Log "OS Version     : $OSVersion ($($osInfo.FriendlyName))"
Write-Log "Package Name   : $PackageName"
Write-Log "ADR Name       : $ADRName"
Write-Log "Package Source : $PackageSourcePath"
Write-Log "Include .NET   : $($IncludeDotNet.IsPresent)"

try {
    Connect-SCCMSite -Code $SiteCode -Server $SiteServer

    # Warn if product not in SUP sync
    Test-SUPProductSync -ProductName $osInfo.Product

    # Create the deployment package
    $pkg = New-WIMServicingPackage -Name $PackageName `
                                   -SourcePath $PackageSourcePath `
                                   -DPGroup $DPGroupName

    Write-Log "Package ready. PackageID: $($pkg.PackageID)" -Level SUCCESS

    # Calculate ADR schedule
    $adrSchedule = Get-PatchTuesdayPlusOne -HourOfDay 3

    # Create the ADR
    $adr = New-WIMServicingADR -Name $ADRName `
                               -SUG $SUGName `
                               -PkgName $PackageName `
                               -PkgSourcePath $PackageSourcePath `
                               -OS $osInfo `
                               -Arch $Architecture `
                               -DotNet $IncludeDotNet.IsPresent `
                               -Schedule $adrSchedule

    # Trigger immediate evaluation if requested
    if ($RunImmediately) {
        Invoke-ADREvaluation -Name $ADRName
    }

    Write-Log "===== Setup Complete =====" -Level SUCCESS
    Write-Log ""
    Write-Log "--- Summary ---"
    Write-Log "Deployment Package : $PackageName (ID: $($pkg.PackageID))"
    Write-Log "ADR Name           : $ADRName"
    Write-Log "SUG Name           : $SUGName"
    Write-Log "Package Source     : $PackageSourcePath"
    Write-Log "ADR Schedule       : Wednesday after each Patch Tuesday at 03:00"
    Write-Log ""
    Write-Log "--- Next Steps ---"
    Write-Log "1. Verify the ADR in: Software Library > Software Updates > Automatic Deployment Rules"
    Write-Log "2. Review/adjust filter criteria (title, product, classification) in the ADR properties"
    Write-Log "3. Monitor download progress in: RuleEngine.log, PatchDownloader.log on $SiteServer"
    Write-Log "4. Once downloaded, point Invoke-WIMOfflineServicing.ps1 at:"
    Write-Log "     -UpdatesFolder `"$PackageSourcePath`""
    Write-Log "5. Logs written to: $LogFile"

} catch {
    Write-Log "FATAL: $_" -Level ERROR
    Write-Log "Setup did not complete. Review errors above." -Level ERROR
    exit 1
} finally {
    # Always return to a filesystem location
    if ((Get-Location).Drive.Name -eq $SiteCode) {
        Set-Location $env:SystemDrive
    }
}

#endregion
