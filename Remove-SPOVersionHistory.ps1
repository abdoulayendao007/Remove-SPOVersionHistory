# ============================================================
# Remove-SPOVersionHistory.ps1
# SharePoint Online File Version History Cleanup Script
# Author  : Abdoulaye Ndao
# Version : 1.0.4
# License : MIT
# GitHub  : https://github.com/abdoulayendao007/Remove-SPOVersionHistory
# ============================================================
# DESCRIPTION :
# Automatically cleans up file version history across all
# SharePoint Online sites in your tenant.
#
# TWO AUTHENTICATION MODES in one script :
#
#   MODE 1 -- Interactive (laptop / local testing)
#   No certificate needed. Opens browser once to authenticate.
#   Usage : .\Remove-SPOVersionHistory.ps1 -UseInteractive
#
#   MODE 2 -- Certificate (server / unattended / production)
#   Runs fully unattended. No popup. Ideal for scheduled tasks.
#   Usage : .\Remove-SPOVersionHistory.ps1
#           (requires SP_CERT_PATH and SP_CERT_PASSWORD env vars)
#
# QUICK START :
#
#   Test on one site (laptop) :
#     $env:SP_TENANT    = "your-tenant.onmicrosoft.com"
#     $env:SP_CLIENT_ID = "your-client-id"
#     .\Remove-SPOVersionHistory.ps1 -UseInteractive -TestSite "AuditReports"
#
#   Full tenant scan -- progressive rollout recommended :
#     Step 1 : .\Remove-SPOVersionHistory.ps1 -TestSites @("Site1","Site2","Site3")
#     Step 2 : .\Remove-SPOVersionHistory.ps1 -TestSites @("Site1","Site2"..."Site20")
#     Step 3 : .\Remove-SPOVersionHistory.ps1
#
#   Resume after crash :
#     .\Remove-SPOVersionHistory.ps1 -StartSite "RH"
#
# PRODUCTION SAFETY (v1.0.4) :
#
#   BatchSize (default 20) :
#     Versions are deleted in small batches to prevent OutOfMemoryException.
#     Lower value = safer on low RAM servers, but slower.
#     Recommended values :
#       4 GB RAM  -> BatchSize = 10
#       6-8 GB    -> BatchSize = 20  (default)
#       16 GB+    -> BatchSize = 50
#
#   MaxVersionsPerFile (default 300) :
#     Files with more versions than this limit are skipped (SKIP_TOO_MANY).
#     This prevents loading thousands of version objects into memory at once.
#     Set to 0 to disable (not recommended on low RAM servers).
#     Recommended values :
#       4 GB RAM  -> MaxVersionsPerFile = 150
#       6-8 GB    -> MaxVersionsPerFile = 300  (default)
#       16 GB+    -> MaxVersionsPerFile = 500
#
#   MaxFilesPerSite (default 0 = no limit) :
#     By default, all files in a site are processed.
#     For servers with less than 8 GB RAM and sites with 10000+ files,
#     consider setting MaxFilesPerSite = 1000 and running multiple times.
#     Each run will process the next batch of files needing cleanup.
#
#   GCInterval (default 200) :
#     Forces garbage collection every N files to free accumulated memory.
#     Lower value = more frequent GC but higher CPU overhead.
#
#   -StartSite parameter :
#     Use to resume a run after a crash without reprocessing already done sites.
#     Example : .\Remove-SPOVersionHistory.ps1 -StartSite "RH"
#     Sites processed before "RH" (alphabetically) will be skipped.
#
# RETENTION POLICY :
# Normal sites   : Option A -- $VersionsNormal TOTAL (including current)
#                  Get-PnPFileVersion returns history only (not current).
#                  "10 total" = 9 history versions + current version.
#                  => history threshold = $VersionsNormal - 1
# Critical sites : Option B -- $VersionsCritical HISTORY (current not counted)
#                  => history threshold = $VersionsCritical
#
# NOTE : This script reduces storage space only.
#        It does NOT replace Microsoft Purview retention labels
#        or records management for regulatory compliance.
#
# REQUIREMENTS :
# - PowerShell 7.4+
# - PnP.PowerShell 3.x+
# - 8 GB RAM recommended for full tenant scans (4 GB minimum)
# - Entra ID App Registration with :
#
#   Interactive mode (-UseInteractive) :
#   - Redirect URI : https://login.microsoftonline.com/common/oauth2/nativeclient
#   - Delegated permission : SharePoint > AllSites.FullControl
#   - Account must have SharePoint Administrator or Global Administrator role
#
#   Certificate mode (default / server) :
#   - Application permission : Sites.FullControl.All (admin consent required)
#   - Certificate uploaded to the App Registration (.cer file)
#   - Environment variables SP_CERT_PATH and SP_CERT_PASSWORD set
# ============================================================

# ============================================================
# PARAMETERS
# ============================================================
param(
    [string]$Tenant          = $env:SP_TENANT,
    [string]$ClientId        = $env:SP_CLIENT_ID,
    [string]$TenantUrl       = $env:SP_TENANT_URL,
    [string]$AdminUrl        = $env:SP_ADMIN_URL,
    [string]$CertificatePath = $env:SP_CERT_PATH,
    [string]$CertificatePass = $env:SP_CERT_PASSWORD,
    [switch]$UseInteractive,
    [string]$TestSite        = "",
    [string[]]$TestSites     = @(),

    # Resume after crash -- start processing from this site keyword
    # All sites before this keyword (alphabetically) will be skipped
    # Usage : .\Remove-SPOVersionHistory.ps1 -StartSite "RH"
    [string]$StartSite       = ""
)

# ============================================================
# PARAMETER VALIDATION
# ============================================================

if ($TestSite -ne "" -and $TestSites.Count -gt 0) {
    Write-Host ""
    Write-Host "ERROR : Use -TestSite OR -TestSites, not both." -ForegroundColor Red
    Write-Host "  -TestSite  : .\Remove-SPOVersionHistory.ps1 -TestSite 'AuditReports'" -ForegroundColor Yellow
    Write-Host "  -TestSites : .\Remove-SPOVersionHistory.ps1 -TestSites @('Accounting','Legal')" -ForegroundColor Yellow
    exit 1
}

$missing = @()
if (-not $Tenant)   { $missing += "SP_TENANT" }
if (-not $ClientId) { $missing += "SP_CLIENT_ID" }

if (-not $UseInteractive) {
    if (-not $CertificatePath) { $missing += "SP_CERT_PATH" }
    if (-not $CertificatePass) { $missing += "SP_CERT_PASSWORD" }
}

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "ERROR : Missing required configuration values." -ForegroundColor Red
    foreach ($m in $missing) { Write-Host "        - $m" -ForegroundColor Yellow }
    Write-Host ""
    if ($UseInteractive) {
        Write-Host "Interactive mode example :" -ForegroundColor Cyan
        Write-Host "  `$env:SP_TENANT    = 'your-tenant.onmicrosoft.com'" -ForegroundColor Cyan
        Write-Host "  `$env:SP_CLIENT_ID = 'your-client-id'" -ForegroundColor Cyan
        Write-Host "  .\Remove-SPOVersionHistory.ps1 -UseInteractive" -ForegroundColor Cyan
    } else {
        Write-Host "Certificate mode example :" -ForegroundColor Cyan
        Write-Host "  `$env:SP_TENANT        = 'your-tenant.onmicrosoft.com'" -ForegroundColor Cyan
        Write-Host "  `$env:SP_CLIENT_ID     = 'your-client-id'" -ForegroundColor Cyan
        Write-Host "  `$env:SP_CERT_PATH     = 'C:\Certs\your-cert.pfx'" -ForegroundColor Cyan
        Write-Host "  `$env:SP_CERT_PASSWORD = 'your-cert-password'" -ForegroundColor Cyan
        Write-Host "  .\Remove-SPOVersionHistory.ps1" -ForegroundColor Cyan
    }
    Write-Host ""
    Write-Host "See README.md for full setup instructions." -ForegroundColor Cyan
    exit 1
}

$SecureCertPassword = $null
if (-not $UseInteractive) {
    $SecureCertPassword = ConvertTo-SecureString $CertificatePass -AsPlainText -Force
}

if (-not $TenantUrl -or -not $AdminUrl) {
    $tenantShort = $Tenant -replace '\.onmicrosoft\.com$', '' -replace '\.sharepoint\.com$', ''
    if (-not $TenantUrl) {
        $TenantUrl = "https://$tenantShort.sharepoint.com"
        Write-Host "INFO : SP_TENANT_URL derived as : $TenantUrl" -ForegroundColor DarkGray
    }
    if (-not $AdminUrl) {
        $AdminUrl = "https://$tenantShort-admin.sharepoint.com"
        Write-Host "INFO : SP_ADMIN_URL derived as : $AdminUrl" -ForegroundColor DarkGray
    }
}

# ============================================================
# CONFIGURATION
# ============================================================

$VersionsNormal   = 10
$VersionsCritical = 50

$ModeTest    = $true
$ModeRecycle = $true

$DaysInactiveMinimum = 30
$MaxRetries          = 3

$MaxFileSizeMB = if ($UseInteractive) { 800 } else { 0 }

# ============================================================
# v1.0.4 : Production safety settings
# See header comments for recommended values per RAM profile
# ============================================================

# Batch size for version deletion
# Prevents OutOfMemoryException on files with many versions
# Set to 0 to disable batching (not recommended)
$BatchSize = 20

# Max versions per file before skipping (SKIP_TOO_MANY)
# Prevents loading too many version objects into memory
# Set to 0 to disable (not recommended on low RAM servers)
$MaxVersionsPerFile = 300

# Max files to process per site (0 = no limit)
# For servers with < 8 GB RAM and sites with 10000+ files,
# consider setting to 1000 and running multiple times
$MaxFilesPerSite = 0

# Force garbage collection every N files
# Prevents memory accumulation on long runs
$GCInterval = 200

# Pause between batches in seconds
$PauseBetweenBatches = 2

$LogFolder  = "C:\Temp\SPOVersionCleanup"
$RunId      = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile    = "$LogFolder\Cleanup_$RunId.log"
$ReportCSV  = "$LogFolder\Report_$RunId.csv"
$ReportJSON = "$LogFolder\Summary_$RunId.json"
$RunMode    = if ($ModeTest) { "SIMULATION" } elseif ($ModeRecycle) { "PRODUCTION-RECYCLE" } else { "PRODUCTION-DELETE" }
$RunDate    = Get-Date -Format 'yyyy-MM-dd'

# ============================================================
# CRITICAL SITES
# ============================================================
$CriticalSites = @(
    "Accounting",
    "LegalAffairs",
    "PeopleOps",
    "RegulatoryDocs",
    "AuditReports"
)

$ExcludedLibraries = @(
    "Form Templates", "Site Assets", "Style Library", "Pages",
    "Site Pages", "App Files", "App Packages",
    "Preservation Hold Library", "Social", "Images", "MicroFeed",
    "Converted Forms", "Master Page Gallery", "Workflows", "wfpub",
    "Bibliotheque de styles", "Bibliothèque de styles",
    "Modèles de formulaire", "Modeles de formulaire",
    "Pièces jointes", "Pieces jointes",
    "Documents de la collection de sites", "AppPages", "Teams Wiki Data"
)

# ============================================================
# INITIALIZATION
# ============================================================

if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder | Out-Null }

$ReportData        = [System.Collections.Generic.List[PSObject]]::new()
$SitesSummary      = [System.Collections.Generic.List[PSObject]]::new()
$SitesAccessDenied = [System.Collections.Generic.List[string]]::new()

$TotalFilesCleaned      = 0
$TotalFilesRecent       = 0
$TotalFilesUnderLimit   = 0
$TotalFilesAccessDenied = 0
$TotalFilesSkippedLarge = 0
$TotalFilesSkippedTooMany = 0
$TotalVersionsRemoved   = 0
$TotalSpaceMB           = 0
$TotalErrors            = 0
$ConnectionCache        = @{}
$GlobalFileCounter      = 0

# ============================================================
# FUNCTIONS
# ============================================================

function Write-Log {
    param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Type] $Message"
    switch ($Type) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        "TEST"    { Write-Host $line -ForegroundColor Cyan }
        "SKIP"    { Write-Host $line -ForegroundColor DarkGray }
        default   { Write-Host $line }
    }
    Add-Content -Path $LogFile -Value $line
}

function Install-PnPIfNeeded {
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Log "PnP.PowerShell not found. Installing..." "WARN"
        try {
            Install-Module -Name "PnP.PowerShell" -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Log "PnP.PowerShell installed" "SUCCESS"
        } catch {
            Write-Log "Installation failed : $_" "ERROR"
            exit 1
        }
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Log "PnP.PowerShell loaded" "SUCCESS"
}

function Get-SPConnection {
    param([string]$Url, [bool]$ForceReconnect = $false)

    if ($ForceReconnect -and $ConnectionCache.ContainsKey($Url)) {
        $ConnectionCache.Remove($Url)
        Write-Log "Forced reconnection : $Url" "WARN"
    }
    if ($ConnectionCache.ContainsKey($Url)) { return $ConnectionCache[$Url] }

    try {
        if ($UseInteractive) {
            $conn = Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
        } else {
            $conn = Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant `
                -CertificatePath $CertificatePath -CertificatePassword $SecureCertPassword `
                -ReturnConnection -ErrorAction Stop
        }
        $ConnectionCache[$Url] = $conn
        Write-Log "Connection established : $Url" "SUCCESS"
        return $conn
    } catch {
        Write-Log "Connection FAILED $Url : $_" "ERROR"
        return $null
    }
}

function Invoke-PnPWithRetry {
    param([scriptblock]$Command, [string]$SiteUrl, [ref]$Connection)

    try {
        return & $Command $Connection.Value
    } catch {
        $errMsg = $_.ToString()

        if ($errMsg -match "HttpClient\.Timeout|request was canceled|TaskCanceledException|timeout of 100 seconds") {
            $retryDelay = 10
            for ($i = 1; $i -le $MaxRetries; $i++) {
                Write-Log "PnP 100s timeout -- retry $i/$MaxRetries (waiting ${retryDelay}s) : $SiteUrl" "WARN"
                Start-Sleep -Seconds $retryDelay
                $retryDelay += 10
                try {
                    return & $Command $Connection.Value
                } catch {
                    $errRetry = $_.ToString()
                    if ($errRetry -notmatch "HttpClient\.Timeout|request was canceled|TaskCanceledException|timeout of 100 seconds") { throw $_ }
                    if ($i -eq $MaxRetries) {
                        Write-Log "Timeout persists after $MaxRetries retries -- skipping" "ERROR"
                        return "TIMEOUT"
                    }
                }
            }
        }

        if ($errMsg -match "401|token.*expired|expired.*token|invalid.*token|authentication.*failed|AADSTS") {
            Write-Log "Token expired (401) -- reconnecting : $SiteUrl" "WARN"
            $newConn = Get-SPConnection -Url $SiteUrl -ForceReconnect $true
            if ($null -ne $newConn) {
                $Connection.Value = $newConn
                try { return & $Command $Connection.Value } catch { return "AUTH_ERROR" }
            }
            return "AUTH_ERROR"
        }

        if ($errMsg -match "403|Unauthorized|AccessDenied|unauthorized operation|Access denied") {
            Write-Log "Access denied (403) -- skipping : $SiteUrl" "WARN"
            return "ACCESS_DENIED"
        }

        throw $_
    }
}

function Test-IsErrorStatus {
    param($Value)
    return ($Value -is [string] -and $Value -in @("ACCESS_DENIED", "TIMEOUT", "AUTH_ERROR"))
}

function Resolve-VersionIdentity {
    param($Version)
    if ($null -ne $Version.Id -and $Version.Id -gt 0) { return $Version.Id }
    if ($Version.VersionLabel -match '^Version\s') { return $Version.VersionLabel }
    if ($null -ne $Version.VersionLabel -and $Version.VersionLabel -ne "") { return "Version $($Version.VersionLabel)" }
    Write-Log "Unable to resolve version identity" "WARN"
    return $null
}

function Get-HistoryThreshold {
    param([bool]$IsCritical)
    if ($IsCritical) { return $VersionsCritical } else { return ($VersionsNormal - 1) }
}

function Invoke-VersionCleanup {
    param(
        [string]$FileUrl, [int]$HistoryThreshold, [datetime]$LastModified,
        [string]$SiteUrl, [string]$LibraryName, [string]$SiteType, [ref]$Connection
    )

    if ($LastModified -gt (Get-Date).AddDays(-$DaysInactiveMinimum)) { return "RECENT" }

    try {
        $rawResult = Invoke-PnPWithRetry -SiteUrl $SiteUrl -Connection $Connection -Command {
            param($conn)
            Get-PnPFileVersion -Url $FileUrl -Connection $conn -ErrorAction Stop
        }

        if (Test-IsErrorStatus $rawResult) { Write-Log "      Skipping ($rawResult) : $FileUrl" "WARN"; return $rawResult }
        if ($null -eq $rawResult) { return "UNDER_LIMIT" }

        $Versions = @($rawResult)
        if ($Versions.Count -le $HistoryThreshold) { return "UNDER_LIMIT" }

        # v1.0.4 : Hard limit on versions per file to prevent OOM
        if ($MaxVersionsPerFile -gt 0 -and $Versions.Count -gt $MaxVersionsPerFile) {
            Write-Log "      SKIP_TOO_MANY : $($Versions.Count) versions > limit $MaxVersionsPerFile : $FileUrl" "WARN"
            Write-Log "      TIP : Increase MaxVersionsPerFile or process this file manually" "WARN"
            return "SKIP_TOO_MANY"
        }

        if ($MaxFileSizeMB -gt 0) {
            $SizeMB = [math]::Round((($Versions | ForEach-Object { if ($null -ne $_.Size) { $_.Size } else { 0 } } | Measure-Object -Sum).Sum / 1MB), 2)
            if ($SizeMB -gt $MaxFileSizeMB) {
                Write-Log "      SKIP_LARGE : $SizeMB MB > limit $MaxFileSizeMB MB : $FileUrl" "WARN"
                return "SKIP_LARGE"
            }
        }

        $VersionsToRemove = @($Versions | Sort-Object Created -Descending | Select-Object -Skip $HistoryThreshold)
        $RemoveCount      = $VersionsToRemove.Count
        $SpaceMB          = [math]::Round((($VersionsToRemove | ForEach-Object { if ($null -ne $_.Size) { $_.Size } else { 0 } } | Measure-Object -Sum).Sum / 1MB), 2)

        Write-Log "      File : $FileUrl" "INFO"
        Write-Log "      History : $($Versions.Count) | To remove : $RemoveCount | Space : $SpaceMB MB" "INFO"

        if (-not $ModeTest) {
            if ($BatchSize -gt 0) {
                $removedCount = 0
                $batchNum     = 0
                for ($b = 0; $b -lt $VersionsToRemove.Count; $b += $BatchSize) {
                    $batchNum++
                    $end   = [math]::Min($b + $BatchSize - 1, $VersionsToRemove.Count - 1)
                    $batch = $VersionsToRemove[$b..$end]

                    foreach ($Version in $batch) {
                        try {
                            $Identity = Resolve-VersionIdentity -Version $Version
                            if ($null -eq $Identity) { Write-Log "      SKIP version -- identity unresolved" "WARN"; continue }
                            $r = Invoke-PnPWithRetry -SiteUrl $SiteUrl -Connection $Connection -Command {
                                param($conn)
                                if ($ModeRecycle) {
                                    Remove-PnPFileVersion -Url $FileUrl -Identity $Identity -Recycle -Force -Connection $conn -ErrorAction Stop
                                } else {
                                    Remove-PnPFileVersion -Url $FileUrl -Identity $Identity -Force -Connection $conn -ErrorAction Stop
                                }
                            }
                            if (-not (Test-IsErrorStatus $r)) { $removedCount++ }
                        } catch { Write-Log "      Error removing version : $_" "ERROR" }
                    }

                    if ($b + $BatchSize -lt $VersionsToRemove.Count) {
                        Start-Sleep -Seconds $PauseBetweenBatches
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                    }
                }
            } else {
                foreach ($Version in $VersionsToRemove) {
                    try {
                        $Identity = Resolve-VersionIdentity -Version $Version
                        if ($null -eq $Identity) { Write-Log "      SKIP version -- identity unresolved" "WARN"; continue }
                        $r = Invoke-PnPWithRetry -SiteUrl $SiteUrl -Connection $Connection -Command {
                            param($conn)
                            if ($ModeRecycle) {
                                Remove-PnPFileVersion -Url $FileUrl -Identity $Identity -Recycle -Force -Connection $conn -ErrorAction Stop
                            } else {
                                Remove-PnPFileVersion -Url $FileUrl -Identity $Identity -Force -Connection $conn -ErrorAction Stop
                            }
                        }
                    } catch { Write-Log "      Error removing version : $_" "ERROR" }
                }
            }
            Write-Log "      $(if ($ModeRecycle) { 'RECYCLED' } else { 'DELETED' }) $RemoveCount versions" "SUCCESS"
        } else {
            Write-Log "      [SIMULATION] $RemoveCount versions would be removed" "TEST"
        }

        return [PSCustomObject]@{
            Site            = $SiteUrl
            Library         = $LibraryName
            File            = $FileUrl
            SiteType        = $SiteType
            HistoryVersions = $Versions.Count
            ThresholdKept   = $HistoryThreshold
            Removed         = $RemoveCount
            SpaceMB         = $SpaceMB
            LastModified    = $LastModified
            Action          = if ($ModeTest) { "SIMULATION" } elseif ($ModeRecycle) { "RECYCLED" } else { "DELETED" }
            RunId           = $RunId
            RunMode         = $RunMode
            RunDate         = $RunDate
        }
    } catch {
        Write-Log "      ERROR on file $FileUrl : $_" "ERROR"
        return $null
    }
}

# ============================================================
# SCRIPT START
# ============================================================
$authMode = if ($UseInteractive) { "INTERACTIVE (browser)" } else { "CERTIFICATE (unattended)" }

Write-Log "=========================================="
Write-Log "SPO VERSION HISTORY CLEANUP v1.0.4"
Write-Log "=========================================="
Write-Log "Auth mode           : $authMode"
Write-Log "MaxFileSizeMB       : $(if ($MaxFileSizeMB -eq 0) { 'No limit' } else { "$MaxFileSizeMB MB" })"
Write-Log "BatchSize           : $(if ($BatchSize -eq 0) { 'Disabled' } else { "$BatchSize versions/batch" })"
Write-Log "MaxVersionsPerFile  : $(if ($MaxVersionsPerFile -eq 0) { 'No limit' } else { $MaxVersionsPerFile })"
Write-Log "MaxFilesPerSite     : $(if ($MaxFilesPerSite -eq 0) { 'No limit' } else { $MaxFilesPerSite })"
Write-Log "GCInterval          : every $GCInterval files"

# RAM check
$totalRAMMB = [math]::Round((Get-CimInstance Win32_OperatingSystem).TotalVisibleMemorySize / 1KB, 0)
$freeRAMMB  = [math]::Round((Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1KB, 0)
Write-Log "RAM                 : $freeRAMMB MB free / $totalRAMMB MB total"

if ($totalRAMMB -lt 8192) {
    Write-Log "==========================================" "WARN"
    Write-Log "WARNING : Server RAM is below 8 GB ($totalRAMMB MB total)." "WARN"
    Write-Log "          Current settings (BatchSize=$BatchSize, MaxVersionsPerFile=$MaxVersionsPerFile)" "WARN"
    Write-Log "          are safe for this RAM profile." "WARN"
    Write-Log "          Consider using -TestSites for progressive rollout." "WARN"
    Write-Log "          See README.md for RAM tuning guide." "WARN"
    Write-Log "==========================================" "WARN"
}

if ($StartSite -ne "") {
    Write-Log "Resume mode         : starting from site matching '$StartSite'" "WARN"
}

if ($UseInteractive) {
    Write-Log "==========================================" "WARN"
    Write-Log "WARNING : Interactive mode is for testing only." "WARN"
    Write-Log "          Use certificate mode for full tenant scans." "WARN"
    Write-Log "==========================================" "WARN"
}

Write-Log "Run ID          : $RunId"
Write-Log "Tenant          : $Tenant"
Write-Log "Mode            : $(if ($ModeTest) { 'SIMULATION (no deletions)' } else { 'PRODUCTION' })"
Write-Log "Deletion mode   : $(if ($ModeRecycle) { 'RECYCLE BIN (~93 days)' } else { 'PERMANENT' })"
Write-Log "Normal sites    : Option A -- $VersionsNormal total (9 history + current)"
Write-Log "Critical sites  : Option B -- $VersionsCritical history + current"
Write-Log "Recent files    : skipped if modified within last $DaysInactiveMinimum days"
Write-Log "Log file        : $LogFile"
Write-Log "Report CSV      : $ReportCSV"
Write-Log "=========================================="

if (-not $ModeTest) {
    Write-Host "WARNING -- PRODUCTION MODE ACTIVE" -ForegroundColor Red
    Write-Host "Deletion : $(if ($ModeRecycle) { 'RECYCLE BIN' } else { 'PERMANENT' })" -ForegroundColor Red
    $Confirm = Read-Host "Type CONFIRM to proceed"
    if ($Confirm -ne "CONFIRM") { Write-Log "Script cancelled by user" "WARN"; exit 0 }
}

Install-PnPIfNeeded

# ============================================================
# SITE LIST BUILDING
# ============================================================
if ($TestSite -ne "" -or $TestSites.Count -gt 0) {
    Write-Log "TEST MODE : Building site list directly (skipping Get-PnPTenantSite)" "WARN"
    $TestSiteUrls = @()
    if ($TestSite -ne "") { $TestSiteUrls += "$TenantUrl/sites/$($TestSite.TrimStart('/'))" }
    foreach ($t in $TestSites) { $TestSiteUrls += "$TenantUrl/sites/$($t.TrimStart('/'))" }
    $Sites = $TestSiteUrls | ForEach-Object { [PSCustomObject]@{ Url = $_; Status = "Active" } }
    Write-Log "TEST MODE : $($Sites.Count) site(s) to process" "WARN"
    Write-Log "TIP : For full tenant scan, remove -TestSite/-TestSites and use certificate mode" "WARN"
} else {
    Write-Log "Connecting to admin : $AdminUrl"
    $ConnAdmin = Get-SPConnection -Url $AdminUrl
    if ($null -eq $ConnAdmin) { Write-Log "Admin connection failed -- aborting" "ERROR"; exit 1 }
    Write-Log "Retrieving sites..."
    try {
        $ConnAdminRef = [ref]$ConnAdmin
        $rawSites = Invoke-PnPWithRetry -SiteUrl $AdminUrl -Connection $ConnAdminRef -Command {
            param($conn)
            Get-PnPTenantSite -IncludeOneDriveSites:$false -Connection $conn |
            Where-Object { $_.Status -eq "Active" } | Sort-Object Url
        }
        if (Test-IsErrorStatus $rawSites) { Write-Log "Failed to retrieve sites ($rawSites)" "ERROR"; exit 1 }
        $Sites = @($rawSites | Where-Object { $null -ne $_ })
        Write-Log "Sites found : $($Sites.Count)" "SUCCESS"
    } catch { Write-Log "Error retrieving sites : $_" "ERROR"; exit 1 }
}

# ============================================================
# v1.0.4 : Resume mode
# ============================================================
$ResumeMode = ($StartSite -ne "")
if ($ResumeMode) {
    Write-Log "RESUME MODE : Will skip sites until matching '$StartSite'" "WARN"
}

# ============================================================
# MAIN LOOP
# ============================================================
foreach ($Site in $Sites) {

    # Resume mode : skip until StartSite found
    if ($ResumeMode) {
        if ($Site.Url -like "*$StartSite*") {
            $ResumeMode = $false
            Write-Log "RESUME MODE : Found -- resuming from $($Site.Url)" "WARN"
        } else {
            Write-Log "RESUME : Skipping $($Site.Url)" "SKIP"
            continue
        }
    }

    Write-Log ""
    Write-Log "=========================================="
    Write-Log "SITE : $($Site.Url)"

    $u = $Site.Url.TrimEnd('/')
    if ($u -match '-my\.sharepoint\.com$') { Write-Log "SKIP OneDrive root : $u" "SKIP"; continue }
    if ($u -match '/search$' -or $u -match '/portals/') { Write-Log "SKIP system site : $u" "SKIP"; continue }

    $IsCritical = $false
    foreach ($keyword in $CriticalSites) { if ($Site.Url -like "*$keyword*") { $IsCritical = $true; break } }

    $Threshold = Get-HistoryThreshold -IsCritical $IsCritical
    Write-Log "Type : $(if ($IsCritical) { 'CRITICAL' } else { 'Normal' }) | Threshold : $Threshold history versions"

    $SiteCleaned = 0; $SiteRecent = 0; $SiteUnderLimit = 0
    $SiteAccessDenied = 0; $SiteVersions = 0; $SiteSpaceMB = 0
    $SiteFileCounter = 0

    $ConnSite = Get-SPConnection -Url $Site.Url
    if ($null -eq $ConnSite) { $TotalErrors++; continue }
    $ConnSiteRef = [ref]$ConnSite

    try {
        $rawLibraries = Invoke-PnPWithRetry -SiteUrl $Site.Url -Connection $ConnSiteRef -Command {
            param($conn) Get-PnPList -Connection $conn -ErrorAction Stop
        }
        if ($rawLibraries -eq "ACCESS_DENIED") {
            $SitesAccessDenied.Add($Site.Url); $TotalFilesAccessDenied++; continue
        }
        if (Test-IsErrorStatus $rawLibraries -or $null -eq $rawLibraries) { $TotalErrors++; continue }
        $Libraries = @($rawLibraries | Where-Object {
            $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -notin $ExcludedLibraries
        })
        Write-Log "Libraries : $($Libraries.Count)"
    } catch {
        $errLib = $_.ToString()
        if ($errLib -match "403|Unauthorized|AccessDenied") {
            $SitesAccessDenied.Add($Site.Url); $TotalFilesAccessDenied++
        } else { $TotalErrors++ }
        continue
    }

    foreach ($Library in $Libraries) {
        Write-Log "  Library : $($Library.Title)"
        try {
            $rawItems = Invoke-PnPWithRetry -SiteUrl $Site.Url -Connection $ConnSiteRef -Command {
                param($conn)
                Get-PnPListItem -List $Library -PageSize 200 -Fields "FileRef","Modified","FSObjType" -Connection $conn -ErrorAction Stop
            }
            if ($rawItems -eq "ACCESS_DENIED") { $SiteAccessDenied++; $TotalFilesAccessDenied++; continue }
            if (Test-IsErrorStatus $rawItems -or $null -eq $rawItems) { $TotalErrors++; continue }

            $Items = @($rawItems | Where-Object { $_.FieldValues.FSObjType -eq 0 })
            Write-Log "    Files : $($Items.Count)"

            foreach ($Item in $Items) {

                # v1.0.4 : MaxFilesPerSite limit
                if ($MaxFilesPerSite -gt 0 -and $SiteFileCounter -ge $MaxFilesPerSite) {
                    Write-Log "    MaxFilesPerSite ($MaxFilesPerSite) reached -- skipping remaining files" "WARN"
                    Write-Log "    TIP : Run again to process remaining files in this site" "WARN"
                    break
                }

                try {
                    $CurrentSiteType = if ($IsCritical) { "Critical" } else { "Normal" }
                    $Result = Invoke-VersionCleanup `
                        -FileUrl $Item.FieldValues.FileRef `
                        -HistoryThreshold $Threshold `
                        -LastModified ([datetime]$Item.FieldValues.Modified) `
                        -SiteUrl $Site.Url `
                        -LibraryName $Library.Title `
                        -SiteType $CurrentSiteType `
                        -Connection $ConnSiteRef

                    $SiteFileCounter++
                    $GlobalFileCounter++

                    switch ($Result) {
                        "RECENT"         { $SiteRecent++; $TotalFilesRecent++ }
                        "UNDER_LIMIT"    { $SiteUnderLimit++; $TotalFilesUnderLimit++ }
                        "ACCESS_DENIED"  { $SiteAccessDenied++; $TotalFilesAccessDenied++ }
                        "SKIP_LARGE"     { $TotalFilesSkippedLarge++ }
                        "SKIP_TOO_MANY"  { $TotalFilesSkippedTooMany++ }
                        "TIMEOUT"        { $TotalErrors++ }
                        "AUTH_ERROR"     { $TotalErrors++ }
                        $null            { $TotalErrors++ }
                        default {
                            $ReportData.Add($Result)
                            $SiteCleaned++; $SiteVersions += $Result.Removed; $SiteSpaceMB += $Result.SpaceMB
                            $TotalFilesCleaned++; $TotalVersionsRemoved += $Result.Removed; $TotalSpaceMB += $Result.SpaceMB
                        }
                    }

                    # v1.0.4 : Periodic GC
                    if ($GCInterval -gt 0 -and $GlobalFileCounter % $GCInterval -eq 0) {
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                    }

                    # v1.0.4 : Flush CSV every 100 files to prevent data loss on crash
                    if ($GlobalFileCounter % 100 -eq 0 -and $ReportData.Count -gt 0) {
                        try { $ReportData | Export-Csv -Path $ReportCSV -NoTypeInformation -Encoding UTF8 -Force } catch { }
                    }

                } catch { Write-Log "    ERROR on file : $_" "ERROR"; $TotalErrors++ }
            }
        } catch { Write-Log "  ERROR on library $($Library.Title) : $_" "ERROR"; $TotalErrors++ }
    }

    $SitesSummary.Add([PSCustomObject]@{
        Site = $Site.Url; Type = if ($IsCritical) { "Critical" } else { "Normal" }
        Cleaned = $SiteCleaned; Recent = $SiteRecent; UnderLimit = $SiteUnderLimit
        AccessDenied = $SiteAccessDenied; Versions = $SiteVersions
        SpaceMB = [math]::Round($SiteSpaceMB, 2); SpaceGB = [math]::Round($SiteSpaceMB / 1024, 2)
    })
    Write-Log "  SITE SUMMARY : Cleaned=$SiteCleaned | Recent=$SiteRecent | UnderLimit=$SiteUnderLimit | AccessDenied=$SiteAccessDenied | Versions=$SiteVersions | $([math]::Round($SiteSpaceMB,2)) MB" "SUCCESS"

    # v1.0.4 : GC after each site
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Export CSV
try {
    $ReportData | Export-Csv -Path $ReportCSV -NoTypeInformation -Encoding UTF8
    Write-Log "CSV report exported : $ReportCSV" "SUCCESS"
} catch { Write-Log "Error exporting CSV : $_" "ERROR" }

# Export JSON Summary
try {
    $JsonSummary = [PSCustomObject]@{
        RunId = $RunId; RunMode = $RunMode; RunDate = $RunDate; AuthMode = $authMode; Tenant = $Tenant
        TotalSitesScanned = $SitesSummary.Count
        TotalSitesCleaned = ($SitesSummary | Where-Object { $_.Cleaned -gt 0 }).Count
        TotalFiles = $TotalFilesCleaned; TotalFilesRecent = $TotalFilesRecent
        TotalFilesUnderLimit = $TotalFilesUnderLimit; TotalFilesAccessDenied = $TotalFilesAccessDenied
        TotalFilesSkippedLarge = $TotalFilesSkippedLarge
        TotalFilesSkippedTooMany = $TotalFilesSkippedTooMany
        TotalVersions = $TotalVersionsRemoved
        TotalSpaceMB = [math]::Round($TotalSpaceMB, 2); TotalSpaceGB = [math]::Round($TotalSpaceMB / 1024, 2)
        TotalErrors = $TotalErrors; RetentionNormal = $VersionsNormal; RetentionCritical = $VersionsCritical
        MaxFileSizeMB = $MaxFileSizeMB; BatchSize = $BatchSize
        MaxVersionsPerFile = $MaxVersionsPerFile; MaxFilesPerSite = $MaxFilesPerSite
        TopSites = ($SitesSummary | Sort-Object SpaceGB -Descending | Select-Object -First 10 |
            ForEach-Object { [PSCustomObject]@{ Site=$_.Site; Type=$_.Type; SpaceGB=$_.SpaceGB; Files=$_.Cleaned; Versions=$_.Versions } })
        AccessDeniedSites = $SitesAccessDenied; CSVReport = $ReportCSV; LogFile = $LogFile
    }
    $JsonSummary | ConvertTo-Json -Depth 5 | Out-File -FilePath $ReportJSON -Encoding UTF8
    Write-Log "JSON summary exported : $ReportJSON" "SUCCESS"
} catch { Write-Log "Error exporting JSON summary : $_" "ERROR" }

# ============================================================
# FINAL SUMMARY
# ============================================================
$TotalGB       = [math]::Round($TotalSpaceMB / 1024, 2)
$LabelVersions = if ($ModeTest) { "Versions to remove (simulation)" } else { "Versions removed" }
$LabelSpace    = if ($ModeTest) { "Recoverable space (simulation)" } else { "Space freed" }

Write-Log ""
Write-Log "=========================================="
Write-Log "FINAL SUMMARY"
Write-Log "=========================================="
Write-Log "Auth mode               : $authMode"
Write-Log "Mode                    : $(if ($ModeTest) { 'SIMULATION' } else { 'PRODUCTION' })"
Write-Log "Deletion mode           : $(if ($ModeRecycle) { 'RECYCLE BIN (~93 days)' } else { 'PERMANENT' })"
Write-Log "Retention policy        : Normal=$($VersionsNormal) total | Critical=$($VersionsCritical) history"
Write-Log "BatchSize               : $(if ($BatchSize -eq 0) { 'Disabled' } else { "$BatchSize versions/batch" })"
Write-Log "MaxVersionsPerFile      : $(if ($MaxVersionsPerFile -eq 0) { 'No limit' } else { $MaxVersionsPerFile })"
Write-Log "MaxFilesPerSite         : $(if ($MaxFilesPerSite -eq 0) { 'No limit' } else { $MaxFilesPerSite })"
Write-Log "Files cleaned           : $TotalFilesCleaned"
Write-Log "Files skipped (recent)  : $TotalFilesRecent"
Write-Log "Files under limit       : $TotalFilesUnderLimit"
Write-Log "Files access denied     : $TotalFilesAccessDenied"
Write-Log "Files skipped (large)   : $TotalFilesSkippedLarge"
Write-Log "Files skipped (too many): $TotalFilesSkippedTooMany"
if ($TotalFilesSkippedTooMany -gt 0) {
    Write-Log "TIP : $TotalFilesSkippedTooMany file(s) had more than $MaxVersionsPerFile versions." "WARN"
    Write-Log "      Increase MaxVersionsPerFile or process these files manually." "WARN"
}
Write-Log "${LabelVersions}        : $TotalVersionsRemoved"
Write-Log "${LabelSpace}           : $TotalSpaceMB MB ($TotalGB GB)"
Write-Log "Errors                  : $TotalErrors"
Write-Log ""

if ($SitesAccessDenied.Count -gt 0) {
    Write-Log "ACCESS DENIED SITES ($($SitesAccessDenied.Count)) :" "WARN"
    foreach ($s in $SitesAccessDenied) { Write-Log "  - $s" "WARN" }
    Write-Log ""
}

Write-Log "TOP 10 SITES (by recoverable space) :"
foreach ($r in ($SitesSummary | Sort-Object SpaceGB -Descending | Select-Object -First 10)) {
    Write-Log "  [$($r.Type)] $($r.Site)"
    Write-Log "    Cleaned=$($r.Cleaned) | Recent=$($r.Recent) | UnderLimit=$($r.UnderLimit) | $($r.SpaceGB) GB"
}
Write-Log "=========================================="
Write-Log "Log    : $LogFile"
Write-Log "Report : $ReportCSV"
Write-Log "=========================================="

if ($ModeTest) {
    Write-Log ""
    Write-Log "SIMULATION COMPLETE -- No deletions performed" "WARN"
    Write-Log "Next steps :" "WARN"
    Write-Log "  1. Review CSV : $ReportCSV" "WARN"
    Write-Log "  2. Validate with stakeholders" "WARN"
    Write-Log "  3. Production run 1 : ModeTest=false + ModeRecycle=true" "WARN"
    Write-Log "  4. Wait 2-3 weeks -- verify no issues" "WARN"
    Write-Log "  5. Production run 2 : ModeRecycle=false (if approved)" "WARN"
    Write-Log ""
    Write-Log "REMINDER : This script reduces storage space only." "WARN"
    Write-Log "           It does NOT replace Purview retention labels." "WARN"
}
