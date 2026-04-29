# ============================================================
# Remove-SPOVersionHistory.ps1
# SharePoint Online File Version History Cleanup Script
# Author  : Abdoulaye Ndao
# Version : 1.0.3
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
#   Full tenant scan (server with certificate) :
#     $env:SP_TENANT        = "your-tenant.onmicrosoft.com"
#     $env:SP_CLIENT_ID     = "your-client-id"
#     $env:SP_CERT_PATH     = "C:\Certs\your-cert.pfx"
#     $env:SP_CERT_PASSWORD = "your-cert-password"
#     .\Remove-SPOVersionHistory.ps1
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
    [string[]]$TestSites     = @()
)

# ============================================================
# PARAMETER VALIDATION
# ============================================================

# Mutual exclusion : -TestSite OR -TestSites, not both
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

# Automatic based on auth mode
# Interactive : 800 MB limit (avoids PnP 100s timeouts on large files)
# Certificate : 0 = no limit (stable connection)
$MaxFileSizeMB = if ($UseInteractive) { 800 } else { 0 }

$LogFolder  = "C:\Temp\SPOVersionCleanup"
$RunId      = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile    = "$LogFolder\Cleanup_$RunId.log"
$ReportCSV  = "$LogFolder\Report_$RunId.csv"
$ReportJSON = "$LogFolder\Summary_$RunId.json"
$RunMode    = if ($ModeTest) { "SIMULATION" } elseif ($ModeRecycle) { "PRODUCTION-RECYCLE" } else { "PRODUCTION-DELETE" }
$RunDate    = Get-Date -Format 'yyyy-MM-dd'

# ============================================================
# CRITICAL SITES
# IMPORTANT : This is an EXAMPLE list.
# Each organization MUST define its own based on
# business and regulatory requirements.
# Review with your compliance team before production.
# ============================================================
$CriticalSites = @(
    "Accounting",     # Example : Accounting, financial records
    "LegalAffairs",   # Example : Contracts, legal documents
    "PeopleOps",      # Example : Employee records, HR
    "RegulatoryDocs", # Example : Regulatory compliance
    "AuditReports"    # Example : Audit records
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
$TotalVersionsRemoved   = 0
$TotalSpaceMB           = 0
$TotalErrors            = 0
$ConnectionCache        = @{}

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

        # PnP 100s HTTP timeout -- progressive retry with backoff
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

        if ($MaxFileSizeMB -gt 0) {
            $SizeMB = [math]::Round((($Versions | ForEach-Object { if ($null -ne $_.Size) { $_.Size } else { 0 } } | Measure-Object -Sum).Sum / 1MB), 2)
            if ($SizeMB -gt $MaxFileSizeMB) {
                Write-Log "      SKIP_LARGE : $SizeMB MB > limit $MaxFileSizeMB MB : $FileUrl" "WARN"
                Write-Log "      TIP : Set MaxFileSizeMB=0 with certificate mode to process this file" "WARN"
                return "SKIP_LARGE"
            }
        }

        $VersionsToRemove = @($Versions | Sort-Object Created -Descending | Select-Object -Skip $HistoryThreshold)
        $RemoveCount      = $VersionsToRemove.Count
        $SpaceMB          = [math]::Round((($VersionsToRemove | ForEach-Object { if ($null -ne $_.Size) { $_.Size } else { 0 } } | Measure-Object -Sum).Sum / 1MB), 2)

        Write-Log "      File : $FileUrl" "INFO"
        Write-Log "      History : $($Versions.Count) | To remove : $RemoveCount | Space : $SpaceMB MB" "INFO"

        if (-not $ModeTest) {
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
                    if (Test-IsErrorStatus $r) { Write-Log "      Version removal skipped ($r)" "WARN" }
                } catch { Write-Log "      Error removing version : $_" "ERROR" }
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
Write-Log "SPO VERSION HISTORY CLEANUP v1.0.3"
Write-Log "=========================================="
Write-Log "Auth mode       : $authMode"
Write-Log "MaxFileSizeMB   : $(if ($MaxFileSizeMB -eq 0) { 'No limit' } else { "$MaxFileSizeMB MB" })"
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
# If -TestSite or -TestSites : connect directly without
# calling Get-PnPTenantSite (avoids 403 in interactive mode)
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
# MAIN LOOP
# ============================================================
foreach ($Site in $Sites) {
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

                    switch ($Result) {
                        "RECENT"        { $SiteRecent++; $TotalFilesRecent++ }
                        "UNDER_LIMIT"   { $SiteUnderLimit++; $TotalFilesUnderLimit++ }
                        "ACCESS_DENIED" { $SiteAccessDenied++; $TotalFilesAccessDenied++ }
                        "SKIP_LARGE"    { $TotalFilesSkippedLarge++ }
                        "TIMEOUT"       { $TotalErrors++ }
                        "AUTH_ERROR"    { $TotalErrors++ }
                        $null           { $TotalErrors++ }
                        default {
                            $ReportData.Add($Result)
                            $SiteCleaned++; $SiteVersions += $Result.Removed; $SiteSpaceMB += $Result.SpaceMB
                            $TotalFilesCleaned++; $TotalVersionsRemoved += $Result.Removed; $TotalSpaceMB += $Result.SpaceMB
                        }
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
        TotalFilesSkippedLarge = $TotalFilesSkippedLarge; TotalVersions = $TotalVersionsRemoved
        TotalSpaceMB = [math]::Round($TotalSpaceMB, 2); TotalSpaceGB = [math]::Round($TotalSpaceMB / 1024, 2)
        TotalErrors = $TotalErrors; RetentionNormal = $VersionsNormal; RetentionCritical = $VersionsCritical
        MaxFileSizeMB = $MaxFileSizeMB
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
Write-Log "Files cleaned           : $TotalFilesCleaned"
Write-Log "Files skipped (recent)  : $TotalFilesRecent"
Write-Log "Files under limit       : $TotalFilesUnderLimit"
Write-Log "Files access denied     : $TotalFilesAccessDenied"
Write-Log "Files skipped (large)   : $TotalFilesSkippedLarge"
if ($TotalFilesSkippedLarge -gt 0) { Write-Log "TIP : Set MaxFileSizeMB=0 with certificate mode to process skipped files" "WARN" }
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
