# ============================================================
# config.example.ps1
# Example configuration for Remove-SPOVersionHistory
# ============================================================
# INSTRUCTIONS :
# Option 1 : Copy this file to config.ps1, fill in your values
#            then run : . .\config.ps1
#
# Option 2 : Set environment variables directly in PowerShell
#            (recommended)
# ============================================================

# Required for both modes
$env:SP_TENANT    = "your-tenant.onmicrosoft.com"
$env:SP_CLIENT_ID = "your-entra-app-client-id"

# Certificate mode only (server / production)
$env:SP_CERT_PATH     = "C:\Certs\your-cert.pfx"
$env:SP_CERT_PASSWORD = "your-cert-password"

# Optional -- auto-derived from SP_TENANT if not set
# $env:SP_TENANT_URL = "https://your-tenant.sharepoint.com"
# $env:SP_ADMIN_URL  = "https://your-tenant-admin.sharepoint.com"

# ============================================================
# USAGE EXAMPLES
# ============================================================

# Certificate mode (server / full tenant scan) :
# .\Remove-SPOVersionHistory.ps1

# Interactive mode (laptop / test) :
# .\Remove-SPOVersionHistory.ps1 -UseInteractive

# Test on a single site :
# .\Remove-SPOVersionHistory.ps1 -UseInteractive -TestSite "AuditReports"

# Test on multiple sites :
# .\Remove-SPOVersionHistory.ps1 -UseInteractive -TestSites @("Accounting","LegalAffairs")
