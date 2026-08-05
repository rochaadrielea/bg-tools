# ============================================================
# open_firewall.ps1
# Opens port 8501 for the Quality BRM dashboard,
# but ONLY on the private (BG) network profile.
#
# Public networks (e.g. hotel wifi, cafe) will NOT be able to reach it.
#
# Run as Administrator:
#   Right-click -> Run with PowerShell (as Administrator)
# ============================================================

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    Write-Host "Right-click the file and choose 'Run with PowerShell (as Administrator)'"
    exit 1
}

$RuleName = "Quality BRM Dashboard (port 8501, private only)"

# Remove existing rule if it exists
$existing = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing firewall rule..." -ForegroundColor Yellow
    Remove-NetFirewallRule -DisplayName $RuleName
}

# Create firewall rule - private profile only
New-NetFirewallRule `
    -DisplayName $RuleName `
    -Direction Inbound `
    -Action Allow `
    -Protocol TCP `
    -LocalPort 8501 `
    -Profile Private `
    -Description "Allows inbound connections to the Quality BRM Streamlit dashboard. Restricted to Private network profile (BG office WiFi / VPN). Public networks are blocked."

Write-Host ""
Write-Host "SUCCESS: Firewall rule created." -ForegroundColor Green
Write-Host ""
Write-Host "Port 8501 is now open on Private networks only."
Write-Host "This means:"
Write-Host "  ALLOWED: BG office WiFi, BG VPN, home network marked as Private" -ForegroundColor Green
Write-Host "  BLOCKED: Public WiFi (cafes, hotels, airports)" -ForegroundColor Red
Write-Host ""
Write-Host "IMPORTANT: When you connect to a new WiFi, Windows will ask:"
Write-Host "  'Do you want your PC to be discoverable on this network?'"
Write-Host "  Yes = Private profile (dashboard reachable)"
Write-Host "  No  = Public profile (dashboard blocked)"
Write-Host ""
Write-Host "To check the rule:"
Write-Host "    Get-NetFirewallRule -DisplayName '$RuleName'" -ForegroundColor Cyan
Write-Host ""
Write-Host "To remove it later:"
Write-Host "    Remove-NetFirewallRule -DisplayName '$RuleName'" -ForegroundColor Cyan
