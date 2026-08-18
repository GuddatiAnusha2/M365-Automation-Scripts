<#
.SYNOPSIS
Entra ID User Export - connects to Microsoft Graph and exports a CSV of users
with UPN, display name, account state, user type, and department.

.TENANT
InnovaticT Technologies (InnovaticT.onmicrosoft.com)

.WHY MICROSOFT GRAPH
Uses the modern, supported Microsoft Graph PowerShell SDK - not the retired
AzureAD/MSOnline modules - consistent with the rest of this automation set.

.PREREQUISITES
Install-Module Microsoft.Graph.Users -Scope CurrentUser

Needs an account that can read directory user data. This script only READS
data (no writes), so it requests the read-only User.Read.All scope - the
least-privilege option for this task. Like any directory-read scope, it may
require one-time admin consent the first time it's used in the tenant.

.OUTPUT
A CSV file with one row per user:
DisplayName, UserPrincipalName, AccountState (Active/Disabled), UserType
(Member/Guest), Department

.USAGE
    .\Entra_ActiveUsers_Export.ps1
    (by default, saves the CSV and log file straight to your Downloads folder)
    .\Entra_ActiveUsers_Export.ps1 -OutputFolder "C:\Reports"
#>

param(
    [string]$Upn = "Aguddati@InnovaticT.onmicrosoft.com",
    [string]$OutputFolder = "$env:USERPROFILE\Downloads"
)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

$stamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = "$OutputFolder\Entra_UserExport_$stamp.log"
$csvFile   = "$OutputFolder\Entra_UserExport_$stamp.csv"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    $color = switch ($Level) { "SUCCESS" { "Green" } "WARN" { "Yellow" } "ERROR" { "Red" } default { "Gray" } }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $logFile -Value $line
}

# ---------------------------------------------------------------------------
# Connect to Microsoft Graph
# ---------------------------------------------------------------------------
try {
    Write-Log "Connecting to Microsoft Graph as $Upn (MFA prompt expected here)..."
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome -ErrorAction Stop
    $context = Get-MgContext
    Write-Log "Connected as $($context.Account)." "SUCCESS"
}
catch {
    Write-Log "Could not connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
    Write-Log "If you saw an 'Approval required' screen, User.Read.All needs one-time admin consent from your tenant admin." "ERROR"
    exit 1
}

# ---------------------------------------------------------------------------
# Fetch users
# ---------------------------------------------------------------------------
try {
    Write-Log "Fetching users from Entra ID (this may take a moment for large tenants)..."
    $users = Get-MgUser -All -Property "DisplayName,UserPrincipalName,AccountEnabled,UserType,Department" `
        -ErrorAction Stop |
        Select-Object DisplayName, UserPrincipalName, AccountEnabled, UserType, Department

    Write-Log "Retrieved $($users.Count) user(s)." "SUCCESS"
}
catch {
    Write-Log "Failed to fetch users: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ---------------------------------------------------------------------------
# Build the export - map raw fields to the requested column names
# ---------------------------------------------------------------------------
$export = $users | ForEach-Object {
    [PSCustomObject]@{
        DisplayName       = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        AccountState      = if ($_.AccountEnabled) { "Active" } else { "Disabled" }
        UserType          = $_.UserType
        Department        = if ($_.Department) { $_.Department } else { "(Not set)" }
    }
}

$export | Sort-Object DisplayName | Export-Csv -Path $csvFile -NoTypeInformation

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$activeCount   = ($export | Where-Object { $_.AccountState -eq "Active" }).Count
$disabledCount = ($export | Where-Object { $_.AccountState -eq "Disabled" }).Count
$guestCount    = ($export | Where-Object { $_.UserType -eq "Guest" }).Count
$memberCount   = ($export | Where-Object { $_.UserType -eq "Member" }).Count

Write-Log "==================== SUMMARY ===================="
Write-Log "Total users : $($export.Count)"
Write-Log "Active      : $activeCount" "SUCCESS"
Write-Log "Disabled    : $disabledCount" $(if ($disabledCount -gt 0) { "WARN" } else { "INFO" })
Write-Log "Members     : $memberCount"
Write-Log "Guests      : $guestCount"
Write-Log "CSV file    : $csvFile"
Write-Log "==================================================="

Write-Host "`nDone. Opening the output folder..." -ForegroundColor Cyan
Start-Process explorer.exe $OutputFolder
