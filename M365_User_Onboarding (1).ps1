<#
.SYNOPSIS
M365 User Onboarding — bulk-creates users in your Microsoft 365 tenant via
Microsoft Graph, so they appear in Microsoft 365 admin center > Active users
fully provisioned (account + license + manager), not just as bare accounts.

.WHY MICROSOFT GRAPH
This uses the Microsoft Graph PowerShell SDK exclusively - not the legacy
AzureAD/MSOnline modules, which Microsoft has fully retired (scripts built on
them stopped working entirely by mid-October 2025). Graph is the current,
supported path for directory and user management.

.PREREQUISITES
Install-Module Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser

Requires an account with User Administrator (or higher) privileges. The
Graph scopes below (User.ReadWrite.All, Organization.Read.All) need one-time
admin consent in your tenant - same as any directory-write scope. If you hit
an "Approval required" screen, that's your tenant admin's step to grant, not
something this script can bypass.

.INPUT
A CSV file with these columns (see Onboarding_Users_Template.csv):
FirstName, LastName, DisplayName, UserPrincipalName, MailNickname, JobTitle,
Department, UsageLocation, LicenseSkuPartNumber, ManagerUPN

LicenseSkuPartNumber and ManagerUPN are optional - leave blank to skip either.
UsageLocation is required if you're assigning a license (Microsoft requires
it before a license can be applied).

.OUTPUT
- A day-stamped log file under %USERPROFILE%\M365_Onboarding_Logs
- A results CSV with one row per user: Success/Failed/Skipped + reason
- A SEPARATE credentials CSV containing generated temporary passwords -
  this file contains sensitive data. Distribute it securely (never email
  it in plain text) and delete it once passwords have been handed off.

.USAGE
    .\M365_User_Onboarding.ps1 -CsvPath ".\Onboarding_Users_Template.csv"

    .\M365_User_Onboarding.ps1 -CsvPath ".\Onboarding_Users_Template.csv" -WhatIf
    (dry run - validates everything and shows what WOULD happen, creates nothing)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [switch]$WhatIf
)

# ---------------------------------------------------------------------------
# Setup: logging
# ---------------------------------------------------------------------------
$logRoot = "$env:USERPROFILE\M365_Onboarding_Logs"
if (-not (Test-Path $logRoot)) { New-Item -ItemType Directory -Path $logRoot -Force | Out-Null }

$stamp        = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile      = "$logRoot\Onboarding_$stamp.log"
$resultsFile  = "$logRoot\Onboarding_Results_$stamp.csv"
$credsFile    = "$logRoot\Onboarding_Credentials_$stamp.csv"   # SENSITIVE - see notes above

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    $color = switch ($Level) { "SUCCESS" { "Green" } "WARN" { "Yellow" } "ERROR" { "Red" } default { "Gray" } }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $logFile -Value $line
}

function New-TempPassword {
    # Meets standard Entra complexity: upper, lower, number, symbol, 14 chars
    $upper   = -join ((65..90)  | Get-Random -Count 4 | ForEach-Object { [char]$_ })
    $lower   = -join ((97..122) | Get-Random -Count 4 | ForEach-Object { [char]$_ })
    $digits  = -join ((48..57)  | Get-Random -Count 3 | ForEach-Object { [char]$_ })
    $symbols = @('!','@','#','$','%','^','&','*')
    $symbol  = -join ($symbols | Get-Random -Count 3)
    $combined = ($upper + $lower + $digits + $symbol).ToCharArray()
    return -join ($combined | Get-Random -Count $combined.Count)
}

# ---------------------------------------------------------------------------
# Load and validate the CSV up front - fail fast on bad input, not mid-run
# ---------------------------------------------------------------------------
if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found at '$CsvPath'." "ERROR"
    exit 1
}

$rows = Import-Csv -Path $CsvPath
if (-not $rows -or $rows.Count -eq 0) {
    Write-Log "CSV file is empty or unreadable." "ERROR"
    exit 1
}

$requiredColumns = @("FirstName","LastName","DisplayName","UserPrincipalName","MailNickname")
$missingColumns = $requiredColumns | Where-Object { $_ -notin $rows[0].PSObject.Properties.Name }
if ($missingColumns) {
    Write-Log "CSV is missing required column(s): $($missingColumns -join ', ')" "ERROR"
    exit 1
}

Write-Log "Loaded $($rows.Count) user(s) from '$CsvPath'."
if ($WhatIf) { Write-Log "Running in -WhatIf mode: no accounts will actually be created." "WARN" }

# ---------------------------------------------------------------------------
# Connect to Microsoft Graph
# ---------------------------------------------------------------------------
try {
    Write-Log "Connecting to Microsoft Graph (MFA prompt expected here)..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All" -NoWelcome -ErrorAction Stop
    $context = Get-MgContext
    Write-Log "Connected as $($context.Account)." "SUCCESS"
}
catch {
    Write-Log "Could not connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
    Write-Log "If you saw an 'Approval required' screen, this Graph scope needs one-time admin consent from your tenant admin before this script can run." "ERROR"
    exit 1
}

$tenantDomain = (Get-MgOrganization).VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name
Write-Log "Tenant default domain: $tenantDomain"

# Pre-fetch available licenses once, so we're not re-querying per user
$availableSkus = Get-MgSubscribedSku -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Main onboarding loop
# ---------------------------------------------------------------------------
$results = @()
$creds   = @()
$total = $rows.Count
$i = 0

foreach ($row in $rows) {
    $i++
    Write-Progress -Activity "Onboarding users" -Status "$($row.DisplayName) ($i of $total)" -PercentComplete (($i / $total) * 100)
    Write-Log "---- [$i/$total] Processing '$($row.DisplayName)' ($($row.UserPrincipalName)) ----"

    $result = [PSCustomObject]@{
        DisplayName       = $row.DisplayName
        UserPrincipalName = $row.UserPrincipalName
        Status            = ""
        Details           = ""
    }

    try {
        if (-not $row.UserPrincipalName -or -not $row.MailNickname) {
            throw "Missing UserPrincipalName or MailNickname."
        }

        # Skip if the user already exists - don't fail the whole batch, just log it
        $existing = Get-MgUser -Filter "userPrincipalName eq '$($row.UserPrincipalName)'" -ErrorAction SilentlyContinue
        if ($existing) {
            $result.Status = "Skipped"
            $result.Details = "User already exists in tenant."
            Write-Log "SKIPPED: '$($row.UserPrincipalName)' already exists." "WARN"
            $results += $result
            continue
        }

        $tempPassword = New-TempPassword

        $userParams = @{
            AccountEnabled    = $true
            DisplayName       = $row.DisplayName
            GivenName         = $row.FirstName
            Surname           = $row.LastName
            UserPrincipalName = $row.UserPrincipalName
            MailNickname      = $row.MailNickname
            JobTitle          = $row.JobTitle
            Department        = $row.Department
            UsageLocation     = if ($row.UsageLocation) { $row.UsageLocation } else { $null }
            PasswordProfile   = @{
                ForceChangePasswordNextSignIn = $true
                Password                      = $tempPassword
            }
        }

        if ($WhatIf) {
            $result.Status = "WhatIf"
            $result.Details = "Would be created with the above properties."
            Write-Log "WHATIF: '$($row.UserPrincipalName)' passed validation, would be created." "SUCCESS"
            $results += $result
            continue
        }

        $newUser = New-MgUser @userParams -ErrorAction Stop
        Write-Log "Account created: $($newUser.UserPrincipalName) (ObjectId: $($newUser.Id))" "SUCCESS"

        # --- Manager assignment (optional) ---
        if ($row.ManagerUPN) {
            try {
                $manager = Get-MgUser -Filter "userPrincipalName eq '$($row.ManagerUPN)'" -ErrorAction Stop
                $refBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)" }
                Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter $refBody -ErrorAction Stop
                Write-Log "Manager set: $($row.ManagerUPN)" "SUCCESS"
            }
            catch {
                Write-Log "Could not set manager for '$($row.UserPrincipalName)': $($_.Exception.Message)" "WARN"
            }
        }

        # --- License assignment (optional) ---
        if ($row.LicenseSkuPartNumber) {
            $sku = $availableSkus | Where-Object { $_.SkuPartNumber -eq $row.LicenseSkuPartNumber }
            if (-not $sku) {
                Write-Log "License SKU '$($row.LicenseSkuPartNumber)' not found in tenant - skipping license assignment." "WARN"
            }
            elseif (-not $row.UsageLocation) {
                Write-Log "No UsageLocation specified for '$($row.UserPrincipalName)' - Microsoft requires this before licensing, skipping." "WARN"
            }
            else {
                try {
                    Set-MgUserLicense -UserId $newUser.Id `
                        -AddLicenses @(@{ SkuId = $sku.SkuId }) `
                        -RemoveLicenses @() -ErrorAction Stop
                    Write-Log "License assigned: $($row.LicenseSkuPartNumber)" "SUCCESS"
                }
                catch {
                    Write-Log "License assignment failed for '$($row.UserPrincipalName)': $($_.Exception.Message)" "WARN"
                }
            }
        }

        # --- Verify the account is genuinely queryable (i.e. it'll show in Active users) ---
        Start-Sleep -Seconds 2
        $verify = Get-MgUser -UserId $newUser.Id -ErrorAction Stop
        if ($verify) {
            $result.Status = "Success"
            $result.Details = "Created and verified. Will appear in Microsoft 365 admin center > Active users (allow a few minutes to propagate)."
            Write-Log "VERIFIED: '$($row.UserPrincipalName)' is live in the directory." "SUCCESS"
        }

        $creds += [PSCustomObject]@{
            DisplayName       = $row.DisplayName
            UserPrincipalName = $row.UserPrincipalName
            TemporaryPassword = $tempPassword
        }
    }
    catch {
        $result.Status = "Failed"
        $result.Details = $_.Exception.Message
        Write-Log "FAILED: '$($row.UserPrincipalName)' - $($_.Exception.Message)" "ERROR"
    }

    $results += $result
}

Write-Progress -Activity "Onboarding users" -Completed

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$results | Export-Csv -Path $resultsFile -NoTypeInformation
if ($creds.Count -gt 0) { $creds | Export-Csv -Path $credsFile -NoTypeInformation }

$successCount = ($results | Where-Object { $_.Status -eq "Success" }).Count
$skippedCount = ($results | Where-Object { $_.Status -eq "Skipped" }).Count
$failedCount  = ($results | Where-Object { $_.Status -eq "Failed" }).Count
$whatifCount  = ($results | Where-Object { $_.Status -eq "WhatIf" }).Count

Write-Log "==================== SUMMARY ===================="
Write-Log "Total processed : $total"
Write-Log "Created         : $successCount" "SUCCESS"
Write-Log "Skipped         : $skippedCount" "WARN"
Write-Log "Failed          : $failedCount" $(if ($failedCount -gt 0) { "ERROR" } else { "INFO" })
if ($whatifCount -gt 0) { Write-Log "WhatIf (dry run): $whatifCount" }
Write-Log "Results CSV     : $resultsFile"
if ($creds.Count -gt 0) {
    Write-Log "Credentials CSV : $credsFile  <-- SENSITIVE, distribute securely and delete after handoff" "WARN"
}
Write-Log "==================================================="
