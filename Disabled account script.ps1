#requires -Version 7.0

[CmdletBinding()]
param(
    [string]$OutputPath = ".\Downloads\DisabledUsers_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"),

    # 0 = search all audit records retained by Entra ID.
    # Use a positive number (e.g. 90) to reduce audit-log volume.
    [int]$AuditLookbackDays = 90,

    # For unattended execution, provide these values and use certificate auth.
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'

try {
    # Required modules
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Reports

    # Authenticate
    if ($TenantId -and $ClientId -and $CertificateThumbprint) {
        Connect-MgGraph -TenantId $TenantId `
                        -ClientId $ClientId `
                        -CertificateThumbprint $CertificateThumbprint `
                        -NoWelcome
    }
    else {
        Connect-MgGraph -Scopes @(
            'User.Read.All',
            'AuditLog.Read.All',
            'User-LifeCycleInfo.Read.All'
        ) -NoWelcome
    }

    # Retrieve only disabled users and only required attributes.
    # Manager is expanded in the same request to avoid N additional manager calls.
    $users = Get-MgUser `
        -Filter "accountEnabled eq false" `
        -Property @(
            'id',
            'mail',
            'userPrincipalName',
            'displayName',
            'accountEnabled',
            'createdDateTime',
            'department',
            'userType',
            'onPremisesSyncEnabled',
            'employeeLeaveDateTime',
            'signInActivity'
        ) `
        -ExpandProperty 'manager($select=displayName,userPrincipalName,mail)' `
        -PageSize 500 `
        -All

    # Build a lookup of disabled dates from audit logs.
    $disabledDate = @{}

    $auditFilter = "activityDisplayName eq 'Disable account' and result eq 'success'"

    if ($AuditLookbackDays -gt 0) {
        $from = (Get-Date).ToUniversalTime().AddDays(-$AuditLookbackDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $auditFilter += " and activityDateTime ge $from"
    }

    Get-MgAuditLogDirectoryAudit -Filter $auditFilter -All |
        ForEach-Object {
            foreach ($target in $_.TargetResources) {
                if ($target.Type -eq 'User' -and $target.Id) {
                    $id = $target.Id.ToString()
                    if (-not $disabledDate.ContainsKey($id) -or
                        $_.ActivityDateTime -gt $disabledDate[$id]) {
                        $disabledDate[$id] = $_.ActivityDateTime
                    }
                }
            }
        }

    # Transform to the final CSV schema.
    $report = $users | ForEach-Object {
        $lastLogin = $_.SignInActivity.LastSuccessfulSignInDateTime
        if (-not $lastLogin) {
            $lastLogin = $_.SignInActivity.LastSignInDateTime
        }

        $manager = $_.Manager
        $managerEmail = if ($manager) {
            $manager.Mail ?? $manager.UserPrincipalName
        }

        $lifecycle = if ($_.EmployeeLeaveDateTime) {
            if ([datetime]$_.EmployeeLeaveDateTime -le (Get-Date)) {
                'Leaver'
            }
            else {
                'Scheduled Leaver'
            }
        }
        else {
            'No Leave Date'
        }

        [pscustomobject]@{
            AccountId        = $_.Id
            UsernameEmail    = $_.Mail ?? $_.UserPrincipalName
            DisplayName      = $_.DisplayName
            AccountStatus    = if ($_.AccountEnabled) { 'Enabled' } else { 'Disabled' }
            DisabledDate    = if ($disabledDate.ContainsKey($_.Id)) {
                                   $disabledDate[$_.Id].ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                               } else { $null }
            LastLoginDate    = if ($lastLogin) {
                                   ([datetime]$lastLogin).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                               } else { $null }
            CreatedDate      = if ($_.CreatedDateTime) {
                                   ([datetime]$_.CreatedDateTime).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                               } else { $null }
            Department       = $_.Department
            Manager          = $managerEmail
            AccountType      = $_.UserType
            SourceSystem     = if ($_.OnPremisesSyncEnabled) {
                                   'On-Premises AD (Synced)'
                               } else {
                                   'Microsoft Entra ID (Cloud)'
                               }
            EmployeeLeaveDate = if ($_.EmployeeLeaveDateTime) {
                                    ([datetime]$_.EmployeeLeaveDateTime).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                                } else { $null }
            LifecycleState   = $lifecycle
        }
    }

    # Ensure the destination directory exists.
    $directory = Split-Path -Parent $OutputPath
    if ($directory -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    # Export final report.
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "Disabled-user report generated successfully."
    Write-Host "Disabled accounts : $($report.Count)"
    Write-Host "Output            : $OutputPath"
}
catch {
    Write-Error "Disabled-user report failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
