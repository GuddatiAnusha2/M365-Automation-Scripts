#requires -Version 7.2
<#
.SYNOPSIS
Audits Exchange Online MRM retention-policy assignments and exports a CSV.

.DESCRIPTION
Read-only. Retrieves every Exchange Online mailbox, including user, shared,
room, equipment, discovery, arbitration, and other mailbox types returned by
Get-EXOMailbox. Mailboxes with no assigned MRM policy are Noncompliant.
Comparisons are performed locally (O(M + T)).
#>

[CmdletBinding()]
param(
    [string]$ExpectedPolicy = 'Default MRM Policy',
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('UserProfile')) ("Downloads\Exchange_RetentionPolicy_Audit_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),
    [switch]$DeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$connected = $false

try {
    # Install Exchange Online Management module if required
    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        Write-Host 'Installing ExchangeOnlineManagement...' -ForegroundColor Yellow

        Install-Module ExchangeOnlineManagement `
            -Scope CurrentUser `
            -Repository PSGallery `
            -Force `
            -AllowClobber
    }

    Import-Module ExchangeOnlineManagement

    # Connect to Exchange Online
    $connect = @{
        ShowBanner = $false
    }

    if ($DeviceCode) {
        $connect.Device = $true
    }

    Connect-ExchangeOnline @connect
    $connected = $true

    Write-Host 'Connected to Exchange Online.' -ForegroundColor Green

    # Validate expected retention policy
    $policy = Get-RetentionPolicy `
        -Identity $ExpectedPolicy `
        -ErrorAction SilentlyContinue

    if (-not $policy) {
        throw "Expected policy '$ExpectedPolicy' does not exist."
    }

    Write-Host "Retention policy found: $ExpectedPolicy" -ForegroundColor Green

    # Get retention policy tag names
    $tagNames = @(
        $policy.RetentionPolicyTagLinks | ForEach-Object {
            if ($_.PSObject.Properties['Name']) {
                $_.Name
            }
            else {
                [string]$_
            }
        }
    )

    # Get retention policy tags
    $policyTags = @(
        Get-RetentionPolicyTag |
            Where-Object { $_.Name -in $tagNames }
    )

    $hasArchiveTag = [bool](
        $policyTags |
            Where-Object RetentionAction -EQ 'MoveToArchive'
    )

    # Check organization-level ELC processing
    $orgElcDisabled = [bool](
        (Get-OrganizationConfig).ElcProcessingDisabled
    )

    # Mailbox properties required by the audit.
    # ExchangeGuid has been explicitly added to fix the
    # "property ExchangeGuid cannot be found" error.
    $properties = @(
        'RetentionPolicy',
        'RetentionHoldEnabled',
        'ElcProcessingDisabled',
        'ArchiveStatus',
        'LitigationHoldEnabled',
        'InPlaceHolds',
        'ExchangeGuid'
    )

    Write-Host 'Retrieving Exchange Online mailboxes...' -ForegroundColor Cyan

    $mailboxes = @(
        Get-EXOMailbox `
            -ResultSize Unlimited `
            -Properties $properties
    )

    if ($mailboxes.Count -eq 0) {
        throw 'No Exchange Online mailboxes were returned.'
    }

    Write-Host "Mailboxes retrieved: $($mailboxes.Count)" -ForegroundColor Green

    # Build audit report
    $report = foreach ($m in $mailboxes) {

        $actual = [string]$m.RetentionPolicy

        $status, $gap, $action = if ([string]::IsNullOrWhiteSpace($actual)) {

            'Noncompliant',
            'No retention policy assigned',
            "Assign '$ExpectedPolicy'"

        }
        elseif ($actual -ne $ExpectedPolicy) {

            'Noncompliant',
            "Incorrect policy: $actual",
            "Replace with '$ExpectedPolicy'"

        }
        elseif ($policyTags.Count -eq 0) {

            'Noncompliant',
            'Expected policy contains no retention tags',
            'Add and validate retention tags'

        }
        elseif ($orgElcDisabled -or $m.ElcProcessingDisabled) {

            'Warning',
            'Managed Folder Assistant processing is disabled',
            'Review organization and mailbox ELC settings'

        }
        elseif ($m.RetentionHoldEnabled) {

            'Warning',
            'Retention hold is enabled',
            'Confirm that the retention hold is intentional'

        }
        elseif (
            $hasArchiveTag -and
            [string]$m.ArchiveStatus -ne 'Active'
        ) {

            'Warning',
            'Policy has an archive tag but archive is not active',
            'Enable archive or review the archive tag'

        }
        else {

            'Compliant',
            'None',
            'No action required'
        }

        [pscustomobject][ordered]@{

            DisplayName = $m.DisplayName

            UserPrincipalName = $m.UserPrincipalName

            PrimarySmtpAddress = $m.PrimarySmtpAddress

            ExchangeGuid = $m.ExchangeGuid

            RecipientType = $m.RecipientTypeDetails

            ExpectedPolicy = $ExpectedPolicy

            ActualPolicy = $actual

            PolicyTagCount = $policyTags.Count

            ArchiveStatus = $m.ArchiveStatus

            RetentionHoldEnabled = $m.RetentionHoldEnabled

            MailboxElcProcessingDisabled = $m.ElcProcessingDisabled

            OrganizationElcProcessingDisabled = $orgElcDisabled

            LitigationHoldEnabled = $m.LitigationHoldEnabled

            ComplianceStatus = $status

            GapFound = $gap

            RecommendedAction = $action

            ReportGeneratedUtc = (
                Get-Date
            ).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        }
    }

    # Ensure output folder exists
    $folder = Split-Path $OutputPath -Parent

    if ([string]::IsNullOrWhiteSpace($folder)) {

        $folder = (Get-Location).Path

        $OutputPath = Join-Path $folder $OutputPath
    }

    if (-not (Test-Path -LiteralPath $folder)) {

        New-Item `
            $folder `
            -ItemType Directory `
            -Force |
            Out-Null
    }

    # Export CSV
    $report |
        Sort-Object ComplianceStatus, DisplayName |
        Export-Csv `
            -LiteralPath $OutputPath `
            -NoTypeInformation `
            -Encoding utf8BOM

    # Display summary
    Write-Host ''
    Write-Host 'Audit Summary' -ForegroundColor Cyan
    Write-Host '-------------' -ForegroundColor Cyan

    $report |
        Group-Object ComplianceStatus |
        Sort-Object Name |
        Select-Object Name, Count |
        Format-Table -AutoSize

    Write-Host "Mailboxes audited: $($report.Count)" -ForegroundColor Cyan
    Write-Host "CSV report: $OutputPath" -ForegroundColor Green
}
catch {

    Write-Error "Retention-policy audit failed: $($_.Exception.Message)"
}
finally {

    if ($connected) {

        Disconnect-ExchangeOnline -Confirm:$false
    }
}
