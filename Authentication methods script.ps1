<#
.SYNOPSIS
Inventories Microsoft Entra ID authentication methods for users defined in a CSV.

.DESCRIPTION
Reads users from CSV, reports configured authentication methods, handles permission/user errors,
and exports the results to a CSV report.
#>

# CSV containing UserPrincipalName column
$CsvPath = Read-Host "Enter the full path of the input CSV file"
$OutPath = Read-Host "Enter the full path where the output CSV should be saved"

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All","UserAuthenticationMethod.Read.All" -NoWelcome

$Users = Import-Csv $CsvPath
$Report = @()

foreach ($User in $Users) {

    try {
        # Check user
        $U = Get-MgUser -UserId $User.UserPrincipalName -ErrorAction Stop

        # Get authentication methods
        $Methods = Get-MgUserAuthenticationMethod `
            -UserId $U.Id `
            -ErrorAction Stop

        if ($Methods.Count -eq 0) {
            $Report += [PSCustomObject]@{
                DisplayName = $U.DisplayName
                UserPrincipalName = $U.UserPrincipalName
                MethodType = "None"
                Status = "Success"
                Error = ""
            }
        }
        else {
            foreach ($Method in $Methods) {

                $Type = $Method.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.',''

                $Report += [PSCustomObject]@{
                    DisplayName = $U.DisplayName
                    UserPrincipalName = $U.UserPrincipalName
                    MethodType = $Type
                    Status = "Success"
                    Error = ""
                }
            }
        }
    }
    catch {
        $ErrorMessage = $_.Exception.Message

        if ($ErrorMessage -match "403|Forbidden|Authorization") {
            $Status = "Permission Denied"
        }
        elseif ($ErrorMessage -match "404|NotFound|does not exist") {
            $Status = "User Not Found"
        }
        else {
            $Status = "Error"
        }

        $Report += [PSCustomObject]@{
            DisplayName = ""
            UserPrincipalName = $User.UserPrincipalName
            MethodType = ""
            Status = $Status
            Error = $ErrorMessage
        }

        continue
    }
}

$Report | Export-Csv $OutPath -NoTypeInformation -Encoding UTF8

Write-Host "`nAuthentication-method inventory completed." -ForegroundColor Green
Write-Host "Report: $OutPath" -ForegroundColor Cyan
