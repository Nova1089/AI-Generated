<#
.SYNOPSIS
    Reads a list of UPNs from the clipboard, looks each one up in Entra ID via
    Microsoft Graph, and exports key account/license details to a CSV.

.DESCRIPTION
    Copy a list of UPNs (one per line, or comma-separated) to your clipboard,
    then run this script. It will:
      1. Connect to Microsoft Graph (if not already connected)
      2. Parse the clipboard content into a clean list of UPNs
      3. Look up each user, including manager and assigned license info
      4. Export the results to a CSV file

.NOTES
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Users,
              Microsoft.Graph.Identity.DirectoryManagement
    Scopes:   User.Read.All, Directory.Read.All
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\EntraUserExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#region Module check / connect
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Module '$mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -ErrorAction Stop
}

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" | Out-Null
}
#endregion

#region Friendly license names (extend as needed)
# Maps SkuPartNumber -> a human-readable name. Anything not listed here
# will just fall back to the raw SkuPartNumber.
$licenseFriendlyNames = @{
    'SPE_E3'                       = 'Microsoft 365 E3'
    'SPE_E5'                       = 'Microsoft 365 E5'
    'ENTERPRISEPACK'               = 'Office 365 E3'
    'ENTERPRISEPREMIUM'            = 'Office 365 E5'
    'ENTERPRISEPACKWITHOUTPROPLUS' = 'Office 365 E3 (No ProPlus)'
    'O365_BUSINESS_PREMIUM'        = 'Microsoft 365 Business Standard'
    'SPB'                          = 'Microsoft 365 Business Premium'
    'SPE_F1'                       = 'Microsoft 365 F3'
    'FLOW_FREE'                    = 'Power Automate Free'
    'POWER_BI_STANDARD'            = 'Power BI (Free)'
    'AAD_PREMIUM'                  = 'Entra ID P1'
    'AAD_PREMIUM_P2'               = 'Entra ID P2'
    'EMS'                          = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                   = 'Enterprise Mobility + Security E5'
    'TEAMS_EXPLORATORY'            = 'Teams Exploratory'
    'MCOEV'                        = 'Microsoft Teams Phone Standard'
    'STREAM'                       = 'Microsoft Stream'
    'VISIOCLIENT'                  = 'Visio Plan 2'
    'PROJECTPROFESSIONAL'          = 'Project Plan 3'
}

Write-Host "Retrieving tenant SKU catalog..." -ForegroundColor Cyan
$skuLookup = @{}
Get-MgSubscribedSku -All | ForEach-Object {
    $skuLookup[$_.SkuId] = $_.SkuPartNumber
}
#endregion

#region Parse clipboard
$clipboardRaw = Get-Clipboard -Raw
if ([string]::IsNullOrWhiteSpace($clipboardRaw)) {
    Write-Error "Clipboard is empty. Copy a list of UPNs and try again."
    return
}

$upns = $clipboardRaw -split "[\r\n,;]+" |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -match '^\S+@\S+\.\S+$' } |
    Select-Object -Unique

if ($upns.Count -eq 0) {
    Write-Error "No valid UPNs found in clipboard content."
    return
}

Write-Host "Found $($upns.Count) unique UPN(s) in clipboard." -ForegroundColor Cyan
#endregion

#region Lookup users
$results = [System.Collections.Generic.List[object]]::new()
$i = 0

foreach ($upn in $upns) {
    $i++
    Write-Progress -Activity "Looking up users in Entra ID" -Status $upn -PercentComplete (($i / $upns.Count) * 100)

    try {
        $user = Get-MgUser -UserId $upn -Property `
            Id, UserPrincipalName, DisplayName, UserType, AccountEnabled, `
            JobTitle, Department, OfficeLocation, CreatedDateTime, AssignedLicenses `
            -ExpandProperty Manager -ErrorAction Stop

        # Manager display name
        $managerName = $null
        if ($user.Manager -and $user.Manager.AdditionalProperties) {
            $managerName = $user.Manager.AdditionalProperties['displayName']
        }

        # Resolve assigned licenses to friendly names
        $licenseNames = foreach ($lic in $user.AssignedLicenses) {
            $partNumber = $skuLookup[$lic.SkuId]
            if ($partNumber) {
                if ($licenseFriendlyNames.ContainsKey($partNumber)) {
                    $licenseFriendlyNames[$partNumber]
                } else {
                    $partNumber
                }
            }
        }
        $licenseString = ($licenseNames | Sort-Object -Unique) -join '; '

        $results.Add([PSCustomObject]@{
            UPN               = $user.UserPrincipalName
            'Display Name'    = $user.DisplayName
            'User Type'       = $user.UserType
            'Account Enabled' = $user.AccountEnabled
            Title             = $user.JobTitle
            Department        = $user.Department
            'Office Location' = $user.OfficeLocation
            Manager           = $managerName
            'Created Date'    = $user.CreatedDateTime
            'Microsoft Licenses' = $licenseString
        })
    }
    catch {
        Write-Warning "Could not find or retrieve user '$upn': $($_.Exception.Message)"
        $results.Add([PSCustomObject]@{
            UPN               = $upn
            'Display Name'    = 'NOT FOUND'
            'User Type'       = $null
            'Account Enabled' = $null
            Title             = $null
            Department        = $null
            'Office Location' = $null
            Manager           = $null
            'Created Date'    = $null
            'Microsoft Licenses' = $null
        })
    }
}

Write-Progress -Activity "Looking up users in Entra ID" -Completed
#endregion

#region Export
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Export complete: $OutputPath" -ForegroundColor Green
Write-Host "$($results.Count) row(s) written ($($results.Where({$_.'Display Name' -eq 'NOT FOUND'}).Count) not found)." -ForegroundColor Cyan
#endregion
