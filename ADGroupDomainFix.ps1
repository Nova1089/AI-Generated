<#
.SYNOPSIS
    Migrates AD group mail from @OldDomain to @NewDomain, keeps the old
    address as an alias, sets the new address as primary SMTP.
#>

$WhatIfMode = $false
$OldDomain = "liveuncertain.com"
$NewDomain = "ocdanxietycenters.com"
$ReportPath = ".\ADGroupDomainFix_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$groups = Get-ADGroup -LDAPFilter "(mail=*@$OldDomain)" -Properties mail, proxyAddresses
$report = foreach ($group in $groups) {
    Write-Progress -Activity "Updating..." -Status $group.Name
    $oldMail = $group.mail
    $newMail = $oldMail -replace [regex]::Escape("@$OldDomain"), "@$NewDomain"
    $proxies = @($group.proxyAddresses)

    $oldPrimarySmtp = ($proxies | Where-Object { $_ -cmatch '^SMTP:' }) -replace '^SMTP:', ''

    $row = [ordered]@{
        GroupName  = $group.Name
        DistinguishedName = $group.DistinguishedName
        SamAccountName = $group.SamAccountName
        ObjectGUID = $group.ObjectGUID
        OldMail    = $oldMail
        OldPrimarySmtp    = $oldPrimarySmtp
        NewMail    = $newMail        
        OldProxies = $proxies -join ';'
        NewProxies = $null
        Status     = $null
    }

    try {
        # keep old address as an alias if it's not already listed

        if (-not ($proxies -icontains "smtp:$oldMail")) {
            $proxies += "smtp:$oldMail"
        }

        # strip any existing primary flag, drop the new address if it's already present so we don't duplicate it
        $proxies = @( $proxies | ForEach-Object { $_ -replace '^SMTP:', 'smtp:' } )
        $proxies = @( $proxies | Where-Object { $_ -ine "smtp:$newMail" } )
        
        # add new address as the primary
        $proxies += "SMTP:$newMail"
        
        # Ensure no duplicates in proxies and cast to string[]
        [string[]]$proxies = @( $proxies | Select-Object -Unique | ForEach-Object { [string]$_ } )

        $row.NewProxies = $proxies -join ';'
        
        if ($WhatIfMode) {
            $row.Status = 'WhatIf'
        } else {
            Set-ADGroup -Identity $group -Replace @{ mail = $newMail; proxyAddresses = $proxies } -ErrorAction Stop
            $row.Status = 'Success'
        }
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor "Yellow"
        $row.Status = "Failed: $($_.Exception.Message)"
    }

    [PSCustomObject]$row
}

$report | Export-Csv -Path $ReportPath -NoTypeInformation
Write-Host "Report saved to $ReportPath"