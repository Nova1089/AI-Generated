<#
.SYNOPSIS
    Migrates AD group mail from @OldDomain to @NewDomain, keeps the old
    address as an alias, sets the new address as primary SMTP.
#>

$WhatIfMode = $true
$OldDomain = "liveuncertain.com"
$NewDomain = "ocdanxietycenters.com"
$ReportPath = ".\ADGroupDomainFix_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$groups = Get-ADGroup -LDAPFilter "(mail=*@$OldDomain)" -Properties mail, proxyAddresses | Select -First 5
$report = foreach ($group in $groups) {

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

        Write-Host "`nAfter loading:"
        $proxies | ForEach-Object { Write-Host "[$_]" }

        if (-not ($proxies -icontains "smtp:$oldMail")) {
            $proxies += "smtp:$oldMail"
        }

        Write-Host "`nAfter adding old alias:"
        $proxies | ForEach-Object { Write-Host "[$_]" }

        # strip any existing primary flag, drop the new address if it's already present so we don't duplicate it
        $proxies = @( $proxies | ForEach-Object { $_ -replace '^SMTP:', 'smtp:' } )

        Write-Host "`nAfter removing primary flag:"
        $proxies | ForEach-Object { Write-Host "[$_]" }


        $proxies = @( $proxies | Where-Object { $_ -ine "smtp:$newMail" } )

        Write-Host "`nAfter removing new alias:"
        $proxies | ForEach-Object { Write-Host "[$_]" }

        # add new address as the primary
        Write-Host $proxies.GetType().FullName
        $proxies += "SMTP:$newMail"

        Write-Host "`nAfter adding new primary:"
        $proxies | ForEach-Object { Write-Host "[$_]" }

        $proxies = @( $proxies | Select-Object -Unique )

        $row.NewProxies = $proxies -join ';'
        
        if ($WhatIfMode) {
            $row.Status = 'WhatIf'
        } else {
            Set-ADGroup -Identity $group -Replace @{ mail = $newMail; proxyAddresses = $proxies } -ErrorAction Stop
            $row.Status = 'Success'
        }
    }
    catch {
        $row.Status = "Failed: $($_.Exception.Message)"
    }

    [PSCustomObject]$row
}

$report | Export-Csv -Path $ReportPath -NoTypeInformation
Write-Host "Report saved to $ReportPath"