__________________________________


$report = foreach ($group in Get-ADGroup -Filter * -Properties mail, proxyAddresses) {

    $proxies = @($group.proxyAddresses | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $hasPrimary = $proxies -cmatch '^SMTP:'
    $newProxies = $proxies

    if (-not $hasPrimary -and $group.mail) {
        $candidateProxies = [string[]](@(
            $proxies | Where-Object { $_ -cne "smtp:$($group.mail)" }
        ) + "SMTP:$($group.mail)")

        try {
            Set-ADGroup $group -Replace @{proxyAddresses = $candidateProxies} -ErrorAction Stop
            $newProxies = $candidateProxies
        }
        catch {
            Write-Warning "Failed to update $($group.Name): $($_.Exception.Message)"
        }
    }

    [PSCustomObject]@{
        Name                 = $group.Name
        Mail                 = $group.mail
        ProxyAddressesBefore = $proxies -join ';'
        ProxyAddressesAfter  = $newProxies -join ';'
        HadPrimarySMTP       = [bool]$hasPrimary
    }
}

$report | Export-Csv .\ADGroupSmtpAudit.csv -NoTypeInformation




_______________________________________________________________________________________
$report = foreach ($group in Get-ADGroup -Filter * -Properties mail, proxyAddresses) {

    $proxies = @($group.proxyAddresses | ForEach-Object { [string]$_ } | Where-Object { $_ })

    $badValues  = $proxies | Where-Object { ($_ -split '(?i)(?=smtp:)').Count -gt 2 }
    $goodValues = $proxies | Where-Object { ($_ -split '(?i)(?=smtp:)').Count -le 2 }

    if ($badValues) {
        # Split each corrupted value back into its individual addresses
        $splitValues = $badValues | ForEach-Object { $_ -split '(?i)(?=smtp:)' } | Where-Object { $_ }

        $fixedProxies = @($goodValues) + @($splitValues) |
            Select-Object -Unique |
            ForEach-Object { [string]$_ }
            $fixedProxies = @($fixedProxies)

        # Ensure only one entry keeps the uppercase (primary) SMTP: prefix
        $primaries = $fixedProxies | Where-Object { $_ -cmatch '^SMTP:' }
        if ($primaries.Count -gt 1) {
            $keepPrimary = $primaries[0]
            $fixedProxies = $fixedProxies | ForEach-Object {
                if ($_ -cmatch '^SMTP:' -and $_ -ne $keepPrimary) {
                    'smtp:' + $_.Substring(5)   # demote extra primaries to secondary
                } else {
                    $_
                }
            }
        }

        try {
            Set-ADGroup $group -Replace @{proxyAddresses = $fixedProxies} -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to fix $($group.Name): $($_.Exception.Message)"
        }
    }

    [PSCustomObject]@{
        Name            = $group.Name
        BadValuesFound  = ($badValues -join ' | ')
        ProxiesBefore   = $proxies -join ';'
        ProxiesAfter    = if ($badValues) { $fixedProxies -join ';' } else { $proxies -join ';' }
    }
}

$report | Where-Object { $_.BadValuesFound } | Export-Csv .\ADGroupSmtpRepair.csv -NoTypeInformation



_______________________________________________________________________________________
$groups = Get-ADGroup -Filter * -Properties proxyAddresses |
    ForEach-Object {
        $bad = $_.proxyAddresses | Where-Object { ($_ -split '(?i)(?=smtp:)').Count -gt 2 }
        if ($bad) {
            [PSCustomObject]@{ Name = $_.Name; BadValue = $bad -join ' | ' }
        }
    }