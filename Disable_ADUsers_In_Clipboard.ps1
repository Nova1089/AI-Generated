<#
.SYNOPSIS
    Disables a list of Entra ID users (UPNs from clipboard) and generates a before/after report.

.NOTES
    Requires: Microsoft.Graph.Users module
    Scope needed: User.ReadWrite.All
#>

# ---- Setup ----
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

# Grab UPNs from clipboard, one per line
$upns = Get-Clipboard | Where-Object { $_.Trim() -ne "" }

if (-not $upns) {
    Write-Host "Clipboard is empty or contains no UPNs. Exiting." -ForegroundColor Yellow
    return
}

Write-Host "Found $($upns.Count) UPNs on clipboard." -ForegroundColor Cyan

# ---- Process ----
$report = foreach ($upn in $upns) {
    $i = [array]::IndexOf($upns, $upn) + 1
    Write-Progress -Activity "Disabling users" -Status "$upn ($i of $($upns.Count))" -PercentComplete (($i / $upns.Count) * 100)

    $result = [PSCustomObject]@{
        UPN            = $upn
        BeforeEnabled  = $null
        AfterEnabled   = $null
        Status         = "Success"
        Error          = ""
    }

    try {
        # Before state
        $user = Get-MgUser -UserId $upn -Property Id, UserPrincipalName, AccountEnabled -ErrorAction Stop
        $result.BeforeEnabled = $user.AccountEnabled

        # Disable
        Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop

        # After state (confirm)
        $confirm = Get-MgUser -UserId $user.Id -Property AccountEnabled -ErrorAction Stop
        $result.AfterEnabled = $confirm.AccountEnabled
    }
    catch {
        $result.Status = "Failed"
        $result.Error  = $_.Exception.Message
    }

    $result
}

Write-Progress -Activity "Disabling users" -Completed

# ---- Report ----
$reportPath = Join-Path -Path $PWD -ChildPath "DisableUsersReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$report | Export-Csv -Path $reportPath -NoTypeInformation

$failCount = ($report | Where-Object Status -eq "Failed").Count
Write-Host "`nDone. $($report.Count - $failCount) succeeded, $failCount failed." -ForegroundColor Cyan
Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan

$report | Format-Table -AutoSize