<#
.SYNOPSIS
    Exports Exchange Online mailbox information to a CSV file.

.DESCRIPTION
    Connects to Exchange Online, lets you choose which mailbox type to report on
    (User, Shared, or All), and exports the following fields per mailbox:

        UserPrincipalName
        Type
        StorageConsumed
        StorageConsumedBytes
        StorageLimit
        StorageLimitBytes
        StoragePercentUsed
        ArchiveStatus
        AutoExpandingArchiveEnabled
        ArchiveStorageConsumed
        ArchiveStorageConsumedBytes
        ArchiveStorageQuota
        ArchiveStorageQuotaBytes
        ArchiveStoragePercentUsed
        RetentionPolicy
        ForwardingSMTPAddress
        ForwardingAddress

.PARAMETER MailboxType
    Which mailboxes to report on: 'User', 'Shared', or 'All'. If omitted, you'll be
    prompted interactively.

.PARAMETER OutputPath
    Path to the CSV file to create. Defaults to a timestamped file in the current
    directory.

.PARAMETER UserPrincipalName
    UPN to use when connecting to Exchange Online. If omitted, Connect-ExchangeOnline
    will prompt for auth (modern auth / MFA).

.EXAMPLE
    .\Export-EXOMailboxReport.ps1 -MailboxType All -OutputPath C:\Reports\mailboxes.csv

.EXAMPLE
    .\Export-EXOMailboxReport.ps1
    (Prompts for mailbox type and uses default output path)

.NOTES
    Requires the ExchangeOnlineManagement (EXO V3) module:
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
    Requires a role with mailbox/statistics read access (e.g. View-Only Recipients +
    View-Only Configuration, or Global Reader / Exchange Admin).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('User', 'Shared', 'All')]
    [string]$MailboxType,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName
)

#region Setup

$ErrorActionPreference = 'Stop'

function Write-Status {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host $Message -ForegroundColor $Color
}

# Pulls the raw byte count out of a ByteQuantifiedSize-type value (mailbox size /
# quota objects). Returns $null for unlimited/unset quotas or unparsable values.
function Get-BytesFromQuota {
    param($QuotaValue)

    if (-not $QuotaValue) { return $null }

    $text = $QuotaValue.ToString()
    if ($text -match 'Unlimited') { return $null }

    # Format is like "1.234 GB (1,325,000,000 bytes)" - pull the raw byte count
    if ($text -match '\(([\d,]+)\s*bytes\)') {
        return [int64]($matches[1] -replace ',', '')
    }

    return $null
}

function Get-PercentUsed {
    param($ConsumedBytes, $LimitBytes)

    if (-not $ConsumedBytes -or -not $LimitBytes -or $LimitBytes -eq 0) { return $null }
    return [math]::Round((($ConsumedBytes / $LimitBytes) * 100), 1)
}

# Ensure the EXO module is present
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Status "The ExchangeOnlineManagement module is not installed." -Color Yellow
    $install = Read-Host "Install it now for the current user? (Y/N)"
    if ($install -match '^[Yy]') {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    }
    else {
        Write-Status "Cannot continue without the ExchangeOnlineManagement module. Exiting." -Color Red
        return
    }
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

#endregion

#region Connect

Write-Status "Connecting to Exchange Online..."
try {
    if ($UserPrincipalName) {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}
catch {
    Write-Status "Failed to connect to Exchange Online: $($_.Exception.Message)" -Color Red
    return
}

#endregion

#region Choose mailbox type

if (-not $MailboxType) {
    Write-Host ""
    Write-Host "Which mailboxes do you want to export?"
    Write-Host "  1) User mailboxes"
    Write-Host "  2) Shared mailboxes"
    Write-Host "  3) All mailboxes"
    $choice = Read-Host "Enter 1, 2, or 3"

    switch ($choice) {
        '1' { $MailboxType = 'User' }
        '2' { $MailboxType = 'Shared' }
        '3' { $MailboxType = 'All' }
        default {
            Write-Status "Invalid selection. Exiting." -Color Red
            Disconnect-ExchangeOnline -Confirm:$false | Out-Null
            return
        }
    }
}

$recipientTypeDetails = switch ($MailboxType) {
    'User'   { @('UserMailbox') }
    'Shared' { @('SharedMailbox') }
    'All'    { @('UserMailbox', 'SharedMailbox') }
}

#endregion

#region Output path

if (-not $OutputPath) {
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath "EXOMailboxReport_${MailboxType}_$timestamp.csv"
}

#endregion

#region Gather mailboxes

Write-Status "Retrieving $MailboxType mailbox list..."

try {
    $mailboxes = Get-EXOMailbox -RecipientTypeDetails $recipientTypeDetails -ResultSize Unlimited `
        -Properties ArchiveStatus, AutoExpandingArchiveEnabled, RetentionPolicy, `
                    ForwardingSmtpAddress, ForwardingAddress, ProhibitSendQuota
}
catch {
    Write-Status "Failed to retrieve mailboxes: $($_.Exception.Message)" -Color Red
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    return
}

if (-not $mailboxes -or $mailboxes.Count -eq 0) {
    Write-Status "No mailboxes found for type '$MailboxType'." -Color Yellow
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    return
}

Write-Status "Found $($mailboxes.Count) mailbox(es). Gathering statistics (this may take a while)..."

#endregion

#region Build report

$report = New-Object System.Collections.Generic.List[Object]
$total  = $mailboxes.Count
$i      = 0

foreach ($mbx in $mailboxes) {

    $i++
    Write-Progress -Activity "Exporting mailbox report" `
        -Status "$i of $total : $($mbx.UserPrincipalName)" `
        -PercentComplete (($i / $total) * 100)

    # --- Primary mailbox statistics (size) ---
    $storageConsumed      = $null
    $storageConsumedBytes = $null
    try {
        $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Properties TotalItemSize
        if ($stats) {
            $storageConsumed      = $stats.TotalItemSize.ToString()
            $storageConsumedBytes = Get-BytesFromQuota -QuotaValue $stats.TotalItemSize
        }
    }
    catch {
        Write-Warning "Could not get mailbox statistics for $($mbx.UserPrincipalName): $($_.Exception.Message)"
    }

    $storageLimitBytes = Get-BytesFromQuota -QuotaValue $mbx.ProhibitSendQuota
    $storagePercentUsed = Get-PercentUsed -ConsumedBytes $storageConsumedBytes -LimitBytes $storageLimitBytes

    # --- Archive mailbox statistics (only if an archive exists) ---
    $archiveStorageConsumed      = $null
    $archiveStorageConsumedBytes = $null
    $archiveStorageQuota         = $null
    $archiveStorageQuotaBytes    = $null
    $archiveStoragePercentUsed   = $null

    if ($mbx.ArchiveStatus -and $mbx.ArchiveStatus -ne 'None') {
        try {
            $archiveStats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Archive -Properties TotalItemSize
            if ($archiveStats) {
                $archiveStorageConsumed      = $archiveStats.TotalItemSize.ToString()
                $archiveStorageConsumedBytes = Get-BytesFromQuota -QuotaValue $archiveStats.TotalItemSize
            }
        }
        catch {
            Write-Warning "Could not get archive statistics for $($mbx.UserPrincipalName): $($_.Exception.Message)"
        }

        try {
            # Archive quota lives on the full Get-Mailbox object
            $mbxDetail = Get-Mailbox -Identity $mbx.UserPrincipalName
            $archiveStorageQuota      = $mbxDetail.ArchiveQuota
            $archiveStorageQuotaBytes = Get-BytesFromQuota -QuotaValue $mbxDetail.ArchiveQuota
        }
        catch {
            Write-Warning "Could not get archive quota for $($mbx.UserPrincipalName): $($_.Exception.Message)"
        }

        $archiveStoragePercentUsed = Get-PercentUsed -ConsumedBytes $archiveStorageConsumedBytes -LimitBytes $archiveStorageQuotaBytes
    }

    # --- Forwarding address (resolve ForwardingAddress GUID/recipient to a readable value) ---
    $forwardingAddressDisplay = $null
    if ($mbx.ForwardingAddress) {
        $forwardingAddressDisplay = $mbx.ForwardingAddress.ToString()
    }

    $report.Add([PSCustomObject]@{
        UserPrincipalName           = $mbx.UserPrincipalName
        Type                        = $mbx.RecipientTypeDetails
        StorageConsumed             = $storageConsumed
        StorageConsumedBytes        = $storageConsumedBytes
        StorageLimit                = $mbx.ProhibitSendQuota
        StorageLimitBytes           = $storageLimitBytes
        StoragePercentUsed          = $storagePercentUsed
        ArchiveStatus               = $mbx.ArchiveStatus
        AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
        ArchiveStorageConsumed      = $archiveStorageConsumed
        ArchiveStorageConsumedBytes = $archiveStorageConsumedBytes
        ArchiveStorageQuota         = $archiveStorageQuota
        ArchiveStorageQuotaBytes    = $archiveStorageQuotaBytes
        ArchiveStoragePercentUsed   = $archiveStoragePercentUsed
        RetentionPolicy             = $mbx.RetentionPolicy
        ForwardingSMTPAddress       = $mbx.ForwardingSmtpAddress
        ForwardingAddress           = $forwardingAddressDisplay
    })
}

Write-Progress -Activity "Exporting mailbox report" -Completed

#endregion

#region Export

try {
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Status "Report exported successfully to: $OutputPath" -Color Green
}
catch {
    Write-Status "Failed to write CSV: $($_.Exception.Message)" -Color Red
}

#endregion
