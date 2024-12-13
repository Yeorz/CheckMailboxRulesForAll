Import-Module ExchangeOnlineManagement

Function Connect-EXOnline {
    Param ([switch]$ForceReconnect)
    $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($ForceReconnect -or -not $connection.ConnectionUri -like "*outlook.office365.com*") {
        Write-Output "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction Stop
    } else {
        Write-Output "Already connected to Exchange Online."
    }
}

Function Get-ExternalForwardingRules {
    Param ([PSObject]$mailbox, [Array]$domains)

    $rules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
    $rules | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo } | ForEach-Object {
        $externalRecipients = $_.ForwardTo + $_.ForwardAsAttachmentTo |
            Where-Object { ($_.ToString() -split "SMTP:")[1].Trim("]") -notmatch "^($domains.DomainName)$" }
        
        if ($externalRecipients.Count -gt 0) {
            [PSCustomObject]@{
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                DisplayName        = $mailbox.DisplayName
                RuleId             = $_.Identity
                RuleName           = $_.Name
                RuleDescription    = $_.Description
                ExternalRecipients = $externalRecipients -join ", "
            }
        }
    }
}

# Main Script Execution
Connect-EXOnline
$domains = Get-AcceptedDomain
$mailboxes = Get-Mailbox -ResultSize Unlimited

$results = foreach ($mailbox in $mailboxes) {
    Write-Host "Checking rules for $($mailbox.DisplayName)" -ForegroundColor Green
    Get-ExternalForwardingRules -mailbox $mailbox -domains $domains
}

$results | Export-Csv -Path "$HOME/Documents/externalrules.csv" -NoTypeInformation
Disconnect-ExchangeOnline -Confirm:$false
