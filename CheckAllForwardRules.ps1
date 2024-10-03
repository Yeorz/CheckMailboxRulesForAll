# Import the ExchangeOnlineManagement module
Import-Module ExchangeOnlineManagement

# Function to connect to Exchange Online only if not already connected
Function Connect-EXOnline {
    # Check if the Exchange Online module is already imported and if there's an active session
    $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($connection.ConnectionUri -like "*outlook.office365.com*") {
        Write-Output "Already connected to Exchange Online. No need to re-authenticate."
    } else {
        Write-Output "Connecting to Exchange Online..."
        # Replace with your desired connection method. This example uses interactive authentication.
        Connect-ExchangeOnline
    }
}

# Establish the connection to Exchange Online
Connect-EXOnline

# Get the accepted domains in the organization
$domains = Get-AcceptedDomain
# Retrieve all mailboxes in the organization
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Set the path to the Documents folder for macOS
$exportPath = "$HOME/Documents/externalrules.csv"

# Iterate through each mailbox to check for forwarding rules
foreach ($mailbox in $mailboxes) {
    $forwardingRules = $null
    Write-Host "Checking rules for $($mailbox.DisplayName) - $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green

    # Retrieve the inbox rules for each mailbox
    $rules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress

    # Filter the rules to find those that forward emails
    $forwardingRules = $rules | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo }

    foreach ($rule in $forwardingRules) {
        $recipients = @()
        $recipients = $rule.ForwardTo | Where-Object { $_ -match "SMTP" }
        $recipients += $rule.ForwardAsAttachmentTo | Where-Object { $_ -match "SMTP" }

        $externalRecipients = @()

        foreach ($recipient in $recipients) {
            $email = ($recipient -split "SMTP:")[1].Trim("]")
            $domain = ($email -split "@")[1]

            if ($domains.DomainName -notcontains $domain) {
                $externalRecipients += $email
            }
        }

        if ($externalRecipients) {
            $extRecString = $externalRecipients -join ", "
            Write-Host "$($rule.Name) forwards to $extRecString" -ForegroundColor Yellow

            # Create a new PSObject to store rule information
            $ruleHash = [ordered]@{
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                DisplayName        = $mailbox.DisplayName
                RuleId             = $rule.Identity
                RuleName           = $rule.Name
                RuleDescription    = $rule.Description
                ExternalRecipients = $extRecString
            }
            $ruleObject = New-Object PSObject -Property $ruleHash

            # Export the rule information to the Documents folder on macOS
            $ruleObject | Export-Csv -Path $exportPath -NoTypeInformation -Append
        }
    }
}

# Disconnect the session if you authenticated in this session
if ($connection.ConnectionUri -notlike "*outlook.office365.com*") {
    Disconnect-ExchangeOnline -Confirm:$false
}
