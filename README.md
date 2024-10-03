# Exchange Online Forwarding Rules Check Script

## Overview

This PowerShell script connects to Exchange Online and retrieves mailbox forwarding rules for all user mailboxes within the organization. It identifies and reports any inbox rules that automatically forward emails to external recipients, exporting this information to a CSV file. This helps administrators maintain security and compliance by monitoring for unauthorized email forwarding configurations.

## Features

- Modern Authentication: Uses OAuth 2.0 with the Connect-ExchangeOnline cmdlet for secure and modern authentication.
- Session Check: Automatically detects if there's an existing connection to Exchange Online to avoid unnecessary re-authentication.
- Forwarding Rule Analysis: Scans all user mailboxes and identifies inbox rules that forward or redirect messages.
- External Recipient Detection: Checks if the forwarding addresses are external to the organization based on the domain names defined in the accepted domains.
- CSV Export: Exports details of detected external forwarding rules to a CSV file located in the user's Documents folder on macOS.

## Requirements

- PowerShell Core: This script is designed to run on PowerShell Core (e.g., pwsh), which is compatible with macOS, Linux, and Windows.

- ExchangeOnlineManagement Module: The script requires the ExchangeOnlineManagement module to be installed. Install it using:


   ``` Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser```

- Exchange Online Permissions: The executing user must have appropriate permissions to retrieve mailbox rules and mailbox details in Exchange Online.

## Installation

    Clone or download this repository to your local system.
    Make sure PowerShell Core (pwsh) is installed and available on your system.
    Install the ExchangeOnlineManagement module if not already installed.

## Usage

Open a terminal and start PowerShell Core by typing:

**bash**

**pwsh**

Navigate to the directory where the script is located.

Run the script using:

    ./Check-ForwardingRules.ps1

The script will connect to Exchange Online and start checking all user mailboxes for forwarding rules. If it detects external forwarding, it will output the results to a CSV file named externalrules.csv in the user's Documents folder.

## Script Flow

    Connection Check: The script first checks if there's an existing Exchange Online session. If connected, it proceeds; otherwise, it initiates a new connection using modern authentication.
    Mailbox Retrieval: Retrieves a list of all mailboxes in the organization.
    Forwarding Rule Analysis: For each mailbox, the script retrieves inbox rules and checks for any rules that forward or redirect messages to external addresses.
    CSV Export: Any external forwarding rules found are exported to a CSV file.

## Output

The script outputs a CSV file named externalrules.csv in the ~/Documents folder. This file contains the following columns:

    PrimarySmtpAddress: The primary SMTP address of the mailbox.
    DisplayName: The display name of the mailbox user.
    RuleId: The unique identifier of the forwarding rule.
    RuleName: The name of the forwarding rule.
    RuleDescription: The description of the rule, if available.
    ExternalRecipients: The list of external email addresses the rule forwards to.

## License

This project is licensed under the GNU General Public License v3.0. See the LICENSE file for details.
Contributing

If you would like to contribute to this project, please fork the repository and submit a pull request with your changes. Feel free to open issues for any bugs or feature requests.
Disclaimer

This script is provided as-is without any warranties. Use at your own risk. Make sure to review the script and test it in a non-production environment before deploying it in production.

