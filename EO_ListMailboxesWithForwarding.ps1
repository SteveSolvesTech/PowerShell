# EO_ListMailboxesWithForwarding.ps1

<#
.SYNOPSIS
    Script to list mailboxes with forwarding enabled in Exchange Online and export the results to a CSV file.

.DESCRIPTION
    This script connects to Exchange Online and lists all mailboxes with forwarding enabled.
    The results are exported to a CSV file. The user is prompted for an output path, with "C:\Temp" as the default.
    The script checks for the required ExchangeOnlineManagement module and installs it if needed.

.NOTES
    Author: Steve Springall
    Date: 26/06/2024
	Version: 1.0
#>

# Clear Host
Clear-Host

# Display script title
Write-Host "------------------------" -ForegroundColor Green
Write-Host " Mail Forwarding Report " -ForegroundColor Green
Write-Host "------------------------" -ForegroundColor Green
Write-Host ""

# Function to check and install prerequisites
function Install-ExchangeOnlineManagement {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found. Installing..."
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    else {
        #Write-Host "ExchangeOnlineManagement module is already installed."
    }
}

# Import the module
Write-Host "Importing module ExchangeOnlineManagement." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement
Write-Host ""

# Connect to Exchange Online
Write-Host "Authentication. Please sign in as a Global Administrator with your browser. " -ForegroundColor White -BackgroundColor DarkYellow
Connect-ExchangeOnline

# Get user mailboxes with forwarding enabled
$userMailboxesWithForwarding = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox |
    Select Name, RecipientType, UserPrincipalName, ForwardingSmtpAddress, DeliverToMailboxAndForward,
    @{Name="IsMailboxEnabled";Expression={($_.ArchiveStatus -ne 'None')}} |
    Where-Object { $_.DeliverToMailboxAndForward -eq $true }

# Get shared mailboxes with forwarding enabled
$sharedMailboxesWithForwarding = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox |
    Select Name, RecipientType, UserPrincipalName, ForwardingSmtpAddress, DeliverToMailboxAndForward,
    @{Name="IsMailboxEnabled";Expression={($_.ArchiveStatus -ne 'None')}} |
    Where-Object { $_.DeliverToMailboxAndForward -eq $true }

# Join the results
$mailboxesWithForwarding = $($userMailboxesWithForwarding; $sharedMailboxesWithForwarding) 

# Display the results
$mailboxesWithForwarding | Format-Table
Write-Host ""

# Prompt user for output path
$outputPath = Read-Host -Prompt "Enter the output path for the CSV file (default is C:\Temp)"
if ([string]::IsNullOrEmpty($outputPath)) {
    $outputPath = "C:\Temp"
}

# Ensure the path exists
if (-not (Test-Path -Path $outputPath)) {
    Write-Host "The specified path does not exist. Creating path..."-ForegroundColor White -BackgroundColor DarkYellow
    New-Item -Path $outputPath -ItemType Directory -Force
}

# Generate a timestamp for the output file
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path -Path $outputPath -ChildPath "MailboxesWithForwarding_$timestamp.csv"

# Export the results to a CSV file
$mailboxesWithForwarding | Export-Csv -Path $outputFile -NoTypeInformation


# Provide feedback to the user
Write-Host ""
Write-Host "The results have been successfully exported to $outputFile"-ForegroundColor White -BackgroundColor DarkGreen
Write-Host ""