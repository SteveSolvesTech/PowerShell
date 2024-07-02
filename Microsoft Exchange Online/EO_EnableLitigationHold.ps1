<#
.SYNOPSIS
    Enable Litigation Hold on all mailboxes in an Exchange Online environment.
    
.DESCRIPTION
    This script connects to Exchange Online, retrieves all mailboxes, and attempts to enable Litigation Hold on each one.
    If a mailbox does not have the required license for Litigation Hold, it handles the error and continues with the next mailbox.
    After enabling Litigation Hold, the script loops through the mailboxes and lists those with Litigation Hold enabled.
    Successful results are displayed in a table format.
    
.AUTHOR
    Steve Springall
    
.VERSION
    1.6
#>

# Clear Host
Clear-Host

# Display script title
Write-Host "------------------------" -ForegroundColor Green
Write-Host " Enable Litigation Hold " -ForegroundColor Green
Write-Host "------------------------" -ForegroundColor Green
Write-Host ""

# Suppress error output to screen
$ErrorActionPreference = "SilentlyContinue"

# Check if ExchangeOnlineManagement module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
        Write-Host "ExchangeOnlineManagement module installed successfully." -ForegroundColor Green
        Write-Host ""
    } catch {
        Write-Host "Failed to install ExchangeOnlineManagement module. Exiting script." -ForegroundColor Red
        exit
    }
}

# Import the ExchangeOnlineManagement module
try {
    Import-Module ExchangeOnlineManagement
    Write-Host "ExchangeOnlineManagement module imported successfully." -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Failed to import ExchangeOnlineManagement module. Exiting script." -ForegroundColor Red
    exit
}

# Connect to Exchange Online
try {
    Connect-ExchangeOnline
    Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Failed to connect to Exchange Online. Exiting script." -ForegroundColor Red
    exit
}

# Retrieve all mailboxes
try {
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    Write-Host "Retrieved all mailboxes." -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Failed to retrieve mailboxes. Exiting script." -ForegroundColor Red
    exit
}

# Initialize an array to store successful results
$SuccessResults = @()

# Loop through each mailbox and enable Litigation Hold
foreach ($mailbox in $mailboxes) {
    # Skip mailboxes starting with "DiscoverySearchMailbox"
    if ($mailbox.UserPrincipalName -like "DiscoverySearchMailbox*") {
        Write-Host "Skipping mailbox: $($mailbox.UserPrincipalName)" -ForegroundColor Yellow
        continue
    }

    try {
        Set-Mailbox -Identity $mailbox.UserPrincipalName -LitigationHoldEnabled $true
        Write-Host "Attempting for mailbox: $($mailbox.UserPrincipalName)" -ForegroundColor Green
        $SuccessResults += [PSCustomObject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            DisplayName       = $mailbox.DisplayName
        }
    } catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*license doesn't permit*") {
            Write-Host "License error for mailbox: $($mailbox.UserPrincipalName). Litigation Hold not enabled." -ForegroundColor Yellow
        } elseif ($errorMessage -like "*RecipientTaskException*") {
            Write-Host "Recipient task error for mailbox: $($mailbox.UserPrincipalName). Error: $errorMessage" -ForegroundColor Red
        } else {
            Write-Host "Failed to enable Litigation Hold for mailbox: $($mailbox.UserPrincipalName). Error: $errorMessage" -ForegroundColor Red
        }
    }
}

# Loop through mailboxes to list those with Litigation Hold enabled
$LitigationHoldEnabledMailboxes = @()
foreach ($mailbox in $mailboxes) {
    # Skip mailboxes starting with "DiscoverySearchMailbox"
    if ($mailbox.UserPrincipalName -like "DiscoverySearchMailbox*") {
        continue
    }

    $mailboxDetails = Get-Mailbox -Identity $mailbox.UserPrincipalName
    if ($mailboxDetails.LitigationHoldEnabled -eq $true) {
        $LitigationHoldEnabledMailboxes += [PSCustomObject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            DisplayName       = $mailbox.DisplayName
        }
    }
}

# Display successful results in a table format
Write-Host ""
Write-Host ""

if ($LitigationHoldEnabledMailboxes.Count -gt 0) {
    Write-Host "The following mailboxes have Litigation Hold enabled:" -ForegroundColor DarkYellow
    $LitigationHoldEnabledMailboxes | Format-Table -Property UserPrincipalName, DisplayName -AutoSize
} else {
    Write-Host "No mailboxes have Litigation Hold enabled." -ForegroundColor Yellow
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
