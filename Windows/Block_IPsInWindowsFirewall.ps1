<#
    .SYNOPSIS
    Script to block IP address ranges from a user-provided file using Windows Firewall.
    
    .DESCRIPTION
    This script prompts the user to select a file containing IP ranges in CIDR notation.
    It then iterates through each entry in the file and creates a Windows Firewall rule 
    to block the specified range for inbound connections to the server.

    .VERSION
    1.4

    .AUTHOR
    Steve Springall

    .NOTES
    Get domain IP lists from here: https://www.ipdeny.com/ipblocks/
    Firewall rules are stored in the registry:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
#>

# Display script title
Write-Host "===============================" -ForegroundColor Cyan
Write-Host " Windows Firewall Blocker Tool " -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

# Start time tracking
$startTime = Get-Date

# Initialize counter for rules added
$rulesAdded = 0

# Prompt the user to select the file
$filePath = Read-Host "Please provide the full path to the file containing IP ranges"

# Check if file exists
if (-Not (Test-Path $filePath)) {
    Write-Host "File not found. Please check the path and try again." -ForegroundColor Red
    exit
}

# Get just the filename without extension for easier filtering
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)

# Define the group name for firewall rules
$groupName = "GeoBlock $fileName"

# Import IP ranges from the file
$ipRanges = Get-Content -Path $filePath

# Loop through each IP range and create a firewall rule
foreach ($ipRange in $ipRanges) {
    # Trim any whitespace
    $ipRange = $ipRange.Trim()

    # Generate a unique name for the firewall rule including the filename
    $ruleName = "Block IP Range ($fileName) - $ipRange"

    # Add the firewall rule with a specified group and suppress the output
    try {
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Block -RemoteAddress $ipRange -Group $groupName -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Successfully blocked IP range: $ipRange" -ForegroundColor Green
        $rulesAdded++
    }
    catch {
        Write-Host "Failed to block IP range: $ipRange. Error: $_" -ForegroundColor Red
    }
}

# Calculate time taken
$endTime = Get-Date
$timeTaken = $endTime - $startTime

# Display completion message and stats
Write-Host ""
Write-Host "IP Blocking Completed!" -ForegroundColor Cyan
Write-Host "Time taken: $($timeTaken.ToString())" -ForegroundColor Cyan
Write-Host "Number of rules added: $rulesAdded" -ForegroundColor Cyan
Write-Host "===============================" -ForegroundColor Cyan
