<#
    .SYNOPSIS
    Script to block all remote IP address ranges except those from a user-provided file using Windows Firewall.
    
    .DESCRIPTION
    This script prompts the user to select a file containing IP ranges in CIDR notation.
    It then iterates through each entry in the file and creates a Windows Firewall rule 
    to allow the specified range for inbound connections to the server. Once the allow rules are created, all 
	other remote connections will be blocked. Existing allow rule are not affected.

    .VERSION
    1.7

    .AUTHOR
    Steve Springall

    .NOTES
    
	Get domain IP lists from here: https://www.ipdeny.com/ipblocks/
    Prefix the file you wish to allow with "allow-" and place in the same directory as the script.
	
	Firewall rules are stored in the registry:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
	
#>

# Display script title
Write-Host "===============================" -ForegroundColor Cyan
Write-Host " Windows Firewall Blocker Tool " -ForegroundColor Yellow
Write-Host "         Allow List            " -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

# Start time tracking
$startTime = Get-Date

# Initialize counter for rules added
$rulesAdded = 0

# Get the directory of the script
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Look for files starting with "allow-" in the script directory
$allowFile = Get-ChildItem -Path $scriptDirectory -Filter "allow-*" | Select-Object -First 1

if ($allowFile) {
    # If a file is found, use it
    $filePath = $allowFile.FullName
    Write-Host "Using file: $filePath" -ForegroundColor Green
} else {
    # If no file is found, prompt the user to select a file
    $filePath = Read-Host "Please provide the full path to the file containing IP ranges to allow."

    # Check if the provided file exists
    if (-Not (Test-Path $filePath)) {
        Write-Host "File not found. Please check the path and try again." -ForegroundColor Red
        exit
    }
}

# Get just the filename without extension for easier filtering
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)

# Define the group name for firewall rules
$groupName = "GeoBlock Allow $fileName"

# Import IP ranges from the file
$ipRanges = Get-Content -Path $filePath

# Loop through each IP range and create a firewall rule
foreach ($ipRange in $ipRanges) {
    # Trim any whitespace
    $ipRange = $ipRange.Trim()

    # Generate a unique name for the firewall rule including the filename
    $ruleName = "Allow IP Range ($fileName) - $ipRange"

    # Add the firewall rule with a specified group and suppress the output
    try {
		New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -RemoteAddress $ipRange -Group $groupName -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Successfully allowed IP range: $ipRange" -ForegroundColor Green
        $rulesAdded++
    }
    catch {
        Write-Host "Failed to block IP range: $ipRange. Error: $_" -ForegroundColor Red
    }
}

# Block all inbound connections from remote networks
$ruleName = "Block All Remote Inbound"

# Add the firewall rule with a specified group and suppress the output
try {
	New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Block -RemoteAddress "any" -Group $groupName -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Successfully blocked all remote IP ranges" -ForegroundColor Green
    $rulesAdded++
}
catch {
    Write-Host "Failed to block all remote IP ranges: $_" -ForegroundColor Red
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
