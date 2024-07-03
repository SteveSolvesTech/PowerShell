<#
.SYNOPSIS
    This script searches for .ost and .pst files on the computer and reports their location and size.

.DESCRIPTION
    This PowerShell script searches the entire computer for .ost and .pst files, calculates their size, 
    and reports the file path and size. The user is prompted for the output folder, 
    defaulting to "C:\Temp" if none is specified. If the folder does not exist, it is created.

.NOTES
    Author: Steve Springall
    Version: 1.9

#>

# Clear the console
Clear-Host

# Display a pretty title
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "  Search for .ost and .pst Files on Computer  " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# Prompt for the output folder, defaulting to C:\Temp
$outputFolder = Read-Host "Please enter the output folder (default is C:\Temp)" 
if ([string]::IsNullOrEmpty($outputFolder)) {
    $outputFolder = "C:\Temp"
}

# Create the output folder if it doesn't exist
if (-not (Test-Path -Path $outputFolder)) {
    Write-Host "Output folder does not exist. Creating $outputFolder..." -ForegroundColor Yellow
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}
Write-Host ""

# Define the file types to search for
$fileTypes = @("*.ost", "*.pst")

# Initialize a global array to hold file information
$fileList = @()

# Function to search for files and gather information
function Get-FileInformation {
    param (
        [string]$path,
        [string]$pattern
    )
    
    Write-Host "Searching for $pattern files in $path..." -ForegroundColor Yellow

    $files = Get-ChildItem -Path $path -Recurse -Filter $pattern -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        Write-Host "Found file: $($file.FullName)" -ForegroundColor Cyan
        $fileSizeGB = [math]::Round(($file.Length / 1GB), 2)
        $global:fileList += [PSCustomObject]@{
            FileName = $file.Name
            SizeGB   = "{0:N2}" -f $fileSizeGB
            FilePath = $file.FullName
        }
    }
}

# Search for files in all drives
$drives = Get-PSDrive -PSProvider FileSystem
foreach ($drive in $drives) {
    Write-Host "Searching drive: $($drive.Root)" -ForegroundColor Magenta
    foreach ($fileType in $fileTypes) {
        Get-FileInformation -path $drive.Root -pattern $fileType
    }
    Write-Host ""
}

# Search specifically in the Outlook folder for the current user
$userProfile = [Environment]::GetFolderPath('UserProfile')
$outlookPath = "$userProfile\AppData\Local\Microsoft\Outlook"
Write-Host "Searching in Outlook path: $outlookPath" -ForegroundColor Magenta
foreach ($fileType in $fileTypes) {
    Get-FileInformation -path $outlookPath -pattern $fileType
}

# Check if fileList is populated
if ($fileList.Count -eq 0) {
    Write-Host "No files found." -ForegroundColor Red
} else {
    Write-Host "Total files found: $($fileList.Count)" -ForegroundColor Green
    $fileList | Sort-Object -Property SizeGB -Descending | Format-Table -Property FileName, SizeGB, FilePath #-AutoSize
}

# Create the output file name with computer name and timestamp
$computerName = $env:COMPUTERNAME
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$outputFile = Join-Path -Path $outputFolder -ChildPath "${computerName}_file_search_results_$timestamp.csv"

# Save the results to a CSV file
Write-Host "Saving results to $outputFile..." -ForegroundColor Green

# Verify the data before exporting
if ($fileList.Count -gt 0) {
    $fileList | Export-Csv -Path $outputFile -NoTypeInformation
    # Verify the CSV file content
    if ((Get-Content $outputFile).Length -eq 0) {
        Write-Host "Warning: The CSV file is empty!" -ForegroundColor Red
    } else {
        Write-Host "Results have been saved to $outputFile" -ForegroundColor Green
    }
} else {
    Write-Host "No data to export." -ForegroundColor Red
}
