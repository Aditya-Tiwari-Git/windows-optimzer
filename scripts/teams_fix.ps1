<#
.SYNOPSIS
Microsoft Teams Quick Fix Script for Windows System Optimizer
.DESCRIPTION
This script fixes common issues with Microsoft Teams by cleaning caches and resetting configuration.
#>

Write-Output "Microsoft Teams Quick Fix Script"
Write-Output "================================="

# Check if Teams is running and stop it
$teamsProcess = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
if ($teamsProcess) {
    Write-Output "Stopping Microsoft Teams processes..."
    Stop-Process -Name "Teams" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# Also stop Teams Updater if running
$updaterProcess = Get-Process -Name "Update" -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "*Teams*" }
if ($updaterProcess) {
    Write-Output "Stopping Teams Update process..."
    Stop-Process -Id $updaterProcess.Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# Get user profile path
$userProfile = $env:USERPROFILE
$teamsAppData = Join-Path -Path $userProfile -ChildPath "AppData\Roaming\Microsoft\Teams"
$teamsLocalAppData = Join-Path -Path $userProfile -ChildPath "AppData\Local\Microsoft\Teams"

# Check if Teams is installed
if (-not (Test-Path -Path $teamsAppData)) {
    Write-Output "Microsoft Teams does not appear to be installed for this user."
    exit
}

# Clear cache directories
$cacheDirs = @(
    "Cache",
    "blob_storage",
    "databases",
    "GPUCache",
    "IndexedDB",
    "Local Storage",
    "tmp"
)

foreach ($dir in $cacheDirs) {
    $dirPath = Join-Path -Path $teamsAppData -ChildPath $dir
    if (Test-Path -Path $dirPath) {
        Write-Output "Clearing $dir..."
        try {
            Get-ChildItem -Path $dirPath -File -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $dirPath -Directory -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Output "Successfully cleared $dir."
        }
        catch {
            Write-Output "Error clearing $dir: $_"
        }
    }
}

# Backup and reset configuration
$configFile = Join-Path -Path $teamsAppData -ChildPath "desktop-config.json"
if (Test-Path -Path $configFile) {
    try {
        $backupFile = "$configFile.bak"
        Write-Output "Backing up configuration to $backupFile..."
        Copy-Item -Path $configFile -Destination $backupFile -Force
        
        Write-Output "Removing configuration file..."
        Remove-Item -Path $configFile -Force
        Write-Output "Configuration file removed successfully. Teams will recreate it on next start."
    }
    catch {
        Write-Output "Error resetting configuration: $_"
    }
}

# Clear local storage cache
$localStoragePath = Join-Path -Path $teamsAppData -ChildPath "Local Storage\leveldb"
if (Test-Path -Path $localStoragePath) {
    Write-Output "Clearing local storage cache..."
    try {
        Get-ChildItem -Path $localStoragePath -File -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Output "Successfully cleared local storage cache."
    }
    catch {
        Write-Output "Error clearing local storage cache: $_"
    }
}

# Clear background gorilla directory (can cause crashes)
$gorillaPath = Join-Path -Path $userProfile -ChildPath "AppData\Local\Microsoft\TeamsMeetingAddin\Cookies"
if (Test-Path -Path $gorillaPath) {
    Write-Output "Clearing Teams Meeting Add-in cookies..."
    try {
        Get-ChildItem -Path $gorillaPath -File -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Output "Successfully cleared Teams Meeting Add-in cookies."
    }
    catch {
        Write-Output "Error clearing Teams Meeting Add-in cookies: $_"
    }
}

# Reset Teams settings in registry
try {
    $teamsRegPath = "HKCU:\Software\Microsoft\Teams"
    if (Test-Path -Path $teamsRegPath) {
        Write-Output "Clearing Teams registry settings..."
        Remove-Item -Path "$teamsRegPath\PersistedState" -Force -ErrorAction SilentlyContinue
        Write-Output "Teams registry settings cleared."
    }
}
catch {
    Write-Output "Error clearing Teams registry settings: $_"
}

Write-Output ""
Write-Output "Microsoft Teams has been reset successfully."
Write-Output "You can now restart Teams and it should work correctly."
Write-Output ""
