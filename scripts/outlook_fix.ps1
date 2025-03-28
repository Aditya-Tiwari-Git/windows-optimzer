<#
.SYNOPSIS
Microsoft Outlook Quick Fix Script for Windows System Optimizer
.DESCRIPTION
This script fixes common issues with Microsoft Outlook by resetting profile settings and clearing caches.
#>

Write-Output "Microsoft Outlook Quick Fix Script"
Write-Output "=================================="

# Check if Outlook is running and stop it
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Output "Stopping Microsoft Outlook process..."
    Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# Get user profile path
$userProfile = $env:USERPROFILE
$outlookAppData = Join-Path -Path $userProfile -ChildPath "AppData\Roaming\Microsoft\Outlook"
$outlookLocalAppData = Join-Path -Path $userProfile -ChildPath "AppData\Local\Microsoft\Outlook"

# Check if Outlook is installed
if (-not (Test-Path -Path $outlookAppData)) {
    Write-Output "Microsoft Outlook does not appear to be installed for this user."
    exit
}

# Backup Outlook profile registry settings
Write-Output "Backing up Outlook registry settings..."
$backupDir = Join-Path -Path $env:TEMP -ChildPath "OutlookBackup"
if (-not (Test-Path -Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
}

$regBackupFile = Join-Path -Path $backupDir -ChildPath "outlook_profiles_backup.reg"
$regExportCommand = "reg export `"HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles`" `"$regBackupFile`" /y"
Invoke-Expression $regExportCommand | Out-Null

if (Test-Path -Path $regBackupFile) {
    Write-Output "Registry settings backed up to $regBackupFile"
}
else {
    Write-Output "Warning: Failed to back up registry settings."
}

# Fix 1: Clear Navigation Pane settings
try {
    $navPanePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
    if (Test-Path -Path $navPanePath) {
        Write-Output "Resetting Outlook Navigation Pane settings..."
        Remove-ItemProperty -Path $navPanePath -Name "NavigationPaneViewState" -ErrorAction SilentlyContinue
        Write-Output "Navigation Pane settings reset."
    }
    
    # Also try for Office 15.0 (Outlook 2013)
    $navPanePath2013 = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Preferences"
    if (Test-Path -Path $navPanePath2013) {
        Remove-ItemProperty -Path $navPanePath2013 -Name "NavigationPaneViewState" -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Output "Error resetting Navigation Pane settings: $_"
}

# Fix 2: Reset Outlook views
try {
    $viewsPath = Join-Path -Path $outlookAppData -ChildPath "*.xml"
    Write-Output "Resetting Outlook views..."
    Get-Item -Path $viewsPath -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "outcmd*.xml" -or $_.Name -like "OutlookView*.xml" } | ForEach-Object {
        Rename-Item -Path $_.FullName -NewName "$($_.Name).bak" -Force -ErrorAction SilentlyContinue
    }
    Write-Output "Outlook views reset."
}
catch {
    Write-Output "Error resetting Outlook views: $_"
}

# Fix 3: Reset Outlook roaming signatures
try {
    $signaturesPath = Join-Path -Path $outlookAppData -ChildPath "Signatures"
    if (Test-Path -Path $signaturesPath) {
        Write-Output "Backing up Outlook signatures..."
        $signaturesBackupPath = Join-Path -Path $backupDir -ChildPath "Signatures"
        if (Test-Path -Path $signaturesBackupPath) {
            Remove-Item -Path $signaturesBackupPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        Copy-Item -Path $signaturesPath -Destination $signaturesBackupPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Signatures backed up to $signaturesBackupPath"
    }
}
catch {
    Write-Output "Error backing up signatures: $_"
}

# Fix 4: Clear Outlook auto-complete cache
try {
    $autocompleteFile = Join-Path -Path $outlookAppData -ChildPath "RoamCache\Stream_Autocomplete*.dat"
    $autocompleteItems = Get-Item -Path $autocompleteFile -ErrorAction SilentlyContinue
    if ($autocompleteItems) {
        Write-Output "Clearing Outlook auto-complete cache..."
        foreach ($item in $autocompleteItems) {
            Rename-Item -Path $item.FullName -NewName "$($item.Name).bak" -Force -ErrorAction SilentlyContinue
        }
        Write-Output "Auto-complete cache cleared."
    }
}
catch {
    Write-Output "Error clearing auto-complete cache: $_"
}

# Fix 5: Clear Outlook temporary files
try {
    $tempOutlookPath = Join-Path -Path $env:TEMP -ChildPath "Outlook Temp"
    if (Test-Path -Path $tempOutlookPath) {
        Write-Output "Clearing Outlook temporary files..."
        Remove-Item -Path $tempOutlookPath\* -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Temporary files cleared."
    }
}
catch {
    Write-Output "Error clearing temporary files: $_"
}

# Fix 6: Reset offline cache settings
try {
    $cachePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Cached Mode"
    if (Test-Path -Path $cachePath) {
        Write-Output "Resetting Outlook offline cache settings..."
        Set-ItemProperty -Path $cachePath -Name "Enable" -Value 1 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $cachePath -Name "Sync Settings" -Value 1 -Type DWord -ErrorAction SilentlyContinue
        Write-Output "Offline cache settings reset."
    }
    
    # Also try for Office 15.0 (Outlook 2013)
    $cachePath2013 = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Cached Mode"
    if (Test-Path -Path $cachePath2013) {
        Set-ItemProperty -Path $cachePath2013 -Name "Enable" -Value 1 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $cachePath2013 -Name "Sync Settings" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Output "Error resetting offline cache settings: $_"
}

# Fix 7: Fix corrupt profile settings
try {
    $profilesPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
    if (Test-Path -Path $profilesPath) {
        Write-Output "Checking for corrupt Outlook profile settings..."
        
        $profiles = Get-ChildItem -Path $profilesPath
        foreach ($profile in $profiles) {
            $profileName = Split-Path -Path $profile.PSPath -Leaf
            $corruptKeyPath = Join-Path -Path $profile.PSPath -ChildPath "9375CFF0413111d3B88A00104B2A6676"
            
            if (Test-Path -Path $corruptKeyPath) {
                Write-Output "Removing corrupt settings in profile: $profileName"
                Remove-Item -Path $corruptKeyPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Output "Corrupt settings removed."
            }
        }
    }
}
catch {
    Write-Output "Error fixing corrupt profile settings: $_"
}

# Fix 8: Reset Search Index for Outlook
try {
    Write-Output "Rebuilding search index for Outlook..."
    $searchPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Search"
    if (-not (Test-Path -Path $searchPath)) {
        New-Item -Path $searchPath -Force | Out-Null
    }
    New-ItemProperty -Path $searchPath -Name "ResetSearchCriteria" -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Output "Search index rebuild initialized."
}
catch {
    Write-Output "Error resetting search index: $_"
}

# Fix 9: Reset Outlook safe mode flag
try {
    $bootPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security"
    if (Test-Path -Path $bootPath) {
        Remove-ItemProperty -Path $bootPath -Name "OutlookBootMode" -ErrorAction SilentlyContinue
    }
    
    # Also try for Office 15.0 (Outlook 2013)
    $bootPath2013 = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security"
    if (Test-Path -Path $bootPath2013) {
        Remove-ItemProperty -Path $bootPath2013 -Name "OutlookBootMode" -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Output "Error resetting safe mode flag: $_"
}

Write-Output ""
Write-Output "Microsoft Outlook has been reset successfully."
Write-Output "You can now restart Outlook and it should work correctly."
Write-Output ""
