# PowerShell Uninstaller Script for FastMail Search Classic Outlook Add-In
[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$DeleteData,
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\FastMailSearch"
)

$ErrorActionPreference = "Continue"

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host " Uninstalling FastMail Search Outlook Add-In..." -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

# 1. Detect if Classic Outlook is currently running
$outlookProcesses = Get-Process outlook -ErrorAction SilentlyContinue
if ($outlookProcesses) {
    if ($Force) {
        Write-Host "Closing running Classic Outlook process..." -ForegroundColor Yellow
        Stop-Process -Name outlook -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    } else {
        Write-Host "`nWARNING: Classic Outlook is currently running." -ForegroundColor Yellow
        $response = Read-Host "Would you like to close Outlook now to complete uninstallation? (Y/N)"
        if ($response -eq "Y" -or $response -eq "y") {
            Stop-Process -Name outlook -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        } else {
            Write-Error "Uninstallation cancelled: Outlook must be closed."
            exit 1
        }
    }
}

$progId = "FastMailSearch.AddIn"
$clsid = "{4A10F4D0-7E55-442C-B23B-983F78BA48D2}"
$smokeProgId = "FastMailSearch.SmokeTest"
$smokeClsid = "{9B1C2A3D-4E5F-6789-0A1B-2C3D4E5F6789}"

# 2. Remove Outlook Addins key
$outlookAddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
if (Test-Path $outlookAddinKey) {
    Write-Host "Removing Outlook Addin Registry Key: $outlookAddinKey..." -ForegroundColor Yellow
    Remove-Item -Path $outlookAddinKey -Recurse -Force | Out-Null
}

# 3. Remove COM ProgID key
$progIdKey = "HKCU:\Software\Classes\$progId"
if (Test-Path $progIdKey) {
    Write-Host "Removing ProgID Registry Key: $progIdKey..." -ForegroundColor Yellow
    Remove-Item -Path $progIdKey -Recurse -Force | Out-Null
}

# 4. Remove COM CLSID key
$clsidKey = "HKCU:\Software\Classes\CLSID\$clsid"
if (Test-Path $clsidKey) {
    Write-Host "Removing CLSID Registry Key: $clsidKey..." -ForegroundColor Yellow
    Remove-Item -Path $clsidKey -Recurse -Force | Out-Null
}

# 5. Remove Smoke Test keys
$smokeProgIdKey = "HKCU:\Software\Classes\$smokeProgId"
$smokeClsidKey = "HKCU:\Software\Classes\CLSID\$smokeClsid"
if (Test-Path $smokeProgIdKey) { Remove-Item -Path $smokeProgIdKey -Recurse -Force | Out-Null }
if (Test-Path $smokeClsidKey) { Remove-Item -Path $smokeClsidKey -Recurse -Force | Out-Null }

# 6. Remove Resiliency Entry
$resiliencyKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"
if (Test-Path $resiliencyKey) {
    Remove-ItemProperty -Path $resiliencyKey -Name $progId -ErrorAction SilentlyContinue
}

# 7. Remove Installed Application Files
if (Test-Path $InstallDir) {
    Write-Host "Removing installed binaries from $InstallDir..." -ForegroundColor Yellow
    try {
        Remove-Item -Path $InstallDir -Recurse -Force | Out-Null
    } catch {
        Write-Host "Notice: Some binary files will be removed on next system restart." -ForegroundColor Yellow
    }
}

# 8. Optional Data Deletion
$dbDir = "$env:LOCALAPPDATA\FastMailSearch"
if (Test-Path $dbDir) {
    if ($DeleteData) {
        Write-Host "Deleting indexed database and user settings from $dbDir (-DeleteData specified)..." -ForegroundColor Yellow
        Remove-Item -Path $dbDir -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host "`nNotice: Your indexed search database and logs have been preserved at:" -ForegroundColor Cyan
        Write-Host "  $dbDir" -ForegroundColor White
        Write-Host "To delete all indexed data manually, remove the directory above." -ForegroundColor Gray
    }
}

Write-Host "`n=================================================" -ForegroundColor Green
Write-Host " FastMail Search Add-In Uninstalled Successfully!" -ForegroundColor Green
Write-Host "=================================================" -ForegroundColor Green
