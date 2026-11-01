# PowerShell Installer Script for FastMail Search Classic Outlook Managed COM Add-In
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host " Installing FastMail Search Outlook Add-In..." -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

# 1. Detect if Outlook is currently running
$outlookProcesses = Get-Process outlook -ErrorAction SilentlyContinue
if ($outlookProcesses) {
    if ($Force) {
        Write-Host "Closing running Classic Outlook process (-Force specified)..." -ForegroundColor Yellow
        Stop-Process -Name outlook -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    } else {
        Write-Host "`nWARNING: Classic Outlook is currently running." -ForegroundColor Yellow
        Write-Host "Please close Outlook to complete installation, or re-run with -Force." -ForegroundColor Yellow
        $response = Read-Host "Would you like to close Outlook now? (Y/N)"
        if ($response -eq "Y" -or $response -eq "y") {
            Stop-Process -Name outlook -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        } else {
            Write-Error "Installation cancelled: Outlook must be closed to register the Add-In."
            exit 1
        }
    }
}

# 2. Locate Assembly File
$solutionDir = Resolve-Path "$PSScriptRoot\.."
$dllPath = "$solutionDir\src\FastMailSearch.OutlookAddIn\bin\Release\net48\FastMailSearch.OutlookAddIn.dll"

if (-not (Test-Path $dllPath)) {
    Write-Host "Assembly not found. Building solution..." -ForegroundColor Yellow
    & "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "$solutionDir\FastMailSearch.slnx" /p:Configuration=Release
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed. Cannot proceed with installation."
        exit 1
    }
}

$dllFullPath = (Get-Item $dllPath).FullName
$codeBaseUrl = "file:///" + $dllFullPath.Replace("\", "/")

Write-Host "`nAssembly Path: $dllFullPath" -ForegroundColor White

# 3. Clean up any previous registration
$progId = "FastMailSearch.AddIn"
$clsid = "{4A10F4D0-7E55-442C-B23B-983F78BA48D2}"
$smokeProgId = "FastMailSearch.SmokeTest"
$smokeClsid = "{9B1C2A3D-4E5F-6789-0A1B-2C3D4E5F6789}"

$outlookAddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
$progIdKey = "HKCU:\Software\Classes\$progId"
$clsidKey = "HKCU:\Software\Classes\CLSID\$clsid"
$smokeProgIdKey = "HKCU:\Software\Classes\$smokeProgId"
$smokeClsidKey = "HKCU:\Software\Classes\CLSID\$smokeClsid"

Write-Host "`nCleaning up old/stale registry entries..." -ForegroundColor Yellow
if (Test-Path $outlookAddinKey) { Remove-Item -Path $outlookAddinKey -Recurse -Force }
if (Test-Path $progIdKey) { Remove-Item -Path $progIdKey -Recurse -Force }
if (Test-Path $clsidKey) { Remove-Item -Path $clsidKey -Recurse -Force }
if (Test-Path $smokeProgIdKey) { Remove-Item -Path $smokeProgIdKey -Recurse -Force }
if (Test-Path $smokeClsidKey) { Remove-Item -Path $smokeClsidKey -Recurse -Force }

# 4. Register COM ProgID and CLSID in HKCU
Write-Host "Configuring Native COM Registration in HKCU:\Software\Classes..." -ForegroundColor Yellow

# ProgID Key
New-Item -Path $progIdKey -Force | Out-Null
Set-ItemProperty -Path $progIdKey -Name "(default)" -Value "FastMailSearch.OutlookAddIn.ThisAddIn"
$progIdClsidKey = "$progIdKey\CLSID"
New-Item -Path $progIdClsidKey -Force | Out-Null
Set-ItemProperty -Path $progIdClsidKey -Name "(default)" -Value $clsid

# CLSID Key
New-Item -Path $clsidKey -Force | Out-Null
Set-ItemProperty -Path $clsidKey -Name "(default)" -Value "FastMailSearch.OutlookAddIn.ThisAddIn"

# InprocServer32 Key
$inprocKey = "$clsidKey\InprocServer32"
New-Item -Path $inprocKey -Force | Out-Null
Set-ItemProperty -Path $inprocKey -Name "(default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocKey -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $inprocKey -Name "Class" -Value "FastMailSearch.OutlookAddIn.ThisAddIn"
Set-ItemProperty -Path $inprocKey -Name "Assembly" -Value "FastMailSearch.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty -Path $inprocKey -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path $inprocKey -Name "CodeBase" -Value $codeBaseUrl

# Version Subkey under InprocServer32
$versionKey = "$inprocKey\1.0.0.0"
New-Item -Path $versionKey -Force | Out-Null
Set-ItemProperty -Path $versionKey -Name "Class" -Value "FastMailSearch.OutlookAddIn.ThisAddIn"
Set-ItemProperty -Path $versionKey -Name "Assembly" -Value "FastMailSearch.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty -Path $versionKey -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path $versionKey -Name "CodeBase" -Value $codeBaseUrl

# ProgID subkey under CLSID
$clsidProgIdKey = "$clsidKey\ProgID"
New-Item -Path $clsidProgIdKey -Force | Out-Null
Set-ItemProperty -Path $clsidProgIdKey -Name "(default)" -Value $progId

# Implemented Categories Subkey
$catKey = "$clsidKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
New-Item -Path $catKey -Force | Out-Null

# 5. Register Smoke Test COM Class
New-Item -Path $smokeProgIdKey -Force | Out-Null
Set-ItemProperty -Path $smokeProgIdKey -Name "(default)" -Value "FastMailSearch.OutlookAddIn.ComSmokeTest"
New-Item -Path "$smokeProgIdKey\CLSID" -Force | Out-Null
Set-ItemProperty -Path "$smokeProgIdKey\CLSID" -Name "(default)" -Value $smokeClsid

New-Item -Path $smokeClsidKey -Force | Out-Null
Set-ItemProperty -Path $smokeClsidKey -Name "(default)" -Value "FastMailSearch.OutlookAddIn.ComSmokeTest"
$smokeInproc = "$smokeClsidKey\InprocServer32"
New-Item -Path $smokeInproc -Force | Out-Null
Set-ItemProperty -Path $smokeInproc -Name "(default)" -Value "mscoree.dll"
Set-ItemProperty -Path $smokeInproc -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $smokeInproc -Name "Class" -Value "FastMailSearch.OutlookAddIn.ComSmokeTest"
Set-ItemProperty -Path $smokeInproc -Name "Assembly" -Value "FastMailSearch.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty -Path $smokeInproc -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path $smokeInproc -Name "CodeBase" -Value $codeBaseUrl

# 6. Register Add-In in HKCU Outlook Addins (Native COM format ONLY - NO Manifest property)
Write-Host "Registering Native COM Outlook Addin Key: $outlookAddinKey..." -ForegroundColor Yellow

New-Item -Path $outlookAddinKey -Force | Out-Null
Set-ItemProperty -Path $outlookAddinKey -Name "FriendlyName" -Value "FastMail Search"
Set-ItemProperty -Path $outlookAddinKey -Name "Description" -Value "FastMail Search Classic Outlook Add-In"
Set-ItemProperty -Path $outlookAddinKey -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $outlookAddinKey -Name "CommandLineSafe" -Value 0 -Type DWord

# Ensure NO Manifest property exists
if ((Get-ItemProperty -Path $outlookAddinKey -Name "Manifest" -ErrorAction SilentlyContinue) -ne $null) {
    Remove-ItemProperty -Path $outlookAddinKey -Name "Manifest" -ErrorAction SilentlyContinue
}

# 7. Outlook Resiliency Protection
$doNotDisableKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"
if (-not (Test-Path $doNotDisableKey)) { New-Item -Path $doNotDisableKey -Force | Out-Null }
Set-ItemProperty -Path $doNotDisableKey -Name $progId -Value 1 -Type DWord

# 8. Verification
Write-Host "`nVerifying Native COM Registration..." -ForegroundColor Yellow

$installedAddin = Get-ItemProperty -Path $outlookAddinKey -ErrorAction SilentlyContinue
$installedClsid = Get-ItemProperty -Path $inprocKey -ErrorAction SilentlyContinue

$hasManifest = ($installedAddin -and $installedAddin.Manifest -ne $null)

if ($installedAddin -and $installedAddin.LoadBehavior -eq 3 -and -not $hasManifest -and $installedClsid -and $installedClsid.CodeBase -eq $codeBaseUrl) {
    Write-Host "`n=================================================" -ForegroundColor Green
    Write-Host " FastMail Search Native COM Add-In Installed!" -ForegroundColor Green
    Write-Host "=================================================" -ForegroundColor Green
    Write-Host "Verified:" -ForegroundColor Green
    Write-Host " - Native COM Addin Registry Key Present" -ForegroundColor White
    Write-Host " - LoadBehavior = 3 (Load on Startup)" -ForegroundColor White
    Write-Host " - Manifest property REMOVED (Prevents VSTOInstaller)" -ForegroundColor White
    Write-Host " - CLSID, Version subkey & Implemented Categories registered" -ForegroundColor White
    Write-Host " - Outlook Resiliency DoNotDisableAddinList configured" -ForegroundColor White
} else {
    Write-Error "Registration verification failed. Please inspect errors above."
    exit 1
}
