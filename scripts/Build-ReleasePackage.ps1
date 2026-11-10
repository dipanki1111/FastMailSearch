# PowerShell Script to Build Phase 27 Microsoft Store Distributable Standalone Installer EXE
# Hardened for v1.0.6 Isolated Workspace Packaging (Stage 1 Safety Pipeline)
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Release version must be explicitly provided (e.g. 1.0.6)")]
    [ValidateNotNullOrEmpty()]
    [string]$Version,

    [Parameter(Mandatory = $false, HelpMessage = "If specified, stops after solution build, staging, and payload audit, skipping installer packaging")]
    [switch]$BuildOnly,

    [Parameter(Mandatory = $false, HelpMessage = "If specified, skips solution build and staging harvesting, packaging the existing audited staging payload")]
    [switch]$PackageOnly
)

$ErrorActionPreference = "Stop"

# 1. Strict Version Validation (v1.0.6 / v1.0.7 isolated release workflow)
if ($Version -ne "1.0.6" -and $Version -ne "1.0.7") {
    Write-Error "PACKAGING ABORTED: This safety-hardened workflow mandates -Version '1.0.6' or '1.0.7'. Received: '$Version'."
    exit 1
}

if ($BuildOnly -and $PackageOnly) {
    Write-Error "PACKAGING ABORTED: Cannot specify both -BuildOnly and -PackageOnly simultaneously."
    exit 1
}

$solutionDir = Resolve-Path "$PSScriptRoot\.."
$scratchDir = Join-Path $solutionDir "scratch"
$workspaceDir = Join-Path $scratchDir "release-v$Version"
$releaseDir = Join-Path $solutionDir "Release"

# Isolated subfolders under the versioned scratch release workspace.
$stagingDir = Join-Path $workspaceDir "staging"
$setupBuildDir = Join-Path $workspaceDir "setup-build"
$packageDir = Join-Path $workspaceDir "package"
$metadataDir = Join-Path $workspaceDir "metadata"
$checksumsDir = Join-Path $workspaceDir "checksums"
$auditDir = Join-Path $workspaceDir "audit"

$setupProjectDir = Join-Path $solutionDir "src\FastMailSearch.Setup"
$sourcePayloadZip = Join-Path $setupProjectDir "Payload.zip"

# Pre-declared output and artifact paths for isolation and collision protection
$isolatedPayloadZip = Join-Path $setupBuildDir "Payload.zip"
$setupCompiledExe = Join-Path $setupBuildDir "bin\FastMailSearch-v$Version-Store-Setup.exe"
$setupCompiledPdb = Join-Path $setupBuildDir "bin\FastMailSearch-v$Version-Store-Setup.pdb"
$targetPackageExe = Join-Path $packageDir "FastMailSearch-v$Version-Store-Setup.exe"
$targetChecksumFile = Join-Path $checksumsDir "FastMailSearch-v$Version-Store-Setup.sha256"
$workspaceShaFile = Join-Path $checksumsDir "SHA256SUMS-v$Version.txt"
$sharedReleaseExe = Join-Path $releaseDir "FastMailSearch-v$Version-Store-Setup.exe"
$sharedReleaseSha = Join-Path $releaseDir "FastMailSearch-v$Version-Store-Setup.sha256"
$sigLog = Join-Path $auditDir "signature-audit.log"
$versionTxt = Join-Path $metadataDir "VERSION.txt"
$readmeStoreTxt = Join-Path $metadataDir "README-STORE.txt"
$storeNotesTxt = Join-Path $metadataDir "STORE_SUBMISSION_NOTES.txt"

$scriptsToCopy = @(
    "$solutionDir\scripts\Uninstall-FastMailSearch.ps1",
    "$solutionDir\scripts\Verify-FastMailSearchInstallation.ps1",
    "$solutionDir\scripts\Verify-FastMailSearchPayloadAudit.ps1",
    "$solutionDir\scripts\Verify-FastMailSearchSignature.ps1"
)

# Resolve MSBuild upfront so it is guaranteed available in all modes
$msbuild = & "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe" -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe
if (-not $msbuild -or -not (Test-Path $msbuild)) {
    Write-Error "MSBuild.exe not found via vswhere."
    exit 1
}

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host " Building FastMail Search v$Version Store Installer" -ForegroundColor Cyan
Write-Host " Isolated Workspace: $workspaceDir" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

# 2. Preflight Collision & State Guards
if (-not $PackageOnly) {
    if (Test-Path $workspaceDir) {
        Write-Error "PREFLIGHT COLLISION GUARD: Workspace directory already exists: $workspaceDir. Aborting to avoid overwrite or deletion."
        exit 1
    }
} else {
    if (-not (Test-Path $workspaceDir)) {
        Write-Error "PREFLIGHT GUARD: Workspace directory does not exist for -PackageOnly: $workspaceDir. Run -BuildOnly first."
        exit 1
    }
    if (-not (Test-Path $stagingDir)) {
        Write-Error "PREFLIGHT GUARD: Staging directory does not exist for -PackageOnly: $stagingDir. Aborting."
        exit 1
    }

    # Validate required staging binaries exist
    $requiredStagingFiles = @(
        "FastMailSearch.OutlookAddIn.dll",
        "FastMailSearch.Core.dll",
        "FastMailSearch.Infrastructure.dll",
        "FastMailSearch.UI.exe",
        "FastMailSearch.UI.exe.config"
    )
    foreach ($reqFile in $requiredStagingFiles) {
        $reqPath = Join-Path $stagingDir $reqFile
        if (-not (Test-Path $reqPath)) {
            Write-Error "PREFLIGHT GUARD: Required staging file missing for -PackageOnly: $reqFile at '$reqPath'. Aborting."
            exit 1
        }
    }

    # Validate all three SQLite binding redirects in staged FastMailSearch.UI.exe.config point to 2.1.6.2060
    $stagedConfig = Join-Path $stagingDir "FastMailSearch.UI.exe.config"
    $configXml = [xml](Get-Content -Path $stagedConfig)
    $deps = $configXml.configuration.runtime.assemblyBinding.dependentAssembly

    $requiredSqliteAssemblies = @(
        "SQLitePCLRaw.core",
        "SQLitePCLRaw.batteries_v2",
        "SQLitePCLRaw.provider.dynamic_cdecl"
    )

    foreach ($asmName in $requiredSqliteAssemblies) {
        $matches = @($deps | Where-Object { $_.assemblyIdentity -and $_.assemblyIdentity.name -eq $asmName })
        if ($matches.Count -ne 1) {
            Write-Error "PREFLIGHT GUARD: Staged FastMailSearch.UI.exe.config must contain exactly 1 dependentAssembly entry for '$asmName', but found $($matches.Count). Aborting."
            exit 1
        }

        $asmEntry = $matches[0]
        if (-not $asmEntry.assemblyIdentity) {
            Write-Error "PREFLIGHT GUARD: Staged FastMailSearch.UI.exe.config entry for '$asmName' is missing assemblyIdentity. Aborting."
            exit 1
        }

        if (-not $asmEntry.bindingRedirect) {
            Write-Error "PREFLIGHT GUARD: Staged FastMailSearch.UI.exe.config entry for '$asmName' is missing bindingRedirect. Aborting."
            exit 1
        }

        $oldVer = $asmEntry.bindingRedirect.oldVersion
        if ([string]::IsNullOrWhiteSpace($oldVer)) {
            Write-Error "PREFLIGHT GUARD: Staged FastMailSearch.UI.exe.config entry for '$asmName' has empty or missing oldVersion. Aborting."
            exit 1
        }

        $newVer = $asmEntry.bindingRedirect.newVersion
        if ($newVer -ne "2.1.6.2060") {
            Write-Error "PREFLIGHT GUARD: Staged FastMailSearch.UI.exe.config entry for '$asmName' specifies newVersion '$newVer' instead of expected '2.1.6.2060'. Aborting."
            exit 1
        }
    }
}

# Collision guards for every output that could be overwritten
if (Test-Path $isolatedPayloadZip) {
    Write-Error "PREFLIGHT COLLISION GUARD: Target payload ZIP already exists in setup-build: $isolatedPayloadZip. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $setupCompiledExe) {
    Write-Error "PREFLIGHT COLLISION GUARD: Compiled setup EXE already exists in setup-build: $setupCompiledExe. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $setupCompiledPdb) {
    Write-Error "PREFLIGHT COLLISION GUARD: Compiled setup PDB already exists in setup-build: $setupCompiledPdb. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $targetPackageExe) {
    Write-Error "PREFLIGHT COLLISION GUARD: Target package already exists in workspace: $targetPackageExe. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $sigLog) {
    Write-Error "PREFLIGHT COLLISION GUARD: Signature audit log already exists in workspace: $sigLog. Aborting to avoid overwrite."
    exit 1
}
if ((Test-Path $versionTxt) -or (Test-Path $readmeStoreTxt) -or (Test-Path $storeNotesTxt)) {
    Write-Error "PREFLIGHT COLLISION GUARD: Metadata documentation files already exist in $metadataDir. Aborting to avoid overwrite."
    exit 1
}
foreach ($script in $scriptsToCopy) {
    $pkgScript = Join-Path $packageDir (Split-Path $script -Leaf)
    if (Test-Path $pkgScript) {
        Write-Error "PREFLIGHT COLLISION GUARD: Package utility script already exists in package directory: $pkgScript. Aborting to avoid overwrite."
        exit 1
    }
}
if (Test-Path $targetChecksumFile) {
    Write-Error "PREFLIGHT COLLISION GUARD: Target checksum file already exists in workspace: $targetChecksumFile. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $workspaceShaFile) {
    Write-Error "PREFLIGHT COLLISION GUARD: Workspace checksum manifest already exists in workspace: $workspaceShaFile. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $sharedReleaseExe) {
    Write-Error "PREFLIGHT COLLISION GUARD: Release executable already exists in shared Release directory: $sharedReleaseExe. Aborting to avoid overwrite."
    exit 1
}
if (Test-Path $sharedReleaseSha) {
    Write-Error "PREFLIGHT COLLISION GUARD: Release checksum file already exists in shared Release directory: $sharedReleaseSha. Aborting to avoid overwrite."
    exit 1
}

# 3. Path Validation Guard & Workspace Initialization
$expectedWorkspace = [System.IO.Path]::GetFullPath((Join-Path $scratchDir "release-v$Version"))
$resolvedWorkspace = [System.IO.Path]::GetFullPath($workspaceDir)

if (-not $resolvedWorkspace.Equals($expectedWorkspace, [StringComparison]::OrdinalIgnoreCase)) {
    Write-Error "PATH VALIDATION GUARD: Workspace path '$resolvedWorkspace' does not match expected isolated path '$expectedWorkspace'. Aborting."
    exit 1
}

$resolvedScratch = [System.IO.Path]::GetFullPath($scratchDir)
if ($resolvedWorkspace.Equals($resolvedScratch, [StringComparison]::OrdinalIgnoreCase) -or -not $resolvedWorkspace.StartsWith($resolvedScratch, [StringComparison]::OrdinalIgnoreCase)) {
    Write-Error "PATH VALIDATION GUARD: Workspace path must be a dedicated subdirectory under '$resolvedScratch'. Aborting."
    exit 1
}

Write-Host "Initializing isolated workspace subfolders..." -ForegroundColor Yellow
$subfolders = @($stagingDir, $setupBuildDir, $packageDir, $metadataDir, $checksumsDir, $auditDir)
foreach ($folder in $subfolders) {
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
}

# Record baseline hash of source Payload.zip to ensure it is never modified
$initialPayloadHash = $null
if (Test-Path $sourcePayloadZip) {
    $initialPayloadHash = (Get-FileHash -Path $sourcePayloadZip -Algorithm SHA256).Hash
    Write-Host "Source Payload.zip baseline SHA256: $initialPayloadHash" -ForegroundColor Gray
}

if (-not $PackageOnly) {
    # 4. Compile Solution Binaries in Release Mode
    Write-Host "`nRestoring and compiling solution binaries in Release mode..." -ForegroundColor Yellow
    & dotnet restore "$solutionDir\FastMailSearch.slnx"
    & $msbuild "$solutionDir\FastMailSearch.slnx" /p:Configuration=Release /p:Version="$Version" /p:FileVersion="$Version.0" /p:AssemblyVersion="$Version.0"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "MSBuild failed on solution."
        exit 1
    }

    # 5. Populate Staging Payload Directory
    Write-Host "`nPreparing staging payload at $stagingDir..." -ForegroundColor Yellow
    $net48Out = Join-Path $solutionDir "src\FastMailSearch.OutlookAddIn\bin\Release\net48"
    $uiOut = Join-Path $solutionDir "src\FastMailSearch.UI\bin\Release\net48"

    Get-ChildItem -Path $net48Out -File | Where-Object { 
        $_.Extension -ne ".pdb" -and $_.Extension -ne ".xml" -and $_.Name -notlike "*Test*"
    } | Copy-Item -Destination $stagingDir -Force

    if (Test-Path $uiOut) {
        Get-ChildItem -Path $uiOut -File | Where-Object { 
            $_.Extension -ne ".pdb" -and $_.Extension -ne ".xml" -and $_.Name -notlike "*Test*"
        } | Copy-Item -Destination $stagingDir -Force
    }

    $nativeSubDir = "$net48Out\runtimes"
    if (Test-Path $nativeSubDir) {
        Copy-Item -Path $nativeSubDir -Destination $stagingDir -Recurse -Force
    }

    foreach ($script in $scriptsToCopy) {
        if (Test-Path $script) { Copy-Item -Path $script -Destination $stagingDir -Force }
    }

    # Create appsettings.json in Staging
    $appSettingsPath = Join-Path $stagingDir "appsettings.json"
    if (-not (Test-Path $appSettingsPath)) {
        @'
{
  "Indexer": {
    "BatchSize": 100,
    "MaxRetryAttempts": 3,
    "IsBetaModeEnabled": false
  }
}
'@ | Out-File -FilePath $appSettingsPath -Encoding utf8
    }
}

# 6. Release Content & Payload Safety Audit
Write-Host "`nRunning Release Payload Audit on Staging Payload..." -ForegroundColor Yellow
$auditScript = Join-Path $solutionDir "scripts\Verify-FastMailSearchPayloadAudit.ps1"
$auditLog = Join-Path $auditDir "payload-audit.log"
& powershell -ExecutionPolicy Bypass -File $auditScript -StagingDir $stagingDir *>&1 | Tee-Object -FilePath $auditLog
$auditExitCode = $LASTEXITCODE
if ($auditExitCode -ne 0) {
    Write-Error "RELEASE PAYLOAD AUDIT FAILED! Setup payload contains forbidden files or invalid assemblies (Exit code: $auditExitCode). See $auditLog"
    exit 1
}

# Early exit for Stage 3A Controlled Build (-BuildOnly)
if ($BuildOnly) {
    if ($initialPayloadHash) {
        $currentPayloadHash = (Get-FileHash -Path $sourcePayloadZip -Algorithm SHA256).Hash
        if ($currentPayloadHash -ne $initialPayloadHash) {
            Write-Error "INTEGRITY VIOLATION: Source Payload.zip was modified! Baseline: $initialPayloadHash, Current: $currentPayloadHash"
            exit 1
        }
        Write-Host "Source Payload.zip integrity verified (hash unchanged)." -ForegroundColor Green
    }
    Write-Host "`n=================================================" -ForegroundColor Green
    Write-Host " Stage 3A: Solution Build & Staging Payload Audit COMPLETED" -ForegroundColor Green
    Write-Host " Staged Binaries: $stagingDir" -ForegroundColor White
    Write-Host " Payload Audit:  $auditLog (Audit exit code: $auditExitCode)" -ForegroundColor White
    Write-Host " STOPPED: Final installer packaging skipped (-BuildOnly)." -ForegroundColor Yellow
    Write-Host "=================================================" -ForegroundColor Green
    exit 0
}

# 7. Zip Staging Payload into Isolated setup-build Directory (PROTECTS source Payload.zip)
Write-Host "`nZipping staging payload into isolated setup-build directory..." -ForegroundColor Yellow
$isolatedPayloadZip = Join-Path $setupBuildDir "Payload.zip"
Compress-Archive -Path "$stagingDir\*" -DestinationPath $isolatedPayloadZip -Force

# 8. Build Standalone Installer Executable into Isolated Setup-Build
Write-Host "`nCompiling standalone Store installer executable FastMailSearch-v$Version-Store-Setup.exe..." -ForegroundColor Yellow
& dotnet restore "$setupProjectDir\FastMailSearch.Setup.csproj"
& $msbuild "$setupProjectDir\FastMailSearch.Setup.csproj" `
    /p:Configuration=Release `
    /p:AssemblyName="FastMailSearch-v$Version-Store-Setup" `
    /p:Version="$Version" `
    /p:FileVersion="$Version.0" `
    /p:InformationalVersion="$Version" `
    /p:PayloadZipPath="$isolatedPayloadZip" `
    /p:OutputPath="$setupBuildDir\bin\" `
    /p:IntermediateOutputPath="$setupBuildDir\obj\"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to compile FastMailSearch.Setup.csproj"
    exit 1
}

$setupCompiledExe = Join-Path $setupBuildDir "bin\FastMailSearch-v$Version-Store-Setup.exe"
if (-not (Test-Path $setupCompiledExe)) {
    # Fallback search if output path had framework subfolder
    $setupCompiledExe = (Get-ChildItem -Path "$setupBuildDir\bin" -Filter "FastMailSearch-v$Version-Store-Setup.exe" -Recurse | Select-Object -First 1).FullName
}

if (-not (Test-Path $setupCompiledExe)) {
    Write-Error "Compiled setup EXE not found under $setupBuildDir\bin"
    exit 1
}

# Copy compiled installer to isolated package directory
Copy-Item -Path $setupCompiledExe -Destination $targetPackageExe -Force

# 9. Verify Source Payload.zip Remained Completely Untouched
if ($initialPayloadHash) {
    $currentPayloadHash = (Get-FileHash -Path $sourcePayloadZip -Algorithm SHA256).Hash
    if ($currentPayloadHash -ne $initialPayloadHash) {
        Write-Error "INTEGRITY VIOLATION: Source Payload.zip was modified! Baseline: $initialPayloadHash, Current: $currentPayloadHash"
        exit 1
    }
    Write-Host "Source Payload.zip integrity verified (hash unchanged)." -ForegroundColor Green
}

# 10. Installer Size Validation (> 1 MB)
$exeItem = Get-Item $targetPackageExe
$exeSizeMb = [math]::Round($exeItem.Length / 1MB, 2)
Write-Host "`n-----------------------------------------" -ForegroundColor White
Write-Host " STORE INSTALLER EXE AUDIT:" -ForegroundColor White
Write-Host " Name:   $($exeItem.Name)" -ForegroundColor White
Write-Host " Length: $($exeItem.Length) bytes ($exeSizeMb MB)" -ForegroundColor White
Write-Host " Path:   $($exeItem.FullName)" -ForegroundColor White
Write-Host "-----------------------------------------`n" -ForegroundColor White

if ($exeItem.Length -lt 1048576) {
    Write-Error "INSTALLER SIZE FAILURE: Setup.exe is less than 1 MB ($($exeItem.Length) bytes). Must contain embedded payload."
    exit 1
}

# 11. Signature Audit (Isolated Log)
Write-Host "Auditing Authenticode Signature Status..." -ForegroundColor Yellow
$sigScript = Join-Path $solutionDir "scripts\Verify-FastMailSearchSignature.ps1"
$sigLog = Join-Path $auditDir "signature-audit.log"
& powershell -ExecutionPolicy Bypass -File $sigScript -FilePath $targetPackageExe *>&1 | Tee-Object -FilePath $sigLog

# 12. Generate Isolated Metadata (inside workspace\metadata, PRESERVING Release\)
Write-Host "`nGenerating metadata inside isolated workspace metadata directory..." -ForegroundColor Yellow

$versionTxt = Join-Path $metadataDir "VERSION.txt"
"FastMail Search Version: $Version`nBuild Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`nTarget: Classic Outlook (x86/x64)`nRelease Type: Microsoft Store Unpackaged Win32 EXE" | Out-File -FilePath $versionTxt -Encoding utf8

$readmeStoreTxt = Join-Path $metadataDir "README-STORE.txt"
@"
=================================================
 FastMail Search v$Version Microsoft Store Release
=================================================

FastMail Search is a high-performance local full-text search add-in for Classic Microsoft Outlook.

INSTALLATION INSTRUCTIONS (SILENT STORE INSTALL):
Run the installer with silent flag:
FastMailSearch-v$Version-Store-Setup.exe /quiet

REQUIREMENTS:
- Classic Microsoft Outlook (x86 or x64)
- Windows 10/11
- .NET Framework 4.8

UNINSTALLATION INSTRUCTIONS:
Run Uninstall-FastMailSearch.ps1 from PowerShell.

VERIFICATION:
Run Verify-FastMailSearchInstallation.ps1 to test installation integrity.
"@ | Out-File -FilePath $readmeStoreTxt -Encoding utf8

$storeNotesTxt = Join-Path $metadataDir "STORE_SUBMISSION_NOTES.txt"
@"
=================================================
 Microsoft Store Submission Notes (Unpackaged Win32)
=================================================

Application Name: FastMail Search
Version: $Version
Packaging Method: Unpackaged Win32 (EXE/MSI)
Architecture: x86 / x64 (Classic Outlook Add-in)

SILENT INSTALL COMMAND:
FastMailSearch-v$Version-Store-Setup.exe /quiet

INSTALL LOCATION:
%LOCALAPPDATA%\Programs\FastMailSearch\

REGISTRATION:
HKCU\Software\Microsoft\Office\Outlook\Addins\FastMailSearch.AddIn (LoadBehavior = 3)
HKCU\Software\Classes\CLSID\{4A10F4D0-7E55-442C-B23B-983F78BA48D2}

DATA PRESERVATION:
User database is preserved at %USERPROFILE%\Documents\FastMailSearch\FastMailSearch.db upon upgrade or uninstall.

OFFLINE INSTALL GUARANTEE:
All binaries, managed assemblies (.NET Framework 4.8), and native SQLite dependencies (e_sqlite3.dll) are fully embedded inside the setup EXE. No network downloads occur during setup.
"@ | Out-File -FilePath $storeNotesTxt -Encoding utf8

# Copy utility scripts into package directory for distributable bundle
foreach ($script in $scriptsToCopy) {
    if (Test-Path $script) { Copy-Item -Path $script -Destination $packageDir -Force }
}

# 13. Generate Isolated Checksums (inside workspace\checksums, PRESERVING Release\SHA256SUMS.txt)
Write-Host "Generating SHA-256 checksums in isolated checksums directory..." -ForegroundColor Yellow
$singleShaFile = Join-Path $checksumsDir "FastMailSearch-v$Version-Store-Setup.sha256"
$exeHash = (Get-FileHash -Path $targetPackageExe -Algorithm SHA256).Hash
"$exeHash  FastMailSearch-v$Version-Store-Setup.exe" | Out-File -FilePath $singleShaFile -Encoding utf8

$workspaceShaFile = Join-Path $checksumsDir "SHA256SUMS-v$Version.txt"
$allWorkspaceFiles = Get-ChildItem -Path $workspaceDir -Recurse -File | Where-Object { $_.FullName -ne $workspaceShaFile }
$hashes = foreach ($f in $allWorkspaceFiles) {
    $h = (Get-FileHash -Path $f.FullName -Algorithm SHA256).Hash
    $rel = $f.FullName.Substring($workspaceDir.Length + 1)
    "$h  $rel"
}
$hashes | Out-File -FilePath $workspaceShaFile -Encoding utf8

Write-Host "`n=================================================" -ForegroundColor Green
Write-Host " Isolated Release Package v$Version Generated Successfully!" -ForegroundColor Green
Write-Host " Installer Package:    $targetPackageExe ($exeSizeMb MB)" -ForegroundColor Green
Write-Host " Isolated Workspace:   $workspaceDir" -ForegroundColor White
Write-Host " Metadata Folder:      $metadataDir" -ForegroundColor White
Write-Host " Checksums Folder:     $checksumsDir" -ForegroundColor White
Write-Host " Shared Release Dir:   UNTOUCHED (Preserved)" -ForegroundColor Green
Write-Host "=================================================" -ForegroundColor Green
