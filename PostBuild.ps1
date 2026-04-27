# Post-build deployment script for dwg2rvt plugin
param(
    [string]$ProjectDir,
    [string]$TargetDir,
    [string]$TargetFileName,
    [switch]$SkipDeploy
)

Write-Host "========================================"
Write-Host "Annotatix Post-Build Deployment"
Write-Host "========================================"

# CRITICAL FIX: MSBuild mangles Cyrillic paths - always use script location
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Detect correct output folder
# Priority: use the directory with the MOST RECENT PluginsManager.dll
# This avoids the NuGet cache problem where bin/R25 has stale DLLs
$R25Dir = [System.IO.Path]::Combine($ProjectDir, "bin", "R25")
$ReleaseDir = [System.IO.Path]::Combine($ProjectDir, "bin", "Release")
$Net48Dir = [System.IO.Path]::Combine($ProjectDir, "bin", "Release", "net48")
$DebugDir = [System.IO.Path]::Combine($ProjectDir, "bin", "Debug")

# Find the most recent build output by checking PluginsManager.dll timestamp
$candidates = @()
foreach ($dir in @($DebugDir, $Net48Dir, $ReleaseDir, $R25Dir)) {
    $testFile = [System.IO.Path]::Combine($dir, "PluginsManager.dll")
    if (Test-Path $testFile) {
        $ts = (Get-Item $testFile).LastWriteTime
        $candidates += [PSCustomObject]@{ Dir = $dir; Timestamp = $ts }
        Write-Host "Found build output: $dir (PluginsManager.dll last written: $ts)"
    }
}

if ($candidates.Count -gt 0) {
    # Sort by timestamp descending, pick the most recent
    $best = $candidates | Sort-Object -Property Timestamp -Descending | Select-Object -First 1
    $TargetDir = $best.Dir
    Write-Host "Using MOST RECENT build output: $TargetDir (timestamp: $($best.Timestamp))"
} else {
    $TargetDir = $ReleaseDir
    Write-Host "WARNING: No build output found, defaulting to: $TargetDir"
}
$TargetFileName = "PluginsManager.dll"

Write-Host "Project Directory: $ProjectDir"
Write-Host "Target Directory: $TargetDir"
Write-Host "Target File: $TargetFileName"

# Read build number
$BuildNumberFile = [System.IO.Path]::Combine($ProjectDir, "BuildNumber.txt")
if (Test-Path $BuildNumberFile) {
    $BuildNumber = (Get-Content $BuildNumberFile -Raw).Trim()
} else {
    $BuildNumber = "1"
}

Write-Host "Build Number: $BuildNumber"

# Format build number with leading zeros (001, 002, etc.)
$FormattedBuild = $BuildNumber.PadLeft(3, '0')

# Define base output directory
$BaseOutput = [System.IO.Path]::GetDirectoryName($ProjectDir.TrimEnd('\'))
$VersionFolder = "annotatix_ver3.$FormattedBuild"
$OutputDir = [System.IO.Path]::Combine($BaseOutput, $VersionFolder)

Write-Host "Output Directory: $OutputDir"

# CRITICAL: Check if version folder exists and auto-increment if necessary
while (Test-Path $OutputDir) {
    Write-Host "WARNING: Version folder already exists: $OutputDir"
    Write-Host "Auto-incrementing build number..."
    
    $BuildNumber = [int]$BuildNumber + 1
    $FormattedBuild = $BuildNumber.ToString().PadLeft(3, '0')
    $VersionFolder = "annotatix_ver3.$FormattedBuild"
    $OutputDir = [System.IO.Path]::Combine($BaseOutput, $VersionFolder)
    
    Write-Host "New output directory: $OutputDir"
}

# Update BuildNumber.txt with new incremented value
Set-Content -Path $BuildNumberFile -Value $BuildNumber -NoNewline
Write-Host "Build number updated to: $BuildNumber"

# Create version folder
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "Created NEW version directory: $OutputDir"
}

# ── Define module file lists (used inside the per-version loop) ──────────────
$MainFiles = @(
    "PluginsManager.dll",
    "PluginsManager.pdb"
)

$Dwg2rvtFiles = @(
    "dwg2rvt.Module.dll",
    "dwg2rvt.Module.pdb"
)

$HvacFiles = @(
    "HVAC.Module.dll",
    "HVAC.Module.pdb"
)

$FamilySyncFiles = @(
    "FamilySync.Module.dll",
    "FamilySync.Module.pdb"
)

$AutoNumberingFiles = @(
    "AutoNumbering.Module.dll",
    "AutoNumbering.Module.pdb"
)

$ClashResolveFiles = @(
    "ClashResolve.Module.dll",
    "ClashResolve.Module.pdb"
)

$TracerFiles = @(
    "Tracer.Module.dll",
    "Tracer.Module.pdb"
)

$AnnotatixFiles = @(
    "Annotatix.Module.dll",
    "Annotatix.Module.pdb"
)

$ExcludePatterns = $MainFiles + $Dwg2rvtFiles + $HvacFiles + $FamilySyncFiles + $AutoNumberingFiles + $ClashResolveFiles + $TracerFiles + $AnnotatixFiles + @("dwg2rvt.dll", "dwg2rvt.pdb")

# ── Multi-version structure: 2024 and 2025 subfolders ────────────────────────
# Each subfolder is a complete, standalone installation:
#   annotatix_ver3.XXX/
#     2024/
#       Annotatix.addin
#       annotatix_dependencies/
#         main/ (PluginsManager.dll + shared deps + icons)
#         dwg2rvt/ ... hvac/ ... family_sync/ ... autonumbering/ ... clash_resolve/ ... tracer/ ... logs/
#     2025/   <-- same layout, same DLLs from same build output
$RevitVersions = @("2024", "2025")

foreach ($RevitVer in $RevitVersions) {
    $VerSubDir           = [System.IO.Path]::Combine($OutputDir, $RevitVer)
    $DependenciesFolder  = [System.IO.Path]::Combine($VerSubDir, "annotatix_dependencies")
    $MainFolder          = [System.IO.Path]::Combine($DependenciesFolder, "main")
    $Dwg2rvtFolder       = [System.IO.Path]::Combine($DependenciesFolder, "dwg2rvt")
    $HvacFolder          = [System.IO.Path]::Combine($DependenciesFolder, "hvac")
    $FamilySyncFolder    = [System.IO.Path]::Combine($DependenciesFolder, "family_sync")
    $AutoNumberingFolder = [System.IO.Path]::Combine($DependenciesFolder, "autonumbering")
    $ClashResolveFolder  = [System.IO.Path]::Combine($DependenciesFolder, "clash_resolve")
    $TracerFolder        = [System.IO.Path]::Combine($DependenciesFolder, "tracer")
    $AnnotatixFolder     = [System.IO.Path]::Combine($DependenciesFolder, "annotatix")
    $LogsFolder          = [System.IO.Path]::Combine($DependenciesFolder, "logs")

    foreach ($folder in @($VerSubDir, $DependenciesFolder, $MainFolder, $Dwg2rvtFolder, $HvacFolder, $FamilySyncFolder, $AutoNumberingFolder, $ClashResolveFolder, $TracerFolder, $AnnotatixFolder, $LogsFolder)) {
        if (-not (Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            Write-Host "Created directory: $folder"
        }
    }

    Write-Host ""
    Write-Host "Organizing modules for Revit $RevitVer..."
    Write-Host "----------------"

    # Copy MAIN module files
    Write-Host "Copying MAIN module..."
    foreach ($file in $MainFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $MainFolder -Force
            Write-Host "  → $RevitVer/main/$file"
        }
    }

    # Copy DWG2RVT module files
    Write-Host "Copying DWG2RVT module..."
    foreach ($file in $Dwg2rvtFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $Dwg2rvtFolder -Force
            Write-Host "  → $RevitVer/dwg2rvt/$file"
        }
    }

    # Copy HVAC module files
    Write-Host "Copying HVAC module..."
    foreach ($file in $HvacFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $HvacFolder -Force
            Write-Host "  → $RevitVer/hvac/$file"
        }
    }

    # Copy FamilySync module files
    Write-Host "Copying FamilySync module..."
    foreach ($file in $FamilySyncFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $FamilySyncFolder -Force
            Write-Host "  → $RevitVer/family_sync/$file"
        }
    }

    # Copy AutoNumbering module files
    Write-Host "Copying AutoNumbering module..."
    foreach ($file in $AutoNumberingFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $AutoNumberingFolder -Force
            Write-Host "  → $RevitVer/autonumbering/$file"
        }
    }

    # Copy ClashResolve module files
    Write-Host "Copying ClashResolve module..."
    foreach ($file in $ClashResolveFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $ClashResolveFolder -Force
            Write-Host "  → $RevitVer/clash_resolve/$file"
        }
    }

    # Copy Tracer module files
    Write-Host "Copying Tracer module..."
    foreach ($file in $TracerFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $TracerFolder -Force
            Write-Host "  → $RevitVer/tracer/$file"
        }
    }

    # Copy Annotatix module files
    Write-Host "Copying Annotatix module..."
    foreach ($file in $AnnotatixFiles) {
        $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $AnnotatixFolder -Force
            Write-Host "  → $RevitVer/annotatix/$file"
        }
    }

    # Copy shared dependencies to main folder
    Write-Host "Copying shared dependencies to $RevitVer/main/..."
    Get-ChildItem -Path $TargetDir -File | Where-Object {
        $fileName = $_.Name
        $ExcludePatterns -notcontains $fileName
    } | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $MainFolder -Force
    }
    Write-Host "  → Shared dependencies copied to $RevitVer/main/"

    # Copy BuildNumber.txt for version display
    $BuildNumberSource = [System.IO.Path]::Combine($ProjectDir, "BuildNumber.txt")
    if (Test-Path $BuildNumberSource) {
        Copy-Item -Path $BuildNumberSource -Destination $MainFolder -Force
        Write-Host "  → $RevitVer/main/BuildNumber.txt"
    }

    # Copy icons to main folder
    $SourceIcons = [System.IO.Path]::Combine($ProjectDir, "PluginsManager", "UI", "icons")
    $DestIcons   = [System.IO.Path]::Combine($MainFolder, "UI", "icons")
    if (Test-Path $SourceIcons) {
        if (-not (Test-Path $DestIcons)) {
            New-Item -ItemType Directory -Path $DestIcons -Force | Out-Null
        }
        Copy-Item -Path "$SourceIcons\*" -Destination $DestIcons -Force
        Write-Host "  → $RevitVer/main/UI/icons/"
    }
    
    # Copy Tracer icons folder (for TracerPanel.xaml)
    $TracerIconsSource = [System.IO.Path]::Combine($ProjectDir, "Tracer.Module", "UI", "icons", "Tracer_icons")
    $TracerIconsDest   = [System.IO.Path]::Combine($DestIcons, "Tracer_icons")
    if (Test-Path $TracerIconsSource) {
        if (-not (Test-Path $TracerIconsDest)) {
            New-Item -ItemType Directory -Path $TracerIconsDest -Force | Out-Null
        }
        Copy-Item -Path "$TracerIconsSource\*" -Destination $TracerIconsDest -Force
        Write-Host "  → $RevitVer/main/UI/icons/Tracer_icons/"
    }

    # Copy Tracer icons to main folder (for ribbon buttons)
    $TracerIcons = @("Tracer_45.png", "Tracer_L.png", "Tracer_Bottom.png", "Tracer_Z.png")
    foreach ($icon in $TracerIcons) {
        $iconSource = [System.IO.Path]::Combine($SourceIcons, $icon)
        if (Test-Path $iconSource) {
            Copy-Item -Path $iconSource -Destination $MainFolder -Force
            Write-Host "  → $RevitVer/main/$icon"
        }
    }

    # Copy Annotatix icons folder
    $AnnotatixIconsSource = [System.IO.Path]::Combine($ProjectDir, "UI", "icons", "Annotatix_icons")
    $AnnotatixIconsDest   = [System.IO.Path]::Combine($DestIcons, "Annotatix_icons")
    if (Test-Path $AnnotatixIconsSource) {
        if (-not (Test-Path $AnnotatixIconsDest)) {
            New-Item -ItemType Directory -Path $AnnotatixIconsDest -Force | Out-Null
        }
        Copy-Item -Path "$AnnotatixIconsSource\*" -Destination $AnnotatixIconsDest -Force -Recurse
        Write-Host "  → $RevitVer/main/UI/icons/Annotatix_icons/"
    }

    # Copy Annotatix icon to main folder (for ribbon button, same pattern as Tracer)
    # Use Get-ChildItem to avoid encoding issues with Cyrillic folder name
    # Note: Annotatix icons are in $ProjectDir\UI\icons, not $SourceIcons (PluginsManager)
    $AnnotatixIconsPath = [System.IO.Path]::Combine($ProjectDir, "UI", "icons", "Annotatix_icons")
    if (Test-Path $AnnotatixIconsPath) {
        $AnnotatixIconFile = Get-ChildItem -Path $AnnotatixIconsPath -Recurse -Filter "*32x*.png" | Select-Object -First 1
        if ($AnnotatixIconFile) {
            Copy-Item -Path $AnnotatixIconFile.FullName -Destination "$MainFolder\Annotatix_32x.png" -Force
            Write-Host "  → $RevitVer/main/Annotatix_32x.png"
        }
    }

    # Create .addin file inside the versioned subfolder
    $AddinFile    = [System.IO.Path]::Combine($VerSubDir, "Annotatix.addin")
    $AssemblyPath = "annotatix_dependencies\main\PluginsManager.dll"

    $AddinContent  = "<?xml version=`"1.0`" encoding=`"utf-8`"?>`n"
    $AddinContent += "<RevitAddIns>`n"
    $AddinContent += "  <AddIn Type=`"Application`">`n"
    $AddinContent += "    <Name>Annotatix</Name>`n"
    $AddinContent += "    <Assembly>$AssemblyPath</Assembly>`n"
    $AddinContent += "    <ClientId>b1c2d3e4-f5a6-7890-1234-567890abcdef</ClientId>`n"
    $AddinContent += "    <FullClassName>PluginsManager.App</FullClassName>`n"
    $AddinContent += "    <VendorId>ANNOTATIX</VendorId>`n"
    $AddinContent += "    <VendorDescription>Annotatix - Revit Plugin Manager with Dynamic Module Loading</VendorDescription>`n"
    $AddinContent += "  </AddIn>`n"
    $AddinContent += "</RevitAddIns>"

    [System.IO.File]::WriteAllText($AddinFile, $AddinContent, [System.Text.Encoding]::UTF8)
    Write-Host "- Created $RevitVer/Annotatix.addin"
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "========================================"
Write-Host "VERIFICATION: Checking DLL timestamps..."
Write-Host "========================================"

# Verify all module DLLs have the same (or very close) timestamp
$moduleDlls = @(
    "PluginsManager.dll",
    "dwg2rvt.Module.dll",
    "HVAC.Module.dll",
    "FamilySync.Module.dll",
    "AutoNumbering.Module.dll",
    "ClashResolve.Module.dll",
    "Tracer.Module.dll",
    "Annotatix.Module.dll"
)

$latestTimestamp = [DateTime]::MinValue
$staleDlls = @()

foreach ($dll in $moduleDlls) {
    $dllPath = [System.IO.Path]::Combine($TargetDir, $dll)
    if (Test-Path $dllPath) {
        $ts = (Get-Item $dllPath).LastWriteTime
        if ($ts -gt $latestTimestamp) {
            $latestTimestamp = $ts
        }
        Write-Host "  $dll : $ts"
    } else {
        Write-Host "  $dll : MISSING!"
        $staleDlls += $dll
    }
}

# Check if any DLL is more than 2 minutes older than the latest
$warningThreshold = [TimeSpan]::FromMinutes(2)
foreach ($dll in $moduleDlls) {
    $dllPath = [System.IO.Path]::Combine($TargetDir, $dll)
    if (Test-Path $dllPath) {
        $ts = (Get-Item $dllPath).LastWriteTime
        $age = $latestTimestamp - $ts
        if ($age -gt $warningThreshold) {
            $staleDlls += "$dll (age: {0:F1} min)" -f $age.TotalMinutes
        }
    }
}

if ($staleDlls.Count -gt 0) {
    Write-Host ""
    Write-Host "WARNING: Some DLLs may be stale (older than 2 min from newest):"
    foreach ($stale in $staleDlls) {
        Write-Host "  - $stale"
    }
    Write-Host "This usually means the build was not complete. Run:"
    Write-Host "  dotnet build PostBuildTrigger.csproj --configuration Release"
    Write-Host "to rebuild ALL modules before deploying."
} else {
    Write-Host ""
    Write-Host "All DLLs are fresh! Build is consistent."
}

Write-Host ""
Write-Host "========================================"
Write-Host "AUTO-DEPLOYMENT TO REVIT ADDINS"
Write-Host "========================================"

if ($SkipDeploy) {
    Write-Host "Skipping deployment (SkipDeploy flag set)."
} else {
    $deployResults = @()
    $lockedFiles = @()
    $successFiles = @()

    foreach ($RevitVer in $RevitVersions) {
        $sourceDir = [System.IO.Path]::Combine($OutputDir, $RevitVer, "annotatix_dependencies")
        $addinsDir = [System.IO.Path]::Combine($env:APPDATA, "Autodesk", "Revit", "Addins", $RevitVer, "annotatix_dependencies")
        $addinFile = [System.IO.Path]::Combine($OutputDir, $RevitVer, "Annotatix.addin")
        $addinDest = [System.IO.Path]::Combine($env:APPDATA, "Autodesk", "Revit", "Addins", $RevitVer, "Annotatix.addin")

        Write-Host ""
        Write-Host "Deploying to Revit $RevitVer..."
        Write-Host "  Source: $sourceDir"
        Write-Host "  Target: $addinsDir"

        # Copy all annotatix_dependencies files
        if (Test-Path $sourceDir) {
            $allSourceFiles = Get-ChildItem -Path $sourceDir -File -Recurse
            foreach ($srcFile in $allSourceFiles) {
                $relativePath = $srcFile.FullName.Substring($sourceDir.Length + 1)
                $destFile = [System.IO.Path]::Combine($addinsDir, $relativePath)
                $destDir = [System.IO.Path]::GetDirectoryName($destFile)

                if (-not (Test-Path $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }

                try {
                    $srcSize = $srcFile.Length
                    Copy-Item -Path $srcFile.FullName -Destination $destFile -Force -ErrorAction Stop

                    # VERIFY: check that the file was actually updated
                    if (Test-Path $destFile) {
                        $destSize = (Get-Item $destFile).Length
                        if ($destSize -eq $srcSize) {
                            $successFiles += "$RevitVer/$relativePath"
                        } else {
                            # File was not updated (likely locked by Revit)
                            $lockedFiles += "$RevitVer/$relativePath (expected $srcSize bytes, got $destSize bytes)"
                        }
                    } else {
                        $lockedFiles += "$RevitVer/$relativePath (file missing after copy)"
                    }
                } catch {
                    $lockedFiles += "$RevitVer/$relativePath (LOCKED: $($_.Exception.Message))"
                }
            }
        } else {
            Write-Host "  WARNING: Source directory not found: $sourceDir"
        }

        # Copy .addin file
        if (Test-Path $addinFile) {
            try {
                Copy-Item -Path $addinFile -Destination $addinDest -Force -ErrorAction Stop
                Write-Host "  Deployed: $RevitVer/Annotatix.addin"
            } catch {
                Write-Host "  WARNING: Could not copy .addin file: $($_.Exception.Message)"
            }
        }
    }

    # Deployment summary
    Write-Host ""
    Write-Host "----------------------------------------"
    Write-Host "DEPLOYMENT SUMMARY"
    Write-Host "----------------------------------------"
    Write-Host "  Successfully deployed: $($successFiles.Count) file(s)"
    if ($lockedFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "  *** LOCKED FILES - NOT UPDATED ***"
        foreach ($locked in $lockedFiles) {
            Write-Host "    X $locked" -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "  CAUSE: Revit is likely running and has these files locked." -ForegroundColor Yellow
        Write-Host "  FIX: Close Revit, then run this script again or copy manually:" -ForegroundColor Yellow
        Write-Host "    xcopy \"$($OutputDir)\2025\annotatix_dependencies\" \"$addinsDir\" /E /Y /I" -ForegroundColor Yellow
    } else {
        Write-Host "  All files deployed successfully!" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "========================================"
Write-Host "Build Deployment Complete!"
Write-Host "========================================"
Write-Host "Version: 3.$FormattedBuild"
Write-Host "Output: $OutputDir"
Write-Host ""
Write-Host "Structure:"
Write-Host "- 2024/  (for Revit 2024)"
Write-Host "  + Annotatix.addin"
Write-Host "  + annotatix_dependencies/"
Write-Host "    + main/ (PluginsManager.dll + ALL deps + icons)"
Write-Host "    + dwg2rvt/ hvac/ family_sync/ autonumbering/ clash_resolve/ logs/"
Write-Host "- 2025/  (for Revit 2025 - same DLLs, correct log path auto-detected)"
Write-Host "  + Annotatix.addin"
Write-Host "  + annotatix_dependencies/"
Write-Host "    + main/ dwg2rvt/ hvac/ family_sync/ autonumbering/ clash_resolve/ logs/"
Write-Host "========================================"
