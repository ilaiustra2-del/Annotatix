# Post-build deployment script for dwg2rvt plugin
param(
    [string]$ProjectDir,
    [string]$TargetDir,
    [string]$TargetFileName
)

Write-Host "========================================"
Write-Host "Annotatix Post-Build Deployment"
Write-Host "========================================"

# CRITICAL FIX: MSBuild mangles Cyrillic paths - always use script location
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TargetDir = [System.IO.Path]::Combine($ProjectDir, "bin", "Release")
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
    
    # Increment build number
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

# Create annotatix_dependencies folder with modular structure
$DependenciesFolder = [System.IO.Path]::Combine($OutputDir, "annotatix_dependencies")
$MainFolder = [System.IO.Path]::Combine($DependenciesFolder, "main")
$Dwg2rvtFolder = [System.IO.Path]::Combine($DependenciesFolder, "dwg2rvt")
$HvacFolder = [System.IO.Path]::Combine($DependenciesFolder, "hvac")
$FamilySyncFolder = [System.IO.Path]::Combine($DependenciesFolder, "family_sync")

foreach ($folder in @($DependenciesFolder, $MainFolder, $Dwg2rvtFolder, $HvacFolder, $FamilySyncFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Write-Host "Created directory: $folder"
    }
}

Write-Host ""
Write-Host "Organizing modules into folders..."
Write-Host "----------------"

# Define module-specific files
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

# Copy MAIN module files
Write-Host "Copying MAIN module..."
foreach ($file in $MainFiles) {
    $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $MainFolder -Force
        Write-Host "  → main/$file"
    }
}

# Copy DWG2RVT module files
Write-Host "Copying DWG2RVT module..."
foreach ($file in $Dwg2rvtFiles) {
    $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $Dwg2rvtFolder -Force
        Write-Host "  → dwg2rvt/$file"
    }
}

# Copy HVAC module files
Write-Host "Copying HVAC module..."
foreach ($file in $HvacFiles) {
    $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $HvacFolder -Force
        Write-Host "  → hvac/$file"
    }
}

# Copy FamilySync module files
Write-Host "Copying FamilySync module..."
foreach ($file in $FamilySyncFiles) {
    $sourcePath = [System.IO.Path]::Combine($TargetDir, $file)
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $FamilySyncFolder -Force
        Write-Host "  → family_sync/$file"
    }
}

# Copy shared dependencies to main folder (all modules will find them there)
Write-Host "Copying shared dependencies to main/..."
$ExcludePatterns = $MainFiles + $Dwg2rvtFiles + $HvacFiles + $FamilySyncFiles + @("dwg2rvt.dll", "dwg2rvt.pdb")  # Exclude old artifacts
Get-ChildItem -Path $TargetDir -File | Where-Object {
    $fileName = $_.Name
    $ExcludePatterns -notcontains $fileName
} | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $MainFolder -Force
}
Write-Host "  → Shared dependencies copied to main/"

# Copy BuildNumber.txt for version display
$BuildNumberSource = [System.IO.Path]::Combine($ProjectDir, "BuildNumber.txt")
if (Test-Path $BuildNumberSource) {
    Copy-Item -Path $BuildNumberSource -Destination $MainFolder -Force
    Write-Host "  → main/BuildNumber.txt"
}

# Copy icons to main folder
$SourceIcons = [System.IO.Path]::Combine($ProjectDir, "UI", "icons")
$DestIcons = [System.IO.Path]::Combine($MainFolder, "UI", "icons")
if (Test-Path $SourceIcons) {
    if (-not (Test-Path $DestIcons)) {
        New-Item -ItemType Directory -Path $DestIcons -Force | Out-Null
    }
    Copy-Item -Path "$SourceIcons\*" -Destination $DestIcons -Force
    Write-Host "  → main/UI/icons/ (all icons)"
}

# Create .addin file
$AddinFile = [System.IO.Path]::Combine($OutputDir, "Annotatix.addin")
Write-Host "Creating .addin file..."

# Build the Assembly path dynamically to avoid encoding issues
$UserProfile = [Environment]::GetFolderPath('ApplicationData')
$AssemblyPath = [System.IO.Path]::Combine($UserProfile, "Autodesk", "Revit", "Addins", "2024", "annotatix_dependencies", "main", "PluginsManager.dll")

# Create XML content
$AddinContent = "<?xml version=`"1.0`" encoding=`"utf-8`"?>`n"
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

# Save with UTF-8 encoding (with BOM for better compatibility)
[System.IO.File]::WriteAllText($AddinFile, $AddinContent, [System.Text.Encoding]::UTF8)
Write-Host "- Created Annotatix.addin"
Write-Host "  Assembly path: $AssemblyPath"

# MANUAL DEPLOYMENT MODE - Files are NOT automatically copied to Revit
# User will manually copy files from the output folder to Revit Addins
Write-Host ""
Write-Host "========================================"
Write-Host "MANUAL DEPLOYMENT MODE"
Write-Host "========================================"
Write-Host "Files are ready in: $OutputDir"
Write-Host ""
Write-Host "To install manually:"
Write-Host "1. Close Revit completely"
Write-Host "2. Copy contents of '$DependenciesFolder' to:"
Write-Host "   %APPDATA%\Autodesk\Revit\Addins\2024\annotatix_dependencies\"
Write-Host "3. Copy '$AddinFile' to:"
Write-Host "   %APPDATA%\Autodesk\Revit\Addins\2024\"
Write-Host "4. Restart Revit"
Write-Host "========================================"

Write-Host ""
Write-Host "========================================"
Write-Host "Build Deployment Complete!"
Write-Host "========================================"
Write-Host "Version: 3.$FormattedBuild"
Write-Host "Output: $OutputDir"
Write-Host ""
Write-Host "Structure:"
Write-Host "- Annotatix.addin (manifest file)"
Write-Host "- annotatix_dependencies/"
Write-Host "  + main/ (Core: PluginsManager.dll + ALL dependencies + UI + icons)"
Write-Host "  + dwg2rvt/ (Module: dwg2rvt.Module.dll ONLY)"
Write-Host "  + hvac/ (Module: HVAC.Module.dll ONLY)"
Write-Host "  + family_sync/ (Module: FamilySync.Module.dll ONLY)"
Write-Host ""
Write-Host "NOTE: All shared dependencies are in main/ for optimal loading"
Write-Host "========================================"
