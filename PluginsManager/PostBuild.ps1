# Post-build deployment script for dwg2rvt plugin
param(
    [string]$ProjectDir,
    [string]$TargetDir,
    [string]$TargetFileName
)

Write-Host "========================================"
Write-Host "DWG2RVT Post-Build Deployment"
Write-Host "========================================"

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
$VersionFolder = "plugins_manager_ver3.$FormattedBuild"
$OutputDir = [System.IO.Path]::Combine($BaseOutput, $VersionFolder)

Write-Host "Output Directory: $OutputDir"

# Create version folder
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "Created directory: $OutputDir"
}

# Create plugins-manager-dll folder
$DllFolder = [System.IO.Path]::Combine($OutputDir, "plugins-manager-dll")
if (-not (Test-Path $DllFolder)) {
    New-Item -ItemType Directory -Path $DllFolder -Force | Out-Null
    Write-Host "Created directory: $DllFolder"
}

# Create plugins_manager_dependencies folder
$DependenciesFolder = [System.IO.Path]::Combine($OutputDir, "plugins_manager_dependencies")
if (-not (Test-Path $DependenciesFolder)) {
    New-Item -ItemType Directory -Path $DependenciesFolder -Force | Out-Null
    Write-Host "Created directory: $DependenciesFolder"
}

Write-Host ""
Write-Host "Copying files..."
Write-Host "----------------"

# Copy all files to plugins-manager-dll folder
Write-Host "Copying to plugins-manager-dll..."
Copy-Item -Path "$TargetDir*" -Destination $DllFolder -Force -Recurse
Write-Host "- Copied all build output to plugins-manager-dll"

# Copy all files to plugins_manager_dependencies folder (ensure all dependencies are present)
Write-Host "Copying to plugins_manager_dependencies..."
Copy-Item -Path "$TargetDir*" -Destination $DependenciesFolder -Force -Recurse
Write-Host "- Copied all build output to plugins_manager_dependencies"

# Copy BuildNumber.txt for version display
$BuildNumberSource = [System.IO.Path]::Combine($ProjectDir, "BuildNumber.txt")
if (Test-Path $BuildNumberSource) {
    Copy-Item -Path $BuildNumberSource -Destination $DependenciesFolder -Force
    Write-Host "- Copied BuildNumber.txt to plugins_manager_dependencies"
}

# Copy icons to the root of dependencies for easier access as per ADN-CIS recommendation
$SourceIcon32 = [System.IO.Path]::Combine($ProjectDir, "UI", "icons", "dwg2rvt32.png")
if (Test-Path $SourceIcon32) {
    Copy-Item -Path $SourceIcon32 -Destination $DependenciesFolder -Force
    Write-Host "- Copied dwg2rvt32.png to root of dependencies"
}
$SourceIcon80 = [System.IO.Path]::Combine($ProjectDir, "UI", "icons", "dwg2rvt80.png")
if (Test-Path $SourceIcon80) {
    Copy-Item -Path $SourceIcon80 -Destination $DependenciesFolder -Force
    Write-Host "- Copied dwg2rvt80.png to root of dependencies"
}
$SourceIcon64 = [System.IO.Path]::Combine($ProjectDir, "UI", "icons", "dwg2rvt.png")
if (Test-Path $SourceIcon64) {
    Copy-Item -Path $SourceIcon64 -Destination $DependenciesFolder -Force
    Write-Host "- Copied dwg2rvt.png to root of dependencies"
}

# Explicitly copy icons folder to dependencies
$SourceIcons = [System.IO.Path]::Combine($ProjectDir, "UI", "icons")
$DestIcons = [System.IO.Path]::Combine($DependenciesFolder, "UI", "icons")
if (Test-Path $SourceIcons) {
    if (-not (Test-Path $DestIcons)) {
        New-Item -ItemType Directory -Path $DestIcons -Force | Out-Null
    }
    Copy-Item -Path "$SourceIcons\*" -Destination $DestIcons -Force
    Write-Host "- Explicitly copied icons to $DestIcons"
}

# Create .addin file
$AddinFile = [System.IO.Path]::Combine($OutputDir, "PluginsManager.addin")
Write-Host "Creating .addin file..."

# Build the Assembly path dynamically to avoid encoding issues
$UserProfile = [Environment]::GetFolderPath('ApplicationData')
$AssemblyPath = [System.IO.Path]::Combine($UserProfile, "Autodesk", "Revit", "Addins", "2024", "plugins_manager_dependencies", "PluginsManager.dll")

# Create XML content
$AddinContent = "<?xml version=`"1.0`" encoding=`"utf-8`"?>`n"
$AddinContent += "<RevitAddIns>`n"
$AddinContent += "  <AddIn Type=`"Application`">`n"
$AddinContent += "    <Name>PluginsManager</Name>`n"
$AddinContent += "    <Assembly>$AssemblyPath</Assembly>`n"
$AddinContent += "    <ClientId>b1c2d3e4-f5a6-7890-1234-567890abcdef</ClientId>`n"
$AddinContent += "    <FullClassName>PluginsManager.App</FullClassName>`n"
$AddinContent += "    <VendorId>PLUGINS_MANAGER</VendorId>`n"
$AddinContent += "    <VendorDescription>Revit Plugins Manager with Dynamic Module Loading</VendorDescription>`n"
$AddinContent += "  </AddIn>`n"
$AddinContent += "</RevitAddIns>"

# Save with UTF-8 encoding (with BOM for better compatibility)
[System.IO.File]::WriteAllText($AddinFile, $AddinContent, [System.Text.Encoding]::UTF8)
Write-Host "- Created PluginsManager.addin"
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
Write-Host "   C:\Users\Свеж как огурец\AppData\Roaming\Autodesk\Revit\Addins\2024\plugins_manager_dependencies\"
Write-Host "3. Copy '$AddinFile' to:"
Write-Host "   C:\Users\Свеж как огурец\AppData\Roaming\Autodesk\Revit\Addins\2024\"
Write-Host "4. Restart Revit"
Write-Host "========================================"

Write-Host ""
Write-Host "========================================"
Write-Host "Build Deployment Complete!"
Write-Host "========================================"
Write-Host "Version: 3.$FormattedBuild"
Write-Host "Output: $OutputDir"
Write-Host ""
Write-Host "Root folder contents:"
Write-Host "- PluginsManager.addin (manifest file)"
Write-Host "- plugins_manager_dependencies\ (folder with all DLLs)"
Write-Host "- plugins-manager-dll\ (folder)"
Write-Host "========================================"
