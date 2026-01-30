# Build and Deploy Script for PluginsManager
# Automatically increments version, builds solution and creates deployment package

param(
    [switch]$SkipBuild
)

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PluginsManager Build & Deploy" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Increment build number
if (-not $SkipBuild) {
    Write-Host "[1/4] Incrementing build number..." -ForegroundColor Yellow
    $buildFile = "BuildNumber.txt"
    $currentBuild = Get-Content $buildFile -Raw
    $newBuild = [int]$currentBuild.Trim() + 1
    $newBuild.ToString() | Set-Content $buildFile -NoNewline
    Write-Host "      Build: $currentBuild -> $newBuild" -ForegroundColor Green
    Write-Host ""
}

# Step 2: Build solution
if (-not $SkipBuild) {
    Write-Host "[2/4] Building solution..." -ForegroundColor Yellow
    $buildResult = dotnet msbuild dwg2rvt.sln /p:Configuration=Release /t:Rebuild /v:minimal 2>&1
    
    # Check if build succeeded
    if ($LASTEXITCODE -eq 0) {
        Write-Host "      ✓ Build succeeded" -ForegroundColor Green
    } else {
        Write-Host "      ✗ Build failed!" -ForegroundColor Red
        Write-Host $buildResult
        exit 1
    }
    Write-Host ""
}

# Step 3: Run PostBuild script
Write-Host "[3/4] Creating deployment package..." -ForegroundColor Yellow
$projectDir = (Get-Location).Path
.\PostBuild.ps1 -ProjectDir $projectDir -TargetDir ".\bin\Release\" -TargetFileName "PluginsManager.dll" 2>&1 | Out-Null

$buildNumber = (Get-Content "BuildNumber.txt" -Raw).Trim()
$versionFolder = "plugins_manager_ver3.$($buildNumber.PadLeft(3, '0'))"
$fullPath = Join-Path (Split-Path $projectDir -Parent) $versionFolder

if (Test-Path $fullPath) {
    Write-Host "      ✓ Package created: $versionFolder" -ForegroundColor Green
} else {
    Write-Host "      ✗ Package creation failed!" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Step 4: Show completion message (no auto-deploy to Revit)
Write-Host "[4/4] Deployment package ready" -ForegroundColor Yellow
Write-Host "      📦 Package location: $versionFolder" -ForegroundColor Cyan
Write-Host ""

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ✅ Version 3.$($buildNumber.PadLeft(3, '0')) ready!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📁 Location: $fullPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠️  Manual installation required:" -ForegroundColor Yellow
Write-Host "   1. Copy contents of 'plugins_manager_dependencies\' to:" -ForegroundColor Gray
Write-Host "      %APPDATA%\Autodesk\Revit\Addins\2024\plugins_manager_dependencies\" -ForegroundColor Gray
Write-Host "   2. Copy 'PluginsManager.addin' to:" -ForegroundColor Gray
Write-Host "      %APPDATA%\Autodesk\Revit\Addins\2024\" -ForegroundColor Gray
Write-Host "   3. Restart Revit" -ForegroundColor Gray
Write-Host ""
