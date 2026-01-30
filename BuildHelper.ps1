# BuildHelper.ps1 - Helper to run build with proper encoding
$ErrorActionPreference = "Stop"
$OriginalEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    # Navigate to project directory
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    Set-Location $scriptPath
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Annotatix Quick Build" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Increment build number
    Write-Host "Incrementing version..." -ForegroundColor Yellow
    $buildNum = [int](Get-Content "BuildNumber.txt" -Raw).Trim()
    $buildNum++
    $buildNum | Set-Content "BuildNumber.txt" -NoNewline
    Write-Host "Version: 3.$buildNum" -ForegroundColor Green
    Write-Host ""
    
    # Build solution
    Write-Host "Building solution..." -ForegroundColor Yellow
    Write-Host "Running: dotnet msbuild dwg2rvt.sln /p:Configuration=Release /t:Rebuild /v:minimal" -ForegroundColor Gray
    Write-Host ""
    
    dotnet msbuild dwg2rvt.sln /p:Configuration=Release /t:Rebuild /v:minimal
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Build SUCCESSFUL! Version 3.$buildNum" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        $outputFolder = Join-Path (Split-Path $scriptPath) "annotatix_ver3.$($buildNum.ToString().PadLeft(3, '0'))"
        Write-Host "Output folder: $outputFolder" -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "Build FAILED!" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
    }
}
finally {
    [Console]::OutputEncoding = $OriginalEncoding
}
