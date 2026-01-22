# Build number management script for dwg2rvt plugin
Write-Host "========================================"
Write-Host "DWG2RVT Build Number Manager"
Write-Host "========================================"

# Path to build number file
$BuildNumberFile = Join-Path $PSScriptRoot "BuildNumber.txt"

# Read current build number
if (Test-Path $BuildNumberFile) {
    $BuildNumber = [int](Get-Content $BuildNumberFile -Raw).Trim()
} else {
    $BuildNumber = 1
}

Write-Host "Current build number: $BuildNumber"

# New build number
$NewBuildNumber = $BuildNumber + 1
Set-Content -Path $BuildNumberFile -Value $NewBuildNumber

# Update AssemblyInfo.cs
$AssemblyInfoPath = Join-Path -Path $PSScriptRoot -ChildPath "Properties\AssemblyInfo.cs"
if (Test-Path $AssemblyInfoPath) {
    $Content = Get-Content $AssemblyInfoPath
    $NewVersion = "2.$NewBuildNumber.0.0"
    $Content = $Content -replace '\[assembly: AssemblyVersion\(.*\)\]', "[assembly: AssemblyVersion(""$NewVersion"")]"
    $Content = $Content -replace '\[assembly: AssemblyFileVersion\(.*\)\]', "[assembly: AssemblyFileVersion(""$NewVersion"")]"
    
    # Fallback to simple replace if regex fails or for different line endings
    if ($Content -notcontains $NewVersion) {
        $Content = $Content.Replace('1.0.0.0', $NewVersion)
    }
    
    Set-Content $AssemblyInfoPath $Content -Encoding UTF8
    Write-Host "Updated AssemblyInfo.cs to version $NewVersion"
}

Write-Host "New build number: $NewBuildNumber"
Write-Host "Build number updated successfully."
Write-Host "========================================"
