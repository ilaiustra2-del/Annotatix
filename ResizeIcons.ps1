Add-Type -AssemblyName System.Drawing

$sourceDir = "UI\icons\new"
$targetDir = "PluginsManager\UI\icons"

# Ensure target directory exists
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

# Process each icon
$icons = @(
    @{Source="dwg2rvt.png"; Targets=@("dwg2rvt80.png")},
    @{Source="hvac.png"; Targets=@("hvac80.png")},
    @{Source="familysync.png"; Targets=@("familysync80.png")}
)

foreach ($icon in $icons) {
    $sourcePath = Join-Path $sourceDir $icon.Source
    
    if (Test-Path $sourcePath) {
        Write-Host "Processing: $($icon.Source)"
        
        # Load image
        $img = [System.Drawing.Image]::FromFile($sourcePath)
        
        foreach ($targetName in $icon.Targets) {
            $targetPath = Join-Path $targetDir $targetName
            
            # Create 80x80 bitmap
            $newImg = New-Object System.Drawing.Bitmap(80, 80)
            $graphics = [System.Drawing.Graphics]::FromImage($newImg)
            
            # Set high quality resize
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            
            # Draw resized image
            $graphics.DrawImage($img, 0, 0, 80, 80)
            
            # Save
            $newImg.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png)
            
            Write-Host "  Created: $targetName (80x80)"
            
            # Cleanup
            $graphics.Dispose()
            $newImg.Dispose()
        }
        
        $img.Dispose()
    } else {
        Write-Host "NOT FOUND: $sourcePath" -ForegroundColor Red
    }
}

Write-Host "`nIcon resize complete!" -ForegroundColor Green
