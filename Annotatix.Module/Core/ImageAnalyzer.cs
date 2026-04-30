using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Result of cell occupancy analysis.
    /// </summary>
    public class CellOccupancy
    {
        public string CellId { get; set; }
        public int Row { get; set; }
        public int Col { get; set; }
        /// <summary>true = cell has visible content (non-white pixels above threshold)</summary>
        public bool IsOccupied { get; set; }
        /// <summary>Percentage of non-white pixels in the cell (0.0 – 1.0)</summary>
        public double FillRatio { get; set; }
    }

    /// <summary>
    /// Analyzes an exported raster image against its grid metadata.
    /// Detects which grid cells are occupied vs empty based on pixel content.
    /// </summary>
    public class ImageAnalyzer
    {
        private readonly GridMetadata _meta;

        /// <summary>Pixel white threshold (0-255). Pixels with all channels &gt;= this are "white".</summary>
        public byte WhiteThreshold { get; set; } = 240;

        /// <summary>Minimum fill ratio (0.0-1.0) for a cell to be considered occupied.</summary>
        public double OccupancyThreshold { get; set; }

        public ImageAnalyzer(GridMetadata meta)
        {
            _meta = meta ?? throw new ArgumentNullException(nameof(meta));
            // Use threshold from metadata (configurable per-export), fallback to 10%
            OccupancyThreshold = _meta.OccupancyThreshold > 0 ? _meta.OccupancyThreshold : 0.10;
        }

        /// <summary>
        /// Analyzes an exported PNG image and returns occupancy for each grid cell.
        /// </summary>
        public List<CellOccupancy> AnalyzeImage(string imagePath)
        {
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Image not found", imagePath);

            DebugLogger.Log($"[IMAGE-ANALYZER] Analyzing: {imagePath}");
            DebugLogger.Log($"[IMAGE-ANALYZER] Grid: {_meta.Cols}x{_meta.Rows}, " +
                $"Image: {_meta.PixelWidth}x{_meta.PixelHeight}");

            using (Bitmap bitmap = new Bitmap(imagePath))
            {
                // Verify image dimensions match metadata
                if (bitmap.Width != _meta.PixelWidth || bitmap.Height != _meta.PixelHeight)
                {
                    DebugLogger.Log($"[IMAGE-ANALYZER] WARNING: Image dimensions ({bitmap.Width}x{bitmap.Height}) " +
                        $"differ from metadata ({_meta.PixelWidth}x{_meta.PixelHeight}). Using actual dimensions.");
                }

                int imgW = bitmap.Width;
                int imgH = bitmap.Height;

                double cellW = (double)imgW / _meta.Cols;
                double cellH = (double)imgH / _meta.Rows;

                DebugLogger.Log($"[IMAGE-ANALYZER] Cell size: {cellW:F2}x{cellH:F2} px");

                // Lock bitmap and read pixel data
                BitmapData data = bitmap.LockBits(
                    new Rectangle(0, 0, imgW, imgH),
                    ImageLockMode.ReadOnly,
                    PixelFormat.Format24bppRgb);

                try
                {
                    int stride = data.Stride;
                    int bytesPerPixel = 3; // 24bpp
                    byte[] pixels = new byte[stride * imgH];
                    Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);

                    var results = new List<CellOccupancy>(_meta.Cols * _meta.Rows);

                    for (int row = 0; row < _meta.Rows; row++)
                    {
                        for (int col = 0; col < _meta.Cols; col++)
                        {
                            // Calculate pixel bounds for this cell
                            int pxStart = (int)Math.Floor(col * cellW);
                            int pxEnd = (int)Math.Ceiling((col + 1) * cellW);
                            int pyStart = (int)Math.Floor(row * cellH);
                            int pyEnd = (int)Math.Ceiling((row + 1) * cellH);

                            // Clamp to image bounds
                            pxEnd = Math.Min(pxEnd, imgW);
                            pyEnd = Math.Min(pyEnd, imgH);

                            int totalPixels = (pxEnd - pxStart) * (pyEnd - pyStart);
                            int nonWhitePixels = 0;

                            for (int py = pyStart; py < pyEnd; py++)
                            {
                                int rowOffset = py * stride;
                                for (int px = pxStart; px < pxEnd; px++)
                                {
                                    int idx = rowOffset + px * bytesPerPixel;
                                    // PixelFormat.Format24bppRgb: B, G, R order
                                    byte b = pixels[idx];
                                    byte g = pixels[idx + 1];
                                    byte r = pixels[idx + 2];

                                    if (r < WhiteThreshold || g < WhiteThreshold || b < WhiteThreshold)
                                    {
                                        nonWhitePixels++;
                                    }
                                }
                            }

                            double fillRatio = totalPixels > 0 ? (double)nonWhitePixels / totalPixels : 0.0;
                            bool occupied = fillRatio >= OccupancyThreshold;

                            results.Add(new CellOccupancy
                            {
                                CellId = _meta.GetCellId(row, col),
                                Row = row,
                                Col = col,
                                IsOccupied = occupied,
                                FillRatio = fillRatio
                            });
                        }
                    }

                    DebugLogger.Log($"[IMAGE-ANALYZER] Analysis complete: " +
                        $"{results.FindAll(c => c.IsOccupied).Count} occupied / " +
                        $"{results.Count} total cells");

                    return results;
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }
            }
        }

        /// <summary>
        /// Saves the analysis results to a CSV file alongside the image.
        /// </summary>
        public void SaveResults(List<CellOccupancy> results, string imagePath)
        {
            string csvPath = Path.ChangeExtension(imagePath, "_analysis.csv");

            using (var sw = new StreamWriter(csvPath, false))
            {
                sw.WriteLine("CellId,Row,Col,IsOccupied,FillRatio");
                foreach (var cell in results)
                {
                    sw.WriteLine($"{cell.CellId},{cell.Row},{cell.Col},{cell.IsOccupied},{cell.FillRatio:F6}");
                }
            }

            DebugLogger.Log($"[IMAGE-ANALYZER] Results saved: {csvPath}");
        }

        /// <summary>
        /// Creates a visual debug image showing occupied vs empty cells.
        /// Occupied cells are tinted red, empty cells green.
        /// </summary>
        public void CreateDebugOverlay(List<CellOccupancy> results, string imagePath)
        {
            string debugPath = Path.ChangeExtension(imagePath, "_debug.png");

            using (Bitmap original = new Bitmap(imagePath))
            using (Bitmap debug = new Bitmap(original))
            using (Graphics g = Graphics.FromImage(debug))
            {
                int imgW = debug.Width;
                int imgH = debug.Height;
                double cellW = (double)imgW / _meta.Cols;
                double cellH = (double)imgH / _meta.Rows;

                // Semi-transparent brush for occupied cells
                SolidBrush occupiedBrush = new SolidBrush(Color.FromArgb(80, Color.Red));
                SolidBrush emptyBrush = new SolidBrush(Color.FromArgb(40, Color.Green));
                Pen borderPen = new Pen(Color.FromArgb(120, Color.Gray), 1);

                foreach (var cell in results)
                {
                    int px = (int)Math.Floor(cell.Col * cellW);
                    int py = (int)Math.Floor(cell.Row * cellH);
                    int pw = (int)Math.Ceiling((cell.Col + 1) * cellW) - px;
                    int ph = (int)Math.Ceiling((cell.Row + 1) * cellH) - py;

                    Rectangle rect = new Rectangle(px, py, pw, ph);
                    g.FillRectangle(cell.IsOccupied ? occupiedBrush : emptyBrush, rect);
                    g.DrawRectangle(borderPen, rect);
                }

                debug.Save(debugPath, ImageFormat.Png);
                occupiedBrush.Dispose();
                emptyBrush.Dispose();
                borderPen.Dispose();
            }

            DebugLogger.Log($"[IMAGE-ANALYZER] Debug overlay saved: {debugPath}");
        }
    }
}
