using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Stores grid metadata from raster export for later use during image analysis.
    /// Saved as a JSON sidecar file alongside the exported image.
    /// </summary>
    public class GridMetadata
    {
        // ── Grid ──
        public int Cols { get; set; }
        public int Rows { get; set; }
        public double CellSizeMm { get; set; } = 3.0;

        // ── Paper (corrected to grid) ──
        public double PaperWidthMm { get; set; }
        public double PaperHeightMm { get; set; }

        // ── Image ──
        public int PixelWidth { get; set; }
        public int PixelHeight { get; set; }
        public double Dpi { get; set; } = 150.0;

        // ── Crop box in view-space coordinates (feet) ──
        public double ViewXmin { get; set; }
        public double ViewYmin { get; set; }
        public double ViewXmax { get; set; }
        public double ViewYmax { get; set; }

        // ── View origin & projection directions (for mapping grid → model) ──
        public double OriginX { get; set; }
        public double OriginY { get; set; }
        public double OriginZ { get; set; }
        public double RightDirX { get; set; }
        public double RightDirY { get; set; }
        public double RightDirZ { get; set; }
        public double UpDirX { get; set; }
        public double UpDirY { get; set; }
        public double UpDirZ { get; set; }
        public int ViewScale { get; set; }

        // ── ID scheme ──
        public string IdScheme { get; set; } = "row-major";

        // ── Analysis settings ──
        /// <summary>Minimum fill ratio (0.0-1.0) for a cell to be marked occupied.</summary>
        public double OccupancyThreshold { get; set; } = 0.10;

        /// <summary>
        /// Returns the cell ID for a given grid position (row-major).
        /// Cells are numbered left-to-right, top-to-bottom starting from 0.
        /// Format: id000000, id000001, etc. (padded to 6 digits).
        /// </summary>
        public string GetCellId(int row, int col)
        {
            int index = row * Cols + col;
            return $"id{index:D6}";
        }

        /// <summary>
        /// Parses a cell ID back to (row, col). Returns (-1, -1) on failure.
        /// </summary>
        public (int row, int col) ParseCellId(string cellId)
        {
            if (cellId != null && cellId.StartsWith("id") &&
                int.TryParse(cellId.Substring(2), out int index))
            {
                int col = index % Cols;
                int row = index / Cols;
                if (row < Rows)
                    return (row, col);
            }
            return (-1, -1);
        }

        /// <summary>
        /// Returns all cell IDs in row-major order.
        /// </summary>
        public List<string> GetAllCellIds()
        {
            var ids = new List<string>(Cols * Rows);
            for (int r = 0; r < Rows; r++)
                for (int c = 0; c < Cols; c++)
                    ids.Add(GetCellId(r, c));
            return ids;
        }

        // ── Serialization ──

        /// <summary>
        /// Saves metadata to a JSON sidecar file alongside the exported image.
        /// </summary>
        public void Save(string imagePath)
        {
            string jsonPath = GetSidecarPath(imagePath);
            string json = JsonConvert.SerializeObject(this, Formatting.Indented);
            File.WriteAllText(jsonPath, json);
            DebugLogger.Log($"[GRID-METADATA] Saved: {jsonPath}");
        }

        /// <summary>
        /// Loads metadata from the JSON sidecar of an exported image.
        /// </summary>
        public static GridMetadata Load(string imagePath)
        {
            string jsonPath = GetSidecarPath(imagePath);
            if (!File.Exists(jsonPath))
            {
                DebugLogger.Log($"[GRID-METADATA] Not found: {jsonPath}");
                return null;
            }
            string json = File.ReadAllText(jsonPath);
            var meta = JsonConvert.DeserializeObject<GridMetadata>(json);
            DebugLogger.Log($"[GRID-METADATA] Loaded: {jsonPath} ({meta.Cols}x{meta.Rows})");
            return meta;
        }

        private static string GetSidecarPath(string imagePath)
        {
            return Path.ChangeExtension(imagePath, ".grid.json");
        }
    }
}
