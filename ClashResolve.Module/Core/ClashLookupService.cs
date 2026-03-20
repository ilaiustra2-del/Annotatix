using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using PluginsManager.Core;

namespace ClashResolve.Module.Core
{
    /// <summary>
    /// Singleton service managing lookup tables for clash resolution parameters.
    /// Tables are stored in {clash_resolve_dir}/lookup_tables.json and survive Revit restarts.
    /// </summary>
    public sealed class ClashLookupService
    {
        // ----------------------------------------------------------------
        // Singleton
        // ----------------------------------------------------------------
        private static ClashLookupService _instance;
        public static ClashLookupService Instance => _instance ?? (_instance = new ClashLookupService());

        // ----------------------------------------------------------------
        // Constructor — auto-loads persisted data immediately
        // ----------------------------------------------------------------
        private ClashLookupService()
        {
            Load();
        }

        // ----------------------------------------------------------------
        // State
        // ----------------------------------------------------------------
        private ClashLookupData _data = new ClashLookupData();

        /// <summary>Whether the lookup feature is globally enabled (mirrors the options-bar checkbox).</summary>
        public bool GlobalEnabled
        {
            get => _data.GlobalEnabled;
            set { _data.GlobalEnabled = value; Save(); }
        }

        // ----------------------------------------------------------------
        // File path
        // ----------------------------------------------------------------
        private static string GetFilePath()
        {
            // Resolve relative to the running DLL location:
            // {annotatix_dependencies}/clash_resolve/lookup_tables.json
            string dllDir = Path.GetDirectoryName(
                typeof(ClashLookupService).Assembly.Location);
            return Path.Combine(dllDir, "lookup_tables.json");
        }

        // ----------------------------------------------------------------
        // Persistence
        // ----------------------------------------------------------------
        public void Load()
        {
            try
            {
                string path = GetFilePath();
                if (!File.Exists(path))
                {
                    _data = new ClashLookupData();
                    DebugLogger.Log("[LOOKUP-SERVICE] No lookup_tables.json found, starting fresh.");
                    return;
                }
                string json = File.ReadAllText(path);
                _data = JsonConvert.DeserializeObject<ClashLookupData>(json) ?? new ClashLookupData();
                DebugLogger.Log($"[LOOKUP-SERVICE] Loaded {_data.Tables.Count} table(s), GlobalEnabled={_data.GlobalEnabled}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[LOOKUP-SERVICE] Load error: {ex.Message}");
                _data = new ClashLookupData();
            }
        }

        public void Save()
        {
            try
            {
                string path = GetFilePath();
                string dir  = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                string json = JsonConvert.SerializeObject(_data, Formatting.Indented);
                File.WriteAllText(path, json);
                DebugLogger.Log($"[LOOKUP-SERVICE] Saved {_data.Tables.Count} table(s).");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[LOOKUP-SERVICE] Save error: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------
        // Table management (called from UI)
        // ----------------------------------------------------------------

        /// <summary>Returns all tables (max 3).</summary>
        public IReadOnlyList<ClashLookupTable> GetAllTables() => _data.Tables;

        /// <summary>Returns the table for the given MEP type, or null.</summary>
        public ClashLookupTable GetTable(string mepType)
            => _data.Tables.FirstOrDefault(t => t.MepType == mepType);

        /// <summary>
        /// Creates a new empty table for the given MEP type.
        /// Returns false if a table for that type already exists.
        /// </summary>
        public bool CreateTable(string mepType)
        {
            if (_data.Tables.Any(t => t.MepType == mepType))
                return false;
            _data.Tables.Add(new ClashLookupTable { MepType = mepType });
            Save();
            return true;
        }

        /// <summary>Deletes the table for the given MEP type.</summary>
        public void DeleteTable(string mepType)
        {
            _data.Tables.RemoveAll(t => t.MepType == mepType);
            Save();
        }

        /// <summary>Replaces rows for the given table (called when user edits the DataGrid).</summary>
        public void UpdateRows(string mepType, List<ClashLookupRow> rows)
        {
            var table = GetTable(mepType);
            if (table == null) return;
            table.Rows = rows ?? new List<ClashLookupRow>();
            Save();
        }

        /// <summary>Sets the fallback policy for the given table.</summary>
        public void SetFallbackPolicy(string mepType, string policy)
        {
            var table = GetTable(mepType);
            if (table == null) return;
            table.FallbackPolicy = policy;
            Save();
        }

        // ----------------------------------------------------------------
        // Auto-calculation (called when user first enters a size value)
        // ----------------------------------------------------------------

        /// <summary>
        /// Fills Drop/Seg values for a given size using the same formulas as ClashResolver.
        /// Only fills fields that are currently marked as auto (AutoFlags[field]==true) or empty.
        /// Call this once when the user first sets SizeMm; never overwrite manual edits.
        /// </summary>
        public ClashLookupRow AutoCalculateRow(string mepType, double sizeMm)
        {
            // radius = half of size (mm) converted to feet
            double radiusMm = sizeMm / 2.0;
            const double mmPerFt = 304.8;
            double radiusFt = radiusMm / mmPerFt;

            // ---- Drop (clearance) per angle ----
            // Same logic as ResolveClash AutoClearance:
            //   RoundDuct / Pipe angled: 50mm   (+ 1 diameter for Pipe)
            //   RectDuct angled: radiusA * angleMult (0.5H@30, 1.5H@45, 2.5H@60, 9H@90)
            //   90°: 3.5 * max(R) for Round/Pipe,  9 * max(R) for Rect
            double drop30, drop45, drop60, drop90;
            double seg30, seg45, seg60, seg90;

            switch (mepType)
            {
                case "RectDuct":
                    drop30 = Math.Ceiling(radiusFt * 1.0 * mmPerFt);   // angleMult=1.0 @30°
                    drop45 = Math.Ceiling(radiusFt * 3.0 * mmPerFt);   // angleMult=3.0 @45°
                    drop60 = Math.Ceiling(radiusFt * 5.0 * mmPerFt);   // angleMult=5.0 @60°
                    drop90 = Math.Ceiling(radiusFt * 9.0 * mmPerFt);   // @90°
                    break;
                case "Pipe":
                    // angled: 50mm (no diameter bonus in angled mode); 90°: 3.5*R + 1D
                    drop30 = 50.0;
                    drop45 = 50.0;
                    drop60 = 50.0;
                    drop90 = Math.Ceiling(radiusFt * 3.5 * mmPerFt) + sizeMm;
                    break;
                default: // RoundDuct
                    drop30 = 50.0;
                    drop45 = 50.0;
                    drop60 = 50.0;
                    drop90 = Math.Ceiling(radiusFt * 3.5 * mmPerFt);
                    break;
            }

            // ---- Segment (flat section full length = 2 * flatHalf) ----
            // flatHalf = max(autoHalfLength, 2*radiusA)
            // autoHalfLength = radius * multiplier (3.0 angled, 2.5 @90)
            double halfMult   = mepType == "RectDuct" ? 3.0 : 3.0; // all same for angled
            double halfMult90 = mepType == "RectDuct" ? 6.0 : 2.5;
            double pipeBonus  = mepType == "Pipe" ? sizeMm : 0.0;

            double autoHalf30 = Math.Ceiling(radiusFt * halfMult * mmPerFt) + pipeBonus;
            double autoHalf45 = Math.Ceiling(radiusFt * halfMult * mmPerFt) + pipeBonus;
            double autoHalf60 = Math.Ceiling(radiusFt * halfMult * mmPerFt) + pipeBonus;
            double autoHalf90 = Math.Ceiling(radiusFt * halfMult90 * mmPerFt) + pipeBonus;

            double minFlat = 2.0 * radiusMm; // 2*radiusFt in mm = sizeMm (1 diameter)

            // expandFt per angle (using drop as the vertical shift approximation)
            double expand30 = drop30 / Math.Tan(30.0 * Math.PI / 180.0);
            double expand45 = drop45 / Math.Tan(45.0 * Math.PI / 180.0);
            double expand60 = drop60 / Math.Tan(60.0 * Math.PI / 180.0);

            double flatHalf30 = Math.Max(autoHalf30, minFlat);
            double flatHalf45 = Math.Max(autoHalf45, minFlat);
            double flatHalf60 = Math.Max(autoHalf60, minFlat);
            double flatHalf90 = Math.Max(autoHalf90, minFlat);

            // Full segment = 2 * splitOffset = 2 * (flatHalf + expand)
            seg30 = Math.Ceiling(2.0 * (flatHalf30 + expand30));
            seg45 = Math.Ceiling(2.0 * (flatHalf45 + expand45));
            seg60 = Math.Ceiling(2.0 * (flatHalf60 + expand60));
            seg90 = Math.Ceiling(2.0 * flatHalf90);  // @90° no expand

            var row = new ClashLookupRow
            {
                SizeMm = sizeMm,
                Drop30 = drop30, Drop45 = drop45, Drop60 = drop60, Drop90 = drop90,
                Seg30  = seg30,  Seg45  = seg45,  Seg60  = seg60,  Seg90  = seg90,
            };
            // Mark all as auto-calculated
            foreach (var f in new[] { "Drop30","Drop45","Drop60","Drop90","Seg30","Seg45","Seg60","Seg90" })
                row.AutoFlags[f] = true;
            return row;
        }

        // ----------------------------------------------------------------
        // Lookup API (called from ClashResolver)
        // ----------------------------------------------------------------

        /// <summary>
        /// Returns the drop (clearance) value in mm for the given element kind,
        /// size and angle, or null if no table entry applies.
        /// </summary>
        public double? LookupDrop(MepKind kind, double sizeMm, double angleDeg)
        {
            var row = FindRow(MepKindToString(kind), sizeMm);
            if (row == null) return null;
            return SelectDropField(row, angleDeg);
        }

        /// <summary>
        /// Returns the segment length (full, not half) in mm, or null if no entry applies.
        /// </summary>
        public double? LookupSegment(MepKind kind, double sizeMm, double angleDeg)
        {
            var row = FindRow(MepKindToString(kind), sizeMm);
            if (row == null) return null;
            return SelectSegField(row, angleDeg);
        }

        // ----------------------------------------------------------------
        // Internal helpers
        // ----------------------------------------------------------------

        private ClashLookupRow FindRow(string mepType, double sizeMm)
        {
            var table = GetTable(mepType);
            if (table == null || table.Rows.Count == 0) return null;

            // Exact match (±0.5 mm tolerance)
            var exact = table.Rows.FirstOrDefault(r => Math.Abs(r.SizeMm - sizeMm) <= 0.5);
            if (exact != null) return exact;

            // Fallback
            if (table.FallbackPolicy == "Nearest")
            {
                // Nearest larger
                var larger = table.Rows
                    .Where(r => r.SizeMm > sizeMm)
                    .OrderBy(r => r.SizeMm)
                    .FirstOrDefault();
                if (larger != null) return larger;

                // Nearest smaller
                var smaller = table.Rows
                    .Where(r => r.SizeMm < sizeMm)
                    .OrderByDescending(r => r.SizeMm)
                    .FirstOrDefault();
                return smaller;
            }

            // "Auto" policy — no match
            return null;
        }

        private static double? SelectDropField(ClashLookupRow row, double angle)
        {
            if (Math.Abs(angle - 30.0) < 0.5) return row.Drop30;
            if (Math.Abs(angle - 45.0) < 0.5) return row.Drop45;
            if (Math.Abs(angle - 60.0) < 0.5) return row.Drop60;
            return row.Drop90;   // 90° and anything else
        }

        private static double? SelectSegField(ClashLookupRow row, double angle)
        {
            if (Math.Abs(angle - 30.0) < 0.5) return row.Seg30;
            if (Math.Abs(angle - 45.0) < 0.5) return row.Seg45;
            if (Math.Abs(angle - 60.0) < 0.5) return row.Seg60;
            return row.Seg90;
        }

        public static string MepKindToString(MepKind kind)
        {
            switch (kind)
            {
                case MepKind.RectDuct:  return "RectDuct";
                case MepKind.Pipe:      return "Pipe";
                default:                return "RoundDuct";
            }
        }

        public static string MepKindToDisplayName(string mepType)
        {
            switch (mepType)
            {
                case "RectDuct":  return "Прямоугольные воздуховоды";
                case "RoundDuct": return "Круглые воздуховоды";
                case "Pipe":      return "Трубы";
                default:          return mepType;
            }
        }

        public static string DisplayNameToMepType(string displayName)
        {
            switch (displayName)
            {
                case "Прямоугольные воздуховоды": return "RectDuct";
                case "Круглые воздуховоды":       return "RoundDuct";
                case "Трубы":                     return "Pipe";
                default:                          return displayName;
            }
        }
    }
}
