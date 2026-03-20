using System.Collections.Generic;

namespace ClashResolve.Module.Core
{
    /// <summary>
    /// MEP element kind — used by both ClashResolver and ClashLookupService.
    /// </summary>
    public enum MepKind { RoundDuct, RectDuct, Pipe }

    /// <summary>
    /// One row in a lookup table — parameters for a specific size.
    /// SizeMm = H (height) for RectDuct, D (diameter) for RoundDuct and Pipe.
    /// Null fields mean "use auto-calculated value".
    /// </summary>
    public class ClashLookupRow
    {
        public double  SizeMm { get; set; }
        public double? Drop30  { get; set; }
        public double? Drop45  { get; set; }
        public double? Drop60  { get; set; }
        public double? Drop90  { get; set; }
        public double? Seg30   { get; set; }
        public double? Seg45   { get; set; }
        public double? Seg60   { get; set; }
        public double? Seg90   { get; set; }

        /// <summary>
        /// Tracks which value cells were auto-calculated (not manually entered).
        /// Key: field name ("Drop30", "Seg45", etc.). True = auto, False/missing = manual override.
        /// </summary>
        public System.Collections.Generic.Dictionary<string, bool> AutoFlags { get; set; }
            = new System.Collections.Generic.Dictionary<string, bool>();
    }

    /// <summary>
    /// Lookup table for one MEP element type (RectDuct, RoundDuct or Pipe).
    /// </summary>
    public class ClashLookupTable
    {
        /// <summary>"RectDuct" | "RoundDuct" | "Pipe"</summary>
        public string MepType { get; set; }

        /// <summary>
        /// What to do when the exact size is not found in Rows:
        /// "Auto"    — fall back to formula-based calculation (return null).
        /// "Nearest" — use the nearest larger row; if none, nearest smaller.
        /// </summary>
        public string FallbackPolicy { get; set; } = "Auto";

        public List<ClashLookupRow> Rows { get; set; } = new List<ClashLookupRow>();
    }

    /// <summary>
    /// Root object serialised to lookup_tables.json.
    /// </summary>
    public class ClashLookupData
    {
        /// <summary>Up to 3 tables (one per MEP type).</summary>
        public List<ClashLookupTable> Tables { get; set; } = new List<ClashLookupTable>();

        /// <summary>Global enable flag — persisted so it survives Revit restarts.</summary>
        public bool GlobalEnabled { get; set; } = false;
    }
}
