using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Detailed record of a single placement iteration for one annotation
    /// </summary>
    public class PlacementIterationRecord
    {
        public int AnnotationIndex { get; set; }
        public long ElementId { get; set; }
        public int GlobalIteration { get; set; }
        public string PositionName { get; set; } // "Vpravo vverh", "Gorizontalno vpravo", etc.
        public double LeaderEndViewX { get; set; }
        public double LeaderEndViewY { get; set; }
        public double ElbowViewX { get; set; }
        public double ElbowViewY { get; set; }
        public double HeaderViewX { get; set; }
        public double HeaderViewY { get; set; }
        public double ElbowHeightMm { get; set; }
        public bool HeaderCollision { get; set; }
        public bool LeaderCollision { get; set; }
        public string CollidingElementIds { get; set; } // comma-separated
        public string CollisionTypes { get; set; } // "header", "leader", "header+leader"
        public bool PlacementSucceeded { get; set; }
    }
    
    /// <summary>
    /// Summary record of one annotation placement attempt (final result)
    /// </summary>
    public class AnnotationPlacementRecord
    {
        public int AnnotationIndex { get; set; }
        public long ElementId { get; set; }
        public string ElementCategory { get; set; }
        public string ElementFamily { get; set; }
        public string ElementType { get; set; }
        public double LocationModelX { get; set; }
        public double LocationModelY { get; set; }
        public double LocationModelZ { get; set; }
        public double LocationViewX { get; set; }
        public double LocationViewY { get; set; }
        public string OrientationSymbol { get; set; } // "\\", "/", "--", "|"  (ASCII)
        public string OrientationDescription { get; set; } // "(x1<x2), (y1>y2)"
        public string AnnotationFamily { get; set; }
        public string AnnotationType { get; set; }
        public string AnnotationContentType { get; set; } // DuctRoundSizeFlow, etc.
        public string TagText { get; set; }
        public string FinalTypeName { get; set; } // actual type name after shelf length change
        public double ShelfLengthMm { get; set; }
        public double CalculatedWidthMm { get; set; }
        public bool Success { get; set; }
        public string FinalPosition { get; set; } // "TopRight", "HorizontalRight", etc.
        public double FinalElbowHeightMm { get; set; }
        public int TotalIterations { get; set; }
        public string FailureReason { get; set; }
    }

    /// <summary>
    /// Exports all annotation placement data to a single structured CSV file for analysis.
    /// Uses UTF-8 with BOM for proper Excel encoding detection.
    /// All Unicode symbols replaced with ASCII equivalents for maximum compatibility.
    /// Three sections in one file: ELEMENTS, ANNOTATIONS, ITERATIONS.
    /// </summary>
    public static class AnnotationCsvExporter
    {
        // UTF-8 with BOM - forces Excel to recognize encoding correctly
        private static readonly Encoding Utf8Bom = new UTF8Encoding(true);
        
        /// <summary>
        /// Export all placement data to a single CSV file in the logs directory
        /// </summary>
        public static void Export(
            ViewSnapshot snapshot,
            List<AnnotationPlacementRecord> placementRecords,
            List<PlacementIterationRecord> iterationRecords,
            string logsDirectory)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string filePath = Path.Combine(logsDirectory, $"annotatix_report_{timestamp}.csv");
                
                var sb = new StringBuilder();
                
                // ═══════════════════════════════════════════════
                // SECTION 1: VIEW INFO
                // ═══════════════════════════════════════════════
                sb.AppendLine($"# Annotatix Placement Report");
                sb.AppendLine($"# Project: {snapshot.DocumentName}");
                sb.AppendLine($"# View: {snapshot.ViewName}");
                sb.AppendLine($"# ViewType: {snapshot.ViewType}");
                sb.AppendLine($"# ViewScale: {snapshot.ViewScaleString} ({snapshot.ViewScale})");
                sb.AppendLine($"# Timestamp: {snapshot.Timestamp:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"# Elements: {snapshot.Elements.Count}");
                sb.AppendLine($"# Annotations placed: {placementRecords.Where(r => r.Success).Count()}");
                sb.AppendLine($"# Annotations failed: {placementRecords.Where(r => !r.Success).Count()}");
                sb.AppendLine();
                
                // ═══════════════════════════════════════════════
                // SECTION 2: ELEMENTS
                // ═══════════════════════════════════════════════
                sb.AppendLine("=== ELEMENTS ===");
                sb.AppendLine("ElementId;Category;FamilyName;TypeName;" +
                              "StartModelX;StartModelY;StartModelZ;StartViewX;StartViewY;" +
                              "EndModelX;EndModelY;EndModelZ;EndViewX;EndViewY;HasEndPoint;" +
                              "Diameter;Width;Height;SizeDisplay;Slope;SlopeDisplay;" +
                              "SystemId;SystemName;BelongTo");
                
                foreach (var elem in snapshot.Elements)
                {
                    sb.AppendLine($"{elem.ElementId};{Esc(elem.Category)};{Esc(elem.FamilyName)};{Esc(elem.TypeName)};" +
                                  $"{Fmt(elem.ModelStart.X)};{Fmt(elem.ModelStart.Y)};{Fmt(elem.ModelStart.Z)};" +
                                  $"{Fmt(elem.ViewStart.X)};{Fmt(elem.ViewStart.Y)};" +
                                  $"{Fmt(elem.ModelEnd.X)};{Fmt(elem.ModelEnd.Y)};{Fmt(elem.ModelEnd.Z)};" +
                                  $"{Fmt(elem.ViewEnd.X)};{Fmt(elem.ViewEnd.Y)};" +
                                  $"{elem.HasEndPoint};" +
                                  $"{FmtNull(elem.Diameter)};{FmtNull(elem.Width)};{FmtNull(elem.Height)};{Esc(elem.SizeDisplay)};" +
                                  $"{FmtNull(elem.Slope)};{Esc(elem.SlopeDisplay)};" +
                                  $"{FmtNullLong(elem.SystemId)};{Esc(elem.SystemName)};{Esc(elem.BelongTo)}");
                }

                sb.AppendLine();
                
                // ═══════════════════════════════════════════════
                // SECTION 3: ANNOTATION SUMMARY
                // ═══════════════════════════════════════════════
                sb.AppendLine("=== ANNOTATIONS ===");
                sb.AppendLine("AnnotationIndex;ElementId;ElementCategory;ElementFamily;ElementType;" +
                              "LocationModelX;LocationModelY;LocationModelZ;LocationViewX;LocationViewY;" +
                              "Orientation;OrientationDesc;" +
                              "AnnotationFamily;AnnotationType;AnnotationContentType;" +
                              "TagText;FinalTypeName;ShelfLengthMm;CalculatedWidthMm;" +
                              "Success;FinalPosition;FinalElbowHeightMm;TotalIterations;FailureReason");
                
                foreach (var r in placementRecords)
                {
                    sb.AppendLine($"{r.AnnotationIndex};{r.ElementId};{Esc(r.ElementCategory)};{Esc(r.ElementFamily)};{Esc(r.ElementType)};" +
                                  $"{Fmt(r.LocationModelX)};{Fmt(r.LocationModelY)};{Fmt(r.LocationModelZ)};" +
                                  $"{Fmt(r.LocationViewX)};{Fmt(r.LocationViewY)};" +
                                  $"{r.OrientationSymbol};{Esc(r.OrientationDescription)};" +
                                  $"{Esc(r.AnnotationFamily)};{Esc(r.AnnotationType)};{r.AnnotationContentType};" +
                                  $"{Esc(r.TagText)};{Esc(r.FinalTypeName)};{Fmt(r.ShelfLengthMm)};{Fmt(r.CalculatedWidthMm)};" +
                                  $"{r.Success};{Esc(r.FinalPosition)};{Fmt(r.FinalElbowHeightMm)};{r.TotalIterations};{Esc(r.FailureReason)}");
                }

                sb.AppendLine();
                
                // ═══════════════════════════════════════════════
                // SECTION 4: PLACEMENT ITERATIONS
                // ═══════════════════════════════════════════════
                sb.AppendLine("=== ITERATIONS ===");
                sb.AppendLine("AnnotationIndex;ElementId;GlobalIteration;PositionName;" +
                              "LeaderEndViewX;LeaderEndViewY;ElbowViewX;ElbowViewY;HeaderViewX;HeaderViewY;" +
                              "ElbowHeightMm;HeaderCollision;LeaderCollision;" +
                              "CollidingElementIds;CollisionTypes;PlacementSucceeded");
                
                foreach (var r in iterationRecords)
                {
                    sb.AppendLine($"{r.AnnotationIndex};{r.ElementId};{r.GlobalIteration};{Esc(r.PositionName)};" +
                                  $"{Fmt(r.LeaderEndViewX)};{Fmt(r.LeaderEndViewY)};" +
                                  $"{Fmt(r.ElbowViewX)};{Fmt(r.ElbowViewY)};" +
                                  $"{Fmt(r.HeaderViewX)};{Fmt(r.HeaderViewY)};" +
                                  $"{Fmt(r.ElbowHeightMm)};{r.HeaderCollision};{r.LeaderCollision};" +
                                  $"{Esc(r.CollidingElementIds)};{Esc(r.CollisionTypes)};{r.PlacementSucceeded}");
                }
                
                // Write with BOM for Excel UTF-8 detection
                File.WriteAllText(filePath, sb.ToString(), Utf8Bom);
                DebugLogger.Log($"[ANNOTATION-CSV-EXPORTER] Exported report: {filePath}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATION-CSV-EXPORTER] Error: {ex.Message}");
            }
        }
        
        // Formatting helpers
        private static string Fmt(double val) => val.ToString("F3");
        private static string FmtNull(double? val) => val.HasValue ? val.Value.ToString("F4") : "";
        private static string FmtNullLong(long? val) => val.HasValue ? val.Value.ToString() : "";
        private static string Esc(string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            // If contains semicolon, quote, or newline - wrap in quotes
            if (val.Contains(";") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r"))
                return "\"" + val.Replace("\"", "\"\"") + "\"";
            return val;
        }
    }
}
