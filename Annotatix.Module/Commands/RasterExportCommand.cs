using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Annotatix.Module.Core;
using PluginsManager.Core;

namespace Annotatix.Module.Commands
{
    /// <summary>
    /// Raster export command: finds model element bounds, sets crop box,
    /// adjusts to 3mm paper-space grid, exports view as PNG
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class RasterExportCommand : IExternalCommand
    {
        // Constants
        private const double DPI = 150.0;
        private const double MM_PER_FT = 304.8;
        private const double PIXELS_PER_MM = DPI / 25.4;

        // Whitelist of physical MEP/building element categories used for bounds calculation.
        // Excludes system-level categories (DuctSystems, PipeSystems, ProjectBasePoint, etc.)
        // whose bounding boxes/geometry vastly overestimate view extents.
        private static readonly HashSet<BuiltInCategory> s_physicalCategories = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_GenericModel,
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uidoc = uiApp.ActiveUIDocument;
                var doc = uidoc.Document;
                var view = doc.ActiveView;

                if (view == null || view.IsTemplate)
                {
                    TaskDialog.Show("Raster Export", "No active view found.");
                    return Result.Failed;
                }

                // Check view type - only plan views and 3D views supported for now
                if (view.ViewType != ViewType.FloorPlan &&
                    view.ViewType != ViewType.CeilingPlan &&
                    view.ViewType != ViewType.EngineeringPlan &&
                    view.ViewType != ViewType.AreaPlan &&
                    view.ViewType != ViewType.Section &&
                    view.ViewType != ViewType.ThreeD)
                {
                    TaskDialog.Show("Raster Export",
                        "Unsupported view type. Please use a plan, section, or 3D view.");
                    return Result.Failed;
                }

                DebugLogger.Log($"[RASTER-EXPORT] Starting raster export for view: '{view.Name}', Scale: 1:{view.Scale}");
                DebugLogger.Log($"[RASTER-EXPORT] view.Origin=({view.Origin.X:F4},{view.Origin.Y:F4},{view.Origin.Z:F4})");
                DebugLogger.Log($"[RASTER-EXPORT] view.RightDir=({view.RightDirection.X:F6},{view.RightDirection.Y:F6},{view.RightDirection.Z:F6})");
                DebugLogger.Log($"[RASTER-EXPORT] view.UpDir=({view.UpDirection.X:F6},{view.UpDirection.Y:F6},{view.UpDirection.Z:F6})");

                // ── Step 1-2: Find element bounds using geometry projection ──
                // Instead of bounding boxes (which overestimate in 3D views),
                // project actual element geometry vertices to the view plane.
                // This gives tight 2D bounds matching visible element extents.
                double screenXmin = double.MaxValue, screenXmax = double.MinValue;
                double screenYmin = double.MaxValue, screenYmax = double.MinValue;
                bool foundAny = false;

                // Track which elements define each extreme boundary
                string xminElem = null, xmaxElem = null;
                string yminElem = null, ymaxElem = null;

                FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id);
                ICollection<Element> allElements = collector.WhereElementIsNotElementType().ToElements();

                // Get view projection vectors (same as ConvertModelToViewCoordinates)
                XYZ viewOrigin = view.Origin;
                XYZ viewRight = view.RightDirection;
                XYZ viewUp = view.UpDirection;

                // Geometry options for extracting element vertices
                // NOTE: Cannot combine View and DetailLevel in Options —
                // Revit API rejects setting one after the other, regardless of order.
                // DetailLevel.Coarse is used WITHOUT View filtering; the
                // FilteredElementCollector(doc, view.Id) already handles visibility.
                Options geomOptions = new Options
                {
                    DetailLevel = ViewDetailLevel.Coarse,
                    ComputeReferences = false
                };

                int elementCount = 0;
                int filteredCount = 0;
                int bboxFallbackCount = 0;

                foreach (Element elem in allElements)
                {
                    // Skip non-model elements (annotations, datums, etc.)
                    if (elem.Category == null)
                    {
                        filteredCount++;
                        continue;
                    }

                    Category cat = elem.Category;
                    if (cat.CategoryType != CategoryType.Model)
                    {
                        filteredCount++;
                        continue;
                    }

                    // Skip datum elements that have CategoryType.Model but infinite extents
                    BuiltInCategory bic = (BuiltInCategory)cat.Id.IntegerValue;
                    if (bic == BuiltInCategory.OST_Levels ||
                        bic == BuiltInCategory.OST_Grids)
                    {
                        filteredCount++;
                        continue;
                    }

                    // Include only physical MEP/building element categories.
                    // Non-physical elements (DuctSystems, ProjectBasePoint, etc.)
                    // have oversized bounding boxes that would inflate the crop bounds.
                    if (!s_physicalCategories.Contains(bic))
                    {
                        filteredCount++;
                        continue;
                    }

                    // Try extracting actual geometry first
                    bool hasGeometry = false;
                    double elemMinX = double.MaxValue, elemMaxX = double.MinValue;
                    double elemMinY = double.MaxValue, elemMaxY = double.MinValue;
                    try
                    {
                        GeometryElement geomElem = elem.get_Geometry(geomOptions);
                        if (geomElem != null)
                        {
                            int vertexCount = 0;
                            foreach (GeometryObject childObj in geomElem)
                            {
                                ProjectGeometry(childObj, Transform.Identity,
                                    viewOrigin, viewRight, viewUp,
                                    ref elemMinX, ref elemMaxX,
                                    ref elemMinY, ref elemMaxY,
                                    ref vertexCount);
                            }
                            if (vertexCount > 0)
                            {
                                hasGeometry = true;
                                elementCount++;
                                foundAny = true;
                                
                                // Update global bounds with contributor tracking
                                string elemLabel = $"GEOM#{elementCount} Elem {elem.Id}({cat.Name})";
                                if (elemMinX < screenXmin) { screenXmin = elemMinX; xminElem = elemLabel; }
                                if (elemMaxX > screenXmax) { screenXmax = elemMaxX; xmaxElem = elemLabel; }
                                if (elemMinY < screenYmin) { screenYmin = elemMinY; yminElem = elemLabel; }
                                if (elemMaxY > screenYmax) { screenYmax = elemMaxY; ymaxElem = elemLabel; }
                                
                                // Log: first 5 always; then elements far from origin
                                double elemW = elemMaxX - elemMinX;
                                double elemH = elemMaxY - elemMinY;
                                bool isFar = Math.Abs(elemMinX) > 20.0 || Math.Abs(elemMaxX) > 20.0 || 
                                             Math.Abs(elemMinY) > 20.0 || Math.Abs(elemMaxY) > 20.0;
                                if (elementCount <= 5 || (isFar && elementCount <= 82))
                                {
                                    DebugLogger.Log($"[RASTER-EXPORT-GEOM#{elementCount}] Elem {elem.Id} " +
                                        $"(cat:{cat.Name}, bic:{bic}) " +
                                        $"vtx={vertexCount} " +
                                        $"X=[{elemMinX:F2},{elemMaxX:F2}] " +
                                        $"Y=[{elemMinY:F2},{elemMaxY:F2}]");
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Geometry extraction failed, fall through to bounding box
                    }

                    // Fallback: use bounding box if geometry extraction failed
                    if (!hasGeometry)
                    {
                        BoundingBoxXYZ bb = elem.get_BoundingBox(view);
                        if (bb != null && bb.Min != null && bb.Max != null)
                        {
                            if (double.IsInfinity(bb.Min.X) || double.IsInfinity(bb.Max.X) ||
                                double.IsInfinity(bb.Min.Y) || double.IsInfinity(bb.Max.Y) ||
                                double.IsNaN(bb.Min.X) || double.IsNaN(bb.Max.X))
                            {
                                filteredCount++;
                                continue;
                            }

                            elementCount++;
                            bboxFallbackCount++;
                            foundAny = true;

                            double bbMinX = double.MaxValue, bbMaxX = double.MinValue;
                            double bbMinY = double.MaxValue, bbMaxY = double.MinValue;

                            for (int i = 0; i < 8; i++)
                            {
                                double cx = (i & 1) == 0 ? bb.Min.X : bb.Max.X;
                                double cy = (i & 2) == 0 ? bb.Min.Y : bb.Max.Y;
                                double cz = (i & 4) == 0 ? bb.Min.Z : bb.Max.Z;
                                XYZ corner = new XYZ(cx, cy, cz);
                                XYZ relative = corner - viewOrigin;
                                double sx = relative.DotProduct(viewRight);
                                double sy = relative.DotProduct(viewUp);
                                if (sx < bbMinX) bbMinX = sx;
                                if (sx > bbMaxX) bbMaxX = sx;
                                if (sy < bbMinY) bbMinY = sy;
                                if (sy > bbMaxY) bbMaxY = sy;
                            }

                            // Update global bounds with contributor tracking
                            string bbLabel = $"BBOX#{elementCount} Elem {elem.Id}({cat.Name})";
                            if (bbMinX < screenXmin) { screenXmin = bbMinX; xminElem = bbLabel; }
                            if (bbMaxX > screenXmax) { screenXmax = bbMaxX; xmaxElem = bbLabel; }
                            if (bbMinY < screenYmin) { screenYmin = bbMinY; yminElem = bbLabel; }
                            if (bbMaxY > screenYmax) { screenYmax = bbMaxY; ymaxElem = bbLabel; }
                            
                            // Log bbox elements far from origin
                            double bbW = bbMaxX - bbMinX;
                            double bbH = bbMaxY - bbMinY;
                            bool bbFar = Math.Abs(bbMinX) > 20.0 || Math.Abs(bbMaxX) > 20.0 || 
                                         Math.Abs(bbMinY) > 20.0 || Math.Abs(bbMaxY) > 20.0;
                            if (bbFar || elementCount <= 5)
                            {
                                DebugLogger.Log($"[RASTER-EXPORT-BBOX#{elementCount}] Elem {elem.Id} (cat:{cat.Name}) " +
                                    $"bbox bounds: X=[{bbMinX:F2},{bbMaxX:F2}](w={bbW:F2}) " +
                                    $"Y=[{bbMinY:F2},{bbMaxY:F2}](h={bbH:F2})");
                            }
                        }
                        else
                        {
                            filteredCount++;
                        }
                    }
                }

                DebugLogger.Log($"[RASTER-EXPORT] Processed {elementCount} elements (bbox fallback for {bboxFallbackCount}), filtered {filteredCount}");

                if (!foundAny)
                {
                    TaskDialog.Show("Raster Export", "No model elements found with bounding boxes on this view.");
                    return Result.Failed;
                }

                DebugLogger.Log($"[RASTER-EXPORT] Screen bounds (ft): X=[{screenXmin:F6}, {screenXmax:F6}], Y=[{screenYmin:F6}, {screenYmax:F6}]");
                DebugLogger.Log($"[RASTER-EXPORT] Screen width={screenXmax-screenXmin:F4}ft, height={screenYmax-screenYmin:F4}ft");
                DebugLogger.Log($"[RASTER-EXPORT] view.Origin=({viewOrigin.X:F4},{viewOrigin.Y:F4},{viewOrigin.Z:F4})");

                DebugLogger.Log($"[RASTER-EXPORT-BOUNDARIES] Xmin={screenXmin:F6} -> {xminElem ?? "(none)"}");
                DebugLogger.Log($"[RASTER-EXPORT-BOUNDARIES] Xmax={screenXmax:F6} -> {xmaxElem ?? "(none)"}");
                DebugLogger.Log($"[RASTER-EXPORT-BOUNDARIES] Ymin={screenYmin:F6} -> {yminElem ?? "(none)"}");
                DebugLogger.Log($"[RASTER-EXPORT-BOUNDARIES] Ymax={screenYmax:F6} -> {ymaxElem ?? "(none)"}");

                // ── Step 8: Determine view/cropbox coordinate system ──
                // Save original view settings
                bool origCropActive = view.CropBoxActive;
                BoundingBoxXYZ origCropBox = view.CropBox;
                DisplayStyle origDisplayStyle = view.DisplayStyle;

                DebugLogger.Log($"[RASTER-EXPORT] View type: {view.ViewType}, CropBoxActive: {origCropActive}");
                DebugLogger.Log($"[RASTER-EXPORT] Orig CropBox Transform: {origCropBox?.Transform?.BasisX} {origCropBox?.Transform?.BasisY} {origCropBox?.Transform?.BasisZ}");
                DebugLogger.Log($"[RASTER-EXPORT] Orig CropBox null? {origCropBox == null}");

                // Determine the transform for the new crop box
                Transform viewTx = Transform.Identity;
                if (origCropBox != null)
                {
                    viewTx = origCropBox.Transform;
                }
                else if (view is View3D view3D && !view3D.IsPerspective)
                {
                    try
                    {
                        ViewOrientation3D orientation = view3D.GetOrientation();
                        XYZ right = orientation.UpDirection.CrossProduct(orientation.ForwardDirection);
                        viewTx = new Transform(Transform.Identity);
                        viewTx.BasisX = right;
                        viewTx.BasisY = orientation.UpDirection;
                        viewTx.BasisZ = orientation.ForwardDirection;
                        viewTx.Origin = XYZ.Zero;
                        DebugLogger.Log($"[RASTER-EXPORT] Created view transform from 3D orientation");
                    }
                    catch (Exception txEx)
                    {
                        DebugLogger.Log($"[RASTER-EXPORT] Failed to get 3D orientation: {txEx.Message}");
                    }
                }

                // ── Convert screen-space bounds to crop box local coordinates ──
                // screenX = (modelPt - viewOrigin).Dot(viewRight)
                // localX  = (modelPt - viewTx.Origin).Dot(viewRight)
                // localX  = screenX + (viewOrigin - viewTx.Origin).Dot(viewRight)
                XYZ originDiff = viewOrigin - viewTx.Origin;
                double offsetX = originDiff.DotProduct(viewRight);
                double offsetY = originDiff.DotProduct(viewUp);

                double vxMin = screenXmin + offsetX;
                double vxMax = screenXmax + offsetX;
                double vyMin = screenYmin + offsetY;
                double vyMax = screenYmax + offsetY;

                DebugLogger.Log($"[RASTER-EXPORT] View-space bounds (ft): X=[{vxMin:F6}, {vxMax:F6}], Y=[{vyMin:F6}, {vyMax:F6}]");

                // ── Step 3-6: Convert to paper mm and align to 3mm grid ──
                int viewScale = view.Scale;

                // ── Apply edge margin (expands bounds outward on all 4 sides) ──
                if (AnnotatixSettings.EdgeMarginMm > 0)
                {
                    // Convert margin from paper mm to model feet
                    double margin_ft = AnnotatixSettings.EdgeMarginMm / MM_PER_FT * viewScale;

                    DebugLogger.Log($"[RASTER-EXPORT] Applying edge margin: {AnnotatixSettings.EdgeMarginMm} mm = {margin_ft:F6} ft per side");

                    // Always subtract from min, add to max — expands bounds outward
                    vxMin -= margin_ft;
                    vxMax += margin_ft;
                    vyMin -= margin_ft;
                    vyMax += margin_ft;

                    DebugLogger.Log($"[RASTER-EXPORT] After margin (ft): X=[{vxMin:F6}, {vxMax:F6}], Y=[{vyMin:F6}, {vyMax:F6}]");
                }

                double modelW_ft = vxMax - vxMin;
                double modelH_ft = vyMax - vyMin;

                // Paper size from element bounds
                double paperW_mm = modelW_ft * MM_PER_FT / viewScale;
                double paperH_mm = modelH_ft * MM_PER_FT / viewScale;

                DebugLogger.Log($"[RASTER-EXPORT] Paper size (mm): {paperW_mm:F2} x {paperH_mm:F2}");

                // Align to 3mm grid (ceiling ensures crop always >= element bounds)
                int cols = (int)Math.Ceiling(paperW_mm / AnnotatixSettings.GridStepMm);
                int rows = (int)Math.Ceiling(paperH_mm / AnnotatixSettings.GridStepMm);
                
                double adjustedPaperW_mm = cols * AnnotatixSettings.GridStepMm;
                double adjustedPaperH_mm = rows * AnnotatixSettings.GridStepMm;

                DebugLogger.Log($"[RASTER-EXPORT] Grid: {cols} cols x {rows} rows = {cols * rows} cells");
                DebugLogger.Log($"[RASTER-EXPORT] Grid-aligned paper size (mm): {adjustedPaperW_mm:F2} x {adjustedPaperH_mm:F2}");
                DebugLogger.Log($"[RASTER-EXPORT] Grid padding (mm total): X={adjustedPaperW_mm - paperW_mm:F2}, Y={adjustedPaperH_mm - paperH_mm:F2}");

                // ── Step 7: Convert back to view feet ───────────────────
                double adjustedModelW_ft = adjustedPaperW_mm / MM_PER_FT * viewScale;
                double adjustedModelH_ft = adjustedPaperH_mm / MM_PER_FT * viewScale;

                // Center the adjusted crop box on view-space bounding box center
                double centerVx = (vxMin + vxMax) / 2.0;
                double centerVy = (vyMin + vyMax) / 2.0;

                BoundingBoxXYZ cropBox = new BoundingBoxXYZ();
                cropBox.Transform = viewTx;
                cropBox.Min = new XYZ(centerVx - adjustedModelW_ft / 2.0,
                                      centerVy - adjustedModelH_ft / 2.0,
                                      -100);
                cropBox.Max = new XYZ(centerVx + adjustedModelW_ft / 2.0,
                                      centerVy + adjustedModelH_ft / 2.0,
                                      100);

                DebugLogger.Log($"[RASTER-EXPORT] Crop box (ft): Min({cropBox.Min.X:F4},{cropBox.Min.Y:F4}) Max({cropBox.Max.X:F4},{cropBox.Max.Y:F4})");

                // Calculate pixel dimensions
                int pixelWidth = (int)Math.Round(adjustedPaperW_mm * PIXELS_PER_MM);
                int pixelHeight = (int)Math.Round(adjustedPaperH_mm * PIXELS_PER_MM);

                DebugLogger.Log($"[RASTER-EXPORT] Image pixel size: {pixelWidth} x {pixelHeight} at {DPI} DPI");

                // Determine output directory
                string outputDir = GetRasterExportDirectory(doc);

                // Generate export filename with timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"raster_{view.Name.Replace('/', '_').Replace('\\', '_')}_{timestamp}";

                try
                {
                    // Apply crop box
                    using (Transaction t = new Transaction(doc, "Raster Export: Set Crop Box"))
                    {
                        t.Start();

                        view.CropBoxActive = true;
                        view.CropBoxVisible = true;
                        view.CropBox = cropBox;

                        t.Commit();
                    }

                    DebugLogger.Log("[RASTER-EXPORT] Crop box applied");

                    // Export the view as PNG
                    string exportedPath = ViewExporter.ExportViewWithSettings(
                        doc, view, pixelWidth, pixelHeight, outputDir, fileName);

                    if (!string.IsNullOrEmpty(exportedPath))
                    {
                        FileInfo fi = new FileInfo(exportedPath);
                        DebugLogger.Log($"[RASTER-EXPORT] Export successful: {exportedPath} ({fi.Length / 1024} KB)");

                        // Save grid metadata sidecar for later image analysis
                        GridMetadata meta = new GridMetadata
                        {
                            Cols = cols,
                            Rows = rows,
                            CellSizeMm = AnnotatixSettings.GridStepMm,
                            PaperWidthMm = adjustedPaperW_mm,
                            PaperHeightMm = adjustedPaperH_mm,
                            PixelWidth = pixelWidth,
                            PixelHeight = pixelHeight,
                            Dpi = DPI,
                            ViewXmin = vxMin,
                            ViewYmin = vyMin,
                            ViewXmax = vxMax,
                            ViewYmax = vyMax,
                            OriginX = viewOrigin.X,
                            OriginY = viewOrigin.Y,
                            OriginZ = viewOrigin.Z,
                            RightDirX = viewRight.X,
                            RightDirY = viewRight.Y,
                            RightDirZ = viewRight.Z,
                            UpDirX = viewUp.X,
                            UpDirY = viewUp.Y,
                            UpDirZ = viewUp.Z,
                            ViewScale = viewScale,
                            OccupancyThreshold = AnnotatixSettings.OccupancyThreshold,
                        };
                        meta.Save(exportedPath);

                        // ── Auto-analyze the exported image ──
                        int occupiedCount = 0;
                        string analysisSummary = "";
                        try
                        {
                            ImageAnalyzer analyzer = new ImageAnalyzer(meta);
                            var results = analyzer.AnalyzeImage(exportedPath);
                            analyzer.SaveResults(results, exportedPath);
                            analyzer.CreateDebugOverlay(results, exportedPath);
                            occupiedCount = results.FindAll(c => c.IsOccupied).Count;
                            analysisSummary = $"\nAnalysis: {occupiedCount} occupied / {results.Count} total (>{meta.OccupancyThreshold*100:F0}% fill)";
                            DebugLogger.Log($"[RASTER-EXPORT] Auto-analysis complete: {occupiedCount} occupied / {results.Count} total (threshold={meta.OccupancyThreshold})");
                        }
                        catch (Exception anEx)
                        {
                            DebugLogger.Log($"[RASTER-EXPORT] Auto-analysis failed: {anEx.Message}");
                            analysisSummary = $"\nAnalysis failed: {anEx.Message}";
                        }

                        TaskDialog.Show("Raster Export",
                            $"Export successful!\n\n" +
                            $"File: {exportedPath}\n" +
                            $"Size: {pixelWidth} x {pixelHeight} px\n" +
                            $"Grid: {cols} x {rows} = {cols * rows} cells\n" +
                            $"Paper: {adjustedPaperW_mm:F1} x {adjustedPaperH_mm:F1} mm\n" +
                            $"Elements processed: {elementCount}" +
                            analysisSummary);
                    }
                    else
                    {
                        DebugLogger.Log("[RASTER-EXPORT] Export failed - no file created");
                        TaskDialog.Show("Raster Export", "Export failed. Check debug log for details.");
                    }
                }
                finally
                {
                    // User requested to keep crop box visible for debugging
                    DebugLogger.Log("[RASTER-EXPORT] Crop box left visible for inspection");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RASTER-EXPORT] ERROR: {ex.Message}");
                DebugLogger.Log($"[RASTER-EXPORT] StackTrace: {ex.StackTrace}");
                TaskDialog.Show("Raster Export", $"Error: {ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Returns the directory for raster export output.
        /// Uses: [AppData]\Autodesk\Revit\Addins\[year]\annotatix_dependencies\raster_exports
        /// </summary>
        private string GetRasterExportDirectory(Document doc)
        {
            string revitVersion = DetectRevitVersion();
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dir = Path.Combine(appData, "Autodesk", "Revit", "Addins",
                                       revitVersion, "annotatix_dependencies", "raster_exports");

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            return dir;
        }

        /// <summary>
        /// Detects Revit version from the assembly location
        /// </summary>
        private string DetectRevitVersion()
        {
            try
            {
                string dllPath = Assembly.GetExecutingAssembly().Location;
                var dir = new DirectoryInfo(Path.GetDirectoryName(dllPath));
                while (dir != null)
                {
                    if (dir.Name.Length == 4 && int.TryParse(dir.Name, out int year) &&
                        year >= 2020 && year <= 2035)
                        return dir.Name;
                    dir = dir.Parent;
                }
            }
            catch { }
            return "2025";
        }

        /// <summary>
        /// Recursively project geometry vertices to view-plane coordinates.
        /// Handles Solid (via face triangulation), Mesh, GeometryInstance (nested),
        /// Curve (tessellated), and PolyLine geometry types.
        /// </summary>
        private void ProjectGeometry(
            GeometryObject geomObj, Transform transform,
            XYZ viewOrigin, XYZ viewRight, XYZ viewUp,
            ref double minX, ref double maxX,
            ref double minY, ref double maxY,
            ref int vertexCount)
        {
            if (geomObj is Solid solid)
            {
                foreach (Face face in solid.Faces)
                {
                    Mesh mesh = face.Triangulate();
                    if (mesh == null) continue;
                    foreach (XYZ vertex in mesh.Vertices)
                    {
                        XYZ worldPt = transform.OfPoint(vertex);
                        XYZ relative = worldPt - viewOrigin;
                        double sx = relative.DotProduct(viewRight);
                        double sy = relative.DotProduct(viewUp);
                        if (sx < minX) minX = sx;
                        if (sx > maxX) maxX = sx;
                        if (sy < minY) minY = sy;
                        if (sy > maxY) maxY = sy;
                        vertexCount++;
                    }
                }
            }
            else if (geomObj is Mesh mesh)
            {
                foreach (XYZ vertex in mesh.Vertices)
                {
                    XYZ worldPt = transform.OfPoint(vertex);
                    XYZ relative = worldPt - viewOrigin;
                    double sx = relative.DotProduct(viewRight);
                    double sy = relative.DotProduct(viewUp);
                    if (sx < minX) minX = sx;
                    if (sx > maxX) maxX = sx;
                    if (sy < minY) minY = sy;
                    if (sy > maxY) maxY = sy;
                    vertexCount++;
                }
            }
            else if (geomObj is GeometryInstance inst)
            {
                // GetInstanceGeometry() returns geometry already in world coordinates.
                // DO NOT multiply by inst.Transform again — that would double-transform.
                GeometryElement instGeom = inst.GetInstanceGeometry();
                if (instGeom == null) return;

                foreach (GeometryObject child in instGeom)
                {
                    // Use the incoming transform (which handles nested hierarchy),
                    // NOT combined with inst.Transform (already baked into geometry).
                    ProjectGeometry(child, transform, viewOrigin, viewRight, viewUp,
                        ref minX, ref maxX, ref minY, ref maxY, ref vertexCount);
                }
            }
            else if (geomObj is Curve curve)
            {
                IList<XYZ> pts;
                try { pts = curve.Tessellate(); }
                catch { return; }
                foreach (XYZ pt in pts)
                {
                    XYZ worldPt = transform.OfPoint(pt);
                    XYZ relative = worldPt - viewOrigin;
                    double sx = relative.DotProduct(viewRight);
                    double sy = relative.DotProduct(viewUp);
                    if (sx < minX) minX = sx;
                    if (sx > maxX) maxX = sx;
                    if (sy < minY) minY = sy;
                    if (sy > maxY) maxY = sy;
                    vertexCount++;
                }
            }
            else if (geomObj is PolyLine polyLine)
            {
                foreach (XYZ pt in polyLine.GetCoordinates())
                {
                    XYZ worldPt = transform.OfPoint(pt);
                    XYZ relative = worldPt - viewOrigin;
                    double sx = relative.DotProduct(viewRight);
                    double sy = relative.DotProduct(viewUp);
                    if (sx < minX) minX = sx;
                    if (sx > maxX) maxX = sx;
                    if (sy < minY) minY = sy;
                    if (sy > maxY) maxY = sy;
                    vertexCount++;
                }
            }
        }
    }
}
