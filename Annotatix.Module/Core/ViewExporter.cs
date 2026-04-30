using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Exports Revit views to image files (PNG)
    /// </summary>
    public static class ViewExporter
    {
        /// <summary>
        /// Default pixel width for exported images
        /// </summary>
        public static int DefaultPixelWidth { get; set; } = 1920;

        /// <summary>
        /// Exports a view to PNG format
        /// </summary>
        public static string ExportViewToPng(Document doc, View view, string sessionDirectory, string fileName, int? pixelWidth = null)
        {
            try
            {
                int width = pixelWidth ?? DefaultPixelWidth;

                // Ensure session directory exists
                if (!Directory.Exists(sessionDirectory))
                {
                    Directory.CreateDirectory(sessionDirectory);
                }

                // Target path for final file in session directory
                string targetPath = Path.Combine(sessionDirectory, $"{fileName}.png");

                // Get the module directory (where Annotatix.Module.dll is located)
                string moduleDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                // Revit exports to PARENT directory of the module!
                string exportDir = Directory.GetParent(moduleDir)?.FullName ?? moduleDir;

                DebugLogger.Log($"[ANNOTATIX-EXPORT] Exporting view '{view.Name}' to PNG");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] Export directory: {exportDir}");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] Target path: {targetPath}");

                // Log view bounds info
                if (view is View3D view3D)
                {
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] 3D View - SectionBox enabled: {view3D.IsSectionBoxActive}");
                }
                DebugLogger.Log($"[ANNOTATIX-EXPORT] View CropBoxActive: {view.CropBoxActive}");

                // Record files before export in EXPORT directory
                var filesBefore = Directory.GetFiles(exportDir, "*.*")
                    .Where(f => f.EndsWith(".png") || f.EndsWith(".jpg") || f.EndsWith(".jpeg"))
                    .ToHashSet();

                // Configure export options
                // Use SetOfViews to export the ENTIRE view content
                ImageExportOptions options = new ImageExportOptions
                {
                    ExportRange = ExportRange.SetOfViews,
                    FilePath = exportDir,
                    ImageResolution = ImageResolution.DPI_150,
                    ZoomType = ZoomFitType.FitToPage,
                    PixelSize = width
                };

                // Set the views to export
                options.SetViewsAndSheets(new ElementId[] { view.Id });

                // Export the view
                doc.ExportImage(options);

                // Find newly created file in EXPORT directory
                var filesAfter = Directory.GetFiles(exportDir, "*.*")
                    .Where(f => f.EndsWith(".png") || f.EndsWith(".jpg") || f.EndsWith(".jpeg"))
                    .ToHashSet();

                string exportedFile = filesAfter.Except(filesBefore).FirstOrDefault();

                if (exportedFile != null && File.Exists(exportedFile))
                {
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] Revit created file: {exportedFile}");

                    // Move to session directory with correct name
                    if (File.Exists(targetPath))
                    {
                        File.Delete(targetPath);
                    }
                    File.Move(exportedFile, targetPath);
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] Moved to: {targetPath}");
                    return targetPath;
                }
                else
                {
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] WARNING: Could not find exported file in {exportDir}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-EXPORT] ERROR exporting view to PNG: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Exports start snapshot (called when recording begins)
        /// </summary>
        public static string ExportStartSnapshot(Document doc, View view, string sessionDirectory)
        {
            return ExportViewToPng(doc, view, sessionDirectory, "view_start");
        }

        /// <summary>
        /// Exports end snapshot (called when recording ends)
        /// </summary>
        public static string ExportEndSnapshot(Document doc, View view, string sessionDirectory)
        {
            return ExportViewToPng(doc, view, sessionDirectory, "view_end");
        }

        /// <summary>
        /// Exports a view to PNG with explicit pixel dimensions and output directory.
        /// The view should already have its crop box configured before calling this method.
        /// </summary>
        public static string ExportViewWithSettings(Document doc, View view,
            int pixelWidth, int pixelHeight, string outputDir, string fileName)
        {
            try
            {
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                string targetPath = Path.Combine(outputDir, $"{fileName}.png");

                DebugLogger.Log($"[VIEW-EXPORT] Exporting view '{view.Name}' with settings");
                DebugLogger.Log($"[VIEW-EXPORT] Pixel size: {pixelWidth} x {pixelHeight}");
                DebugLogger.Log($"[VIEW-EXPORT] Output directory: {outputDir}");
                DebugLogger.Log($"[VIEW-EXPORT] Target path: {targetPath}");

                // Get the parent of the module DLL directory
                string moduleDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string parentDir = Directory.GetParent(moduleDir)?.FullName ?? moduleDir;

                DebugLogger.Log($"[VIEW-EXPORT] Module dir: {moduleDir}");
                DebugLogger.Log($"[VIEW-EXPORT] Parent dir: {parentDir}");

                // Revit ExportImage with SetOfViews creates files in the PARENT
                // of the FilePath directory. So we set FilePath to a subfolder
                // (outputDir) so the file lands in parentDir where we can find it.
                DebugLogger.Log($"[VIEW-EXPORT] FilePath for export: {outputDir}");

                // Record ALL files in parentDir before export
                var filesBefore = new HashSet<string>(Directory.GetFiles(parentDir, "*.*"),
                    StringComparer.OrdinalIgnoreCase);

                // Configure export options (ONE clean call, no retry)
                ImageExportOptions options = new ImageExportOptions
                {
                    ExportRange = ExportRange.SetOfViews,
                    FilePath = outputDir,
                    ImageResolution = ImageResolution.DPI_150,
                    ZoomType = ZoomFitType.FitToPage,
                    PixelSize = pixelWidth
                };
                options.SetViewsAndSheets(new ElementId[] { view.Id });

                DebugLogger.Log($"[VIEW-EXPORT] Calling ExportImage...");
                long ticksBefore = DateTime.Now.Ticks;
                doc.ExportImage(options);
                long ticksAfter = DateTime.Now.Ticks;
                double elapsedMs = (ticksAfter - ticksBefore) / 10000.0;
                DebugLogger.Log($"[VIEW-EXPORT] ExportImage returned (elapsed: {elapsedMs:F0}ms)");

                // Find new file in parentDir (where Revit actually writes)
                var filesAfter = new HashSet<string>(Directory.GetFiles(parentDir, "*.*"),
                    StringComparer.OrdinalIgnoreCase);
                var newFiles = filesAfter.Except(filesBefore).ToList();

                DebugLogger.Log($"[VIEW-EXPORT] New files in parentDir: {newFiles.Count}");
                foreach (var f in newFiles)
                    DebugLogger.Log($"[VIEW-EXPORT]   NEW: {f}");

                if (newFiles.Count >= 1)
                {
                    // Pick the file that matches the view name (most specific)
                    string exportedFile = newFiles.Count == 1
                        ? newFiles[0]
                        : newFiles.FirstOrDefault(f => f.Contains(view.Name)) ?? newFiles[0];

                    DebugLogger.Log($"[VIEW-EXPORT] Found exported file: {exportedFile}");

                    // Move to target path in outputDir
                    if (File.Exists(targetPath))
                        File.Delete(targetPath);
                    File.Move(exportedFile, targetPath);

                    DebugLogger.Log($"[VIEW-EXPORT] Moved to: {targetPath}");

                    // Clean up any remaining new files (duplicates)
                    foreach (var f in newFiles)
                    {
                        if (f != exportedFile && File.Exists(f))
                        {
                            try { File.Delete(f); } catch { }
                            DebugLogger.Log($"[VIEW-EXPORT] Cleaned up duplicate: {f}");
                        }
                    }

                    return targetPath;
                }

                DebugLogger.Log("[VIEW-EXPORT] Export failed - no file created.");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[VIEW-EXPORT] ERROR exporting view with settings: {ex.Message}");
                DebugLogger.Log($"[VIEW-EXPORT] StackTrace: {ex.StackTrace}");
                return null;
            }
        }
    }
}
