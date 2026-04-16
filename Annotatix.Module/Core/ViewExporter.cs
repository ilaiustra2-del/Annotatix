using System;
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
    }
}
