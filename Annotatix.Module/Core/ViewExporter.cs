using System;
using System.IO;
using System.Linq;
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
        /// <param name="doc">The document containing the view</param>
        /// <param name="view">The view to export</param>
        /// <param name="outputDirectory">Directory to save the PNG file</param>
        /// <param name="fileName">Name of the output file (without extension)</param>
        /// <param name="pixelWidth">Width in pixels (height is calculated proportionally)</param>
        /// <returns>Full path to the exported PNG file, or null if export failed</returns>
        public static string ExportViewToPng(Document doc, View view, string outputDirectory, string fileName, int? pixelWidth = null)
        {
            try
            {
                int width = pixelWidth ?? DefaultPixelWidth;

                // Ensure output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Target path for final file
                string targetPath = Path.Combine(outputDirectory, $"{fileName}.png");

                DebugLogger.Log($"[ANNOTATIX-EXPORT] Exporting view '{view.Name}' to PNG");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] Target path: {targetPath}");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] Pixel width: {width}");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] View crop box active: {view.CropBoxActive}");

                // Record files before export to detect new files
                var filesBefore = Directory.GetFiles(outputDirectory, "*.*")
                    .Where(f => f.EndsWith(".png") || f.EndsWith(".jpg") || f.EndsWith(".jpeg"))
                    .ToHashSet();

                // Configure export options
                // Use SetOfViews to export the ENTIRE view content
                ImageExportOptions options = new ImageExportOptions
                {
                    ExportRange = ExportRange.SetOfViews,
                    FilePath = outputDirectory,
                    ImageResolution = ImageResolution.DPI_150,
                    ZoomType = ZoomFitType.FitToPage,
                    PixelSize = width
                };

                // Set the views to export
                options.SetViewsAndSheets(new ElementId[] { view.Id });

                // Export the view
                doc.ExportImage(options);

                // Find newly created file (Revit generates its own filename)
                var filesAfter = Directory.GetFiles(outputDirectory, "*.*")
                    .Where(f => f.EndsWith(".png") || f.EndsWith(".jpg") || f.EndsWith(".jpeg"))
                    .ToHashSet();

                string exportedFile = filesAfter.Except(filesBefore).FirstOrDefault();

                if (exportedFile != null && File.Exists(exportedFile))
                {
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] Revit created file: {exportedFile}");

                    // Rename to target path
                    if (File.Exists(targetPath))
                    {
                        File.Delete(targetPath);
                    }
                    File.Move(exportedFile, targetPath);
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] Renamed to: {targetPath}");
                    return targetPath;
                }
                else
                {
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] WARNING: Could not find exported file");
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
