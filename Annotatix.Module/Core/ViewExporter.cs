using System;
using System.IO;
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

                // Build output path
                string outputPath = Path.Combine(outputDirectory, $"{fileName}.png");

                DebugLogger.Log($"[ANNOTATIX-EXPORT] Exporting view '{view.Name}' to PNG: {outputPath}");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] Pixel width: {width}");

                // Configure export options
                // Use SetOfViews with the view ID to export the ENTIRE view content
                // (not just the current viewport/zoom state)
                ImageExportOptions options = new ImageExportOptions
                {
                    ExportRange = ExportRange.SetOfViews,
                    FilePath = outputDirectory,
                    ImageResolution = ImageResolution.DPI_150,
                    ZoomType = ZoomFitType.FitToPage,
                    PixelSize = width
                };

                // Set the views to export (using ElementId collection)
                options.SetViewsAndSheets(new ElementId[] { view.Id });

                DebugLogger.Log($"[ANNOTATIX-EXPORT] Using SetOfViews for view ID: {view.Id}");
                DebugLogger.Log($"[ANNOTATIX-EXPORT] View crop box active: {view.CropBoxActive}");

                // Export the view
                doc.ExportImage(options);

                // Revit generates file with view name when using SetOfViews
                // File name format: "FileName - ViewName.png" or similar
                string actualPath = FindExportedFile(outputDirectory, fileName, view.Name);

                if (actualPath != null && File.Exists(actualPath))
                {
                    DebugLogger.Log($"[ANNOTATIX-EXPORT] Successfully exported to: {actualPath}");
                    return actualPath;
                }
                else
                {
                    // Try to find any PNG file created recently in the directory
                    var files = Directory.GetFiles(outputDirectory, "*.png");
                    if (files.Length > 0)
                    {
                        // Return the most recently created file
                        string mostRecent = files[0];
                        DateTime mostRecentTime = File.GetCreationTime(mostRecent);
                        foreach (string file in files)
                        {
                            DateTime fileTime = File.GetCreationTime(file);
                            if (fileTime > mostRecentTime)
                            {
                                mostRecent = file;
                                mostRecentTime = fileTime;
                            }
                        }
                        DebugLogger.Log($"[ANNOTATIX-EXPORT] Found exported file: {mostRecent}");
                        return mostRecent;
                    }

                    DebugLogger.Log($"[ANNOTATIX-EXPORT] WARNING: Could not find exported PNG file");
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
        /// Finds the actual exported file (Revit may append view name to filename)
        /// </summary>
        private static string FindExportedFile(string directory, string baseFileName, string viewName)
        {
            // Revit typically appends " - ViewName" or similar
            string[] possiblePatterns = new string[]
            {
                $"{baseFileName}.png",
                $"{baseFileName} - {viewName}.png",
                $"{baseFileName}_{viewName}.png",
                $"{baseFileName} {viewName}.png"
            };

            foreach (string pattern in possiblePatterns)
            {
                string fullPath = Path.Combine(directory, pattern);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
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
