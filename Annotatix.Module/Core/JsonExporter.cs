using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Exports ViewSnapshot to JSON and XLS files
    /// </summary>
    public static class JsonExporter
    {
        /// <summary>
        /// Export snapshot to JSON and XLS files in a session subfolder
        /// </summary>
        public static string Export(ViewSnapshot snapshot, string baseDirectory)
        {
            try
            {
                // Create session subfolder
                string sessionFolder = Path.Combine(baseDirectory, snapshot.SessionId);
                if (!Directory.Exists(sessionFolder))
                {
                    Directory.CreateDirectory(sessionFolder);
                    DebugLogger.Log($"[ANNOTATIX-EXPORTER] Created session folder: {sessionFolder}");
                }

                // Generate filename: Project-View-Date-Time
                string sanitizedProject = SanitizeFileName(snapshot.DocumentName);
                string sanitizedView = SanitizeFileName(snapshot.ViewName);
                string dateStr = snapshot.Timestamp.ToString("yyyy-MM-dd");
                string timeStr = snapshot.Timestamp.ToString("HH-mm-ss");
                string baseFilename = $"{sanitizedProject}-{sanitizedView}-{dateStr}-{timeStr}-{snapshot.SnapshotType}";

                // Export JSON
                string jsonPath = ExportJson(snapshot, sessionFolder, baseFilename);
                
                // Export XLS (CSV format, opens in Excel)
                string xlsPath = ExportXls(snapshot, sessionFolder, baseFilename);

                return jsonPath;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-EXPORTER] Error exporting snapshot: {ex.Message}");
                throw;
            }
        }

        private static string ExportJson(ViewSnapshot snapshot, string directory, string baseFilename)
        {
            string filePath = Path.Combine(directory, baseFilename + ".json");

            var settings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                NullValueHandling = NullValueHandling.Ignore,
                DateFormatString = "yyyy-MM-ddTHH:mm:ss"
            };

            string json = JsonConvert.SerializeObject(snapshot, settings);
            File.WriteAllText(filePath, json);

            DebugLogger.Log($"[ANNOTATIX-EXPORTER] Exported JSON to: {filePath}");
            return filePath;
        }

        private static string ExportXls(ViewSnapshot snapshot, string directory, string baseFilename)
        {
            string filePath = Path.Combine(directory, baseFilename + ".csv");

            var sb = new StringBuilder();
            
            // Header info
            sb.AppendLine($"# Project: {snapshot.DocumentName}");
            sb.AppendLine($"# View: {snapshot.ViewName}");
            sb.AppendLine($"# Timestamp: {snapshot.Timestamp:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"# Type: {snapshot.SnapshotType}");
            sb.AppendLine();

            // Elements section with start/end coordinates
            sb.AppendLine("=== ELEMENTS ===");
            sb.AppendLine("ElementId;Category;FamilyName;TypeName;" +
                          "StartModelX;StartModelY;StartModelZ;StartViewX;StartViewY;" +
                          "EndModelX;EndModelY;EndModelZ;EndViewX;EndViewY;HasEndPoint;" +
                          "Diameter;Width;Height;SizeDisplay;" +
                          "SystemId;SystemName;BelongTo");
            
            foreach (var elem in snapshot.Elements)
            {
                sb.AppendLine($"{elem.ElementId};{elem.Category};{elem.FamilyName};{elem.TypeName};" +
                              $"{elem.ModelStart.X};{elem.ModelStart.Y};{elem.ModelStart.Z};" +
                              $"{elem.ViewStart.X};{elem.ViewStart.Y};" +
                              $"{elem.ModelEnd.X};{elem.ModelEnd.Y};{elem.ModelEnd.Z};" +
                              $"{elem.ViewEnd.X};{elem.ViewEnd.Y};" +
                              $"{elem.HasEndPoint};" +
                              $"{elem.Diameter};{elem.Width};{elem.Height};{elem.SizeDisplay};" +
                              $"{elem.SystemId};{elem.SystemName};{elem.BelongTo}");
            }

            sb.AppendLine();

            // Annotations section with head/leader positions
            sb.AppendLine("=== ANNOTATIONS ===");
            sb.AppendLine("ElementId;Category;FamilyName;TypeName;" +
                          "HeadModelX;HeadModelY;HeadModelZ;HeadViewX;HeadViewY;" +
                          "LeaderEndModelX;LeaderEndModelY;LeaderEndModelZ;LeaderEndViewX;LeaderEndViewY;" +
                          "LeaderType;HasLeader;HasElbow;TaggedElementId;BelongTo;TagText");
            
            foreach (var ann in snapshot.Annotations)
            {
                sb.AppendLine($"{ann.ElementId};{ann.Category};{ann.FamilyName};{ann.TypeName};" +
                              $"{ann.HeadModelPosition?.X ?? 0};{ann.HeadModelPosition?.Y ?? 0};{ann.HeadModelPosition?.Z ?? 0};" +
                              $"{ann.HeadViewPosition?.X ?? 0};{ann.HeadViewPosition?.Y ?? 0};" +
                              $"{ann.LeaderEndModel?.X ?? 0};{ann.LeaderEndModel?.Y ?? 0};{ann.LeaderEndModel?.Z ?? 0};" +
                              $"{ann.LeaderEndView?.X ?? 0};{ann.LeaderEndView?.Y ?? 0};" +
                              $"{ann.LeaderType};{ann.HasLeader};{ann.HasElbow};" +
                              $"{ann.TaggedElementId};{ann.BelongTo};{ann.TagText}");
            }

            sb.AppendLine();

            // Systems section
            sb.AppendLine("=== SYSTEMS ===");
            sb.AppendLine("SystemId;SystemName;SystemType;ElementIds");
            
            foreach (var sys in snapshot.Systems)
            {
                sb.AppendLine($"{sys.SystemId};{sys.SystemName};{sys.SystemType};{string.Join(", ", sys.ElementIds)}");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
            DebugLogger.Log($"[ANNOTATIX-EXPORTER] Exported CSV to: {filePath}");
            return filePath;
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Unknown";
            
            var invalid = Path.GetInvalidFileNameChars();
            var result = new StringBuilder();
            
            foreach (var c in name)
            {
                if (Array.IndexOf(invalid, c) >= 0 || c == ' ' || c == '(' || c == ')')
                    result.Append('_');
                else
                    result.Append(c);
            }
            
            return result.ToString();
        }

        /// <summary>
        /// Get the default recordings directory
        /// </summary>
        public static string GetDefaultRecordingsDirectory()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string annotatixPath = Path.Combine(
                appData, 
                "Autodesk", 
                "Revit", 
                "Addins", 
                "2025", 
                "annotatix_dependencies", 
                "annotatix", 
                "recordings"
            );

            return annotatixPath;
        }
    }
}
