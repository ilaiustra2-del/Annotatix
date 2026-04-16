using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Annotatix.Module.Core;
using Annotatix.Module.UI;
using PluginsManager.Core;

namespace Annotatix.Module.Commands
{
    /// <summary>
    /// Command to send current view state for ML analysis
    /// Creates -start files in both recordings and ML Processed directories
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public class SendForAnalysisCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uidoc = uiApp.ActiveUIDocument;
                var doc = uidoc.Document;

                // Get active view
                var activeView = doc.ActiveView;
                if (activeView == null)
                {
                    TaskDialog.Show("Annotatix", "Нет активного вида. Откройте вид для анализа.");
                    return Result.Failed;
                }

                // Get UIView for the active view
                UIView uiView = null;
                var uiViews = uidoc.GetOpenUIViews();
                foreach (var uv in uiViews)
                {
                    if (uv.ViewId == activeView.Id)
                    {
                        uiView = uv;
                        break;
                    }
                }

                if (uiView == null)
                {
                    TaskDialog.Show("Annotatix", "Не удалось получить активный UIView. Попробуйте еще раз.");
                    return Result.Failed;
                }

                DebugLogger.Log("[ANNOTATIX-ANALYSIS] Starting send for analysis...");

                // Generate new session ID
                string sessionId = Guid.NewGuid().ToString();

                // Collect snapshot (start type - current state before any changes)
                var collector = new ViewDataCollector(doc, activeView, uiView);
                var snapshot = collector.CollectSnapshot(sessionId, "start");

                // Export to recordings directory
                string recordingsDir = RecordingState.RecordingsDirectory;
                if (string.IsNullOrEmpty(recordingsDir))
                {
                    recordingsDir = JsonExporter.GetDefaultRecordingsDirectory();
                }

                string recordingsPath = JsonExporter.ExportToDirectory(snapshot, recordingsDir);
                DebugLogger.Log($"[ANNOTATIX-ANALYSIS] Exported to recordings: {recordingsPath}");

                // Export to ML Processed directory
                string mlProcessedDir = JsonExporter.MLProcessedDirectory;
                string mlProcessedPath = JsonExporter.ExportToDirectory(snapshot, mlProcessedDir);
                DebugLogger.Log($"[ANNOTATIX-ANALYSIS] Exported to ML Processed: {mlProcessedPath}");

                // Get session folders for message
                string sessionFolderRecordings = JsonExporter.GetSessionDirectory(recordingsDir, sessionId);
                string sessionFolderML = JsonExporter.GetSessionDirectory(mlProcessedDir, sessionId);

                TaskDialog.Show("Annotatix", 
                    "Данные отправлены на анализ.\\n\\n" +
                    $"Session ID: {sessionId}\\n\\n" +
                    $"Сохранено в:\\n{sessionFolderRecordings}\\n\\n" +
                    $"И скопировано в:\\n{sessionFolderML}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-ANALYSIS] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
