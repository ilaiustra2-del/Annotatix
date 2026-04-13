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
    /// Ribbon command for recording view snapshots
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public class RecordingCommand : IExternalCommand
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
                    TaskDialog.Show("Annotatix", "Нет активного вида. Откройте вид для записи.");
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

                // Initialize recordings directory if not set
                if (string.IsNullOrEmpty(RecordingState.RecordingsDirectory))
                {
                    RecordingState.RecordingsDirectory = JsonExporter.GetDefaultRecordingsDirectory();
                }

                // Toggle recording state
                if (!RecordingState.IsRecording)
                {
                    // START RECORDING
                    StartRecording(doc, activeView, uiView);
                }
                else
                {
                    // END RECORDING
                    EndRecording(doc, activeView, uiView);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-CMD] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }

        private void StartRecording(Document doc, View view, UIView uiView)
        {
            DebugLogger.Log("[ANNOTATIX-CMD] Starting recording...");

            // Start new session
            RecordingState.StartNewSession();

            // Get session directory path BEFORE export
            string sessionDirectory = JsonExporter.GetSessionDirectory(
                RecordingState.RecordingsDirectory, 
                RecordingState.SessionId);
            DebugLogger.Log($"[ANNOTATIX-CMD] Session directory: {sessionDirectory}");

            // Collect start snapshot
            var collector = new ViewDataCollector(doc, view, uiView);
            var snapshot = collector.CollectSnapshot(RecordingState.SessionId, "start");

            // Export start snapshot
            RecordingState.StartSnapshotPath = JsonExporter.Export(snapshot, RecordingState.RecordingsDirectory);

            // Export start snapshot as PNG (to session directory)
            try
            {
                string pngPath = ViewExporter.ExportStartSnapshot(doc, view, sessionDirectory);
                if (!string.IsNullOrEmpty(pngPath))
                {
                    DebugLogger.Log($"[ANNOTATIX-CMD] Start snapshot PNG exported: {pngPath}");
                }
            }
            catch (Exception pngEx)
            {
                DebugLogger.Log($"[ANNOTATIX-CMD] WARNING: Failed to export start PNG: {pngEx.Message}");
            }

            // Update button text
            UpdateButtonText(true);

            DebugLogger.Log($"[ANNOTATIX-CMD] Recording started. Session: {RecordingState.SessionId}");
            TaskDialog.Show("Annotatix", 
                "Запись начата.\n\n" +
                "Произведите манипуляции с аннотациями на виде,\n" +
                "затем нажмите кнопку 'Завершить запись'.");
        }

        private void EndRecording(Document doc, View view, UIView uiView)
        {
            DebugLogger.Log("[ANNOTATIX-CMD] Ending recording...");

            // Get session directory path (use stored SessionId)
            string sessionDirectory = JsonExporter.GetSessionDirectory(
                RecordingState.RecordingsDirectory, 
                RecordingState.SessionId);
            DebugLogger.Log($"[ANNOTATIX-CMD] Session directory: {sessionDirectory}");

            // Collect end snapshot
            var collector = new ViewDataCollector(doc, view, uiView);
            var snapshot = collector.CollectSnapshot(RecordingState.SessionId, "end");

            // Export end snapshot
            var endPath = JsonExporter.Export(snapshot, RecordingState.RecordingsDirectory);

            // Export end snapshot as PNG (to session directory)
            try
            {
                string pngPath = ViewExporter.ExportEndSnapshot(doc, view, sessionDirectory);
                if (!string.IsNullOrEmpty(pngPath))
                {
                    DebugLogger.Log($"[ANNOTATIX-CMD] End snapshot PNG exported: {pngPath}");
                }
            }
            catch (Exception pngEx)
            {
                DebugLogger.Log($"[ANNOTATIX-CMD] WARNING: Failed to export end PNG: {pngEx.Message}");
            }

            // Store directory for message before reset
            string recordingsDir = RecordingState.RecordingsDirectory;

            // Reset state
            RecordingState.Reset();

            // Update button text
            UpdateButtonText(false);

            DebugLogger.Log($"[ANNOTATIX-CMD] Recording ended. Files saved to: {recordingsDir}");
            TaskDialog.Show("Annotatix", 
                "Запись завершена.\n\n" +
                $"Файлы сохранены в:\n{recordingsDir}");
        }

        private void UpdateButtonText(bool isRecording)
        {
            if (RecordingState.RecordButton != null)
            {
                RecordingState.RecordButton.ItemText = isRecording 
                    ? "Завершить\nзапись" 
                    : "Начать\nзапись";
                
                DebugLogger.Log($"[ANNOTATIX-CMD] Button text updated: {RecordingState.RecordButton.ItemText}");
            }
            else
            {
                DebugLogger.Log("[ANNOTATIX-CMD] WARNING: RecordButton is null, cannot update text");
            }
        }
    }
}
