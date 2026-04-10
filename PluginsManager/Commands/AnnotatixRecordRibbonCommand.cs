using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Ribbon command for Annotatix recording - wraps Annotatix.Module.Commands.RecordingCommand
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public class AnnotatixRecordRibbonCommand : IExternalCommand
    {
        /// <summary>
        /// Reference to the record button for dynamic text update
        /// </summary>
        public static PushButton RecordButton { get; set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Record command started");

                // Get annotatix module
                var module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");

                // If module not loaded - try to auto-load
                if (module == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Module not loaded, attempting auto-load...");

                    // Path to annotatix module (annotatix_dependencies/main/ -> annotatix_dependencies/)
                    string assemblyPath = typeof(AnnotatixRecordRibbonCommand).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);  // Go up one level
                    string annotatixModulePath = Path.Combine(modulesPath, "annotatix", "Annotatix.Module.dll");

                    Core.DebugLogger.Log($"[ANNOTATIX-RIBBON] Annotatix module path: {annotatixModulePath}");

                    if (!File.Exists(annotatixModulePath))
                    {
                        TaskDialog.Show("Annotatix", "Модуль Annotatix не найден. Откройте Plugins Hub для авторизации.");
                        return Result.Failed;
                    }

                    if (!Core.DynamicModuleLoader.LoadModule("annotatix", annotatixModulePath))
                    {
                        TaskDialog.Show("Annotatix", "Не удалось загрузить модуль Annotatix. Попробуйте открыть Plugins Hub.");
                        return Result.Failed;
                    }

                    module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Module auto-loaded successfully");
                }

                if (module == null)
                {
                    TaskDialog.Show("Annotatix", "Модуль Annotatix не загружен. Откройте Plugins Hub для авторизации.");
                    return Result.Failed;
                }

                // Get module assembly
                var moduleAssembly = module.GetType().Assembly;

                // Find RecordingCommand type
                var recordingCommandType = moduleAssembly.GetType("Annotatix.Module.Commands.RecordingCommand");
                if (recordingCommandType == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] RecordingCommand type not found");
                    TaskDialog.Show("Annotatix", "Не найдена команда RecordingCommand в модуле Annotatix");
                    return Result.Failed;
                }

                // Set RecordButton reference in module's RecordingState
                var recordingStateType = moduleAssembly.GetType("Annotatix.Module.UI.RecordingState");
                if (recordingStateType != null)
                {
                    var recordButtonProperty = recordingStateType.GetProperty("RecordButton");
                    recordButtonProperty?.SetValue(null, RecordButton);
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] RecordButton reference set in RecordingState");
                }

                // Create and execute RecordingCommand
                var recordingCommand = Activator.CreateInstance(recordingCommandType) as IExternalCommand;
                if (recordingCommand == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Failed to create RecordingCommand instance");
                    TaskDialog.Show("Annotatix", "Не удалось создать экземпляр команды записи");
                    return Result.Failed;
                }

                return recordingCommand.Execute(commandData, ref message, elements);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[ANNOTATIX-RIBBON] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
