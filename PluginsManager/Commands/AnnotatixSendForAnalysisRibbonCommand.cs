using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Ribbon command for Send For Analysis - wraps Annotatix.Module.Commands.SendForAnalysisCommand
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public class AnnotatixSendForAnalysisRibbonCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                Core.DebugLogger.Log("[ANNOTATIX-ANALYSIS-RIBBON] Send for analysis command started");

                // Get annotatix module
                var module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");

                // If module not loaded - try to auto-load
                if (module == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-ANALYSIS-RIBBON] Module not loaded, attempting auto-load...");

                    // Path to annotatix module (annotatix_dependencies/main/ -> annotatix_dependencies/)
                    string assemblyPath = typeof(AnnotatixSendForAnalysisRibbonCommand).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);  // Go up one level
                    string annotatixModulePath = Path.Combine(modulesPath, "annotatix", "Annotatix.Module.dll");

                    Core.DebugLogger.Log($"[ANNOTATIX-ANALYSIS-RIBBON] Annotatix module path: {annotatixModulePath}");

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
                    Core.DebugLogger.Log("[ANNOTATIX-ANALYSIS-RIBBON] Module auto-loaded successfully");
                }

                if (module == null)
                {
                    TaskDialog.Show("Annotatix", "Модуль Annotatix не загружен. Откройте Plugins Hub для авторизации.");
                    return Result.Failed;
                }

                // Get module assembly
                var moduleAssembly = module.GetType().Assembly;

                // Find SendForAnalysisCommand type
                var sendForAnalysisCommandType = moduleAssembly.GetType("Annotatix.Module.Commands.SendForAnalysisCommand");
                if (sendForAnalysisCommandType == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-ANALYSIS-RIBBON] SendForAnalysisCommand type not found");
                    TaskDialog.Show("Annotatix", "Не найдена команда SendForAnalysisCommand в модуле Annotatix");
                    return Result.Failed;
                }

                // Create and execute SendForAnalysisCommand
                var sendForAnalysisCommand = Activator.CreateInstance(sendForAnalysisCommandType) as IExternalCommand;
                if (sendForAnalysisCommand == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-ANALYSIS-RIBBON] Failed to create SendForAnalysisCommand instance");
                    TaskDialog.Show("Annotatix", "Не удалось создать экземпляр команды отправки на анализ");
                    return Result.Failed;
                }

                return sendForAnalysisCommand.Execute(commandData, ref message, elements);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[ANNOTATIX-ANALYSIS-RIBBON] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
