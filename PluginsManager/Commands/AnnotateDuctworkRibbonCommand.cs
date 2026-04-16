using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Ribbon command for deterministic ductwork annotation placement
    /// Wraps Annotatix.Module.Commands.AnnotateDuctworkCommand via dynamic loading
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class AnnotateDuctworkRibbonCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                Core.DebugLogger.Log("[ANNOTATE-DUCTWORK-RIBBON] Annotate Ductwork command started");

                // Get annotatix module
                var module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");

                // If module not loaded - try to auto-load
                if (module == null)
                {
                    Core.DebugLogger.Log("[ANNOTATE-DUCTWORK-RIBBON] Module not loaded, attempting auto-load...");

                    // Path to annotatix module
                    string assemblyPath = typeof(AnnotateDuctworkRibbonCommand).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);  // Go up one level
                    string annotatixModulePath = Path.Combine(modulesPath, "annotatix", "Annotatix.Module.dll");

                    Core.DebugLogger.Log($"[ANNOTATE-DUCTWORK-RIBBON] Annotatix module path: {annotatixModulePath}");

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
                    Core.DebugLogger.Log("[ANNOTATE-DUCTWORK-RIBBON] Module auto-loaded successfully");
                }

                if (module == null)
                {
                    TaskDialog.Show("Annotatix", "Модуль Annotatix не загружен. Откройте Plugins Hub для авторизации.");
                    return Result.Failed;
                }

                // Get module assembly
                var moduleAssembly = module.GetType().Assembly;

                // Find AnnotateDuctworkCommand type
                var commandType = moduleAssembly.GetType("Annotatix.Module.Commands.AnnotateDuctworkCommand");
                if (commandType == null)
                {
                    Core.DebugLogger.Log("[ANNOTATE-DUCTWORK-RIBBON] AnnotateDuctworkCommand type not found");
                    TaskDialog.Show("Annotatix", "Не найдена команда AnnotateDuctworkCommand в модуле Annotatix");
                    return Result.Failed;
                }

                // Create and execute command
                var command = Activator.CreateInstance(commandType) as IExternalCommand;
                if (command == null)
                {
                    Core.DebugLogger.Log("[ANNOTATE-DUCTWORK-RIBBON] Failed to create command instance");
                    TaskDialog.Show("Annotatix", "Не удалось создать экземпляр команды");
                    return Result.Failed;
                }

                return command.Execute(commandData, ref message, elements);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[ANNOTATE-DUCTWORK-RIBBON] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
