using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Ribbon command for placing annotations from last recording - wraps Annotatix.Module.Commands.PlaceAnnotationsCommand
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class AnnotatixPlaceAnnotationsRibbonCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                Core.DebugLogger.Log("[ANNOTATIX-PLACE-RIBBON] Place annotations command started");

                // Get annotatix module
                var module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");

                // If module not loaded - try to auto-load
                if (module == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-PLACE-RIBBON] Module not loaded, attempting auto-load...");

                    // Path to annotatix module
                    string assemblyPath = typeof(AnnotatixPlaceAnnotationsRibbonCommand).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);  // Go up one level
                    string annotatixModulePath = Path.Combine(modulesPath, "annotatix", "Annotatix.Module.dll");

                    Core.DebugLogger.Log($"[ANNOTATIX-PLACE-RIBBON] Annotatix module path: {annotatixModulePath}");

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
                    Core.DebugLogger.Log("[ANNOTATIX-PLACE-RIBBON] Module auto-loaded successfully");
                }

                if (module == null)
                {
                    TaskDialog.Show("Annotatix", "Модуль Annotatix не загружен. Откройте Plugins Hub для авторизации.");
                    return Result.Failed;
                }

                // Get module assembly
                var moduleAssembly = module.GetType().Assembly;

                // Find PlaceAnnotationsCommand type
                var placeCommandType = moduleAssembly.GetType("Annotatix.Module.Commands.PlaceAnnotationsCommand");
                if (placeCommandType == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-PLACE-RIBBON] PlaceAnnotationsCommand type not found");
                    TaskDialog.Show("Annotatix", "Не найдена команда PlaceAnnotationsCommand в модуле Annotatix");
                    return Result.Failed;
                }

                // Create and execute PlaceAnnotationsCommand
                var placeCommand = Activator.CreateInstance(placeCommandType) as IExternalCommand;
                if (placeCommand == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-PLACE-RIBBON] Failed to create PlaceAnnotationsCommand instance");
                    TaskDialog.Show("Annotatix", "Не удалось создать экземпляр команды размещения аннотаций");
                    return Result.Failed;
                }

                return placeCommand.Execute(commandData, ref message, elements);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[ANNOTATIX-PLACE-RIBBON] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
