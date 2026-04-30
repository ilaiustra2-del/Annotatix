using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Ribbon command for Raster Export - wraps Annotatix.Module.Commands.RasterExportCommand
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class AnnotatixRasterExportRibbonCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Raster export command started");

                // Get annotatix module
                var module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");

                // If module not loaded - try to auto-load
                if (module == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Module not loaded, attempting auto-load...");

                    string assemblyPath = typeof(AnnotatixRasterExportRibbonCommand).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);
                    string annotatixModulePath = Path.Combine(modulesPath, "annotatix", "Annotatix.Module.dll");

                    Core.DebugLogger.Log($"[ANNOTATIX-RIBBON] Annotatix module path: {annotatixModulePath}");

                    if (!File.Exists(annotatixModulePath))
                    {
                        TaskDialog.Show("Raster Export", "Annotatix module not found. Open Plugins Hub to authorize.");
                        return Result.Failed;
                    }

                    if (!Core.DynamicModuleLoader.LoadModule("annotatix", annotatixModulePath))
                    {
                        TaskDialog.Show("Raster Export", "Failed to load Annotatix module. Try opening Plugins Hub.");
                        return Result.Failed;
                    }

                    module = Core.DynamicModuleLoader.GetModuleInstance("annotatix");
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Module auto-loaded successfully");
                }

                if (module == null)
                {
                    TaskDialog.Show("Raster Export", "Annotatix module not loaded. Open Plugins Hub to authorize.");
                    return Result.Failed;
                }

                // Get module assembly
                var moduleAssembly = module.GetType().Assembly;

                // Find RasterExportCommand type
                var commandType = moduleAssembly.GetType("Annotatix.Module.Commands.RasterExportCommand");
                if (commandType == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] RasterExportCommand type not found");
                    TaskDialog.Show("Raster Export", "RasterExportCommand not found in Annotatix module.");
                    return Result.Failed;
                }

                // Create and execute the command
                var command = Activator.CreateInstance(commandType) as IExternalCommand;
                if (command == null)
                {
                    Core.DebugLogger.Log("[ANNOTATIX-RIBBON] Failed to create RasterExportCommand instance");
                    TaskDialog.Show("Raster Export", "Failed to create RasterExportCommand instance.");
                    return Result.Failed;
                }

                return command.Execute(commandData, ref message, elements);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[ANNOTATIX-RIBBON] ERROR: {ex.Message}");
                TaskDialog.Show("Raster Export", $"Error: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
