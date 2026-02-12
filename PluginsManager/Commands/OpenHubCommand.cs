using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace PluginsManager.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenHubCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                
                // Load FamilySync module first (to make type available)
                LoadFamilySyncModule();
                
                // Create FamilySync ExternalEvent in API context
                object familySyncHandler = null;
                ExternalEvent familySyncEvent = null;
                
                try
                {
                    System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] Attempting to create FamilySync ExternalEvent...");
                    
                    // Get all loaded assemblies to check
                    var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
                    var familySyncAsm = loadedAssemblies.FirstOrDefault(a => a.GetName().Name == "FamilySync.Module");
                    
                    if (familySyncAsm != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Found assembly: {familySyncAsm.FullName}");
                        
                        // Get type from loaded assembly
                        var familySyncHandlerType = familySyncAsm.GetType("FamilySync.Module.UI.FamilySyncHandler");
                        
                        if (familySyncHandlerType != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Found type: {familySyncHandlerType.FullName}");
                            
                            familySyncHandler = Activator.CreateInstance(familySyncHandlerType);
                            System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Created handler instance");
                            
                            var iExternalEventHandler = familySyncHandler as IExternalEventHandler;
                            if (iExternalEventHandler != null)
                            {
                                familySyncEvent = ExternalEvent.Create(iExternalEventHandler);
                                System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] ✓ FamilySync ExternalEvent created successfully!");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] ✗ Handler does not implement IExternalEventHandler");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] ✗ FamilySyncHandler type not found in assembly");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] ✗ FamilySync.Module assembly not loaded");
                        System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Loaded assemblies: {string.Join(", ", loadedAssemblies.Select(a => a.GetName().Name))}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] ✗ Failed to create FamilySync ExternalEvent: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Stack trace: {ex.StackTrace}");
                }
                
                // Create and show the main hub panel with ExternalEvent
                UI.MainHubPanel hub = new UI.MainHubPanel(uiApp, familySyncHandler, familySyncEvent);
                
                // Set owner to Revit window
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(hub);
                helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;
                
                hub.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
        
        private void LoadFamilySyncModule()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = System.IO.Path.GetDirectoryName(assembly.Location);
                var modulesPath = System.IO.Path.GetDirectoryName(assemblyPath);
                var familySyncDllPath = System.IO.Path.Combine(modulesPath, "family_sync", "FamilySync.Module.dll");
                
                System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Looking for module at: {familySyncDllPath}");
                
                if (System.IO.File.Exists(familySyncDllPath))
                {
                    // Load assembly explicitly first
                    var familySyncAssembly = System.Reflection.Assembly.LoadFrom(familySyncDllPath);
                    System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Assembly loaded: {familySyncAssembly.FullName}");
                    
                    // Then register with module loader
                    Core.DynamicModuleLoader.LoadModule("family_sync", familySyncDllPath);
                    System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] FamilySync module registered");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[OPEN-HUB-CMD] FamilySync.Module.dll not found");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Error loading FamilySync: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[OPEN-HUB-CMD] Stack trace: {ex.StackTrace}");
            }
        }
    }
}
