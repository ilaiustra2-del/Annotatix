using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using HVACSuperScheme.Commands.Settings;

namespace dwg2rvt.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenHVACPanelCommand : IExternalCommand
    {
        private static ExternalEvent _createSchemaEvent;
        private static ExternalEvent _completeSchemaEvent;
        private static ExternalEvent _settingsEvent;
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc?.Document;

                if (doc == null)
                {
                    TaskDialog.Show("Error", "No active document found. Please open a Revit document.");
                    return Result.Failed;
                }

                // Get module from loader
                var module = PluginsManager.Core.DynamicModuleLoader.GetModuleInstance("hvac");
                if (module == null)
                {
                    TaskDialog.Show("Error", "HVAC module not loaded. Please authenticate first.");
                    return Result.Failed;
                }

                // Create ExternalEvents (in IExternalCommand context, which is valid)
                if (_createSchemaEvent == null)
                {
                    var createHandler = new HVAC.Module.Commands.CreateSchemaHandler();
                    _createSchemaEvent = ExternalEvent.Create(createHandler);
                }
                
                if (_completeSchemaEvent == null)
                {
                    var completeHandler = new HVAC.Module.Commands.CompleteSchemaHandler();
                    _completeSchemaEvent = ExternalEvent.Create(completeHandler);
                }
                
                if (_settingsEvent == null)
                {
                    var settingsHandler = new HVAC.Module.Commands.SettingsHandler();
                    _settingsEvent = ExternalEvent.Create(settingsHandler);
                }
                
                // Initialize Updater ONCE (only first time module is opened)
                PluginsManager.Core.DebugLogger.Log($"[OpenHvacPanelCommand] Checking Updater initialization. Current: {(HVACSuperScheme.App._updater == null ? "NULL" : "EXISTS")}");
                
                if (HVACSuperScheme.App._updater == null)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Initializing Updater for first time...");
                    HVACSuperScheme.Utils.LoggingUtils.Logging("[HVAC-INIT] Initializing Updater from OpenHVACPanelCommand", "-");
                    
                    // Read settings first
                    SettingStorage.ReadSettings();
                    PluginsManager.Core.DebugLogger.Log($"[HVAC-INIT] Settings loaded. IsUpdaterSync={SettingStorage.Instance.IsUpdaterSync}");
                    HVACSuperScheme.Utils.LoggingUtils.Logging($"[HVAC-INIT] Settings: IsUpdaterSync={SettingStorage.Instance.IsUpdaterSync}", "-");
                    
                    // Store UIApplication for Idling events (plugin mode)
                    HVACSuperScheme.App._uiApplication = uiApp;
                    
                    // Initialize element filters BEFORE creating Updater
                    FilterUtils.InitElementFilters();
                    PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Element filters initialized");
                    
                    // Get AddInId from the current application
                    var addInId = uiApp.ActiveAddInId;
                    
                    // Create Updater instance
                    HVACSuperScheme.App._updater = new Updater(addInId);
                    
                    // Register the updater
                    UpdaterRegistry.RegisterUpdater(HVACSuperScheme.App._updater);
                    PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Updater registered");
                    
                    // Add triggers for element deletion and addition
                    HVACSuperScheme.App.CreateDeletionTriggers();
                    HVACSuperScheme.App.CreateAdditionTriggers();
                    PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Deletion and addition triggers created");
                    
                    // Start IdlingHandler to create parameter change triggers when document is ready
                    // Only if sync is enabled in settings
                    if (SettingStorage.Instance.IsUpdaterSync && !HVACSuperScheme.App._idlingHandlerIsActive)
                    {
                        HVACSuperScheme.App.CreateIdlingHandler();
                        PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] IdlingHandler activated (sync enabled)");
                    }
                    else if (!SettingStorage.Instance.IsUpdaterSync)
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] IdlingHandler NOT started (sync disabled in settings)");
                    }
                    
                    PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Updater initialization complete");
                }
                
                // Call CreatePanel with ExternalEvents
                var window = module.CreatePanel(new object[] { 
                    uiApp, 
                    _createSchemaEvent, 
                    _completeSchemaEvent, 
                    _settingsEvent 
                });
                
                if (window != null)
                {
                    // Find MainHubPanel window
                    var hubWindow = System.Windows.Application.Current.Windows
                        .OfType<System.Windows.Window>()
                        .FirstOrDefault(w => w.Title.Contains("AnnotatiX"));
                    
                    if (hubWindow != null)
                    {
                        // Find MainHubPanel instance and show HVAC content
                        var mainHubPanel = FindMainHubPanel(hubWindow);
                        if (mainHubPanel != null)
                        {
                            var hvacPanelContent = window.Content;
                            var method = mainHubPanel.GetType().GetMethod("ShowHvacPanel");
                            if (method != null)
                            {
                                method.Invoke(mainHubPanel, new object[] { hvacPanelContent });
                            }
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"Failed to open HVAC panel: {ex.Message}");
                return Result.Failed;
            }
        }
        
        private object FindMainHubPanel(System.Windows.Window window)
        {
            // Traverse visual tree to find MainHubPanel instance
            var content = window.Content;
            if (content != null && content.GetType().Name == "MainHubPanel")
            {
                return content;
            }
            return null;
        }
    }
}
