using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Command to open HVAC panel with properly initialized ExternalEvents
    /// This command MUST be executed in IExternalCommand.Execute() context
    /// to allow creation of ExternalEvents
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenHvacPanelCommand : IExternalCommand
    {
        // Store events statically to reuse across invocations
        private static ExternalEvent _createSchemaEvent;
        private static ExternalEvent _completeSchemaEvent;
        private static ExternalEvent _settingsEvent;
        
        // Store MainHubPanel reference to update it after panel creation
        private static UI.MainHubPanel _hubPanel;
        
        public static void SetHubPanel(UI.MainHubPanel hubPanel)
        {
            _hubPanel = hubPanel;
        }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteInternal(commandData.Application, ref message);
        }
        
        /// <summary>
        /// Internal execution logic that can be called from both IExternalCommand and IExternalEventHandler
        /// </summary>
        public Result ExecuteInternal(UIApplication uiApp, ref string message)
        {
            try
            {
                Core.DebugLogger.Log("[OpenHvacPanelCommand] Opening HVAC panel...");
                
                // Get module instance
                var module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                if (module == null)
                {
                    message = "HVAC module not loaded";
                    return Result.Failed;
                }
                
                // Create ExternalEvents if not already created (only once per session)
                // These MUST be created here in IExternalCommand.Execute() context
                if (_createSchemaEvent == null)
                {
                    Core.DebugLogger.Log("[OpenHvacPanelCommand] Creating ExternalEvents...");
                    
                    // Get the assembly of the module
                    var moduleAssembly = module.GetType().Assembly;
                    
                    // Find handler types by name
                    var createSchemaHandlerType = moduleAssembly.GetType("HVAC.Module.Commands.CreateSchemaHandler");
                    var completeSchemaHandlerType = moduleAssembly.GetType("HVAC.Module.Commands.CompleteSchemaHandler");
                    var settingsHandlerType = moduleAssembly.GetType("HVAC.Module.Commands.SettingsHandler");
                    
                    if (createSchemaHandlerType == null || completeSchemaHandlerType == null || settingsHandlerType == null)
                    {
                        message = "Failed to find handler types in HVAC module";
                        Core.DebugLogger.Log($"[OpenHvacPanelCommand] ERROR: Handler types not found");
                        Core.DebugLogger.Log($"[OpenHvacPanelCommand] createSchemaHandlerType: {createSchemaHandlerType != null}");
                        Core.DebugLogger.Log($"[OpenHvacPanelCommand] completeSchemaHandlerType: {completeSchemaHandlerType != null}");
                        Core.DebugLogger.Log($"[OpenHvacPanelCommand] settingsHandlerType: {settingsHandlerType != null}");
                        return Result.Failed;
                    }
                    
                    // Create handler instances (pass UIApplication to constructors)
                    var createSchemaHandler = Activator.CreateInstance(createSchemaHandlerType, uiApp) as IExternalEventHandler;
                    var completeSchemaHandler = Activator.CreateInstance(completeSchemaHandlerType, uiApp) as IExternalEventHandler;
                    var settingsHandler = Activator.CreateInstance(settingsHandlerType, uiApp) as IExternalEventHandler;
                    
                    // Create ExternalEvents
                    _createSchemaEvent = ExternalEvent.Create(createSchemaHandler);
                    _completeSchemaEvent = ExternalEvent.Create(completeSchemaHandler);
                    _settingsEvent = ExternalEvent.Create(settingsHandler);
                    
                    Core.DebugLogger.Log("[OpenHvacPanelCommand] ExternalEvents created successfully");
                    
                    // Initialize Updater ONCE (only first time module is opened)
                    try
                    {
                        // Get HVACSuperScheme.App type via reflection
                        var hvacApp = moduleAssembly.GetType("HVACSuperScheme.App");
                        if (hvacApp != null)
                        {
                            var updaterField = hvacApp.GetField("_updater", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                            var updaterValue = updaterField?.GetValue(null);
                            
                            Core.DebugLogger.Log($"[HVAC-INIT] Checking Updater. Current: {(updaterValue == null ? "NULL" : "EXISTS")}");
                            
                            if (updaterValue == null)
                            {
                                Core.DebugLogger.Log("[HVAC-INIT] Initializing Updater for first time...");
                                
                                // Call LoggingUtils.Logging via reflection
                                try
                                {
                                    var loggingUtilsType = moduleAssembly.GetType("HVACSuperScheme.Utils.LoggingUtils");
                                    var loggingMethod = loggingUtilsType?.GetMethod("Logging", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                    loggingMethod?.Invoke(null, new object[] { "[HVAC-INIT] Initializing Updater from OpenHVACPanelCommand", "-", "", 0, "" });
                                }
                                catch { }
                                
                                // Read settings first
                                var settingStorageType = moduleAssembly.GetType("HVACSuperScheme.Data.SettingStorage");
                                var readSettingsMethod = settingStorageType?.GetMethod("ReadSettings", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                readSettingsMethod?.Invoke(null, null);
                                
                                var instanceProp = settingStorageType?.GetProperty("Instance", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                var settingsInstance = instanceProp?.GetValue(null);
                                var isUpdaterSyncProp = settingsInstance?.GetType().GetProperty("IsUpdaterSync");
                                var isUpdaterSync = (bool?)isUpdaterSyncProp?.GetValue(settingsInstance) ?? false;
                                
                                Core.DebugLogger.Log($"[HVAC-INIT] Settings loaded. IsUpdaterSync={isUpdaterSync}");
                                
                                // Call LoggingUtils.Logging via reflection
                                try
                                {
                                    var loggingUtilsType = moduleAssembly.GetType("HVACSuperScheme.Utils.LoggingUtils");
                                    var loggingMethod = loggingUtilsType?.GetMethod("Logging", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                    loggingMethod?.Invoke(null, new object[] { $"[HVAC-INIT] Settings: IsUpdaterSync={isUpdaterSync}", "-", "", 0, "" });
                                }
                                catch { }
                                
                                // Store UIApplication for Idling events (plugin mode)
                                var uiApplicationField = hvacApp.GetField("_uiApplication", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                                if (uiApplicationField != null)
                                {
                                    uiApplicationField.SetValue(null, uiApp);
                                    Core.DebugLogger.Log($"[HVAC-INIT] _uiApplication field SET via reflection");
                                }
                                else
                                {
                                    Core.DebugLogger.Log($"[HVAC-INIT] ERROR: _uiApplication field NOT FOUND in HVACSuperScheme.App");
                                    // Try to list all fields for debugging
                                    var allFields = hvacApp.GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                                    Core.DebugLogger.Log($"[HVAC-INIT] Available static fields: {string.Join(", ", allFields.Select(f => f.Name))}");
                                }
                                
                                // Initialize element filters BEFORE creating Updater
                                var filterUtilsType = moduleAssembly.GetType("HVACSuperScheme.Data.FilterUtils");
                                var initFiltersMethod = filterUtilsType?.GetMethod("InitElementFilters", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                initFiltersMethod?.Invoke(null, null);
                                Core.DebugLogger.Log("[HVAC-INIT] Element filters initialized");
                                
                                // Get AddInId from the current application
                                var addInId = uiApp.ActiveAddInId;
                                
                                // Create Updater instance
                                var updaterType = moduleAssembly.GetType("HVACSuperScheme.Updaters.Updater");
                                var updater = Activator.CreateInstance(updaterType, addInId);
                                updaterField?.SetValue(null, updater);
                                
                                // Register the updater
                                UpdaterRegistry.RegisterUpdater(updater as IUpdater);
                                Core.DebugLogger.Log("[HVAC-INIT] Updater registered");
                                
                                // Add triggers for element deletion and addition
                                try
                                {
                                    var createDeletionMethod = hvacApp.GetMethod("CreateDeletionTriggers", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                    var createAdditionMethod = hvacApp.GetMethod("CreateAdditionTriggers", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                    createDeletionMethod?.Invoke(null, null);
                                    createAdditionMethod?.Invoke(null, null);
                                    Core.DebugLogger.Log("[HVAC-INIT] Deletion and addition triggers created");
                                }
                                catch (Exception triggerEx)
                                {
                                    // This may fail if document is not ready - triggers will be created by IdlingHandler
                                    Core.DebugLogger.Log($"[HVAC-INIT] WARNING: Could not create triggers now: {triggerEx.InnerException?.Message ?? triggerEx.Message}");
                                    Core.DebugLogger.Log("[HVAC-INIT] Triggers will be created later by IdlingHandler");
                                }
                                
                                // Start IdlingHandler to create parameter change triggers when document is ready
                                // Only if sync is enabled in settings
                                if (isUpdaterSync)
                                {
                                    var idlingActiveField = hvacApp.GetField("_idlingHandlerIsActive", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                    var idlingActive = (bool?)idlingActiveField?.GetValue(null) ?? false;
                                    
                                    if (!idlingActive)
                                    {
                                        var createIdlingMethod = hvacApp.GetMethod("CreateIdlingHandler", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                        createIdlingMethod?.Invoke(null, null);
                                        Core.DebugLogger.Log("[HVAC-INIT] IdlingHandler activated (sync enabled)");
                                    }
                                }
                                else
                                {
                                    Core.DebugLogger.Log("[HVAC-INIT] IdlingHandler NOT started (sync disabled in settings)");
                                }
                                
                                Core.DebugLogger.Log("[HVAC-INIT] Updater initialization complete");
                                
                                // Call LoggingUtils.Logging via reflection
                                try
                                {
                                    var loggingUtilsType = moduleAssembly.GetType("HVACSuperScheme.Utils.LoggingUtils");
                                    var loggingMethod = loggingUtilsType?.GetMethod("Logging", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                                    loggingMethod?.Invoke(null, new object[] { "[HVAC-INIT] Updater initialization complete", "-", "", 0, "" });
                                }
                                catch { }
                            }
                            else
                            {
                                Core.DebugLogger.Log("[HVAC-INIT] Updater already initialized");
                            }
                        }
                    }
                    catch (Exception initEx)
                    {
                        Core.DebugLogger.Log($"[HVAC-INIT] ERROR initializing Updater: {initEx.Message}");
                        Core.DebugLogger.Log($"[HVAC-INIT] Stack trace: {initEx.StackTrace}");
                    }
                }
                
                // Create panel from module with ExternalEvents
                var panel = module.CreatePanel(new object[] { 
                    uiApp, 
                    _createSchemaEvent, 
                    _completeSchemaEvent, 
                    _settingsEvent 
                });
                
                if (panel == null)
                {
                    message = "Failed to create HVAC panel";
                    return Result.Failed;
                }
                
                Core.DebugLogger.Log("[OpenHvacPanelCommand] Panel created successfully");
                
                // Update MainHubPanel with the panel content
                if (_hubPanel != null)
                {
                    try
                    {
                        var panelContent = panel.Content;
                        if (panelContent != null)
                        {
                            _hubPanel.Dispatcher.Invoke(() =>
                            {
                                _hubPanel.ShowHvacPanel(panelContent);
                            });
                            
                            Core.DebugLogger.Log("[OpenHvacPanelCommand] Panel displayed in hub");
                        }
                    }
                    catch (Exception dispatchEx)
                    {
                        Core.DebugLogger.Log($"[OpenHvacPanelCommand] ERROR updating UI: {dispatchEx.Message}");
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error opening HVAC panel: {ex.Message}";
                Core.DebugLogger.Log($"[OpenHvacPanelCommand] ERROR: {ex.Message}");
                Core.DebugLogger.Log($"[OpenHvacPanelCommand] Stack trace: {ex.StackTrace}");
                return Result.Failed;
            }
        }
    }
    
    /// <summary>
    /// ExternalEventHandler wrapper to open HVAC panel
    /// </summary>
    public class OpenHvacPanelHandler : IExternalEventHandler
    {
        private UIApplication _uiApp;
        
        public void SetUIApplication(UIApplication uiApp)
        {
            _uiApp = uiApp;
        }
        
        public void Execute(UIApplication app)
        {
            string message = "";
            var command = new OpenHvacPanelCommand();
            command.ExecuteInternal(app ?? _uiApp, ref message);
        }
        
        public string GetName()
        {
            return "OpenHvacPanelHandler";
        }
    }
}
