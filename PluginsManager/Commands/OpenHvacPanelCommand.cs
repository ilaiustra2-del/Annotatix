using System;
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
