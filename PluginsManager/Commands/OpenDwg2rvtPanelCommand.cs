using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Command to open DWG2RVT panel with properly initialized ExternalEvents
    /// This command MUST be executed in IExternalCommand.Execute() context
    /// to allow creation of ExternalEvents
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenDwg2rvtPanelCommand : IExternalCommand
    {
        // Store events statically to reuse across invocations
        private static ExternalEvent _annotateEvent;
        private static ExternalEvent _placeElementsEvent;
        private static ExternalEvent _placeSingleBlockTypeEvent;
        private static object _placeSingleBlockTypeHandler; // Store as object to avoid circular reference
        
        // Store MainHubPanel reference to update it after panel creation
        private static UI.MainHubPanel _hubPanel;
        
        public static void SetHubPanel(UI.MainHubPanel hubPanel)
        {
            _hubPanel = hubPanel;
        }
        
        /// <summary>
        /// Set the block type name for single block placement
        /// This is called from dwg2rvtPanel before raising the event
        /// </summary>
        public static void SetBlockTypeNameForPlacement(string blockTypeName)
        {
            if (_placeSingleBlockTypeHandler != null)
            {
                // Use reflection to set the property
                var handlerType = _placeSingleBlockTypeHandler.GetType();
                var property = handlerType.GetProperty("BlockTypeName");
                if (property != null)
                {
                    property.SetValue(_placeSingleBlockTypeHandler, blockTypeName);
                    Core.DebugLogger.Log($"[OpenDwg2rvtPanelCommand] BlockTypeName set to: {blockTypeName}");
                }
            }
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
                Core.DebugLogger.Log("[OpenDwg2rvtPanelCommand] Opening DWG2RVT panel...");
                
                // Get module instance
                var module = Core.DynamicModuleLoader.GetModuleInstance("dwg2rvt");
                if (module == null)
                {
                    message = "DWG2RVT module not loaded";
                    return Result.Failed;
                }
                
                // Create ExternalEvents if not already created (only once per session)
                // These MUST be created here in IExternalCommand.Execute() context
                if (_annotateEvent == null)
                {
                    Core.DebugLogger.Log("[OpenDwg2rvtPanelCommand] Creating ExternalEvents...");
                    
                    // We need to create handlers from dwg2rvt.Module
                    // But we can't reference it directly, so we'll use reflection
                    
                    // Get the assembly of the module
                    var moduleAssembly = module.GetType().Assembly;
                    
                    // Find handler types by name
                    var annotateHandlerType = moduleAssembly.GetType("dwg2rvt.Module.UI.dwg2rvtPanel+AnnotateEventHandler");
                    var placeElementsHandlerType = moduleAssembly.GetType("dwg2rvt.Module.UI.dwg2rvtPanel+PlaceElementsEventHandler");
                    var placeSingleHandlerType = moduleAssembly.GetType("dwg2rvt.Module.UI.dwg2rvtPanel+PlaceSingleBlockTypeEventHandler");
                    
                    if (annotateHandlerType == null || placeElementsHandlerType == null || placeSingleHandlerType == null)
                    {
                        message = "Failed to find handler types in module";
                        Core.DebugLogger.Log($"[OpenDwg2rvtPanelCommand] ERROR: Handler types not found");
                        return Result.Failed;
                    }
                    
                    // Create handler instances
                    var annotateHandler = Activator.CreateInstance(annotateHandlerType) as IExternalEventHandler;
                    var placeElementsHandler = Activator.CreateInstance(placeElementsHandlerType) as IExternalEventHandler;
                    var placeSingleHandler = Activator.CreateInstance(placeSingleHandlerType) as IExternalEventHandler;
                    
                    // Store the placeSingleHandler for later access
                    _placeSingleBlockTypeHandler = placeSingleHandler;
                    
                    // Create ExternalEvents
                    _annotateEvent = ExternalEvent.Create(annotateHandler);
                    _placeElementsEvent = ExternalEvent.Create(placeElementsHandler);
                    _placeSingleBlockTypeEvent = ExternalEvent.Create(placeSingleHandler);
                    
                    Core.DebugLogger.Log("[OpenDwg2rvtPanelCommand] ExternalEvents created successfully");
                }
                
                // Create panel from module with ExternalEvents
                var panel = module.CreatePanel(new object[] { 
                    uiApp, 
                    _annotateEvent, 
                    _placeElementsEvent, 
                    _placeSingleBlockTypeEvent 
                });
                
                if (panel == null)
                {
                    message = "Failed to create panel";
                    return Result.Failed;
                }
                
                Core.DebugLogger.Log("[OpenDwg2rvtPanelCommand] Panel created successfully");
                
                // Link the panel to the PlaceSingleBlockTypeEventHandler
                // We need to get the panel instance from the Window.Content and call SetPanel() on the handler
                try
                {
                    var panelContent = panel.Content;
                    if (panelContent != null)
                    {
                        // Use reflection to call SetPanel on the handler
                        var setPanelMethod = _placeSingleBlockTypeHandler.GetType().GetMethod("SetPanel");
                        if (setPanelMethod != null)
                        {
                            setPanelMethod.Invoke(_placeSingleBlockTypeHandler, new object[] { panelContent });
                            Core.DebugLogger.Log("[OpenDwg2rvtPanelCommand] PlaceSingleBlockTypeHandler linked to panel");
                        }
                    }
                }
                catch (Exception linkEx)
                {
                    Core.DebugLogger.Log($"[OpenDwg2rvtPanelCommand] WARNING: Failed to link handler to panel: {linkEx.Message}");
                }
                
                // Update MainHubPanel with the panel content
                if (_hubPanel != null)
                {
                    var panelContent = panel.Content;
                    if (panelContent != null)
                    {
                        _hubPanel.ShowDwg2rvtPanel(panelContent);
                        Core.DebugLogger.Log("[OpenDwg2rvtPanelCommand] Panel shown in Hub");
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Core.DebugLogger.Log($"[OpenDwg2rvtPanelCommand] ERROR: {ex.Message}");
                Core.DebugLogger.Log($"[OpenDwg2rvtPanelCommand] StackTrace: {ex.StackTrace}");
                return Result.Failed;
            }
        }
    }
}
