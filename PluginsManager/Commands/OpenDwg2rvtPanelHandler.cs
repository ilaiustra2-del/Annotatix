using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Handler to invoke OpenDwg2rvtPanelCommand from UI button click
    /// </summary>
    public class OpenDwg2rvtPanelHandler : IExternalEventHandler
    {
        private UIApplication _uiApp;
        
        public void SetUIApplication(UIApplication uiApp)
        {
            _uiApp = uiApp;
        }
        
        public void Execute(UIApplication app)
        {
            try
            {
                Core.DebugLogger.Log("[OpenDwg2rvtPanelHandler] Executing command...");
                
                // Create command instance and call Execute directly
                // We can't create ExternalCommandData, but the command doesn't actually need it
                // because it gets UIApplication from the handler parameter
                var command = new OpenDwg2rvtPanelCommand();
                
                // Call the command's internal execution logic
                // We'll need to refactor OpenDwg2rvtPanelCommand to have a method we can call
                string message = "";
                var result = command.ExecuteInternal(app, ref message);
                
                if (result != Result.Succeeded)
                {
                    Core.DebugLogger.Log($"[OpenDwg2rvtPanelHandler] Command failed: {message}");
                }
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[OpenDwg2rvtPanelHandler] ERROR: {ex.Message}");
            }
        }
        
        public string GetName()
        {
            return "OpenDwg2rvtPanelHandler";
        }
    }
}
