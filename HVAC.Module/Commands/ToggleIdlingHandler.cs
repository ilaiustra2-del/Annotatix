using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace HVAC.Module.Commands
{
    /// <summary>
    /// Handler to start/stop IdlingHandler from UI thread
    /// </summary>
    public class ToggleIdlingHandler : IExternalEventHandler
    {
        private bool _shouldStart = true;

        public void SetStart(bool start)
        {
            _shouldStart = start;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (_shouldStart)
                {
                    DebugLogger.Log("[HVAC-TOGGLE] Starting IdlingHandler via ExternalEvent...");
                    HVACSuperScheme.App.CreateIdlingHandler();
                    DebugLogger.Log("[HVAC-TOGGLE] IdlingHandler started successfully");
                }
                else
                {
                    DebugLogger.Log("[HVAC-TOGGLE] Stopping IdlingHandler via ExternalEvent...");
                    HVACSuperScheme.App.RemoveIdlingHandler();
                    DebugLogger.Log("[HVAC-TOGGLE] IdlingHandler stopped successfully");
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[HVAC-TOGGLE] ERROR: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "ToggleIdlingHandler";
        }
    }
}
