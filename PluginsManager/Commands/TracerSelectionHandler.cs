using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Subscribes to UIApplication.SelectionChanged and forwards events to TracerSession.
    /// Must be registered/unregistered within the Revit API event thread.
    /// </summary>
    public class TracerSelectionHandler
    {
        private UIApplication _uiApp;

        public void Register(UIApplication uiApp)
        {
            _uiApp = uiApp;
            uiApp.SelectionChanged += OnSelectionChanged;
            Core.DebugLogger.Log("[TRACER-SELECTION] SelectionChanged handler registered");
        }

        public void Unregister()
        {
            if (_uiApp != null)
            {
                _uiApp.SelectionChanged -= OnSelectionChanged;
                Core.DebugLogger.Log("[TRACER-SELECTION] SelectionChanged handler unregistered");
            }
            _uiApp = null;
        }

        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e)
        {
            try
            {
                // Route to TracerSession if active
                TracerSession.Current?.OnSelectionChanged(_uiApp);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[TRACER-SELECTION] OnSelectionChanged error: {ex.Message}");
            }
        }
    }
}
