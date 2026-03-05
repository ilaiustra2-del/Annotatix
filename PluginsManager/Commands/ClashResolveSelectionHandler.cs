using System;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Subscribes to UIApplication.SelectionChanged and forwards events to ClashResolveSession.
    /// Must be registered/unregistered within the Revit API event thread.
    /// </summary>
    public class ClashResolveSelectionHandler
    {
        private UIApplication _uiApp;

        public void Register(UIApplication uiApp)
        {
            _uiApp = uiApp;
            uiApp.SelectionChanged += OnSelectionChanged;
            Core.DebugLogger.Log("[CLASH-SELECTION] SelectionChanged handler registered");
        }

        public void Unregister()
        {
            if (_uiApp != null)
            {
                _uiApp.SelectionChanged -= OnSelectionChanged;
                Core.DebugLogger.Log("[CLASH-SELECTION] SelectionChanged handler unregistered");
            }
            _uiApp = null;
        }

        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e)
        {
            try
            {
                ClashResolveSession.Current?.OnSelectionChanged(_uiApp);
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[CLASH-SELECTION] OnSelectionChanged error: {ex.Message}");
            }
        }
    }
}
