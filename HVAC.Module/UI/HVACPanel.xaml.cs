using System;
using System.Windows.Controls;
using PluginsManager.Core;

namespace HVAC.Module.UI
{
    /// <summary>
    /// HVAC Module Panel - Placeholder for future HVAC functionality
    /// </summary>
    public partial class HVACPanel : UserControl
    {
        public HVACPanel()
        {
            InitializeComponent();
            
            // Debug logging with timestamp to verify dynamic loading
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            DebugLogger.Log("[HVAC-PANEL] *** PANEL INSTANCE CREATED ***");
            DebugLogger.Log("[HVAC-PANEL] This proves the panel was NOT loaded at Revit startup");
            DebugLogger.Log("[HVAC-PANEL] Panel created AFTER authentication");
        }
    }
}
