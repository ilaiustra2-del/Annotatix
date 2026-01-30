using System;
using System.Windows;
using PluginsManager.Core;

namespace HVAC.Module
{
    /// <summary>
    /// HVAC Module implementation
    /// </summary>
    public class HVACModule : IModule
    {
        public string ModuleId => "hvac";
        public string ModuleName => "HVAC";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            // Module initialization logic
            System.Diagnostics.Debug.WriteLine("[HVAC-MODULE] HVAC Module initialized");
        }

        public Window CreatePanel(object[] parameters)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[HVAC-MODULE] Creating HVAC panel");
                var panel = new UI.HVACPanel();
                
                // Wrap UserControl in a Window
                var window = new Window
                {
                    Content = panel,
                    Title = "HVAC",
                    Width = 800,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };
                
                return window;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HVAC-MODULE] ERROR creating panel: {ex.Message}");
                return null;
            }
        }
    }
}
