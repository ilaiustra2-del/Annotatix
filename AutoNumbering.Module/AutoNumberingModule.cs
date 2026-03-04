using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace AutoNumbering.Module
{
    /// <summary>
    /// AutoNumbering Module - Automatic numbering for vertical pipes (risers)
    /// Implements IModule interface for dynamic loading
    /// </summary>
    public class AutoNumberingModule : IModule
    {
        public string ModuleId => "autonumbering";
        public string ModuleName => "AutoNumbering";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[AUTONUMBERING-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[AUTONUMBERING-MODULE] AutoNumbering module loaded with main plugin");
        }

        public Window CreatePanel(object[] parameters)
        {
            UIApplication uiApp = null;
            ExternalEvent numberingEvent = null;
            object numberingHandler = null;
            
            if (parameters != null && parameters.Length > 0)
            {
                uiApp = parameters[0] as UIApplication;
            }
            if (parameters != null && parameters.Length > 1)
            {
                numberingEvent = parameters[1] as ExternalEvent;
            }
            if (parameters != null && parameters.Length > 2)
            {
                numberingHandler = parameters[2];
            }

            DebugLogger.Log("[AUTONUMBERING-MODULE] Creating AutoNumberingPanel...");
            DebugLogger.Log($"[AUTONUMBERING-MODULE] Parameters: UIApp={uiApp != null}, Event={numberingEvent != null}, Handler={numberingHandler != null}");
            
            var panel = new UI.AutoNumberingPanel(uiApp, numberingEvent, numberingHandler);
            
            var window = new Window
            {
                Content = panel,
                Title = "AutoNumbering - Автонумерация стояков",
                Width = 700,
                Height = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize
            };

            DebugLogger.Log("[AUTONUMBERING-MODULE] Panel window created successfully");
            
            return window;
        }
    }
}
