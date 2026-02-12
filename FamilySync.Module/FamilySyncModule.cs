using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace FamilySync.Module
{
    /// <summary>
    /// Family Sync Module - Analyzes and synchronizes nested family parameters
    /// Implements IModule interface for dynamic loading
    /// </summary>
    public class FamilySyncModule : IModule
    {
        public string ModuleId => "family_sync";
        public string ModuleName => "Family Sync";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[FAMILY-SYNC-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[FAMILY-SYNC-MODULE] Module loaded with main plugin");
        }

        public Window CreatePanel(object[] parameters)
        {
            UIApplication uiApp = null;
            UI.FamilySyncHandler syncHandler = null;
            ExternalEvent syncEvent = null;
            
            if (parameters != null && parameters.Length > 0)
            {
                uiApp = parameters[0] as UIApplication;
                
                // Check if ExternalEvent was provided
                if (parameters.Length >= 3)
                {
                    syncHandler = parameters[1] as UI.FamilySyncHandler;
                    syncEvent = parameters[2] as ExternalEvent;
                }
            }

            DebugLogger.Log("[FAMILY-SYNC-MODULE] Creating FamilySyncPanel...");
            
            var panel = new UI.FamilySyncPanel(uiApp, syncHandler, syncEvent);
            
            var window = new Window
            {
                Content = panel,
                Title = "Family Sync - Синхронизация вложенных семейств",
                Width = 900,
                Height = 700,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize
            };

            DebugLogger.Log("[FAMILY-SYNC-MODULE] Panel window created successfully");
            
            return window;
        }
    }
}
