using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace Annotatix.Module
{
    /// <summary>
    /// Annotatix Module - Data collection for ML-based annotation automation
    /// Records view state before and after annotation placement
    /// </summary>
    public class AnnotatixModule : IModule
    {
        public string ModuleId => "annotatix";
        public string ModuleName => "Annotatix";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[ANNOTATIX-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[ANNOTATIX-MODULE] Annotatix module loaded dynamically after authentication");
        }

        public Window CreatePanel(object[] parameters)
        {
            UIApplication uiApp = null;
            UI.AnnotatixPanelHandler handler = null;
            ExternalEvent extEvent = null;

            if (parameters != null && parameters.Length > 0)
            {
                uiApp = parameters[0] as UIApplication;

                if (parameters.Length >= 3)
                {
                    handler = parameters[1] as UI.AnnotatixPanelHandler;
                    extEvent = parameters[2] as ExternalEvent;
                }
            }

            DebugLogger.Log("[ANNOTATIX-MODULE] Creating AnnotatixPanel for Plugins Hub...");

            var panel = new UI.AnnotatixPanel(uiApp, handler, extEvent);

            var window = new Window
            {
                Content = panel,
                Title = "Annotatix - Управление аннотациями и анализом",
                Width = 500,
                Height = 450,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize
            };

            DebugLogger.Log("[ANNOTATIX-MODULE] Panel window created successfully");

            return window;
        }
    }
}
