using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace HVAC.Module
{
    /// <summary>
    /// HVAC SuperScheme Module - Dynamic loading wrapper
    /// </summary>
    public class HVACModule : IModule
    {
        public string ModuleId => "hvac";
        public string ModuleName => "HVAC SuperScheme";
        public string ModuleVersion => "3.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[HVAC-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[HVAC-MODULE] HVAC SuperScheme module loaded dynamically after authentication");
        }

        public Window CreatePanel(object[] parameters)
        {
            if (parameters == null || parameters.Length == 0)
            {
                throw new ArgumentException("UIApplication parameter is required");
            }

            var uiApp = parameters[0] as UIApplication;
            if (uiApp == null)
            {
                throw new ArgumentException("First parameter must be UIApplication");
            }
            
            // ExternalEvents are optional (parameters[1], [2], [3])
            // If provided, they MUST be created in IExternalCommand.Execute() context
            ExternalEvent createSchemaEvent = parameters.Length > 1 ? parameters[1] as ExternalEvent : null;
            ExternalEvent completeSchemaEvent = parameters.Length > 2 ? parameters[2] as ExternalEvent : null;
            ExternalEvent settingsEvent = parameters.Length > 3 ? parameters[3] as ExternalEvent : null;

            try
            {
                DebugLogger.Log("[HVAC-MODULE] Creating HVAC SuperScheme panel...");
                DebugLogger.Log($"[HVAC-MODULE] ExternalEvents provided: createSchema={createSchemaEvent != null}, completeSchema={completeSchemaEvent != null}, settings={settingsEvent != null}");
                
                var panel = new UI.HVACPanel(uiApp, createSchemaEvent, completeSchemaEvent, settingsEvent);
                
                var window = new Window
                {
                    Content = panel,
                    Title = "HVAC SuperScheme",
                    Width = 900,
                    Height = 700,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };
                
                DebugLogger.Log("[HVAC-MODULE] Panel window created successfully");
                
                return window;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-MODULE] ERROR creating panel: {ex.Message}");
                DebugLogger.Log($"[HVAC-MODULE] Stack trace: {ex.StackTrace}");
                return null;
            }
        }
    }
}
