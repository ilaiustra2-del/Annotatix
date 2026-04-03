using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace Tracer.Module
{
    /// <summary>
    /// Tracer Module - Sewage pipe connection tool
    /// Connects vertical risers to sloped main lines at 45° angle
    /// </summary>
    public class TracerModule : IModule
    {
        public string ModuleId => "tracer";
        public string ModuleName => "Трассировка канализации";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[TRACER-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[TRACER-MODULE] Tracer module loaded dynamically after authentication");
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
            
            // ExternalEvents for Revit operations
            ExternalEvent selectMainPipeEvent = parameters.Length > 1 ? parameters[1] as ExternalEvent : null;
            ExternalEvent selectRiserEvent = parameters.Length > 2 ? parameters[2] as ExternalEvent : null;
            ExternalEvent createConnectionEvent = parameters.Length > 3 ? parameters[3] as ExternalEvent : null;
            ExternalEvent createLConnectionEvent = parameters.Length > 4 ? parameters[4] as ExternalEvent : null;

            try
            {
                DebugLogger.Log("[TRACER-MODULE] Creating Tracer panel...");
                DebugLogger.Log($"[TRACER-MODULE] ExternalEvents provided: selectMainPipe={selectMainPipeEvent != null}, selectRiser={selectRiserEvent != null}, createConnection={createConnectionEvent != null}, createLConnection={createLConnectionEvent != null}");
                
                var panel = new UI.TracerPanel(uiApp, selectMainPipeEvent, selectRiserEvent, createConnectionEvent, createLConnectionEvent);
                
                var window = new Window
                {
                    Content = panel,
                    Title = "Трассировка канализации",
                    Width = 500,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.CanResize
                };
                
                DebugLogger.Log("[TRACER-MODULE] Panel window created successfully");
                
                return window;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-MODULE] ERROR creating panel: {ex.Message}");
                DebugLogger.Log($"[TRACER-MODULE] Stack trace: {ex.StackTrace}");
                return null;
            }
        }
    }
}
