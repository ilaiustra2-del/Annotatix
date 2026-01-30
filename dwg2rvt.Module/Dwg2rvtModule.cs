using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace dwg2rvt.Module
{
    /// <summary>
    /// DWG to Revit Analysis Module
    /// Implements IModule interface for dynamic loading
    /// </summary>
    public class Dwg2rvtModule : IModule
    {
        public string ModuleId => "dwg2rvt";
        public string ModuleName => "DWG2RVT";
        public string ModuleVersion => "3.004";

        public void Initialize()
        {
            DebugLogger.Log("[DWG2RVT-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[DWG2RVT-MODULE] Module loaded dynamically after authentication");
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
            ExternalEvent annotateEvent = parameters.Length > 1 ? parameters[1] as ExternalEvent : null;
            ExternalEvent placeElementsEvent = parameters.Length > 2 ? parameters[2] as ExternalEvent : null;
            ExternalEvent placeSingleBlockEvent = parameters.Length > 3 ? parameters[3] as ExternalEvent : null;

            DebugLogger.Log("[DWG2RVT-MODULE] Creating dwg2rvtPanel...");
            
            if (annotateEvent != null && placeElementsEvent != null && placeSingleBlockEvent != null)
            {
                DebugLogger.Log("[DWG2RVT-MODULE] ExternalEvents provided - full functionality enabled");
            }
            else
            {
                DebugLogger.Log("[DWG2RVT-MODULE] WARNING: ExternalEvents not provided");
                DebugLogger.Log("[DWG2RVT-MODULE] Buttons 'Annotate' and 'Place Elements' will be disabled");
            }
            
            var panel = new UI.dwg2rvtPanel(uiApp, annotateEvent, placeElementsEvent, placeSingleBlockEvent);
            
            var window = new Window
            {
                Content = panel,
                Title = "DWG2RVT - Анализ блоков",
                Width = 900,
                Height = 700,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            DebugLogger.Log("[DWG2RVT-MODULE] Panel window created successfully");
            
            return window;
        }
    }
}
