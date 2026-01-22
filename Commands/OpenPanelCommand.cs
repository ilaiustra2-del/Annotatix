using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace dwg2rvt.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenPanelCommand : IExternalCommand
    {
        private static ExternalEvent _annotateEvent;
        private static ExternalEvent _placeElementsEvent;
        private static ExternalEvent _placeSingleBlockTypeEvent;
        private static UI.dwg2rvtPanel.PlaceElementsEventHandler _placeElementsHandler;
        private static UI.dwg2rvtPanel.PlaceSingleBlockTypeEventHandler _placeSingleBlockTypeHandler;
        
        public static void SetBlockTypeNameForPlacement(string blockTypeName)
        {
            if (_placeSingleBlockTypeHandler != null)
            {
                _placeSingleBlockTypeHandler.BlockTypeName = blockTypeName;
            }
        }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc?.Document;

                if (doc == null)
                {
                    TaskDialog.Show("Error", "No active document found. Please open a Revit document.");
                    return Result.Failed;
                }

                // Create external events if not already created
                if (_annotateEvent == null)
                {
                    System.Diagnostics.Debug.WriteLine("[OpenPanelCommand] Creating annotateEvent");
                    UI.dwg2rvtPanel.AnnotateEventHandler annotateHandler = new UI.dwg2rvtPanel.AnnotateEventHandler();
                    _annotateEvent = ExternalEvent.Create(annotateHandler);
                }
                
                if (_placeElementsEvent == null)
                {
                    System.Diagnostics.Debug.WriteLine("[OpenPanelCommand] Creating placeElementsEvent");
                    _placeElementsHandler = new UI.dwg2rvtPanel.PlaceElementsEventHandler();
                    _placeElementsEvent = ExternalEvent.Create(_placeElementsHandler);
                }
                
                if (_placeSingleBlockTypeEvent == null)
                {
                    System.Diagnostics.Debug.WriteLine("[OpenPanelCommand] Creating placeSingleBlockTypeEvent");
                    _placeSingleBlockTypeHandler = new UI.dwg2rvtPanel.PlaceSingleBlockTypeEventHandler();
                    _placeSingleBlockTypeEvent = ExternalEvent.Create(_placeSingleBlockTypeHandler);
                    System.Diagnostics.Debug.WriteLine($"[OpenPanelCommand] placeSingleBlockTypeEvent created: {_placeSingleBlockTypeEvent != null}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[OpenPanelCommand] Reusing existing placeSingleBlockTypeEvent");
                }

                // Create panel
                System.Diagnostics.Debug.WriteLine("[OpenPanelCommand] Creating panel");
                var panel = new UI.dwg2rvtPanel(uiApp, _annotateEvent, _placeElementsEvent, _placeSingleBlockTypeEvent);
                
                // Set panel reference in handler
                System.Diagnostics.Debug.WriteLine("[OpenPanelCommand] Setting panel in handler");
                _placeSingleBlockTypeHandler.SetPanel(panel);
                
                // Create a window to host the user control
                System.Windows.Window window = new System.Windows.Window
                {
                    Title = "DWG2RVT - Control Panel",
                    Content = panel,
                    Width = 500,
                    Height = 450,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    ResizeMode = System.Windows.ResizeMode.NoResize
                };
                
                // Set the owner to Revit window to make it modeless but stay on top
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(window);
                helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;
                
                window.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"Failed to open control panel: {ex.Message}\n\nStack trace:\n{ex.StackTrace}");
                return Result.Failed;
            }
        }
    }
}
