using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace dwg2rvt.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenHubCommand : IExternalCommand
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
        
        public static void LinkPanelToHandler(UI.dwg2rvtPanel panel)
        {
            if (_placeSingleBlockTypeHandler != null)
            {
                _placeSingleBlockTypeHandler.SetPanel(panel);
            }
        }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                
                // Create external events if not already created
                if (_annotateEvent == null)
                {
                    UI.dwg2rvtPanel.AnnotateEventHandler annotateHandler = new UI.dwg2rvtPanel.AnnotateEventHandler();
                    _annotateEvent = ExternalEvent.Create(annotateHandler);
                }
                
                if (_placeElementsEvent == null)
                {
                    _placeElementsHandler = new UI.dwg2rvtPanel.PlaceElementsEventHandler();
                    _placeElementsEvent = ExternalEvent.Create(_placeElementsHandler);
                }
                
                if (_placeSingleBlockTypeEvent == null)
                {
                    _placeSingleBlockTypeHandler = new UI.dwg2rvtPanel.PlaceSingleBlockTypeEventHandler();
                    _placeSingleBlockTypeEvent = ExternalEvent.Create(_placeSingleBlockTypeHandler);
                }

                // Create and show the main hub panel
                UI.MainHubPanel hub = new UI.MainHubPanel(uiApp, _annotateEvent, _placeElementsEvent, _placeSingleBlockTypeEvent);
                
                // Set owner to Revit window
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(hub);
                helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;
                
                hub.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
