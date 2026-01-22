using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using dwg2rvt.UI;

namespace dwg2rvt.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class OpenAuthPanelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Open authentication panel as modal dialog
                var authPanel = new AuthPanel();
                bool? result = authPanel.ShowDialog();
                
                if (result == true)
                {
                    // Authentication successful
                    var currentUser = Core.AuthService.CurrentUser;
                    if (currentUser != null && currentUser.IsSuccess)
                    {
                        TaskDialog.Show("Успех", 
                            $"Добро пожаловать!\n" +
                            $"Логин: {currentUser.Login}\n" +
                            $"Тариф: {currentUser.SubscriptionPlan}");
                    }
                }
                
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
