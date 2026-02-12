using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HVACSuperScheme.Utils;
using System;
using System.Windows;

namespace HVACSuperScheme.Commands.Settings
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ViewModel vm = null;
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                vm = new ViewModel(doc);
                vm.View.ShowDialog();
            }
            catch (CustomException ex)
            {
                vm.View.Close();
                MessageBox.Show(ExceptionUtils.Error(ex));
            }
            catch (Exception ex)
            {
                vm.View.Close();
                MessageBox.Show(ExceptionUtils.SystemError(ex));
            }
            return Result.Succeeded;
        }
    }
}
