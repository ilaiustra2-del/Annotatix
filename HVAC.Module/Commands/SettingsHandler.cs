using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;
using HVACSuperScheme.Commands.Settings;
using HVACSuperScheme;

namespace HVAC.Module.Commands
{
    /// <summary>
    /// External Event Handler for Settings command
    /// </summary>
    public class SettingsHandler : IExternalEventHandler
    {
        private UIApplication _uiApp;

        public SettingsHandler(UIApplication uiApp)
        {
            _uiApp = uiApp;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Log("[HVAC-CMD] Executing Settings command...");
                
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                {
                    TaskDialog.Show("Ошибка", "Нет активного документа Revit");
                    return;
                }
                
                Document doc = uidoc.Document;
                
                // Show settings window
                var settingsWindow = new SettingsWindow(doc);
                settingsWindow.ShowDialog();
                
                DebugLogger.Log("[HVAC-CMD] Settings window executed");
            }
            catch (CustomException ex)
            {
                DebugLogger.Log($"[HVAC-CMD] CustomException in Settings: {ex.Message}");
                TaskDialog.Show("Ошибка", ex.Message);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-CMD] ERROR in Settings: {ex.Message}");
                DebugLogger.Log($"[HVAC-CMD] Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Ошибка", $"Ошибка при открытии настроек:\n{ex.Message}");
            }
        }

        public string GetName()
        {
            return "HVAC_SettingsHandler";
        }
    }
}
