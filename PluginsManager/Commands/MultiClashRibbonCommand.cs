using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// IExternalCommand bound to the "Множественное исправление коллизий" ribbon button.
    /// Activates MultiClashSession which injects the multi-step Options Bar control.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MultiClashRibbonCommand : IExternalCommand
    {
        private static ExternalEvent               _execEvent;
        private static ClashResolveExecuteHandler  _execHandler;
        private static ClashResolveSelectionHandler _selectionHandler;

        private static DateTime _lastActivationTime = DateTime.MinValue;
        private static readonly TimeSpan DebounceWindow = TimeSpan.FromMilliseconds(500);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                Core.DebugLogger.Log("[MULTI-CLASH-RIBBON] MultiClashRibbonCommand.Execute()");

                // Ensure ClashResolve module is loaded
                var module = Core.DynamicModuleLoader.GetModuleInstance("clash_resolve");
                if (module == null)
                {
                    message = "Модуль ClashResolve не загружен. Откройте Plugins Hub для авторизации.";
                    Core.DebugLogger.Log("[MULTI-CLASH-RIBBON] Module not loaded");
                    TaskDialog.Show("MultiClashResolve", message);
                    return Result.Failed;
                }

                var now = DateTime.UtcNow;

                // Toggle: if already active, deactivate
                if (MultiClashSession.Current != null)
                {
                    if ((now - _lastActivationTime) < DebounceWindow)
                    {
                        Core.DebugLogger.Log("[MULTI-CLASH-RIBBON] Debounce: ignoring rapid second call");
                        return Result.Succeeded;
                    }
                    Core.DebugLogger.Log("[MULTI-CLASH-RIBBON] Session already active — deactivating");
                    UnregisterSelectionHandler();
                    MultiClashSession.DeactivateCurrent();
                    return Result.Succeeded;
                }

                // Create ExternalEvent (once per session)
                if (_execEvent == null)
                {
                    _execHandler = new ClashResolveExecuteHandler();
                    _execEvent   = ExternalEvent.Create(_execHandler);
                    Core.DebugLogger.Log("[MULTI-CLASH-RIBBON] ExternalEvent created");
                }

                // Activate session
                _lastActivationTime = now;
                MultiClashSession.Activate(uiApp, _execEvent, _execHandler);

                // Register selection tracking
                if (_selectionHandler == null)
                    _selectionHandler = new ClashResolveSelectionHandler();
                _selectionHandler.Register(uiApp);

                Core.DebugLogger.Log("[MULTI-CLASH-RIBBON] Session activated, selection tracking started");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Ошибка запуска MultiClashResolve: {ex.Message}";
                Core.DebugLogger.Log($"[MULTI-CLASH-RIBBON] ERROR: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private static void UnregisterSelectionHandler()
        {
            try { _selectionHandler?.Unregister(); }
            catch { }
        }
    }
}
