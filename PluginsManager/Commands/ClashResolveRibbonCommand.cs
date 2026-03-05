using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// IExternalCommand bound to the "Исправление коллизий" ribbon button.
    /// Activates ClashResolveSession which injects the Options Bar control
    /// and begins tracking sequential Ctrl-selection.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ClashResolveRibbonCommand : IExternalCommand
    {
        // Shared static state — created once per session
        private static ExternalEvent              _execEvent;
        private static ClashResolveExecuteHandler _execHandler;
        private static ClashResolveSelectionHandler _selectionHandler;

        // Debounce: ignore calls within 500ms of the last activation
        private static DateTime _lastActivationTime = DateTime.MinValue;
        private static readonly TimeSpan DebounceWindow = TimeSpan.FromMilliseconds(500);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                Core.DebugLogger.Log("[CLASH-RIBBON] ClashResolveRibbonCommand.Execute()");

                // Ensure ClashResolve module is loaded
                var module = Core.DynamicModuleLoader.GetModuleInstance("clash_resolve");
                if (module == null)
                {
                    message = "Модуль ClashResolve не загружен. Откройте Plugins Hub для авторизации.";
                    Core.DebugLogger.Log("[CLASH-RIBBON] Module not loaded");
                    TaskDialog.Show("ClashResolve", message);
                    return Result.Failed;
                }

                var now = DateTime.UtcNow;

                // Toggle: if already active, deactivate — but only if enough time has passed
                // (prevents the Revit double-fire from immediately deactivating what was just activated)
                if (ClashResolveSession.Current != null)
                {
                    if ((now - _lastActivationTime) < DebounceWindow)
                    {
                        Core.DebugLogger.Log("[CLASH-RIBBON] Debounce: ignoring rapid second call while session is active");
                        return Result.Succeeded;
                    }
                    Core.DebugLogger.Log("[CLASH-RIBBON] Session already active — deactivating");
                    UnregisterSelectionHandler();
                    ClashResolveSession.DeactivateCurrent();
                    return Result.Succeeded;
                }

                // Create ExternalEvent (once per session)
                if (_execEvent == null)
                {
                    _execHandler = new ClashResolveExecuteHandler();
                    _execEvent   = ExternalEvent.Create(_execHandler);
                    Core.DebugLogger.Log("[CLASH-RIBBON] ExternalEvent created");
                }

                // Activate session
                _lastActivationTime = now;
                ClashResolveSession.Activate(uiApp, _execEvent, _execHandler);

                // Register selection tracking
                if (_selectionHandler == null)
                    _selectionHandler = new ClashResolveSelectionHandler();
                _selectionHandler.Register(uiApp);

                Core.DebugLogger.Log("[CLASH-RIBBON] Session activated, selection tracking started");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Ошибка запуска ClashResolve: {ex.Message}";
                Core.DebugLogger.Log($"[CLASH-RIBBON] ERROR: {ex.Message}\n{ex.StackTrace}");
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
