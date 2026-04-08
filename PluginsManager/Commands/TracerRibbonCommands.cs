using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Base class for Tracer ribbon commands
    /// </summary>
    public abstract class TracerRibbonCommandBase : IExternalCommand
    {
        // Debounce: ignore calls within 500ms of the last activation
        private static DateTime _lastActivationTime = DateTime.MinValue;
        private static readonly TimeSpan DebounceWindow = TimeSpan.FromMilliseconds(500);
        private static TracerSelectionHandler _selectionHandler;

        protected abstract TracerSession.TracerConnectionType ConnectionType { get; }
        protected abstract string CommandName { get; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                DebugLogger.Log($"[TRACER-RIBBON] {CommandName}.Execute()");

                // Ensure Tracer module is loaded
                var module = DynamicModuleLoader.GetModuleInstance("tracer");
                if (module == null)
                {
                    message = "Модуль Tracer не загружен. Откройте Plugins Hub для авторизации.";
                    DebugLogger.Log("[TRACER-RIBBON] Module not loaded");
                    TaskDialog.Show("Tracer", message);
                    return Result.Failed;
                }

                var now = DateTime.UtcNow;

                // Toggle: if already active, deactivate — but only if enough time has passed
                if (TracerSession.Current != null)
                {
                    if ((now - _lastActivationTime) < DebounceWindow)
                    {
                        DebugLogger.Log("[TRACER-RIBBON] Debounce: ignoring rapid second call");
                        return Result.Succeeded;
                    }
                    DebugLogger.Log("[TRACER-RIBBON] Session already active — deactivating");
                    TracerSession.DeactivateCurrent();
                    UnregisterSelectionHandler();
                    return Result.Succeeded;
                }

                // Activate session
                _lastActivationTime = now;
                TracerSession.Activate(uiApp, ConnectionType);

                // Register selection tracking
                if (_selectionHandler == null)
                    _selectionHandler = new TracerSelectionHandler();
                _selectionHandler.Register(uiApp);

                DebugLogger.Log($"[TRACER-RIBBON] Session activated for {ConnectionType}, selection tracking started");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Ошибка запуска Tracer: {ex.Message}";
                DebugLogger.Log($"[TRACER-RIBBON] ERROR: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private static void UnregisterSelectionHandler()
        {
            try { _selectionHandler?.Unregister(); }
            catch { }
        }
    }

    /// <summary>
    /// 45-degree connection ribbon button
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Tracer45DegreeRibbonCommand : TracerRibbonCommandBase
    {
        protected override TracerSession.TracerConnectionType ConnectionType => 
            TracerSession.TracerConnectionType.Angle45;
        protected override string CommandName => "Tracer45Degree";
    }

    /// <summary>
    /// L-shaped connection ribbon button
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TracerLShapedRibbonCommand : TracerRibbonCommandBase
    {
        protected override TracerSession.TracerConnectionType ConnectionType => 
            TracerSession.TracerConnectionType.LShaped;
        protected override string CommandName => "TracerLShaped";
    }

    /// <summary>
    /// Bottom connection ribbon button
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TracerBottomRibbonCommand : TracerRibbonCommandBase
    {
        protected override TracerSession.TracerConnectionType ConnectionType => 
            TracerSession.TracerConnectionType.Bottom;
        protected override string CommandName => "TracerBottom";
    }

    /// <summary>
    /// Z-shaped connection ribbon button
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TracerZShapedRibbonCommand : TracerRibbonCommandBase
    {
        protected override TracerSession.TracerConnectionType ConnectionType => 
            TracerSession.TracerConnectionType.ZShaped;
        protected override string CommandName => "TracerZShaped";
    }
}
