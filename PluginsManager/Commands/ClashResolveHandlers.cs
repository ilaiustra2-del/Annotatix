using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Proxy ExternalEventHandler for picking element A in ClashResolve.
    /// Created in PluginsManager so ExternalEvent.Create() can be called from IExternalCommand context.
    /// Delegates to the real PickElementHandler in ClashResolve.Module via Action callbacks.
    /// </summary>
    public class ClashResolvePickAHandler : IExternalEventHandler, PluginsManager.Core.IPickElementProxy
    {
        public string PromptMessage { get; set; } = "Выберите трубу/воздуховод A (будет обходить)";
        public Action<ElementId, string> OnPicked { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                {
                    DebugLogger.Log("[CLASH-PICK-A] No active document");
                    OnPicked?.Invoke(null, null);
                    return;
                }

                var reference = uidoc.Selection.PickObject(
                    Autodesk.Revit.UI.Selection.ObjectType.Element,
                    PromptMessage);

                var id = reference.ElementId;
                var elem = uidoc.Document.GetElement(id);
                string name = $"ID {id.Value}: {elem?.Category?.Name ?? elem?.GetType().Name ?? "?"}";

                DebugLogger.Log($"[CLASH-PICK-A] Picked: {name}");
                OnPicked?.Invoke(id, name);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                DebugLogger.Log("[CLASH-PICK-A] Cancelled by user");
                OnPicked?.Invoke(null, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-PICK-A] Error: {ex.Message}");
                OnPicked?.Invoke(null, null);
            }
        }

        public string GetName() => "ClashResolvePickAHandler";
    }

    /// <summary>
    /// Proxy ExternalEventHandler for picking element B in ClashResolve.
    /// </summary>
    public class ClashResolvePickBHandler : IExternalEventHandler, PluginsManager.Core.IPickElementProxy
    {
        public string PromptMessage { get; set; } = "Выберите трубу/воздуховод B (препятствие, остаётся неподвижным)";
        public Action<ElementId, string> OnPicked { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                {
                    DebugLogger.Log("[CLASH-PICK-B] No active document");
                    OnPicked?.Invoke(null, null);
                    return;
                }

                var reference = uidoc.Selection.PickObject(
                    Autodesk.Revit.UI.Selection.ObjectType.Element,
                    PromptMessage);

                var id = reference.ElementId;
                var elem = uidoc.Document.GetElement(id);
                string name = $"ID {id.Value}: {elem?.Category?.Name ?? elem?.GetType().Name ?? "?"}";

                DebugLogger.Log($"[CLASH-PICK-B] Picked: {name}");
                OnPicked?.Invoke(id, name);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                DebugLogger.Log("[CLASH-PICK-B] Cancelled by user");
                OnPicked?.Invoke(null, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-PICK-B] Error: {ex.Message}");
                OnPicked?.Invoke(null, null);
            }
        }

        public string GetName() => "ClashResolvePickBHandler";
    }

    /// <summary>
    /// Proxy ExternalEventHandler for executing clash resolution.
    /// Holds pending work as an Action set by ClashResolve.Module.
    /// </summary>
    public class ClashResolveExecuteHandler : IExternalEventHandler, PluginsManager.Core.IClashExecuteProxy
    {
        /// <summary>Action to execute inside Revit API context. Set by ClashResolve.Module before raising event.</summary>
        public Action<UIApplication> PendingAction { get; set; }

        public void Execute(UIApplication app)
        {
            if (PendingAction == null)
            {
                DebugLogger.Log("[CLASH-EXEC] Execute called but no PendingAction set");
                return;
            }

            DebugLogger.Log("[CLASH-EXEC] Executing pending clash resolve action");
            try
            {
                PendingAction.Invoke(app);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-EXEC] Unhandled exception: {ex.Message}");
            }
            finally
            {
                PendingAction = null;
            }
        }

        public string GetName() => "ClashResolveExecuteHandler";
    }
}
