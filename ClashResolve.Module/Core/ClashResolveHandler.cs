using System;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace ClashResolve.Module.Core
{
    /// <summary>
    /// ExternalEventHandler that executes clash resolution inside the Revit API context.
    /// </summary>
    public class ClashResolveHandler : IExternalEventHandler
    {
        /// <summary>The clash pair to resolve on next Execute call.</summary>
        public ClashPair PendingPair { get; set; }

        /// <summary>Result of the last Execute call.</summary>
        public ClashResolveResult LastResult { get; private set; }

        /// <summary>Callback invoked on the UI thread after Execute completes.</summary>
        public Action<ClashResolveResult> OnCompleted { get; set; }

        public void Execute(UIApplication app)
        {
            if (PendingPair == null)
            {
                DebugLogger.Log("[CLASH-HANDLER] Execute called with null PendingPair");
                return;
            }

            DebugLogger.Log($"[CLASH-HANDLER] Executing clash resolve for A={PendingPair.PipeAId}, B={PendingPair.PipeBId}");

            try
            {
                var doc = app.ActiveUIDocument.Document;
                var resolver = new ClashResolver();
                LastResult = resolver.ResolveClash(doc, PendingPair);

                DebugLogger.Log($"[CLASH-HANDLER] Result: Success={LastResult.Success}, Msg={LastResult.Message}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-HANDLER] Unhandled exception: {ex.Message}");
                LastResult = new ClashResolveResult
                {
                    Success = false,
                    Message = $"Внутренняя ошибка: {ex.Message}"
                };
            }
            finally
            {
                OnCompleted?.Invoke(LastResult);
            }
        }

        public string GetName() => "ClashResolveHandler";
    }

    /// <summary>
    /// ExternalEventHandler for element picking (selection must happen inside API context).
    /// </summary>
    public class PickElementHandler : IExternalEventHandler
    {
        public string PromptMessage { get; set; }
        public Autodesk.Revit.DB.ElementId PickedElementId { get; private set; }
        public Action<Autodesk.Revit.DB.ElementId, string> OnPicked { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = app.ActiveUIDocument;
                var reference = uidoc.Selection.PickObject(
                    Autodesk.Revit.UI.Selection.ObjectType.Element,
                    PromptMessage ?? "Выберите элемент");

                PickedElementId = reference.ElementId;
                var elem = uidoc.Document.GetElement(PickedElementId);
                string name = $"ID {PickedElementId.Value}: {elem?.Category?.Name ?? elem?.GetType().Name ?? "?"}";

                DebugLogger.Log($"[PICK-HANDLER] Picked element: {name}");
                OnPicked?.Invoke(PickedElementId, name);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                DebugLogger.Log("[PICK-HANDLER] Pick cancelled by user.");
                OnPicked?.Invoke(null, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PICK-HANDLER] Error: {ex.Message}");
                OnPicked?.Invoke(null, null);
            }
        }

        public string GetName() => "PickElementHandler";
    }
}
