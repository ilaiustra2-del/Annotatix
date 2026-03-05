using System;
using Autodesk.Revit.DB;

namespace PluginsManager.Core
{
    /// <summary>
    /// Interface for the pick-element proxy handler exposed to ClashResolve.Module.
    /// Implemented by ClashResolvePickAHandler / ClashResolvePickBHandler in PluginsManager.Commands.
    /// </summary>
    public interface IPickElementProxy
    {
        string PromptMessage { get; set; }
        Action<ElementId, string> OnPicked { get; set; }
    }

    /// <summary>
    /// Interface for the execute proxy handler exposed to ClashResolve.Module.
    /// Implemented by ClashResolveExecuteHandler in PluginsManager.Commands.
    /// </summary>
    public interface IClashExecuteProxy
    {
        Action<Autodesk.Revit.UI.UIApplication> PendingAction { get; set; }
    }
}
