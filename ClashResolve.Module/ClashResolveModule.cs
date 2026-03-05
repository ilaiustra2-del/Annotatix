using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace ClashResolve.Module
{
    /// <summary>
    /// ClashResolve Module — resolves pipe/duct clashes by rerouting one element under the other.
    /// Implements IModule for dynamic loading by PluginsManager.
    /// </summary>
    public class ClashResolveModule : IModule
    {
        public string ModuleId => "clash_resolve";
        public string ModuleName => "ClashResolve";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[CLASHRESOLVE-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[CLASHRESOLVE-MODULE] ClashResolve module loaded with main plugin");
        }

        public Window CreatePanel(object[] parameters)
        {
            UIApplication uiApp = null;
            ExternalEvent clashEvent = null;
            object clashHandler = null;
            ExternalEvent pickEventA = null;
            object pickHandlerA = null;
            ExternalEvent pickEventB = null;
            object pickHandlerB = null;

            if (parameters != null)
            {
                if (parameters.Length > 0) uiApp       = parameters[0] as UIApplication;
                if (parameters.Length > 1) clashEvent  = parameters[1] as ExternalEvent;
                if (parameters.Length > 2) clashHandler = parameters[2];
                if (parameters.Length > 3) pickEventA  = parameters[3] as ExternalEvent;
                if (parameters.Length > 4) pickHandlerA = parameters[4];
                if (parameters.Length > 5) pickEventB  = parameters[5] as ExternalEvent;
                if (parameters.Length > 6) pickHandlerB = parameters[6];
            }

            DebugLogger.Log("[CLASHRESOLVE-MODULE] Creating ClashResolvePanel...");
            DebugLogger.Log($"[CLASHRESOLVE-MODULE] UIApp={uiApp != null}, ClashEvent={clashEvent != null}, PickA={pickEventA != null}, PickB={pickEventB != null}");

            if (clashEvent == null || pickEventA == null || pickEventB == null)
            {
                DebugLogger.Log("[CLASHRESOLVE-MODULE] ERROR: ExternalEvents were not provided — cannot create panel");
                throw new InvalidOperationException("ClashResolve: ExternalEvents must be pre-created and passed via parameters.");
            }

            // Cast proxy handlers to their interfaces (defined in PluginsManager.Core)
            var clashExecProxy = clashHandler as PluginsManager.Core.IClashExecuteProxy;
            var pickProxyA     = pickHandlerA as PluginsManager.Core.IPickElementProxy;
            var pickProxyB     = pickHandlerB as PluginsManager.Core.IPickElementProxy;

            if (clashExecProxy == null || pickProxyA == null || pickProxyB == null)
            {
                DebugLogger.Log("[CLASHRESOLVE-MODULE] ERROR: Handlers do not implement required interfaces");
                throw new InvalidOperationException("ClashResolve: Handlers must implement IClashExecuteProxy / IPickElementProxy.");
            }

            var panel = new UI.ClashResolvePanel(
                uiApp,
                clashEvent, clashExecProxy,
                pickEventA, pickProxyA,
                pickEventB, pickProxyB);

            var window = new Window
            {
                Content = panel,
                Title = "ClashResolve — Обход пересечений",
                Width = 540,
                Height = 460,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize,
                MinWidth = 400,
                MinHeight = 360
            };

            DebugLogger.Log("[CLASHRESOLVE-MODULE] Panel window created successfully");

            return window;
        }
    }
}
