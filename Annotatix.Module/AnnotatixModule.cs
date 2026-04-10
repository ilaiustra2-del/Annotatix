using System;
using System.Windows;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace Annotatix.Module
{
    /// <summary>
    /// Annotatix Module - Data collection for ML-based annotation automation
    /// Records view state before and after annotation placement
    /// </summary>
    public class AnnotatixModule : IModule
    {
        public string ModuleId => "annotatix";
        public string ModuleName => "Annotatix";
        public string ModuleVersion => "1.0.0";

        public void Initialize()
        {
            DebugLogger.Log("[ANNOTATIX-MODULE] *** MODULE INITIALIZED ***");
            DebugLogger.Log("[ANNOTATIX-MODULE] Annotatix module loaded dynamically after authentication");
        }

        public Window CreatePanel(object[] parameters)
        {
            // Annotatix module doesn't need a panel - it uses ribbon buttons directly
            // This method is required by IModule but returns null
            DebugLogger.Log("[ANNOTATIX-MODULE] CreatePanel called - Annotatix uses ribbon buttons directly");
            return null;
        }
    }
}
