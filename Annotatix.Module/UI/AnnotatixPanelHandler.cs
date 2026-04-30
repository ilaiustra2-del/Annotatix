using System;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace Annotatix.Module.UI
{
    /// <summary>
    /// External event handler for executing Annotatix commands from the Plugins Hub panel.
    /// Works by dynamically loading the command from the Annotatix.Module assembly.
    /// </summary>
    public class AnnotatixPanelHandler : IExternalEventHandler
    {
        /// <summary>
        /// Reference to the panel for UI updates.
        /// </summary>
        public AnnotatixPanel Panel { get; set; }

        /// <summary>
        /// Which command to execute: "RasterExport" or "AnnotateDuctwork"
        /// </summary>
        public string CommandName { get; set; }

        public void Execute(UIApplication uiApp)
        {
            try
            {
                Panel?.SetProgress("Выполнение команды...");

                // Build ExternalCommandData via reflection so we can pass it to IExternalCommand
                ExternalCommandData cmdData = BuildCommandData(uiApp);
                if (cmdData == null)
                {
                    Panel?.SetStatus("Ошибка: не удалось создать контекст команды Revit");
                    return;
                }

                string msg = "";
                ElementSet elements = null;

                // Determine which command class to invoke based on CommandName
                if (CommandName == "RasterExport")
                {
                    var cmd = new Commands.RasterExportCommand();
                    cmd.Execute(cmdData, ref msg, elements);
                }
                else if (CommandName == "AnnotateDuctwork")
                {
                    var cmd = new Commands.AnnotateDuctworkCommand();
                    cmd.Execute(cmdData, ref msg, elements);
                }
                else
                {
                    Panel?.SetStatus($"Неизвестная команда: {CommandName}");
                }

                Panel?.SetStatus("Команда выполнена.");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PANEL] Command error: {ex.Message}");
                Panel?.SetStatus($"Ошибка: {ex.Message}");
            }
        }

        public string GetName() => "Annotatix Panel Handler";

        /// <summary>
        /// Creates an ExternalCommandData instance with the given UIApplication set,
        /// using reflection to work around the read-only Application property.
        /// </summary>
        private static ExternalCommandData BuildCommandData(UIApplication uiApp)
        {
            try
            {
                var type = typeof(ExternalCommandData);
                ExternalCommandData cmdData;

                // Try public parameterless constructor first
                var ctor = type.GetConstructor(Type.EmptyTypes);
                if (ctor != null)
                {
                    cmdData = (ExternalCommandData)ctor.Invoke(null);
                }
                else
                {
                    // Fallback: create instance without constructor
                    cmdData = (ExternalCommandData)System.Runtime.Serialization.FormatterServices
                        .GetUninitializedObject(typeof(ExternalCommandData));
                }

                // Try property setter
                var prop = type.GetProperty("Application",
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty);
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(cmdData, uiApp);
                    return cmdData;
                }

                // Try backing field
                var field = type.GetField("m_application",
                    BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                {
                    field.SetValue(cmdData, uiApp);
                    return cmdData;
                }

                // Search for any matching field
                var fields = type.GetFields(BindingFlags.NonPublic | BindingFlags.Instance);
                foreach (var f in fields)
                {
                    if (f.FieldType == typeof(UIApplication))
                    {
                        f.SetValue(cmdData, uiApp);
                        return cmdData;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PANEL] Failed to build ExternalCommandData: {ex.Message}");
            }
            return null;
        }
    }
}
