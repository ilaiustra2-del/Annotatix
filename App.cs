using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace dwg2rvt
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Log Revit startup with timestamp
                var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator();
                Core.DebugLogger.Log("[APP] *** REVIT PLUGIN STARTUP ***");
                Core.DebugLogger.Log("[APP] dwg2rvt v3.x loading...");
                
                // Log where this App class is loaded from
                var appAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                Core.DebugLogger.Log($"[APP] Main assembly: {appAssembly.GetName().Name}");
                Core.DebugLogger.Log($"[APP] Assembly location: {appAssembly.Location}");
                Core.DebugLogger.Log($"[APP] Assembly version: {appAssembly.GetName().Version}");
                Core.DebugLogger.Log("");
                
                // Check what types are already loaded in this assembly
                var types = appAssembly.GetTypes();
                var uiTypes = types.Where(t => t.Namespace == "dwg2rvt.UI").Select(t => t.Name).ToList();
                var coreTypes = types.Where(t => t.Namespace == "dwg2rvt.Core").Select(t => t.Name).ToList();
                
                Core.DebugLogger.Log($"[APP] UI classes in assembly: {string.Join(", ", uiTypes)}");
                Core.DebugLogger.Log($"[APP] Core classes in assembly: {string.Join(", ", coreTypes)}");
                Core.DebugLogger.Log("");
                
                Core.DebugLogger.Log("[APP] IMPORTANT: All classes (dwg2rvtPanel, HVACPanel, etc.) are ALREADY compiled into this DLL");
                Core.DebugLogger.Log("[APP] They are NOT loaded from external files yet - this is the CURRENT implementation");
                Core.DebugLogger.Log("[APP] Future: We will load .cs files dynamically using ModuleLoader");
                Core.DebugLogger.Log($"[APP] Log file: {Core.DebugLogger.GetLogFilePath()}");
                Core.DebugLogger.LogSeparator();
                Core.DebugLogger.Log("");
                
                // Create ribbon tab
                string tabName = "Plugin";
                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch (Exception)
                {
                    // Tab already exists, continue
                }

                // Create ribbon panels
                RibbonPanel panelManagement = application.CreateRibbonPanel(tabName, "Управление");
                RibbonPanel panelDwg2rvt = application.CreateRibbonPanel(tabName, "DWG2RVT");

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyPath);
                
                // Paths for icons
                string icon32 = Path.Combine(assemblyDir, "dwg2rvt32.png");
                string icon80 = Path.Combine(assemblyDir, "dwg2rvt80.png");
                string iconOriginal = Path.Combine(assemblyDir, "dwg2rvt.png");

                // Add button for plugins hub
                PushButtonData hubButtonData = new PushButtonData(
                    "PluginsHub",
                    "Plugins Hub",
                    assemblyPath,
                    "dwg2rvt.Commands.OpenHubCommand"
                );
                hubButtonData.ToolTip = "Manage all plugins";
                
                // Add button for control panel
                PushButtonData buttonData = new PushButtonData(
                    "dwg2rvtPanel",
                    "DWG2RVT",
                    assemblyPath,
                    "dwg2rvt.Commands.OpenPanelCommand"
                );
                buttonData.ToolTip = "Open DWG Analysis Control Panel";
                buttonData.LongDescription = "Opens the control panel to analyze imported DWG files and extract block information";

                PushButton pushButton = panelDwg2rvt.AddItem(buttonData) as PushButton;
                PushButton hubButton = panelManagement.AddItem(hubButtonData) as PushButton;

                // Set icons ONLY for DWG2RVT button
                string bestIconPath = File.Exists(icon32) ? icon32 : (File.Exists(icon80) ? icon80 : (File.Exists(iconOriginal) ? iconOriginal : null));
                
                if (bestIconPath != null)
                {
                    try
                    {
                        BitmapImage bitmap = new BitmapImage(new Uri(bestIconPath));
                        pushButton.LargeImage = bitmap;
                    }
                    catch { }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to initialize plugin: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private BitmapImage GetEmbeddedImage(string resourceName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                Stream stream = assembly.GetManifestResourceStream(resourceName);
                
                if (stream != null)
                {
                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.EndInit();
                    return image;
                }
            }
            catch { }
            
            return null;
        }
    }
}
