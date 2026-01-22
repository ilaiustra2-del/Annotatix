using System;
using System.IO;
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
