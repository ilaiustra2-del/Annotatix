using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace PluginsManager
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
                Core.DebugLogger.Log("[APP] *** REVIT PLUGINS MANAGER STARTUP ***");
                Core.DebugLogger.Log("[APP] PluginsManager v3.x loading...");
                
                // Log where this App class is loaded from
                var appAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                Core.DebugLogger.Log($"[APP] Main assembly: {appAssembly.GetName().Name}");
                Core.DebugLogger.Log($"[APP] Assembly location: {appAssembly.Location}");
                Core.DebugLogger.Log($"[APP] Assembly version: {appAssembly.GetName().Version}");
                Core.DebugLogger.Log("");
                
                // Check what types are already loaded in this assembly
                var types = appAssembly.GetTypes();
                var uiTypes = types.Where(t => t.Namespace == "PluginsManager.UI").Select(t => t.Name).ToList();
                var coreTypes = types.Where(t => t.Namespace == "PluginsManager.Core").Select(t => t.Name).ToList();
                
                Core.DebugLogger.Log($"[APP] UI classes in assembly: {string.Join(", ", uiTypes)}");
                Core.DebugLogger.Log($"[APP] Core classes in assembly: {string.Join(", ", coreTypes)}");
                Core.DebugLogger.Log("");
                
                Core.DebugLogger.Log("[APP] PluginsManager will dynamically load module DLLs based on user authentication");
                Core.DebugLogger.Log($"[APP] Log file: {Core.DebugLogger.GetLogFilePath()}");
                Core.DebugLogger.LogSeparator();
                Core.DebugLogger.Log("");
                
                // Create ribbon tab
                string tabName = "Annotatix";
                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch (Exception)
                {
                    // Tab already exists, continue
                }

                // Create ribbon panel
                RibbonPanel panelManagement = application.CreateRibbonPanel(tabName, "Управление");

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
                    "PluginsManager.Commands.OpenHubCommand"
                );
                hubButtonData.ToolTip = "Manage all plugins";
                
                PushButton hubButton = panelManagement.AddItem(hubButtonData) as PushButton;

                // Set icon for Hub button
                string bestIconPath = File.Exists(icon32) ? icon32 : (File.Exists(icon80) ? icon80 : (File.Exists(iconOriginal) ? iconOriginal : null));
                
                if (bestIconPath != null)
                {
                    try
                    {
                        BitmapImage bitmap = new BitmapImage(new Uri(bestIconPath));
                        hubButton.LargeImage = bitmap;
                    }
                    catch { }
                }

                // ── Clash Resolve ribbon panel ──────────────────────────────
                try
                {
                    RibbonPanel panelClash = application.CreateRibbonPanel(tabName, "Clash Resolve");

                    PushButtonData clashBtnData = new PushButtonData(
                        "ClashResolveRibbon",
                        "Исправление\nколлизий",
                        assemblyPath,
                        "PluginsManager.Commands.ClashResolveRibbonCommand"
                    );
                    clashBtnData.ToolTip =
                        "Исправление коллизий труб.\n" +
                        "Нажмите, затем последовательно выделяйте трубы с зажатой Ctrl.\n" +
                        "Горячие клавиши: CR (EN) / СК (RU) — настраиваются в \"Управление → Сочетания клавиш\".";

                    PushButton clashBtn = panelClash.AddItem(clashBtnData) as PushButton;

                    // Try to use the same icon as the hub button
                    if (bestIconPath != null)
                    {
                        try { clashBtn.LargeImage = new BitmapImage(new Uri(bestIconPath)); } catch { }
                    }

                    Core.DebugLogger.Log("[APP] Clash Resolve ribbon panel created");
                }
                catch (Exception ex)
                {
                    Core.DebugLogger.Log($"[APP] WARNING: Could not create Clash Resolve panel: {ex.Message}");
                }

                // ── Inject keyboard shortcuts (CR / СК) if absent ──────────
                try
                {
                    InjectKeyboardShortcuts();
                }
                catch (Exception ex)
                {
                    Core.DebugLogger.Log($"[APP] WARNING: Could not inject keyboard shortcuts: {ex.Message}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to initialize Plugins Manager: {ex.Message}");
                return Result.Failed;
            }
        }

        // ----------------------------------------------------------------
        // Keyboard shortcut injection
        // ----------------------------------------------------------------
        private static void InjectKeyboardShortcuts()
        {
            // Revit stores user shortcuts in:
            // %APPDATA%\Autodesk\Revit\Autodesk Revit XXXX\KeyboardShortcuts.xml
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string[] revitDirs = Directory.GetDirectories(
                Path.Combine(appData, "Autodesk", "Revit"), "Autodesk Revit *");

            foreach (string revitDir in revitDirs)
            {
                string shortcutFile = Path.Combine(revitDir, "KeyboardShortcuts.xml");
                TryAddShortcuts(shortcutFile);
            }
        }

        private static void TryAddShortcuts(string filePath)
        {
            const string commandId = "CustomCtrl_%CustomCtrl_%Annotatix%Clash Resolve%ClashResolveRibbon";
            const string shortcutEN = "CR";
            const string shortcutRU = "\u0421\u041a"; // СК

            XDocument doc;
            XElement root;

            if (File.Exists(filePath))
            {
                try { doc = XDocument.Load(filePath); }
                catch { return; } // malformed XML — skip
                root = doc.Root;
            }
            else
            {
                // Create minimal file
                doc = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement("Shortcuts"));
                root = doc.Root;
            }

            if (root == null) return;

            // Check if already present
            bool hasEN = root.Descendants("ShortcutItem")
                .Any(e => (string)e.Attribute("CommandId") == commandId
                       && (string)e.Attribute("Shortcuts") != null
                       && ((string)e.Attribute("Shortcuts")).Contains(shortcutEN));

            if (hasEN)
            {
                Core.DebugLogger.Log($"[APP] Keyboard shortcuts already present in {filePath}");
                return;
            }

            // Add entry
            var shortcutsAttr = $"{shortcutEN}#{shortcutRU}";
            root.Add(new XElement("ShortcutItem",
                new XAttribute("CommandId", commandId),
                new XAttribute("Shortcuts", shortcutsAttr)));

            try
            {
                doc.Save(filePath);
                Core.DebugLogger.Log($"[APP] Keyboard shortcuts (CR/СК) added to {filePath}");
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[APP] Could not save shortcuts file: {ex.Message}");
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
