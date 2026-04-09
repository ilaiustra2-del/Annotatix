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
        // Auto-authentication state tracking
        public static bool AutoAuthInProgress { get; private set; } = false;
        public static bool AutoAuthCompleted { get; private set; } = false;
        
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

                // ── Auto-authenticate on startup (background task) ─────────────────
                try
                {
                    if (Core.LocalAuthStorage.HasSavedAuth())
                    {
                        Core.DebugLogger.Log("[APP] Found saved auth data, starting background auto-authentication...");
                        // Start async auto-authentication without blocking startup
                        System.Threading.Tasks.Task.Run(async () => await PerformAutoAuthAsync());
                    }
                    else
                    {
                        Core.DebugLogger.Log("[APP] No saved auth data, skipping auto-authentication");
                    }
                }
                catch (Exception ex)
                {
                    Core.DebugLogger.Log($"[APP] ERROR starting auto-auth: {ex.Message}");
                }
                
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

                    // Path to Clash Resolve icon
                    string clashResolveIconPath = Path.Combine(assemblyDir, "UI", "icons", "Clash_Resolve.png");

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

                    // Second button: Batch / Multi clash resolve
                    PushButtonData multiClashBtnData = new PushButtonData(
                        "MultiClashResolveRibbon",
                        "Множественное\nисправление\nколлизий",
                        assemblyPath,
                        "PluginsManager.Commands.MultiClashRibbonCommand"
                    );
                    multiClashBtnData.ToolTip =
                        "Пакетное исправление коллизий труб.\n" +
                        "Выберите трубы A (обходящие), затем трубы B (препятствия), \n" +
                        "затем настройте параметры и нажмите Готово.";
                    PushButton multiClashBtn = panelClash.AddItem(multiClashBtnData) as PushButton;

                    // Set Clash Resolve icon for both buttons
                    if (File.Exists(clashResolveIconPath))
                    {
                        try
                        {
                            var clashIcon = new BitmapImage(new Uri(clashResolveIconPath));
                            clashBtn.LargeImage = clashIcon;
                            multiClashBtn.LargeImage = clashIcon;
                        }
                        catch { }
                    }

                    Core.DebugLogger.Log("[APP] Clash Resolve ribbon panel created");
                }
                catch (Exception ex)
                {
                    Core.DebugLogger.Log($"[APP] WARNING: Could not create Clash Resolve panel: {ex.Message}");
                }

                // ── Tracer ribbon panel ──────────────────────────────
                try
                {
                    RibbonPanel panelTracer = application.CreateRibbonPanel(tabName, "Tracer");

                    // 45-degree connection button
                    PushButtonData tracer45BtnData = new PushButtonData(
                        "Tracer45DegreeRibbon",
                        "Присоединение\nпод 45°",
                        assemblyPath,
                        "PluginsManager.Commands.Tracer45DegreeRibbonCommand"
                    );
                    tracer45BtnData.ToolTip =
                        "Присоединение стояка к магистрали под углом 45°.\n" +
                        "Выберите магистраль, затем стояки, затем настройте уклон.";

                    PushButton tracer45Btn = panelTracer.AddItem(tracer45BtnData) as PushButton;

                    // L-shaped connection button
                    PushButtonData tracerLBtnData = new PushButtonData(
                        "TracerLShapedRibbon",
                        "L-образное\nприсоединение",
                        assemblyPath,
                        "PluginsManager.Commands.TracerLShapedRibbonCommand"
                    );
                    tracerLBtnData.ToolTip =
                        "L-образное присоединение стояка к магистрали.\n" +
                        "Выберите магистраль, затем стояки, затем настройте уклон.";

                    PushButton tracerLBtn = panelTracer.AddItem(tracerLBtnData) as PushButton;

                    // Bottom connection button
                    PushButtonData tracerBottomBtnData = new PushButtonData(
                        "TracerBottomRibbon",
                        "Присоединение\nснизу",
                        assemblyPath,
                        "PluginsManager.Commands.TracerBottomRibbonCommand"
                    );
                    tracerBottomBtnData.ToolTip =
                        "Присоединение стояка к магистрали снизу.\n" +
                        "Выберите магистраль, затем стояки, затем настройте уклон.";

                    PushButton tracerBottomBtn = panelTracer.AddItem(tracerBottomBtnData) as PushButton;

                    // Z-shaped connection button
                    PushButtonData tracerZBtnData = new PushButtonData(
                        "TracerZShapedRibbon",
                        "Z-образное\nприсоединение",
                        assemblyPath,
                        "PluginsManager.Commands.TracerZShapedRibbonCommand"
                    );
                    tracerZBtnData.ToolTip =
                        "Z-образное присоединение стояка к магистрали.\n" +
                        "Выберите магистраль, затем стояки, затем настройте уклон.";

                    PushButton tracerZBtn = panelTracer.AddItem(tracerZBtnData) as PushButton;

                    // Set icons for Tracer buttons
                    string tracerIcon45 = Path.Combine(assemblyDir, "Tracer_45.png");
                    string tracerIconL = Path.Combine(assemblyDir, "Tracer_L.png");
                    string tracerIconBottom = Path.Combine(assemblyDir, "Tracer_Bottom.png");
                    string tracerIconZ = Path.Combine(assemblyDir, "Tracer_Z.png");

                    if (File.Exists(tracerIcon45))
                    {
                        try { tracer45Btn.LargeImage = new BitmapImage(new Uri(tracerIcon45)); } catch { }
                    }
                    if (File.Exists(tracerIconL))
                    {
                        try { tracerLBtn.LargeImage = new BitmapImage(new Uri(tracerIconL)); } catch { }
                    }
                    if (File.Exists(tracerIconBottom))
                    {
                        try { tracerBottomBtn.LargeImage = new BitmapImage(new Uri(tracerIconBottom)); } catch { }
                    }
                    if (File.Exists(tracerIconZ))
                    {
                        try { tracerZBtn.LargeImage = new BitmapImage(new Uri(tracerIconZ)); Core.DebugLogger.Log($"[APP] Tracer Z icon loaded from: {tracerIconZ}"); }
                        catch (Exception ex) { Core.DebugLogger.Log($"[APP] WARNING: Failed to load Tracer Z icon: {ex.Message}"); }
                    }
                    else
                    {
                        Core.DebugLogger.Log($"[APP] WARNING: Tracer Z icon not found at: {tracerIconZ}");
                    }

                    Core.DebugLogger.Log("[APP] Tracer ribbon panel created");
                }
                catch (Exception ex)
                {
                    Core.DebugLogger.Log($"[APP] WARNING: Could not create Tracer panel: {ex.Message}");
                }

                // ── HVAC ribbon panel ──────────────────────────────
                try
                {
                    RibbonPanel panelHVAC = application.CreateRibbonPanel(tabName, "HVAC");

                    // Построить схему button
                    PushButtonData hvacCreateBtnData = new PushButtonData(
                        "HvacCreateSchemaRibbon",
                        "Построить\nсхему",
                        assemblyPath,
                        "PluginsManager.Commands.HvacCreateSchemaRibbonCommand"
                    );
                    hvacCreateBtnData.ToolTip =
                        "Построение схемы систем HVAC.\n" +
                        "Создаёт чертёж схемы на основе модели.";

                    PushButton hvacCreateBtn = panelHVAC.AddItem(hvacCreateBtnData) as PushButton;

                    // Достроить схему button
                    PushButtonData hvacCompleteBtnData = new PushButtonData(
                        "HvacCompleteSchemaRibbon",
                        "Достроить\nсхему",
                        assemblyPath,
                        "PluginsManager.Commands.HvacCompleteSchemaRibbonCommand"
                    );
                    hvacCompleteBtnData.ToolTip =
                        "Достраивание схемы систем HVAC.\n" +
                        "Обновляет существующий чертёж схемы.";

                    PushButton hvacCompleteBtn = panelHVAC.AddItem(hvacCompleteBtnData) as PushButton;

                    // Sync toggle button - сохраняем ссылку для динамического изменения текста
                    PushButtonData hvacSyncBtnData = new PushButtonData(
                        "HvacSyncToggleRibbon",
                        "Синхронизация\nвключена",
                        assemblyPath,
                        "PluginsManager.Commands.HvacSyncToggleRibbonCommand"
                    );
                    hvacSyncBtnData.ToolTip =
                        "Включить/выключить синхронизацию чертежа с моделью.\n" +
                        "При включении изменения в модели автоматически отражаются на схеме.";

                    PushButton hvacSyncBtn = panelHVAC.AddItem(hvacSyncBtnData) as PushButton;
                    
                    // Сохраняем ссылку на кнопку синхронизации
                    Commands.HvacSyncToggleRibbonCommand.SyncButton = hvacSyncBtn;
                    
                    // Загружаем настройки и обновляем текст кнопки
                    try
                    {
                        string settingsPath = System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                            "Autodesk", "Revit", "Addins", "HVACSuperSchemeSettings.cfg");
                        bool isSyncEnabled = true; // По умолчанию включена
                        
                        if (System.IO.File.Exists(settingsPath))
                        {
                            string json = System.IO.File.ReadAllText(settingsPath);
                            dynamic settings = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                            isSyncEnabled = settings?.IsUpdaterSync ?? true;
                        }
                        
                        hvacSyncBtn.ItemText = isSyncEnabled 
                            ? "Синхронизация\nвключена" 
                            : "Синхронизация\nвыключена";
                        Core.DebugLogger.Log($"[APP-HVAC] Sync button text initialized: {isSyncEnabled}");
                    }
                    catch (Exception ex) 
                    { 
                        Core.DebugLogger.Log($"[APP-HVAC] Error loading settings: {ex.Message}");
                    }

                    // Try to set icons for HVAC buttons
                    string hvacIcon = Path.Combine(assemblyDir, "HVAC.png");
                    if (File.Exists(hvacIcon))
                    {
                        try
                        {
                            var hvacBitmap = new BitmapImage(new Uri(hvacIcon));
                            hvacCreateBtn.LargeImage = hvacBitmap;
                            hvacCompleteBtn.LargeImage = hvacBitmap;
                            hvacSyncBtn.LargeImage = hvacBitmap;
                        }
                        catch { }
                    }

                    Core.DebugLogger.Log("[APP] HVAC ribbon panel created");
                }
                catch (Exception ex)
                {
                    Core.DebugLogger.Log($"[APP] WARNING: Could not create HVAC panel: {ex.Message}");
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

        /// <summary>
        /// Performs auto-authentication in background when Revit starts.
        /// Loads user modules if authentication is successful.
        /// </summary>
        private static async System.Threading.Tasks.Task PerformAutoAuthAsync()
        {
            try
            {
                AutoAuthInProgress = true;
                Core.DebugLogger.Log("[APP-AUTOAUTH] Starting background auto-authentication...");
                
                var authService = new Core.AuthService();
                var result = await authService.TryAutoAuthenticateAsync();
                
                if (result.IsSuccess)
                {
                    Core.DebugLogger.Log("[APP-AUTOAUTH] Auto-authentication successful!");
                    Core.DebugLogger.Log($"[APP-AUTOAUTH] User: {result.Login}");
                    
                    // Set CurrentUser so that other parts of the app know we're authenticated
                    Core.AuthService.CurrentUser = result;
                    Core.DebugLogger.Log("[APP-AUTOAUTH] CurrentUser set in AuthService");
                    
                    var moduleTags = result.Modules?.Select(m => m.ModuleTag).ToList() ?? new System.Collections.Generic.List<string>();
                    Core.DebugLogger.Log($"[APP-AUTOAUTH] Available modules: {string.Join(", ", moduleTags)}");
                    
                    // Load modules in background
                    if (moduleTags.Count > 0)
                    {
                        Core.DebugLogger.Log("[APP-AUTOAUTH] Loading modules...");
                        int loadedCount = 0;
                        
                        // Get modules directory path
                        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        var assemblyPath = Path.GetDirectoryName(assembly.Location);
                        var modulesPath = Path.GetDirectoryName(assemblyPath);
                        Core.DebugLogger.Log($"[APP-AUTOAUTH] Modules path: {modulesPath}");
                        
                        foreach (var moduleTag in moduleTags)
                        {
                            try
                            {
                                // Map module tag to DLL name and folder
                                string moduleFolder = moduleTag.ToLower() switch
                                {
                                    "dwg2rvt" => "dwg2rvt",
                                    "hvac" => "hvac",
                                    "familysync" => "family_sync",
                                    "autonumbering" => "autonumbering",
                                    "clash_resolve" => "clash_resolve",
                                    "tracer" => "tracer",
                                    "full" => null, // Special case - loads all modules
                                    _ => moduleTag.ToLower()
                                };
                                
                                if (moduleFolder == null) continue; // Skip "full" tag
                                
                                string dllName = moduleTag.ToLower() switch
                                {
                                    "dwg2rvt" => "dwg2rvt.Module.dll",
                                    "hvac" => "HVAC.Module.dll",
                                    "familysync" => "FamilySync.Module.dll",
                                    "autonumbering" => "AutoNumbering.Module.dll",
                                    "clash_resolve" => "ClashResolve.Module.dll",
                                    "tracer" => "Tracer.Module.dll",
                                    _ => $"{moduleTag}.Module.dll"
                                };
                                
                                var moduleDllPath = Path.Combine(modulesPath, moduleFolder, dllName);
                                Core.DebugLogger.Log($"[APP-AUTOAUTH] Loading {moduleTag} from: {moduleDllPath}");
                                
                                if (Core.DynamicModuleLoader.LoadModule(moduleTag, moduleDllPath))
                                {
                                    loadedCount++;
                                    Core.DebugLogger.Log($"[APP-AUTOAUTH] ✓ Module loaded: {moduleTag}");
                                }
                                else
                                {
                                    Core.DebugLogger.Log($"[APP-AUTOAUTH] ✗ Failed to load module: {moduleTag}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Core.DebugLogger.Log($"[APP-AUTOAUTH] Error loading module {moduleTag}: {ex.Message}");
                            }
                        }
                        Core.DebugLogger.Log($"[APP-AUTOAUTH] Loaded {loadedCount}/{moduleTags.Count} modules");
                    }
                }
                else
                {
                    Core.DebugLogger.Log($"[APP-AUTOAUTH] Auto-authentication failed: {result.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[APP-AUTOAUTH] ERROR: {ex.Message}");
            }
            finally
            {
                AutoAuthInProgress = false;
                AutoAuthCompleted = true;
                Core.DebugLogger.Log("[APP-AUTOAUTH] Auto-authentication process completed");
            }
        }
    }
}
