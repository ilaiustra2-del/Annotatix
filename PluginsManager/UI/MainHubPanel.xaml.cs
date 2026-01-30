using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace PluginsManager.UI
{
    public partial class MainHubPanel : Window
    {
        private UIApplication _uiApp;
        private object _hubContent;
        private ExternalEvent _openDwg2rvtPanelEvent;
        private Commands.OpenDwg2rvtPanelHandler _openDwg2rvtPanelHandler;

        public MainHubPanel(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _hubContent = MainContent.Content;
            
            // Create ExternalEvent for opening dwg2rvt panel
            // This is created in the constructor which is called from IExternalCommand.Execute()
            _openDwg2rvtPanelHandler = new Commands.OpenDwg2rvtPanelHandler();
            _openDwg2rvtPanelHandler.SetUIApplication(uiApp);
            _openDwg2rvtPanelEvent = ExternalEvent.Create(_openDwg2rvtPanelHandler);
            
            // Register this panel with the command so it can update UI
            Commands.OpenDwg2rvtPanelCommand.SetHubPanel(this);
            
            // Load icon from assembly location
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = Path.GetDirectoryName(assembly.Location);
                var iconPath = Path.Combine(assemblyPath, "UI", "icons", "dwg2rvt80.png");
                
                System.Diagnostics.Debug.WriteLine($"[HUB] Looking for icon at: {iconPath}");
                
                if (File.Exists(iconPath))
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(iconPath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    imgDwg2rvtIcon.Source = bitmap;
                    System.Diagnostics.Debug.WriteLine("[HUB] Icon loaded successfully");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB] Icon file not found: {iconPath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading icon: {ex.Message}");
            }
            
            // Set version number from BuildNumber.txt or assembly
            try
            {
                // Try to read from BuildNumber.txt (more reliable)
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = Path.GetDirectoryName(assembly.Location);
                var buildNumberFile = Path.Combine(assemblyPath, "BuildNumber.txt");
                
                System.Diagnostics.Debug.WriteLine($"[HUB] Looking for BuildNumber.txt at: {buildNumberFile}");
                
                if (File.Exists(buildNumberFile))
                {
                    var buildNumber = File.ReadAllText(buildNumberFile).Trim();
                    txtVersion.Text = $"v 3.{buildNumber}";
                    System.Diagnostics.Debug.WriteLine($"[HUB] Version set from BuildNumber.txt: v 3.{buildNumber}");
                }
                else
                {
                    // Fallback to assembly version
                    var version = assembly.GetName().Version;
                    // Version format is Major.Minor.Build.Revision (e.g. 3.001.0.0)
                    txtVersion.Text = $"v {version.Major}.{version.Minor}";
                    System.Diagnostics.Debug.WriteLine($"[HUB] Version set from assembly: v {version.Major}.{version.Minor}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error getting version: {ex.Message}");
                txtVersion.Text = "v 3.0"; // Fallback
            }
            
            // Try auto-authentication if saved credentials exist
            this.Loaded += async (s, e) => await TryAutoAuthenticateOnStartup();
        }
        
        /// <summary>
        /// Try to authenticate automatically using saved credentials
        /// </summary>
        private async System.Threading.Tasks.Task TryAutoAuthenticateOnStartup()
        {
            try
            {
                // Check if there are saved credentials
                if (!Core.LocalAuthStorage.HasSavedAuth())
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] No saved auth data, skipping auto-login");
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine("[HUB] Found saved auth data, attempting auto-login...");
                
                // Show loading message
                btnAuth.Content = "Авторизация...";
                
                var authService = new Core.AuthService();
                var result = await authService.TryAutoAuthenticateAsync();
                
                if (result.IsSuccess)
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB] Auto-login successful for: {result.Login}");
                    
                    // Store current user
                    Core.AuthService.CurrentUser = result;
                    
                    // Download missing module files
                    await DownloadMissingModules(result.ModuleFiles);
                    
                    // Update UI
                    btnAuth.Content = "Выйти";
                    
                    // Show modules info
                    ShowModulesInfo(result.Modules);
                    
                    // Show refresh button
                    btnRefreshUserData.Visibility = Visibility.Visible;
                    
                    // Activate module buttons
                    ActivateModuleButtons(result.Modules);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB] Auto-login failed: {result.ErrorMessage}");
                    btnAuth.Content = "Войти в учётную запись";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Auto-login exception: {ex.Message}");
                btnAuth.Content = "Войти в учётную запись";
            }
        }

        private void Dwg2rvt_Click(object sender, MouseButtonEventArgs e)
        {
            // Check if button is enabled
            if (!pnlDwg2rvt.IsEnabled)
            {
                System.Diagnostics.Debug.WriteLine("[HUB] DWG2RVT button is disabled");
                MessageBox.Show("Модуль DWG2RVT недоступен. Пожалуйста, войдите в учётную запись.", 
                    "Доступ запрещён", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            try
            {
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("[HUB] User clicked DWG2RVT button");
                Core.DebugLogger.Log("[HUB] Requesting panel open via ExternalEvent...");
                
                // Trigger the ExternalEvent to open panel in proper API context
                _openDwg2rvtPanelEvent.Raise();
                
                Core.DebugLogger.Log("[HUB] ExternalEvent raised");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке панели dwg2rvt:\n{ex.Message}\n\n{ex.StackTrace}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void HVAC_Click(object sender, MouseButtonEventArgs e)
        {
            // Check if button is enabled
            if (!pnlHVAC.IsEnabled)
            {
                System.Diagnostics.Debug.WriteLine("[HUB] HVAC button is disabled");
                MessageBox.Show("Модуль HVAC недоступен. Пожалуйста, войдите в учётную запись.", 
                    "Доступ запрещён", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            try
            {
                // Log panel creation with timestamp
                var clickTimestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("[HUB] User clicked HVAC button");
                Core.DebugLogger.Log("[HUB] Loading HVAC module panel...");
                
                // Get module instance
                var module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                if (module == null)
                {
                    MessageBox.Show("Модуль HVAC не загружен. Попробуйте войти заново.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log($"[HUB] Module found: {module.ModuleName} v{module.ModuleVersion}");
                
                // Create panel from module
                var panel = module.CreatePanel(new object[] {});
                if (panel == null)
                {
                    MessageBox.Show("Не удалось создать панель модуля HVAC.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log("[HUB] Panel created successfully");
                
                // Extract content from Window (modules return Window, but we need UIElement)
                var panelContent = panel.Content;
                if (panelContent == null)
                {
                    MessageBox.Show("Панель модуля не содержит контента.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log($"[HUB] Panel content type: {panelContent.GetType().Name}");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("");
                
                // Switch to HVAC panel
                txtTitle.Text = "HVAC";
                MainContent.Content = panelContent;
                btnBack.Visibility = Visibility.Visible;
                
                // Hide Right Sidebar when HVAC panel is open
                RightSidebar.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке панели HVAC:\n{ex.Message}\n\n{ex.StackTrace}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            // Back to Hub
            txtTitle.Text = "Панель управления плагинами";
            MainContent.Content = _hubContent;
            btnBack.Visibility = Visibility.Collapsed;
            
            // Show Right Sidebar when back to hub
            RightSidebar.Visibility = Visibility.Visible;
        }
        
        /// <summary>
        /// Open debug log file in Notepad
        /// </summary>
        private void BtnOpenLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Core.DebugLogger.OpenLogFile();
                Core.DebugLogger.Log("[HUB] User opened debug log file");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть лог-файл: {ex.Message}\n\nПуть: {Core.DebugLogger.GetLogFilePath()}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        
        /// <summary>
        /// Show loaded modules information
        /// </summary>
        private void BtnModulesInfo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var loadedModules = Core.DynamicModuleLoader.GetLoadedModulesInfo();
                var message = "=== ИНФОРМАЦИЯ О ЗАГРУЖЕННЫХ МОДУЛЯХ ===\n\n" + loadedModules;
                
                Core.DebugLogger.Log("[HUB] User requested modules info");
                Core.DebugLogger.Log(message);
                
                MessageBox.Show(message, "Загруженные модули", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении информации о модулях: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Public method called by OpenDwg2rvtPanelCommand to show the panel
        /// This is called from IExternalCommand context, so it's safe
        /// </summary>
        public void ShowDwg2rvtPanel(object panelContent)
        {
            try
            {
                Core.DebugLogger.Log("[HUB] ShowDwg2rvtPanel called");
                Core.DebugLogger.Log($"[HUB] Panel content type: {panelContent.GetType().Name}");
                
                // Switch to dwg2rvt panel
                txtTitle.Text = "DWG2RVT";
                MainContent.Content = panelContent;
                btnBack.Visibility = Visibility.Visible;
                
                // Hide Right Sidebar when dwg2rvt panel is open
                RightSidebar.Visibility = Visibility.Collapsed;
                
                Core.DebugLogger.Log("[HUB] Panel displayed successfully");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("");
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[HUB] ERROR in ShowDwg2rvtPanel: {ex.Message}");
                MessageBox.Show($"Ошибка при отображении панели:\n{ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private async void BtnAuth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Check if user is already logged in (button shows "Выйти")
                if (Core.AuthService.CurrentUser != null && Core.AuthService.CurrentUser.IsSuccess)
                {
                    // Logout
                    var result = MessageBox.Show("Вы уверены, что хотите выйти?", 
                        "Выход", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        Core.AuthService.Logout();
                        
                        // Reset UI
                        btnAuth.Content = "Войти в учётную запись";
                        txtModulesInfo.Text = "";
                        txtModulesInfo.Visibility = Visibility.Collapsed;
                        btnRefreshUserData.Visibility = Visibility.Collapsed;
                        
                        // Disable module buttons
                        pnlDwg2rvt.IsEnabled = false;
                        pnlDwg2rvt.Opacity = 0.5;
                        pnlHVAC.IsEnabled = false;
                        pnlHVAC.Opacity = 0.5;
                        
                        System.Diagnostics.Debug.WriteLine("[AUTH] User logged out");
                    }
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine("[AUTH] Opening authentication panel...");
                
                // Open authentication panel
                var authPanel = new AuthPanel();
                bool? result2 = authPanel.ShowDialog();
                
                System.Diagnostics.Debug.WriteLine($"[AUTH] Dialog result: {result2}");
                
                if (result2 == true)
                {
                    // Log authentication success with timestamp
                    var authTimestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                    Core.DebugLogger.Log("");
                    Core.DebugLogger.LogSeparator();
                    Core.DebugLogger.Log("[AUTH] *** AUTHENTICATION SUCCESSFUL ***");
                    Core.DebugLogger.Log("[AUTH] Now loading modules dynamically...");
                    Core.DebugLogger.LogSeparator();
                    Core.DebugLogger.Log("");
                    
                    // Authentication successful
                    var currentUser = Core.AuthService.CurrentUser;
                    if (currentUser != null && currentUser.IsSuccess)
                    {
                        System.Diagnostics.Debug.WriteLine($"[AUTH] SUCCESS - User: {currentUser.Login}");
                        
                        // Download missing module files
                        await DownloadMissingModules(currentUser.ModuleFiles);
                        
                        // Update UI
                        btnAuth.Content = "Выйти";
                        
                        // Show modules info
                        ShowModulesInfo(currentUser.Modules);
                        
                        // Show refresh button
                        btnRefreshUserData.Visibility = Visibility.Visible;
                        
                        // Activate module buttons based on user's modules
                        ActivateModuleButtons(currentUser.Modules);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[AUTH] ERROR - Authentication result is null or failed");
                        MessageBox.Show("Ошибка: нет данных о пользователе", 
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH] Authentication cancelled or failed");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AUTH] EXCEPTION - Client side error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[AUTH] Stack trace: {ex.StackTrace}");
                
                MessageBox.Show($"Ошибка на стороне плагина:\n{ex.Message}\n\nТип: {ex.GetType().Name}", 
                    "Ошибка плагина", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Show modules information in the UI
        /// </summary>
        private void ShowModulesInfo(List<Core.UserModule> modules)
        {
            if (modules == null || modules.Count == 0)
            {
                txtModulesInfo.Text = "Модули: нет активных";
                txtModulesInfo.Visibility = Visibility.Visible;
                return;
            }
            
            var activeModules = modules.Where(m => m.IsActive).ToList();
            if (activeModules.Count == 0)
            {
                txtModulesInfo.Text = "Модули: нет активных";
                txtModulesInfo.Visibility = Visibility.Visible;
                return;
            }
            
            var moduleDetails = activeModules.Select(m => 
                $"{m.ModuleTag} (до {m.EndDate:dd.MM.yyyy})").ToList();
            txtModulesInfo.Text = $"Модули: {string.Join(", ", moduleDetails)}";
            txtModulesInfo.Visibility = Visibility.Visible;
        }
        
        /// <summary>
        /// Download missing module files from server
        /// </summary>
        private async System.Threading.Tasks.Task DownloadMissingModules(List<Core.ModuleFileInfo> moduleFiles)
        {
            if (moduleFiles == null || moduleFiles.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[HUB] No module files to download");
                return;
            }
            
            try
            {
                var downloader = new Core.ModuleDownloader();
                
                // Check which modules need to be downloaded
                var toDownload = downloader.GetModulesToDownload(moduleFiles);
                
                if (toDownload.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] All module files already present");
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine($"[HUB] Need to download {toDownload.Count} modules");
                
                // Show progress
                var progress = new Progress<string>(message =>
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB-DOWNLOAD] {message}");
                });
                
                // Download modules
                var (success, total) = await downloader.DownloadAllModules(toDownload, progress);
                
                if (success > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB] Downloaded {success}/{total} modules successfully");
                    
                    // Show success message
                    if (success == total)
                    {
                        MessageBox.Show($"Загружено модулей: {success}",
                            "Загрузка завершена", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Build list of failed modules
                        var failedModules = toDownload.Where(m => 
                            !new Core.ModuleDownloader().IsModuleInstalled(m.ModuleTag)).ToList();
                        var failedNames = string.Join(", ", failedModules.Select(m => m.ModuleTag));
                        
                        MessageBox.Show($"Загружено: {success}/{total}\n\nНе удалось загрузить:\n{failedNames}\n\nВозможные причины:\n- Файлы модулей отсутствуют на сервере\n- Проблемы с интернет-соединением",
                            "Частичная загрузка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else if (total > 0)
                {
                    MessageBox.Show("Не удалось загрузить файлы модулей.\n\nВозможные причины:\n- Отсутствует интернет-соединение\n- Файлы не загружены на сервер",
                        "Ошибка загрузки", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Download error: {ex.Message}");
                MessageBox.Show($"Ошибка при загрузке модулей:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Refresh user data and reload modules
        /// </summary>
        private async void BtnRefreshUserData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Check if user is logged in
                if (Core.AuthService.CurrentUser == null || !Core.AuthService.CurrentUser.IsSuccess)
                {
                    MessageBox.Show("Пожалуйста, войдите в учётную запись.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine("[HUB] User requested data refresh...");
                
                // Disable button during refresh
                btnRefreshUserData.IsEnabled = false;
                btnRefreshUserData.Content = "⏳ Обновление...";
                
                // Get current user email
                var currentEmail = Core.AuthService.CurrentUser.Login;
                
                // Request fresh data from server
                var authService = new Core.AuthService();
                var result = await authService.TryAutoAuthenticateAsync();
                
                if (result.IsSuccess)
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] User data refreshed successfully");
                    
                    // Update current user
                    Core.AuthService.CurrentUser = result;
                    
                    // Download missing module files
                    await DownloadMissingModules(result.ModuleFiles);
                    
                    // Update modules info display
                    ShowModulesInfo(result.Modules);
                    
                    // First, reset all module buttons to disabled state
                    pnlDwg2rvt.IsEnabled = false;
                    pnlDwg2rvt.Opacity = 0.5;
                    pnlHVAC.IsEnabled = false;
                    pnlHVAC.Opacity = 0.5;
                    
                    // Now re-activate based on fresh data and file availability
                    ActivateModuleButtons(result.Modules);
                    
                    MessageBox.Show("Данные пользователя обновлены!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB] Refresh failed: {result.ErrorMessage}");
                    MessageBox.Show($"Не удалось обновить данные:\n{result.ErrorMessage}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Refresh exception: {ex.Message}");
                MessageBox.Show($"Ошибка при обновлении данных:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Re-enable button
                btnRefreshUserData.IsEnabled = true;
                btnRefreshUserData.Content = "🔄 Обновить данные пользователя";
            }
        }
        
        /// <summary>
        /// Load and activate modules based on user's active modules
        /// </summary>
        private void ActivateModuleButtons(List<Core.UserModule> modules)
        {
            if (modules == null || modules.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[HUB] No modules to activate");
                return;
            }
            
            System.Diagnostics.Debug.WriteLine($"[HUB] Loading and activating modules: {modules.Count}");
            
            // Get modules directory path
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var assemblyPath = Path.GetDirectoryName(assembly.Location);
            
            // Go up one level from main/ to annotatix_dependencies/
            var modulesPath = Path.GetDirectoryName(assemblyPath);
            
            System.Diagnostics.Debug.WriteLine($"[HUB] Assembly path: {assemblyPath}");
            System.Diagnostics.Debug.WriteLine($"[HUB] Modules base path: {modulesPath}");
            
            // Check for missing module files
            var missingModules = new List<string>();
            
            foreach (var module in modules)
            {
                if (!module.IsActive)
                    continue;
                
                string moduleTag = module.ModuleTag.ToLower();
                System.Diagnostics.Debug.WriteLine($"[HUB] Processing module: {moduleTag}");
                
                switch (moduleTag)
                {
                    case "dwg2rvt":
                        if (!CheckModuleFilesExist(modulesPath, "dwg2rvt", "dwg2rvt.Module.dll"))
                            missingModules.Add("dwg2rvt");
                        else
                            LoadAndActivateDwg2rvtModule(modulesPath);
                        break;
                        
                    case "hvac":
                        if (!CheckModuleFilesExist(modulesPath, "hvac", "HVAC.Module.dll"))
                            missingModules.Add("hvac");
                        else
                            LoadAndActivateHVACModule(modulesPath);
                        break;
                        
                    case "full":
                        // Check both modules
                        bool dwg2rvtExists = CheckModuleFilesExist(modulesPath, "dwg2rvt", "dwg2rvt.Module.dll");
                        bool hvacExists = CheckModuleFilesExist(modulesPath, "hvac", "HVAC.Module.dll");
                        
                        if (!dwg2rvtExists)
                            missingModules.Add("dwg2rvt");
                        else
                            LoadAndActivateDwg2rvtModule(modulesPath);
                            
                        if (!hvacExists)
                            missingModules.Add("hvac");
                        else
                            LoadAndActivateHVACModule(modulesPath);
                            
                        if (dwg2rvtExists && hvacExists)
                            System.Diagnostics.Debug.WriteLine("[HUB] All modules activated (full access)");
                        break;
                        
                    default:
                        System.Diagnostics.Debug.WriteLine($"[HUB] Unknown module: {moduleTag}");
                        break;
                }
            }
            
            // Show warning if any modules are missing
            if (missingModules.Count > 0)
            {
                var message = "У вас отсутствуют файлы одного или нескольких модулей:\n";
                foreach (var moduleName in missingModules)
                {
                    message += $"- {moduleName}\n";
                }
                message += "\nОбратитесь в поддержку для получения необходимых файлов.";
                
                MessageBox.Show(message, "Отсутствуют файлы модулей", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    
                Core.DebugLogger.Log($"[HUB] WARNING: Missing module files: {string.Join(", ", missingModules)}");
            }
            
            // Log loaded modules info
            Core.DebugLogger.Log("");
            Core.DebugLogger.Log("[HUB] === LOADED MODULES INFO ===");
            Core.DebugLogger.Log(Core.DynamicModuleLoader.GetLoadedModulesInfo());
            Core.DebugLogger.Log("[HUB] ============================");
            Core.DebugLogger.Log("");
        }
        
        /// <summary>
        /// Check if module files exist
        /// </summary>
        private bool CheckModuleFilesExist(string modulesPath, string moduleFolderName, string moduleDllName)
        {
            var moduleFolder = Path.Combine(modulesPath, moduleFolderName);
            var moduleDllPath = Path.Combine(moduleFolder, moduleDllName);
            
            if (!Directory.Exists(moduleFolder))
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Module folder does not exist: {moduleFolder}");
                return false;
            }
            
            if (!File.Exists(moduleDllPath))
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Module DLL does not exist: {moduleDllPath}");
                return false;
            }
            
            return true;
        }
        
        /// <summary>
        /// Load and activate DWG2RVT module
        /// </summary>
        private void LoadAndActivateDwg2rvtModule(string modulesPath)
        {
            try
            {
                var moduleDllPath = Path.Combine(modulesPath, "dwg2rvt", "dwg2rvt.Module.dll");
                System.Diagnostics.Debug.WriteLine($"[HUB] Loading dwg2rvt module from: {moduleDllPath}");
                
                if (Core.DynamicModuleLoader.LoadModule("dwg2rvt", moduleDllPath))
                {
                    pnlDwg2rvt.IsEnabled = true;
                    pnlDwg2rvt.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("[HUB] ✓ DWG2RVT module loaded and activated");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] ✗ Failed to load DWG2RVT module");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading DWG2RVT module: {ex.Message}");
                Core.DebugLogger.Log($"[HUB] ERROR: Failed to load dwg2rvt module - {ex.Message}");
            }
        }
        
        /// <summary>
        /// Load and activate HVAC module
        /// </summary>
        private void LoadAndActivateHVACModule(string modulesPath)
        {
            try
            {
                var moduleDllPath = Path.Combine(modulesPath, "hvac", "HVAC.Module.dll");
                System.Diagnostics.Debug.WriteLine($"[HUB] Loading HVAC module from: {moduleDllPath}");
                
                if (Core.DynamicModuleLoader.LoadModule("hvac", moduleDllPath))
                {
                    pnlHVAC.IsEnabled = true;
                    pnlHVAC.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("[HUB] ✓ HVAC module loaded and activated");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] ✗ Failed to load HVAC module");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading HVAC module: {ex.Message}");
                Core.DebugLogger.Log($"[HUB] ERROR: Failed to load HVAC module - {ex.Message}");
            }
        }
    }
}
