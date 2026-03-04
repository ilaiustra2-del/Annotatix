using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.UI
{
    public partial class MainHubPanel : Window
    {
        private UIApplication _uiApp;
        private object _hubContent;
        private ExternalEvent _openDwg2rvtPanelEvent;
        private Commands.OpenDwg2rvtPanelHandler _openDwg2rvtPanelHandler;
        private ExternalEvent _openHvacPanelEvent;
        private Commands.OpenHvacPanelHandler _openHvacPanelHandler;
        
        // FamilySync ExternalEvent will be created when module is loaded
        private ExternalEvent _familySyncEvent;
        private object _familySyncHandler;
        
        // AutoNumbering ExternalEvent for numbering (created in constructor)
        private ExternalEvent _autoNumberingEvent;
        private Commands.NumberRisersHandler _autoNumberingHandler;
        
        // Cache for module panels to avoid recreating them
        private object _familySyncPanelContent;
        private object _autoNumberingPanelContent;

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
            
            // Create ExternalEvent for opening HVAC panel
            _openHvacPanelHandler = new Commands.OpenHvacPanelHandler();
            _openHvacPanelHandler.SetUIApplication(uiApp);
            _openHvacPanelEvent = ExternalEvent.Create(_openHvacPanelHandler);
            
            // Create ExternalEvent for AutoNumbering numbering
            // Created HERE (in constructor from IExternalCommand context) to avoid API context errors
            _autoNumberingHandler = new Commands.NumberRisersHandler();
            _autoNumberingEvent = ExternalEvent.Create(_autoNumberingHandler);
            System.Diagnostics.Debug.WriteLine("[HUB] AutoNumbering ExternalEvent created in constructor");
            
            // Register this panel with the commands so they can update UI
            Commands.OpenDwg2rvtPanelCommand.SetHubPanel(this);
            Commands.OpenHvacPanelCommand.SetHubPanel(this);
            
            // Load icons from assembly location
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = Path.GetDirectoryName(assembly.Location);
                var iconsPath = Path.Combine(assemblyPath, "UI", "icons");
                
                System.Diagnostics.Debug.WriteLine($"[HUB] Looking for icons at: {iconsPath}");
                DebugLogger.Log($"[HUB] Looking for icons at: {iconsPath}");
                
                // Load DWG2RVT icon
                LoadIcon(Path.Combine(iconsPath, "dwg2rvt80.png"), imgDwg2rvtIcon);
                
                // Load HVAC icon
                LoadIcon(Path.Combine(iconsPath, "hvac80.png"), imgHvacIcon);
                
                // Load FamilySync icon
                LoadIcon(Path.Combine(iconsPath, "familysync80.png"), imgFamilySyncIcon);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading icons: {ex.Message}");
                DebugLogger.Log($"[HUB] Error loading icons: {ex.Message}");
            }
            
            // Set version number from BuildNumber.txt or assembly
            try
            {
                // Try to read from BuildNumber.txt (more reliable)
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = Path.GetDirectoryName(assembly.Location);
                var buildNumberFile = Path.Combine(assemblyPath, "BuildNumber.txt");
                
                System.Diagnostics.Debug.WriteLine($"[HUB] Looking for BuildNumber.txt at: {buildNumberFile}");
                
                string versionString = "";
                if (File.Exists(buildNumberFile))
                {
                    var buildNumber = File.ReadAllText(buildNumberFile).Trim();
                    versionString = $"v 3.{buildNumber}";
                    this.Title = $"Менеджер плагинов AnnotatiX.AI  {versionString}";
                    System.Diagnostics.Debug.WriteLine($"[HUB] Version set from BuildNumber.txt: {versionString}");
                }
                else
                {
                    // Fallback to assembly version
                    var version = assembly.GetName().Version;
                    // Version format is Major.Minor.Build.Revision (e.g. 3.001.0.0)
                    versionString = $"v {version.Major}.{version.Minor}";
                    this.Title = $"Менеджер плагинов AnnotatiX.AI  {versionString}";
                    System.Diagnostics.Debug.WriteLine($"[HUB] Version set from assembly: {versionString}");
                }
                
                // Update footer version display
                txtStatusRight.Text = $"Revit 2024 | Build {versionString}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error getting version: {ex.Message}");
                this.Title = "Менеджер плагинов AnnotatiX.AI  v 3.0"; // Fallback
                txtStatusRight.Text = "Revit 2024 | Build v 3.0";
            }
            
            // Load FamilySync module automatically REMOVED
            // Note: Module loading now happens in ActivateModuleButtons after authentication
            // LoadFamilySyncModule(); // REMOVED - FamilySync is now loaded dynamically like other modules
            
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

        // REMOVED: Window_Close - Close button removed from UI
        
        /// <summary>
        /// Load an icon from file to an Image control
        /// </summary>
        private void LoadIcon(string iconPath, System.Windows.Controls.Image imageControl)
        {
            try
            {
                if (File.Exists(iconPath))
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(iconPath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                    bitmap.EndInit();
                    imageControl.Source = bitmap;
                    System.Diagnostics.Debug.WriteLine($"[HUB] Icon loaded: {Path.GetFileName(iconPath)} to {imageControl.Name}");
                    DebugLogger.Log($"[HUB] Icon loaded: {Path.GetFileName(iconPath)} to {imageControl.Name}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HUB] Icon file not found: {iconPath}");
                    DebugLogger.Log($"[HUB] Icon file not found: {iconPath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading icon {iconPath}: {ex.Message}");
                DebugLogger.Log($"[HUB] Error loading icon {iconPath}: {ex.Message}");
            }
        }
        
        private void BtnDebugLog_Click(object sender, RoutedEventArgs e)
        {
            // Open debug log file in Notepad
            try
            {
                Core.DebugLogger.OpenLogFile();
                Core.DebugLogger.Log("[HUB] User opened debug log file");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть лог-файл:\n{ex.Message}\n\nПуть: {Core.DebugLogger.GetLogFilePath()}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                DebugLogger.Log($"[HUB] ERROR opening log file: {ex.Message}");
            }
        }
        
        private void BtnModulesInfo_Click(object sender, RoutedEventArgs e)
        {
            // Show modules info
            var info = "Установленные модули:\n\n";
            
            if (Core.AuthService.CurrentUser != null && Core.AuthService.CurrentUser.Modules != null)
            {
                foreach (var module in Core.AuthService.CurrentUser.Modules)
                {
                    info += $"• {module.ModuleTag}\n";
                    info += $"  Активен: {(module.IsActive ? "Да" : "Нет")}\n";
                    info += $"  Период: {module.StartDate:dd.MM.yyyy} - {module.EndDate:dd.MM.yyyy}\n\n";
                }
            }
            else
            {
                info += "Нет информации о модулях.\nВыполните авторизацию.";
            }
            
            MessageBox.Show(info, "Информация о модулях", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        // REMOVED: LoadFamilySyncModule() - now loaded dynamically after authentication
        
        private void Hvac_Click(object sender, MouseButtonEventArgs e)
        {
            // Check if button is enabled
            if (!pnlHvac.IsEnabled)
            {
                System.Diagnostics.Debug.WriteLine("[HUB] HVAC button is disabled");
                return;
            }
            
            try
            {
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("[HUB] User clicked HVAC button");
                Core.DebugLogger.Log("[HUB] Triggering OpenHVACPanelCommand...");
                
                // Raise the ExternalEvent to trigger OpenHVACPanelCommand
                // which will create ExternalEvents in proper IExternalCommand context
                if (_openHvacPanelEvent != null)
                {
                    _openHvacPanelEvent.Raise();
                    Core.DebugLogger.Log("[HUB] OpenHVACPanelCommand event raised");
                }
                else
                {
                    Core.DebugLogger.Log("[HUB] ERROR: _openHvacPanelEvent is null");
                    MessageBox.Show("Ошибка открытия модуля HVAC", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[HUB] ERROR loading HVAC: {ex.Message}");
                MessageBox.Show($"Ошибка загрузки модуля HVAC:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void FamilySync_Click(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("[HUB] User clicked Family Sync button");
                
                // Check if panel already created (cached)
                if (_familySyncPanelContent != null)
                {
                    Core.DebugLogger.Log("[HUB] Using cached Family Sync panel content");
                    
                    // Switch to cached FamilySync panel
                    MainContent.Content = _familySyncPanelContent;
                    // Hide sidebar and expand content to full width
                    RightSidebar.Visibility = System.Windows.Visibility.Collapsed;
                    System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 2);
                    TitleBar.Visibility = System.Windows.Visibility.Visible;
                    txtModuleTitle.Text = "Family Sync - Синхронизация параметров вложенных семейств";
                    
                    Core.DebugLogger.Log("[HUB] Family Sync panel switched (from cache)");
                    Core.DebugLogger.LogSeparator('-');
                    Core.DebugLogger.Log("");
                    return;
                }
                
                Core.DebugLogger.Log("[HUB] Loading Family Sync module panel (first time)...");
                
                // Get module instance
                var module = Core.DynamicModuleLoader.GetModuleInstance("family_sync");
                if (module == null)
                {
                    MessageBox.Show("Модуль Family Sync не загружен. Проверьте наличие файла FamilySync.Module.dll.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log($"[HUB] Module found: {module.ModuleName} v{module.ModuleVersion}");
                
                // Create panel from module, passing ExternalEvent
                object[] parameters = (_familySyncEvent != null && _familySyncHandler != null) 
                    ? new object[] { _uiApp, _familySyncHandler, _familySyncEvent }
                    : new object[] { _uiApp };
                    
                var panel = module.CreatePanel(parameters);
                if (panel == null)
                {
                    MessageBox.Show("Не удалось создать панель модуля Family Sync.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log("[HUB] Panel created successfully");
                
                // Extract content from Window
                var panelContent = panel.Content;
                if (panelContent == null)
                {
                    MessageBox.Show("Панель модуля не содержит контента.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log($"[HUB] Panel content type: {panelContent.GetType().Name}");
                
                // Cache the panel content for reuse
                _familySyncPanelContent = panelContent;
                Core.DebugLogger.Log("[HUB] Panel content cached for future use");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("");
                
                // Switch to FamilySync panel
                MainContent.Content = panelContent;
                // Hide sidebar and expand content to full width
                RightSidebar.Visibility = System.Windows.Visibility.Collapsed;
                System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 2);
                TitleBar.Visibility = System.Windows.Visibility.Visible;
                txtModuleTitle.Text = "Family Sync - Синхронизация параметров вложенных семейств";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке панели Family Sync:\n{ex.Message}\n\n{ex.StackTrace}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void AutoNumbering_Click(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("[HUB] User clicked AutoNumbering button");
                
                // Check if panel already created (cached)
                if (_autoNumberingPanelContent != null)
                {
                    Core.DebugLogger.Log("[HUB] Using cached AutoNumbering panel content");
                    
                    // Switch to cached AutoNumbering panel
                    MainContent.Content = _autoNumberingPanelContent;
                    RightSidebar.Visibility = System.Windows.Visibility.Collapsed;
                    System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 2);
                    TitleBar.Visibility = System.Windows.Visibility.Visible;
                    txtModuleTitle.Text = "AutoNumbering - Автонумерация стояков";
                    
                    Core.DebugLogger.Log("[HUB] AutoNumbering panel switched (from cache)");
                    Core.DebugLogger.LogSeparator('-');
                    Core.DebugLogger.Log("");
                    return;
                }
                
                Core.DebugLogger.Log("[HUB] Loading AutoNumbering module panel (first time)...");
                
                // Get module instance
                var module = Core.DynamicModuleLoader.GetModuleInstance("autonumbering");
                if (module == null)
                {
                    MessageBox.Show("Модуль AutoNumbering не загружен. Проверьте наличие файла AutoNumbering.Module.dll.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log($"[HUB] Module found: {module.ModuleName} v{module.ModuleVersion}");
                
                // Pass UIApp AND pre-created ExternalEvent to panel
                Core.DebugLogger.Log($"[HUB] Passing ExternalEvent to panel: Event={_autoNumberingEvent != null}, Handler={_autoNumberingHandler != null}");
                var panel = module.CreatePanel(new object[] { _uiApp, _autoNumberingEvent, _autoNumberingHandler });
                if (panel == null)
                {
                    MessageBox.Show("Не удалось создать панель модуля AutoNumbering.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log("[HUB] Panel created successfully");
                
                // Extract content from Window
                var panelContent = panel.Content;
                if (panelContent == null)
                {
                    MessageBox.Show("Панель модуля не содержит контента.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                Core.DebugLogger.Log($"[HUB] Panel content type: {panelContent.GetType().Name}");
                
                // Cache the panel content for reuse
                _autoNumberingPanelContent = panelContent;
                Core.DebugLogger.Log("[HUB] Panel content cached for future use");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("");
                
                // Switch to AutoNumbering panel
                MainContent.Content = panelContent;
                RightSidebar.Visibility = System.Windows.Visibility.Collapsed;
                System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 2);
                TitleBar.Visibility = System.Windows.Visibility.Visible;
                txtModuleTitle.Text = "AutoNumbering - Автонумерация стояков";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке панели AutoNumbering:\n{ex.Message}\n\n{ex.StackTrace}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnAuth_Click(object sender, RoutedEventArgs e)
        {
            // Check if already authenticated - then logout
            if (Core.AuthService.CurrentUser != null && Core.AuthService.CurrentUser.IsSuccess)
            {
                // Logout
                var result = MessageBox.Show("Вы уверены, что хотите выйти?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    Core.AuthService.Logout();
                    
                    // Reset UI
                    warningMessage.Visibility = Visibility.Visible;
                    successMessage.Visibility = Visibility.Collapsed;
                    pnlLicenseInfo.Visibility = Visibility.Collapsed;
                    btnAuth.Content = "Войти";
                    btnRefreshUserData.Visibility = Visibility.Collapsed;
                    
                    // Disable all module buttons
                    pnlDwg2rvt.IsEnabled = false;
                    pnlDwg2rvt.Opacity = 0.5;
                    pnlHvac.IsEnabled = false;
                    pnlHvac.Opacity = 0.5;
                    pnlFamilySync.IsEnabled = false;
                    pnlFamilySync.Opacity = 0.5;
                }
                return;
            }
            
            // Show auth panel
            var authPanel = new AuthPanel();
            var dialogResult = authPanel.ShowDialog();
            
            if (dialogResult == true && Core.AuthService.CurrentUser != null)
            {
                // Authentication successful
                var currentUser = Core.AuthService.CurrentUser;
                
                // Download missing modules
                await DownloadMissingModules(currentUser.ModuleFiles);
                
                // Update UI
                ShowModulesInfo(currentUser.Modules);
                
                // Show refresh button
                btnRefreshUserData.Visibility = Visibility.Visible;
                
                // Activate module buttons
                ActivateModuleButtons(currentUser.Modules);
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
            if (!pnlHvac.IsEnabled)
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
                Core.DebugLogger.Log("[HUB] Opening HVAC module via ExternalEvent...");
                
                // Use ExternalEvent to open panel (this is the correct way)
                if (_openHvacPanelEvent != null)
                {
                    _openHvacPanelEvent.Raise();
                    Core.DebugLogger.Log("[HUB] OpenHvacPanel event raised");
                }
                else
                {
                    Core.DebugLogger.Log("[HUB] ERROR: OpenHvacPanel ExternalEvent not initialized");
                    MessageBox.Show("Не удалось инициализировать событие открытия HVAC.",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке панели HVAC:\n{ex.Message}\n\n{ex.StackTrace}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Show HVAC panel content (called from OpenHvacPanelCommand)
        /// </summary>
        public void ShowHvacPanel(object panelContent)
        {
            MainContent.Content = panelContent;
            // Hide sidebar and expand content to full width
            RightSidebar.Visibility = System.Windows.Visibility.Collapsed;
            System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 2);
            TitleBar.Visibility = System.Windows.Visibility.Visible;
            txtModuleTitle.Text = "HVAC SuperScheme - Построение принципиальных схем";
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            // Back to Hub
            MainContent.Content = _hubContent;
            // Show sidebar and restore normal layout
            RightSidebar.Visibility = System.Windows.Visibility.Visible;
            System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 1);
            TitleBar.Visibility = System.Windows.Visibility.Collapsed;
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
                MainContent.Content = panelContent;
                // Hide sidebar and expand content to full width
                RightSidebar.Visibility = System.Windows.Visibility.Collapsed;
                System.Windows.Controls.Grid.SetColumnSpan(ContentArea, 2);
                TitleBar.Visibility = System.Windows.Visibility.Visible;
                txtModuleTitle.Text = "DWG2RVT";
                
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
        
        /// <summary>
        /// Show modules information in the UI
        /// </summary>
        private void ShowModulesInfo(List<Core.UserModule> modules)
        {
            // Update UI for authenticated state
            warningMessage.Visibility = Visibility.Collapsed;
            successMessage.Visibility = Visibility.Visible;
            pnlLicenseInfo.Visibility = Visibility.Visible;
            btnAuth.Content = "Выйти";
            
            if (modules == null || modules.Count == 0)
            {
                txtPlan.Text = "Free Plan";
                txtExpiry.Text = "N/A";
                return;
            }
            
            var activeModules = modules.Where(m => m.IsActive).ToList();
            if (activeModules.Count > 0)
            {
                txtPlan.Text = "Pro Plan";
                // Show earliest expiry date
                var earliestExpiry = activeModules.Min(m => m.EndDate);
                txtExpiry.Text = earliestExpiry.ToString("dd.MM.yyyy");
            }
            else
            {
                txtPlan.Text = "Free Plan";
                txtExpiry.Text = "N/A";
            }
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
                    pnlHvac.IsEnabled = false;
                    pnlHvac.Opacity = 0.5;
                    
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
                        
                    case "family_sync":
                        if (!CheckModuleFilesExist(modulesPath, "family_sync", "FamilySync.Module.dll"))
                            missingModules.Add("family_sync");
                        else
                            LoadAndActivateFamilySyncModule(modulesPath);
                        break;
                    
                    case "autonumbering":
                        if (!CheckModuleFilesExist(modulesPath, "autonumbering", "AutoNumbering.Module.dll"))
                            missingModules.Add("autonumbering");
                        else
                            LoadAndActivateAutoNumberingModule(modulesPath);
                        break;
                        
                    case "full":
                        // Check all three modules
                        bool dwg2rvtExists = CheckModuleFilesExist(modulesPath, "dwg2rvt", "dwg2rvt.Module.dll");
                        bool hvacExists = CheckModuleFilesExist(modulesPath, "hvac", "HVAC.Module.dll");
                        bool familySyncExists = CheckModuleFilesExist(modulesPath, "family_sync", "FamilySync.Module.dll");
                        bool autoNumberingExists = CheckModuleFilesExist(modulesPath, "autonumbering", "AutoNumbering.Module.dll");
                        
                        if (!dwg2rvtExists)
                            missingModules.Add("dwg2rvt");
                        else
                            LoadAndActivateDwg2rvtModule(modulesPath);
                            
                        if (!hvacExists)
                            missingModules.Add("hvac");
                        else
                            LoadAndActivateHVACModule(modulesPath);
                            
                        if (!familySyncExists)
                            missingModules.Add("family_sync");
                        else
                            LoadAndActivateFamilySyncModule(modulesPath);
                        
                        if (!autoNumberingExists)
                            missingModules.Add("autonumbering");
                        else
                            LoadAndActivateAutoNumberingModule(modulesPath);
                            
                        if (dwg2rvtExists && hvacExists && familySyncExists && autoNumberingExists)
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
            
            // FamilySync is now loaded dynamically based on user subscription (like other modules)
            // REMOVED: Always activate FamilySync - it's now subscription-based
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
                    pnlHvac.IsEnabled = true;
                    pnlHvac.Opacity = 1.0;
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
        
        /// <summary>
        /// Load and activate FamilySync module
        /// </summary>
        private void LoadAndActivateFamilySyncModule(string modulesPath)
        {
            try
            {
                var moduleDllPath = Path.Combine(modulesPath, "family_sync", "FamilySync.Module.dll");
                System.Diagnostics.Debug.WriteLine($"[HUB] Loading FamilySync module from: {moduleDllPath}");
                
                if (Core.DynamicModuleLoader.LoadModule("family_sync", moduleDllPath))
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] FamilySync module loaded successfully");
                    
                    // Create FamilySync ExternalEvent only if not already created
                    if (_familySyncEvent == null)
                    {
                        try
                        {
                            System.Diagnostics.Debug.WriteLine("[HUB] Attempting to create FamilySync ExternalEvent...");
                            var familySyncHandlerType = Type.GetType("FamilySync.Module.UI.FamilySyncHandler, FamilySync.Module");
                            
                            if (familySyncHandlerType != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[HUB] Found type: {familySyncHandlerType.FullName}");
                                _familySyncHandler = Activator.CreateInstance(familySyncHandlerType);
                                System.Diagnostics.Debug.WriteLine($"[HUB] Created handler instance");
                                
                                var iExternalEventHandler = _familySyncHandler as IExternalEventHandler;
                                if (iExternalEventHandler != null)
                                {
                                    _familySyncEvent = ExternalEvent.Create(iExternalEventHandler);
                                    System.Diagnostics.Debug.WriteLine("[HUB] FamilySync ExternalEvent created successfully");
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine("[HUB] Handler does not implement IExternalEventHandler");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("[HUB] FamilySyncHandler type not found via reflection");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[HUB] Failed to create FamilySync ExternalEvent: {ex.Message}");
                            System.Diagnostics.Debug.WriteLine($"[HUB] Stack trace: {ex.StackTrace}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[HUB] FamilySync ExternalEvent already exists, skipping creation");
                    }
                    
                    // Enable and make FamilySync card visible
                    pnlFamilySync.IsEnabled = true;
                    pnlFamilySync.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("[HUB] ✓ FamilySync module loaded and activated");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] ✗ Failed to load FamilySync module");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading FamilySync module: {ex.Message}");
                Core.DebugLogger.Log($"[HUB] ERROR: Failed to load FamilySync module - {ex.Message}");
            }
        }
        
        /// <summary>
        /// Load and activate AutoNumbering module
        /// </summary>
        private void LoadAndActivateAutoNumberingModule(string modulesPath)
        {
            try
            {
                var moduleDllPath = Path.Combine(modulesPath, "autonumbering", "AutoNumbering.Module.dll");
                System.Diagnostics.Debug.WriteLine($"[HUB] Loading AutoNumbering module from: {moduleDllPath}");
                
                if (Core.DynamicModuleLoader.LoadModule("autonumbering", moduleDllPath))
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] AutoNumbering module loaded successfully");
                    
                    // Enable and make AutoNumbering card visible
                    pnlAutoNumbering.IsEnabled = true;
                    pnlAutoNumbering.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("[HUB] ✓ AutoNumbering module loaded and activated");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] ✗ Failed to load AutoNumbering module");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading AutoNumbering module: {ex.Message}");
                Core.DebugLogger.Log($"[HUB] ERROR: Failed to load AutoNumbering module - {ex.Message}");
            }
        }
    }
}