using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace dwg2rvt.UI
{
    public partial class MainHubPanel : Window
    {
        private UIApplication _uiApp;
        private object _hubContent;
        private ExternalEvent _annotateEvent;
        private ExternalEvent _placeElementsEvent;
        private ExternalEvent _placeSingleBlockTypeEvent;

        public MainHubPanel(UIApplication uiApp, ExternalEvent annotateEvent, ExternalEvent placeElementsEvent, ExternalEvent placeSingleBlockTypeEvent = null)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _annotateEvent = annotateEvent;
            _placeElementsEvent = placeElementsEvent;
            _placeSingleBlockTypeEvent = placeSingleBlockTypeEvent;
            _hubContent = MainContent.Content;
            
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
                // Log panel creation with timestamp
                var clickTimestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                Core.DebugLogger.Log("");
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("[HUB] User clicked DWG2RVT button");
                Core.DebugLogger.Log("[HUB] Creating dwg2rvtPanel instance NOW (not at startup)...");
                
                // Log where the dwg2rvtPanel class is loaded from
                var panelType = typeof(dwg2rvtPanel);
                var panelAssembly = panelType.Assembly;
                Core.DebugLogger.Log($"[HUB] dwg2rvtPanel type: {panelType.FullName}");
                Core.DebugLogger.Log($"[HUB] Loaded from assembly: {panelAssembly.GetName().Name}");
                Core.DebugLogger.Log($"[HUB] Assembly location: {panelAssembly.Location}");
                Core.DebugLogger.Log($"[HUB] Assembly version: {panelAssembly.GetName().Version}");
                
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("");
                
                // Switch to dwg2rvt panel
                txtTitle.Text = "DWG2RVT";
                var panel = new dwg2rvtPanel(_uiApp, _annotateEvent, _placeElementsEvent, _placeSingleBlockTypeEvent);
                
                // Link panel to handler via OpenHubCommand (which has static reference)
                Commands.OpenHubCommand.LinkPanelToHandler(panel);
                
                MainContent.Content = panel;
                btnBack.Visibility = Visibility.Visible;
                
                // Hide Right Sidebar when dwg2rvt panel is open
                RightSidebar.Visibility = Visibility.Collapsed;
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
                Core.DebugLogger.Log("[HUB] Creating HVACPanel instance NOW (not at startup)...");
                
                // Log where the HVACPanel class is loaded from
                var panelType = typeof(HVACPanel);
                var panelAssembly = panelType.Assembly;
                Core.DebugLogger.Log($"[HUB] HVACPanel type: {panelType.FullName}");
                Core.DebugLogger.Log($"[HUB] Loaded from assembly: {panelAssembly.GetName().Name}");
                Core.DebugLogger.Log($"[HUB] Assembly location: {panelAssembly.Location}");
                Core.DebugLogger.Log($"[HUB] Assembly version: {panelAssembly.GetName().Version}");
                
                Core.DebugLogger.LogSeparator('-');
                Core.DebugLogger.Log("");
                
                // Switch to HVAC panel
                txtTitle.Text = "HVAC";
                var hvacPanel = new HVACPanel();
                
                MainContent.Content = hvacPanel;
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
        
        private void BtnAuth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[AUTH] Opening authentication panel...");
                
                // Open authentication panel
                var authPanel = new AuthPanel();
                bool? result = authPanel.ShowDialog();
                
                System.Diagnostics.Debug.WriteLine($"[AUTH] Dialog result: {result}");
                
                if (result == true)
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
                        System.Diagnostics.Debug.WriteLine($"[AUTH] SUCCESS - User: {currentUser.Login}, Plan: {currentUser.SubscriptionPlan}");
                        
                        // Update button state
                        btnAuth.Content = $"Вы вошли: {currentUser.Login}";
                        btnAuth.IsEnabled = false;
                        
                        // Show user status with active modules in bottom right
                        var activeModules = currentUser.Modules?.Where(m => m.IsActive).ToList();
                        string modulesText = "";
                        if (activeModules != null && activeModules.Count > 0)
                        {
                            var moduleDetails = activeModules.Select(m => 
                                $"{m.ModuleTag} (до {m.EndDate:dd.MM.yyyy})").ToList();
                            modulesText = $"\nМодули: {string.Join(", ", moduleDetails)}";
                        }
                        else
                        {
                            modulesText = "\nМодули: нет активных";
                        }
                        
                        txtUserStatus.Text = $"Пользователь: {currentUser.Login}{modulesText}";
                        txtUserStatus.Visibility = Visibility.Visible;
                        
                        // Activate module buttons based on user's modules
                        ActivateModuleButtons(currentUser.Modules);
                        
                        // Show success notification (without subscription plan)
                        string welcomeMsg = "Добро пожаловать!";
                        if (activeModules != null && activeModules.Count > 0)
                        {
                            var moduleInfo = activeModules.Select(m => 
                                $"{m.ModuleTag} (активен до {m.EndDate:dd.MM.yyyy})").ToList();
                            welcomeMsg += $"\nАктивные модули:\n{string.Join("\n", moduleInfo)}";
                        }
                        MessageBox.Show(welcomeMsg, "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
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
            var modulesPath = Path.Combine(assemblyPath, "Modules");
            
            System.Diagnostics.Debug.WriteLine($"[HUB] Modules path: {modulesPath}");
            
            foreach (var module in modules)
            {
                if (!module.IsActive)
                    continue;
                
                string moduleTag = module.ModuleTag.ToLower();
                System.Diagnostics.Debug.WriteLine($"[HUB] Processing module: {moduleTag}");
                
                switch (moduleTag)
                {
                    case "dwg2rvt":
                        LoadAndActivateDwg2rvtModule(modulesPath);
                        break;
                        
                    case "hvac":
                        LoadAndActivateHVACModule(modulesPath);
                        break;
                        
                    case "full":
                        // Load all modules
                        LoadAndActivateDwg2rvtModule(modulesPath);
                        LoadAndActivateHVACModule(modulesPath);
                        System.Diagnostics.Debug.WriteLine("[HUB] All modules activated (full access)");
                        break;
                        
                    default:
                        System.Diagnostics.Debug.WriteLine($"[HUB] Unknown module: {moduleTag}");
                        break;
                }
            }
        }
        
        /// <summary>
        /// Load and activate DWG2RVT module
        /// </summary>
        private void LoadAndActivateDwg2rvtModule(string modulesPath)
        {
            try
            {
                // For now, dwg2rvt module is built-in, so just enable the button
                // In future, this will load external .cs file
                System.Diagnostics.Debug.WriteLine("[HUB] DWG2RVT module is built-in, enabling button");
                pnlDwg2rvt.IsEnabled = true;
                pnlDwg2rvt.Opacity = 1.0;
                System.Diagnostics.Debug.WriteLine("[HUB] DWG2RVT module activated");
                
                /* Future implementation for dynamic loading:
                var moduleFile = Path.Combine(modulesPath, "dwg2rvt.cs");
                if (Core.ModuleLoader.LoadModule("dwg2rvt", moduleFile))
                {
                    pnlDwg2rvt.IsEnabled = true;
                    pnlDwg2rvt.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("[HUB] DWG2RVT module loaded and activated");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] Failed to load DWG2RVT module");
                }
                */
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading DWG2RVT module: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Load and activate HVAC module
        /// </summary>
        private void LoadAndActivateHVACModule(string modulesPath)
        {
            try
            {
                // For now, HVAC module is built-in, so just enable the button
                // In future, this will load external .cs file
                System.Diagnostics.Debug.WriteLine("[HUB] HVAC module is built-in, enabling button");
                pnlHVAC.IsEnabled = true;
                pnlHVAC.Opacity = 1.0;
                System.Diagnostics.Debug.WriteLine("[HUB] HVAC module activated");
                
                /* Future implementation for dynamic loading:
                var moduleFile = Path.Combine(modulesPath, "hvac.cs");
                if (Core.ModuleLoader.LoadModule("hvac", moduleFile))
                {
                    pnlHVAC.IsEnabled = true;
                    pnlHVAC.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("[HUB] HVAC module loaded and activated");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[HUB] Failed to load HVAC module");
                }
                */
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error loading HVAC module: {ex.Message}");
            }
        }
    }
}
