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
                    txtVersion.Text = $"v 2.{buildNumber}";
                    System.Diagnostics.Debug.WriteLine($"[HUB] Version set from BuildNumber.txt: v 2.{buildNumber}");
                }
                else
                {
                    // Fallback to assembly version
                    var version = assembly.GetName().Version;
                    // Version format is Major.Minor.Build.Revision (e.g. 2.65.0.0)
                    txtVersion.Text = $"v {version.Major}.{version.Minor}";
                    System.Diagnostics.Debug.WriteLine($"[HUB] Version set from assembly: v {version.Major}.{version.Minor}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HUB] Error getting version: {ex.Message}");
                txtVersion.Text = "v 2.0"; // Fallback
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
        /// Activate module buttons based on user's active modules
        /// </summary>
        private void ActivateModuleButtons(List<Core.UserModule> modules)
        {
            if (modules == null || modules.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[HUB] No modules to activate");
                return;
            }
            
            System.Diagnostics.Debug.WriteLine($"[HUB] Activating modules: {modules.Count}");
            
            foreach (var module in modules)
            {
                if (!module.IsActive)
                    continue;
                
                string moduleTag = module.ModuleTag.ToLower();
                System.Diagnostics.Debug.WriteLine($"[HUB] Activating module: {moduleTag}");
                
                switch (moduleTag)
                {
                    case "dwg2rvt":
                        pnlDwg2rvt.IsEnabled = true;
                        pnlDwg2rvt.Opacity = 1.0;
                        System.Diagnostics.Debug.WriteLine("[HUB] DWG2RVT module activated");
                        break;
                        
                    case "hvac":
                        pnlHVAC.IsEnabled = true;
                        pnlHVAC.Opacity = 1.0;
                        System.Diagnostics.Debug.WriteLine("[HUB] HVAC module activated");
                        break;
                        
                    case "full":
                        // Activate all modules
                        pnlDwg2rvt.IsEnabled = true;
                        pnlDwg2rvt.Opacity = 1.0;
                        pnlHVAC.IsEnabled = true;
                        pnlHVAC.Opacity = 1.0;
                        System.Diagnostics.Debug.WriteLine("[HUB] All modules activated (full access)");
                        break;
                        
                    default:
                        System.Diagnostics.Debug.WriteLine($"[HUB] Unknown module: {moduleTag}");
                        break;
                }
            }
        }
    }
}
