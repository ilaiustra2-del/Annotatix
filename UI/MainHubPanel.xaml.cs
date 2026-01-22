using System;
using System.IO;
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
            try
            {
                // Switch to dwg2rvt panel
                txtTitle.Text = "dwg2rvt - Анализ DWG";
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
                        
                        // Show user status in bottom right
                        txtUserStatus.Text = $"Пользователь: {currentUser.Login}";
                        txtUserStatus.Visibility = Visibility.Visible;
                        
                        // Show success notification
                        MessageBox.Show($"Добро пожаловать!\nТариф: {currentUser.SubscriptionPlan}", 
                            "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
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
    }
}
