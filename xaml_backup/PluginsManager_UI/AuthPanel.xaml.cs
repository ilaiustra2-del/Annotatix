using System;
using System.Linq;
using System.Windows;
using PluginsManager.Core;

namespace PluginsManager.UI
{
    public partial class AuthPanel : Window
    {
        private readonly AuthService _authService;

        public AuthPanel()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] Initializing AuthPanel...");
                InitializeComponent();
                System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] InitializeComponent completed");
                
                _authService = new AuthService();
                System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] AuthService created successfully");
                
                // Force enable button
                if (btnLogin != null)
                {
                    btnLogin.IsEnabled = true;
                    System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] btnLogin enabled");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] WARNING: btnLogin is NULL!");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AUTH-PANEL] INITIALIZATION ERROR: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[AUTH-PANEL] Stack: {ex.StackTrace}");
                MessageBox.Show($"Ошибка инициализации окна авторизации:\n{ex.Message}\n\nСм. лог для деталей", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private async void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] Login button clicked");
                
                string login = txtLogin.Text.Trim();
                string password = txtPassword.Password;
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-PANEL] Login: '{login}', Password length: {password.Length}");

                // Validation
                if (string.IsNullOrEmpty(login))
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] Validation failed: empty login");
                    txtStatus.Text = "Введите логин";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(220, 53, 69));
                    return;
                }

                if (string.IsNullOrEmpty(password))
                {
                    System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] Validation failed: empty password");
                    txtStatus.Text = "Введите пароль";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(220, 53, 69));
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] Validation passed, disabling button");
                
                btnLogin.IsEnabled = false;
                txtStatus.Text = "Проверка данных...";
                txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0, 123, 255));
                this.UpdateLayout();
                
                System.Diagnostics.Debug.WriteLine("[AUTH-PANEL] Calling AuthService.AuthenticateAsync...");

                // Authenticate
                var result = await _authService.AuthenticateAsync(login, password);
                
                System.Diagnostics.Debug.WriteLine($"[AUTH-PANEL] Auth result: IsSuccess={result.IsSuccess}");

                if (result.IsSuccess)
                {
                    // Build status message with modules
                    string modulesInfo = "";
                    if (result.Modules != null && result.Modules.Count > 0)
                    {
                        var activeModules = result.Modules.Where(m => m.IsActive).ToList();
                        if (activeModules.Count > 0)
                        {
                            var moduleDetails = activeModules.Select(m => 
                                $"{m.ModuleTag} (до {m.EndDate:dd.MM.yyyy})");
                            modulesInfo = $"\nАктивные модули: {string.Join(", ", moduleDetails)}";
                        }
                        else
                        {
                            modulesInfo = "\nНет активных модулей";
                        }
                    }
                    
                    txtStatus.Text = $"Авторизация успешна!{modulesInfo}";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(40, 167, 69));

                    AuthService.CurrentUser = result;

                    await System.Threading.Tasks.Task.Delay(1000);
                    this.DialogResult = true;
                    this.Close();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[AUTH-PANEL] FAILED - Error: {result.ErrorMessage}");
                    
                    txtStatus.Text = result.ErrorMessage ?? "Неверный логин или пароль";
                    txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(220, 53, 69));
                    btnLogin.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ИСКЛЮЧЕНИЕ!\n\nТип: {ex.GetType().Name}\nСообщение: {ex.Message}\n\nStack:\n{ex.StackTrace}", 
                    "Критическая ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                
                txtStatus.Text = $"Ошибка: {ex.Message}";
                txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(220, 53, 69));
                btnLogin.IsEnabled = true;
            }
        }
    }
}
