using System;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace HVAC.Module.UI
{
    /// <summary>
    /// Static state manager for HVAC synchronization
    /// </summary>
    public static class HVACSyncState
    {
        /// <summary>
        /// Enable/disable drawing-model synchronization via Updater
        /// </summary>
        public static bool IsSyncEnabled { get; set; } = true;
    }
    
    /// <summary>
    /// HVAC Module Panel - User interface for HVAC SuperScheme commands
    /// </summary>
    public partial class HVACPanel : UserControl
    {
        private UIApplication _uiApp;
        
        // External events for Revit API commands
        private ExternalEvent _createSchemaEvent;
        private ExternalEvent _completeSchemaEvent;
        private ExternalEvent _settingsEvent;
        private ExternalEvent _toggleIdlingEvent;
        private Commands.ToggleIdlingHandler _toggleIdlingHandler;
        
        // Callback to go back to hub
        public Action OnBackToHub { get; set; }

        public HVACPanel(UIApplication uiApp, ExternalEvent createSchemaEvent, ExternalEvent completeSchemaEvent, ExternalEvent settingsEvent)
        {
            InitializeComponent();
            
            _uiApp = uiApp;
            _createSchemaEvent = createSchemaEvent;
            _completeSchemaEvent = completeSchemaEvent;
            _settingsEvent = settingsEvent;
            
            // Create ToggleIdlingHandler ExternalEvent
            _toggleIdlingHandler = new Commands.ToggleIdlingHandler();
            _toggleIdlingEvent = ExternalEvent.Create(_toggleIdlingHandler);
            
            // Debug logging
            DebugLogger.Log("[HVAC-PANEL] *** PANEL INSTANCE CREATED ***");
            DebugLogger.Log("[HVAC-PANEL] Panel created after authentication");
            DebugLogger.Log($"[HVAC-PANEL] ExternalEvents received: createSchema={_createSchemaEvent != null}, completeSchema={_completeSchemaEvent != null}, settings={_settingsEvent != null}");
            
            // Load settings from file and sync with HVACSyncState
            HVACSuperScheme.Commands.Settings.SettingStorage.ReadSettings();
            HVACSyncState.IsSyncEnabled = HVACSuperScheme.Commands.Settings.SettingStorage.Instance.IsUpdaterSync;
            
            // Initialize checkbox state from loaded settings
            chkSyncEnabled.IsChecked = HVACSyncState.IsSyncEnabled;
            
            DebugLogger.Log($"[HVAC-PANEL] Settings loaded. IsUpdaterSync={HVACSyncState.IsSyncEnabled}");
            
            // If sync is enabled, start IdlingHandler now via ExternalEvent
            if (HVACSyncState.IsSyncEnabled && !HVACSuperScheme.App._idlingHandlerIsActive)
            {
                DebugLogger.Log("[HVAC-PANEL] Requesting IdlingHandler start via ExternalEvent");
                _toggleIdlingHandler.SetStart(true);
                _toggleIdlingEvent.Raise();
            }
        }
        
        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            OnBackToHub?.Invoke();
        }
        
        /// <summary>
        /// Handle synchronization checkbox state change
        /// </summary>
        private void ChkSyncEnabled_CheckedChanged(object sender, RoutedEventArgs e)
        {
            HVACSyncState.IsSyncEnabled = chkSyncEnabled.IsChecked == true;
            
            // Save to SettingStorage file
            HVACSuperScheme.Commands.Settings.SettingStorage.Instance.IsUpdaterSync = HVACSyncState.IsSyncEnabled;
            HVACSuperScheme.Commands.Settings.SettingStorage.SaveSettings();
            
            DebugLogger.Log($"[HVAC-PANEL] Synchronization {(HVACSyncState.IsSyncEnabled ? "enabled" : "disabled")} and saved to settings");
            
            // Toggle IdlingHandler via ExternalEvent
            if (HVACSyncState.IsSyncEnabled && !HVACSuperScheme.App._idlingHandlerIsActive)
            {
                DebugLogger.Log("[HVAC-PANEL] Requesting IdlingHandler start via ExternalEvent");
                _toggleIdlingHandler.SetStart(true);
                _toggleIdlingEvent.Raise();
            }
            else if (!HVACSyncState.IsSyncEnabled && HVACSuperScheme.App._idlingHandlerIsActive)
            {
                DebugLogger.Log("[HVAC-PANEL] Requesting IdlingHandler stop via ExternalEvent");
                _toggleIdlingHandler.SetStart(false);
                _toggleIdlingEvent.Raise();
            }
        }
        
        private void BtnCreateSchema_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtStatus.Text = "Построение схемы начато...";
                DebugLogger.Log("[HVAC-PANEL] User clicked 'Create Schema' button");
                
                if (_createSchemaEvent != null)
                {
                    _createSchemaEvent.Raise();
                    DebugLogger.Log("[HVAC-PANEL] CreateSchema event raised");
                }
                else
                {
                    txtStatus.Text = "Ошибка: ExternalEvent не инициализирован";
                    MessageBox.Show("ExternalEvent не инициализирован", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Ошибка: {ex.Message}";
                DebugLogger.Log($"[HVAC-PANEL] ERROR in Create Schema: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void BtnCompleteSchema_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtStatus.Text = "Достройка схемы начата...";
                DebugLogger.Log("[HVAC-PANEL] User clicked 'Complete Schema' button");
                
                if (_completeSchemaEvent != null)
                {
                    _completeSchemaEvent.Raise();
                    DebugLogger.Log("[HVAC-PANEL] CompleteSchema event raised");
                }
                else
                {
                    txtStatus.Text = "Ошибка: ExternalEvent не инициализирован";
                    MessageBox.Show("ExternalEvent не инициализирован", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Ошибка: {ex.Message}";
                DebugLogger.Log($"[HVAC-PANEL] ERROR in Complete Schema: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[HVAC-PANEL] User clicked 'Settings' button");
                
                if (_settingsEvent != null)
                {
                    _settingsEvent.Raise();
                    DebugLogger.Log("[HVAC-PANEL] Settings event raised");
                }
                else
                {
                    MessageBox.Show("ExternalEvent не инициализирован", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-PANEL] ERROR in Settings: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Для построения принципиальной схемы:\n" +
                "1. Выберите необходимые элементы в модели\n" +
                "2. Нажмите 'Построить схему'\n\n" +
                "Для достройки схемы:\n" +
                "1. Добавьте новые элементы в модель\n" +
                "2. Нажмите 'Достроить схему'\n\n" +
                "По вопросам: support@annotatix.ai",
                "Помощь", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        private void BtnReference_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://annotatix.ai/docs/hvac");
            }
            catch
            {
                MessageBox.Show("Откройте ссылку: https://annotatix.ai/docs/hvac", 
                    "Справка", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
