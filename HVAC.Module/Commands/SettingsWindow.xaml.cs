using System;
using System.Windows;
using Autodesk.Revit.DB;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using PluginsManager.Core;

namespace HVACSuperScheme.Commands.Settings
{
    /// <summary>
    /// Settings Window for HVAC Module (MVVM-free version)
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private Document _document;
        private bool _oldValueUpdaterSync;

        public SettingsWindow(Document document)
        {
            InitializeComponent();
            _document = document;
            
            // Read current settings
            SettingStorage.ReadSettings();
            _oldValueUpdaterSync = SettingStorage.Instance.IsUpdaterSync;
            
            // Set checkbox state
            UpdaterSyncCheckBox.IsChecked = SettingStorage.Instance.IsUpdaterSync;
            
            DebugLogger.Log($"[HVAC-SETTINGS] Settings window opened, current sync state: {_oldValueUpdaterSync}");
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Update setting value
                bool newSyncValue = UpdaterSyncCheckBox.IsChecked ?? true;
                SettingStorage.Instance.IsUpdaterSync = newSyncValue;
                
                // Save to file
                SettingStorage.SaveSettings();
                
                bool changeSyncStatus = _oldValueUpdaterSync != newSyncValue;
                
                DebugLogger.Log($"[HVAC-SETTINGS] Save clicked: old={_oldValueUpdaterSync}, new={newSyncValue}, changed={changeSyncStatus}");
                
                if (changeSyncStatus)
                {
                    // Remove all triggers
                    LoggingUtils.Logging(Warnings.TriggersDeleted(), _document.PathName);
                    Updater.RemoveAllTriggers();
                    
                    // Recreate triggers
                    App.CreateAdditionTriggers();
                    App.CreateDeletionTriggers();
                    
                    DebugLogger.Log($"[HVAC-SETTINGS] Triggers recreated");
                    
                    if (newSyncValue)
                    {
                        // Enable sync - subscribe to Idling event
                        if (!App._idlingHandlerIsActive)
                        {
                            LoggingUtils.Logging(Warnings.SubscribeToIdlingComplete(), _document.PathName);
                            App.CreateIdlingHandler();
                            DebugLogger.Log($"[HVAC-SETTINGS] Idling handler activated");
                        }
                    }
                    else
                    {
                        // Disable sync - unsubscribe from Idling event
                        if (App._idlingHandlerIsActive)
                        {
                            LoggingUtils.Logging(Warnings.UnsubscribeToIdlingComplete(), _document.PathName);
                            App.RemoveIdlingHandler();
                            DebugLogger.Log($"[HVAC-SETTINGS] Idling handler deactivated");
                        }
                    }
                }
                
                DebugLogger.Log($"[HVAC-SETTINGS] Settings saved successfully");
                this.DialogResult = true;
                this.Close();
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-SETTINGS] ERROR saving settings: {ex.Message}");
                DebugLogger.Log($"[HVAC-SETTINGS] Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Ошибка при сохранении настроек:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DebugLogger.Log($"[HVAC-SETTINGS] Cancel clicked, settings not saved");
            this.DialogResult = false;
            this.Close();
        }
    }
}
