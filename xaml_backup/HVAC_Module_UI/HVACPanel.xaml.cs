using System;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace HVAC.Module.UI
{
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

        public HVACPanel(UIApplication uiApp, ExternalEvent createSchemaEvent, ExternalEvent completeSchemaEvent, ExternalEvent settingsEvent)
        {
            InitializeComponent();
            
            _uiApp = uiApp;
            _createSchemaEvent = createSchemaEvent;
            _completeSchemaEvent = completeSchemaEvent;
            _settingsEvent = settingsEvent;
            
            // Debug logging
            DebugLogger.Log("[HVAC-PANEL] *** PANEL INSTANCE CREATED ***");
            DebugLogger.Log("[HVAC-PANEL] Panel created after authentication");
            DebugLogger.Log($"[HVAC-PANEL] ExternalEvents received: createSchema={_createSchemaEvent != null}, completeSchema={_completeSchemaEvent != null}, settings={_settingsEvent != null}");
        }
        
        private void BtnCreateSchema_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[HVAC-PANEL] User clicked 'Create Schema' button");
                
                if (_createSchemaEvent != null)
                {
                    _createSchemaEvent.Raise();
                    DebugLogger.Log("[HVAC-PANEL] CreateSchema event raised");
                }
                else
                {
                    MessageBox.Show("ExternalEvent не инициализирован", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-PANEL] ERROR in Create Schema: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void BtnCompleteSchema_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[HVAC-PANEL] User clicked 'Complete Schema' button");
                
                if (_completeSchemaEvent != null)
                {
                    _completeSchemaEvent.Raise();
                    DebugLogger.Log("[HVAC-PANEL] CompleteSchema event raised");
                }
                else
                {
                    MessageBox.Show("ExternalEvent не инициализирован", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
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
    }
}
