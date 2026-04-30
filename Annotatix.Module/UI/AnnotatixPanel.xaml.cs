using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using Annotatix.Module.Core;
using PluginsManager.Core;

namespace Annotatix.Module.UI
{
    /// <summary>
    /// Interaction logic for AnnotatixPanel.xaml
    /// </summary>
    public partial class AnnotatixPanel : UserControl
    {
        private UIApplication _uiApp;
        private AnnotatixPanelHandler _handler;
        private ExternalEvent _externalEvent;
        private bool _isUpdating = false;
        private bool _externalEventReady = false;

        public AnnotatixPanel(UIApplication uiApp, AnnotatixPanelHandler handler = null, ExternalEvent extEvent = null)
        {
            InitializeComponent();

            _uiApp = uiApp;

            // Use provided ExternalEvent from hub (created in API context) or fall back
            if (handler != null && extEvent != null)
            {
                _handler = handler;
                _handler.Panel = this;
                _externalEvent = extEvent;
                _externalEventReady = true;
                System.Diagnostics.Debug.WriteLine("[ANNOTATIX-PANEL] Using provided ExternalEvent");
            }
            else
            {
                try
                {
                    _handler = new AnnotatixPanelHandler { Panel = this };
                    _externalEvent = ExternalEvent.Create(_handler);
                    _externalEventReady = true;
                    System.Diagnostics.Debug.WriteLine("[ANNOTATIX-PANEL] Created new ExternalEvent");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ANNOTATIX-PANEL] Failed to create ExternalEvent: {ex.Message}");
                    _handler = null;
                    _externalEvent = null;
                    _externalEventReady = false;
                }
            }

            // Load current settings into UI
            LoadSettings();
        }

        // ── Lazy ExternalEvent creation ──

        /// <summary>
        /// Ensures ExternalEvent is created. Must be called from
        /// within Revit API context OR on first button press (will
        /// try to create and show error if not possible).
        /// </summary>
        private bool EnsureExternalEvent()
        {
            if (_externalEventReady && _externalEvent != null)
                return true;

            try
            {
                _externalEvent = ExternalEvent.Create(_handler);
                _externalEventReady = true;
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PANEL] Failed to create ExternalEvent: {ex.Message}");
                SetStatus($"Ошибка: не удалось создать контекст Revit — {ex.Message}");
                return false;
            }
        }

        // ── Settings UI ──

        private void LoadSettings()
        {
            _isUpdating = true;

            // Grid step
            string[] gridValues = { "1", "2", "3", "4", "5", "6", "8", "10", "12", "15", "20" };
            foreach (var v in gridValues)
                cmbGridStep.Items.Add(v);
            cmbGridStep.Text = AnnotatixSettings.GridStepMm.ToString("F0");

            // Occupancy threshold
            string[] thresholdValues = { "5", "10", "15", "20", "25", "30", "40", "50" };
            foreach (var v in thresholdValues)
                cmbOccupancyThreshold.Items.Add(v);
            cmbOccupancyThreshold.Text = (AnnotatixSettings.OccupancyThreshold * 100).ToString("F0");

            // Edge margin
            string[] marginValues = { "0", "5", "10", "15", "20", "25", "30", "40", "50" };
            foreach (var v in marginValues)
                cmbEdgeMargin.Items.Add(v);
            cmbEdgeMargin.Text = AnnotatixSettings.EdgeMarginMm.ToString("F0");

            _isUpdating = false;
        }

        private void UpdateGridStep()
        {
            if (_isUpdating) return;
            if (double.TryParse(cmbGridStep.Text, out double val) && val > 0 && val <= 100)
            {
                AnnotatixSettings.GridStepMm = val;
                AnnotatixSettings.Save();
                SetStatus($"Шаг сетки изменён: {val} мм");
                DebugLogger.Log($"[ANNOTATIX-PANEL] Grid step set to {val} mm");
            }
            else
            {
                cmbGridStep.Text = AnnotatixSettings.GridStepMm.ToString("F0");
                SetStatus("Некорректное значение шага сетки. Используется предыдущее значение.");
            }
        }

        private void UpdateOccupancyThreshold()
        {
            if (_isUpdating) return;
            if (double.TryParse(cmbOccupancyThreshold.Text, out double val) && val > 0 && val <= 100)
            {
                AnnotatixSettings.OccupancyThreshold = val / 100.0;
                AnnotatixSettings.Save();
                SetStatus($"Порог занятости изменён: {val}%");
                DebugLogger.Log($"[ANNOTATIX-PANEL] Occupancy threshold set to {val}%");
            }
            else
            {
                cmbOccupancyThreshold.Text = (AnnotatixSettings.OccupancyThreshold * 100).ToString("F0");
                SetStatus("Некорректное значение порога. Используется предыдущее значение.");
            }
        }

        private void UpdateEdgeMargin()
        {
            if (_isUpdating) return;
            if (double.TryParse(cmbEdgeMargin.Text, out double val) && val >= 0 && val <= 200)
            {
                AnnotatixSettings.EdgeMarginMm = val;
                AnnotatixSettings.Save();
                SetStatus($"Отступ от краев изменён: {val} мм");
                DebugLogger.Log($"[ANNOTATIX-PANEL] Edge margin set to {val} mm");
            }
            else
            {
                cmbEdgeMargin.Text = AnnotatixSettings.EdgeMarginMm.ToString("F0");
                SetStatus("Некорректное значение отступа. Используется предыдущее значение.");
            }
        }

        private void CmbGridStep_LostFocus(object sender, RoutedEventArgs e) => UpdateGridStep();
        private void CmbGridStep_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateGridStep();

        private void CmbOccupancyThreshold_LostFocus(object sender, RoutedEventArgs e) => UpdateOccupancyThreshold();
        private void CmbOccupancyThreshold_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateOccupancyThreshold();

        private void CmbEdgeMargin_LostFocus(object sender, RoutedEventArgs e) => UpdateEdgeMargin();
        private void CmbEdgeMargin_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateEdgeMargin();

        // ── Command Buttons ──

        private void BtnRasterExport_Click(object sender, RoutedEventArgs e)
        {
            if (_uiApp?.ActiveUIDocument == null)
            {
                MessageBox.Show("Нет активного документа Revit.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (_externalEvent == null)
            {
                MessageBox.Show("Не удалось создать контекст выполнения Revit.\nПерезапустите Revit и попробуйте снова.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            btnRasterExport.IsEnabled = false;
            SetProgress("Запуск растрового экспорта и анализа...");

            _handler.CommandName = "RasterExport";
            _externalEvent.Raise();

            var timer = new System.Windows.Forms.Timer { Interval = 2000 };
            timer.Tick += (s, args) =>
            {
                btnRasterExport.IsEnabled = true;
                timer.Stop();
                timer.Dispose();
            };
            timer.Start();
        }

        private void BtnAnnotateDuctwork_Click(object sender, RoutedEventArgs e)
        {
            if (_uiApp?.ActiveUIDocument == null)
            {
                MessageBox.Show("Нет активного документа Revit.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (_externalEvent == null)
            {
                MessageBox.Show("Не удалось создать контекст выполнения Revit.\nПерезапустите Revit и попробуйте снова.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            btnAnnotateDuctwork.IsEnabled = false;
            SetProgress("Запуск авто-разметки воздуховодов...");

            _handler.CommandName = "AnnotateDuctwork";
            _externalEvent.Raise();

            var timer = new System.Windows.Forms.Timer { Interval = 3000 };
            timer.Tick += (s, args) =>
            {
                btnAnnotateDuctwork.IsEnabled = true;
                timer.Stop();
                timer.Dispose();
            };
            timer.Start();
        }

        // ── Status helpers ──

        public void SetProgress(string text)
        {
            Dispatcher.Invoke(() =>
            {
                txtProgress.Text = text;
                txtProgress.Visibility = Visibility.Visible;
                txtStatus.Visibility = Visibility.Collapsed;
            });
        }

        public void SetStatus(string text)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = text;
                txtStatus.Visibility = Visibility.Visible;
                txtProgress.Visibility = Visibility.Collapsed;
            });
        }
    }
}
