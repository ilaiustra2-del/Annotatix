using System;
using System.Windows;
using System.Windows.Controls;
using PluginsManager.Core;

namespace Tracer.Module.UI
{
    /// <summary>
    /// Options Bar control for Tracer ribbon mode.
    /// Displayed under the Revit ribbon via AdWindows ComponentManager.OptionsBar.
    /// </summary>
    public partial class TracerOptionsBar : UserControl
    {
        // ----------------------------------------------------------------
        // Events
        // ----------------------------------------------------------------
        public event Action DoneRequested;
        public event Action CancelRequested;
        public event Action<double> SlopeChanged;
        public event Action<bool> UseMainLineSlopeChanged;

        // Current stage (1=main line, 2=risers, 3=slope)
        private int _currentStage = 1;
        // Whether we have valid selection for current stage
        private bool _hasValidSelection = false;
        
        /// <summary>
        /// Gets whether to add fittings between riser and pipe
        /// </summary>
        public bool AddFittings => chkAddFittings.IsChecked == true;

        // ----------------------------------------------------------------
        // Constructor
        // ----------------------------------------------------------------
        public TracerOptionsBar()
        {
            InitializeComponent();
        }

        // ----------------------------------------------------------------
        // Public API called by TracerSession
        // ----------------------------------------------------------------
        public void SetStage(int stage)
        {
            Dispatcher.Invoke(() =>
            {
                _currentStage = stage;
                _hasValidSelection = false;

                // Hide all stage-specific UI
                txtStage1.Visibility = Visibility.Collapsed;
                txtStage2.Visibility = Visibility.Collapsed;
                txtRiserCount.Visibility = Visibility.Collapsed;
                panelStage3.Visibility = Visibility.Collapsed;

                // Show appropriate UI for current stage
                switch (stage)
                {
                    case 1:
                        txtStage1.Visibility = Visibility.Visible;
                        btnDone.IsEnabled = false;
                        btnDone.Content = "Далее";
                        break;
                    case 2:
                        txtStage2.Visibility = Visibility.Visible;
                        txtRiserCount.Visibility = Visibility.Visible;
                        btnDone.IsEnabled = false;
                        btnDone.Content = "Далее";
                        break;
                    case 3:
                        panelStage3.Visibility = Visibility.Visible;
                        btnDone.IsEnabled = true;
                        btnDone.Content = "Готово";
                        break;
                }
            });
        }

        public void SetConnectionType(int connectionType)
        {
            Dispatcher.Invoke(() =>
            {
                string typeText = connectionType switch
                {
                    0 => "| Присоединение под 45°",
                    1 => "| L-образное присоединение",
                    2 => "| Присоединение снизу",
                    3 => "| Z-образное присоединение",
                    _ => "| Присоединение"
                };
                txtConnectionType.Text = typeText;
            });
        }

        public void SetRiserCount(int count)
        {
            Dispatcher.Invoke(() =>
            {
                txtRiserCount.Text = $"Выбрано: {count}";
                // Enable Done button in stage 2 if at least one riser selected
                if (_currentStage == 2)
                {
                    _hasValidSelection = count > 0;
                    btnDone.IsEnabled = _hasValidSelection;
                }
            });
        }

        /// <summary>
        /// Called by TracerSession when main line is selected
        /// </summary>
        public void SetMainLineSelected(bool selected)
        {
            Dispatcher.Invoke(() =>
            {
                if (_currentStage == 1)
                {
                    _hasValidSelection = selected;
                    btnDone.IsEnabled = selected;
                }
            });
        }

        public void SetSlope(double slope, bool useMainLineSlope)
        {
            Dispatcher.Invoke(() =>
            {
                txtSlope.Text = slope.ToString("F1");
                chkUseMainLineSlope.IsChecked = useMainLineSlope;
                UpdateSlopeInputState();
            });
        }

        /// <summary>
        /// Programmatically trigger the Done action (e.g. from Enter key).
        /// </summary>
        public void TriggerDone()
        {
            if (btnDone.IsEnabled)
            {
                TriggerDoneInternal();
            }
        }

        // ----------------------------------------------------------------
        // Private helpers
        // ----------------------------------------------------------------
        private void UpdateSlopeInputState()
        {
            bool useMainLine = chkUseMainLineSlope.IsChecked == true;
            txtSlope.IsEnabled = !useMainLine;
            txtSlope.Opacity = useMainLine ? 0.5 : 1.0;
        }

        private void TriggerDoneInternal()
        {
            // Stage 3: Validate slope if not using main line slope
            if (_currentStage == 3 && chkUseMainLineSlope.IsChecked != true)
            {
                if (!double.TryParse(txtSlope.Text.Trim().Replace(',', '.'), 
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, 
                    out double slope))
                {
                    MessageBox.Show("Введите корректное значение уклона (%).",
                        "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            DoneRequested?.Invoke();
        }

        // ----------------------------------------------------------------
        // Event handlers
        // ----------------------------------------------------------------
        private void BtnDone_Click(object sender, RoutedEventArgs e)
        {
            TriggerDoneInternal();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            CancelRequested?.Invoke();
        }

        private void TxtSlope_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (double.TryParse(txtSlope.Text.Trim().Replace(',', '.'),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double slope))
            {
                SlopeChanged?.Invoke(slope);
            }
        }

        private void ChkUseMainLineSlope_Changed(object sender, RoutedEventArgs e)
        {
            UpdateSlopeInputState();
            UseMainLineSlopeChanged?.Invoke(chkUseMainLineSlope.IsChecked == true);
        }
    }
}
