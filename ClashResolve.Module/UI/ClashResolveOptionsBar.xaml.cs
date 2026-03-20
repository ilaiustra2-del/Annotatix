using System;
using System.Windows;
using System.Windows.Controls;
using MediaColor = System.Windows.Media.Color;
using Autodesk.Revit.UI;
using PluginsManager.Core;
using ClashResolve.Module.Core;

namespace ClashResolve.Module.UI
{
    /// <summary>
    /// Options Bar control for ClashResolve ribbon mode.
    /// Displayed under the Revit ribbon via AdWindows ComponentManager.OptionsBar.
    /// </summary>
    public partial class ClashResolveOptionsBar : UserControl
    {
        // ----------------------------------------------------------------
        // Events
        // ----------------------------------------------------------------

        /// <summary>Raised when "Перестроить" is clicked with valid parameters.</summary>
        public event Action<ClashResolveOptionsBarParams> ResolveRequested;

        /// <summary>Raised when "Выйти" is clicked.</summary>
        public event Action ExitRequested;

        // ----------------------------------------------------------------
        // Constructor
        // ----------------------------------------------------------------
        public ClashResolveOptionsBar()
        {
            InitializeComponent();
            // Subscribe after InitializeComponent so all named elements exist
            chkAutoClearance.Checked   += ChkAuto_Changed;
            chkAutoClearance.Unchecked += ChkAuto_Changed;
            chkAutoLength.Checked      += ChkAuto_Changed;
            chkAutoLength.Unchecked    += ChkAuto_Changed;
            ApplyAutoState();

            // Restore UseTable state from persisted service
            chkUseTable.IsChecked = ClashLookupService.Instance.GlobalEnabled;
        }

        // ----------------------------------------------------------------
        // Public helpers — called by ClashResolveSession
        // ----------------------------------------------------------------

        /// <summary>Update the hint label with current selection state.</summary>
        public void SetHint(string text)
        {
            Dispatcher.Invoke(() => txtSelectionHint.Text = text);
        }

        /// <summary>
        /// Update auto fields preview when pipes are known.
        /// rMaxMm = max(outerRadiusA, outerRadiusB) in mm.
        /// </summary>
        public void UpdateAutoDisplay(double rMaxMm)
        {
            Dispatcher.Invoke(() =>
            {
                bool isAngled = Math.Abs(GetAngleDegrees() - 90.0) > 0.5;

                if (chkAutoClearance.IsChecked == true)
                {
                    double clearanceMm = isAngled ? 50.0 : Math.Ceiling(rMaxMm * 3.5);
                    txtClearance.Text = ((int)clearanceMm).ToString();
                }

                if (chkAutoLength.IsChecked == true)
                {
                    // full segment = 2 × halfLength; halfLength = rMax * coeff
                    double halfMm = Math.Ceiling(rMaxMm * (isAngled ? 3.0 : 2.5));
                    txtSegmentLength.Text = ((int)(halfMm * 2)).ToString();
                }
            });
        }

        // ----------------------------------------------------------------
        // Private helpers
        // ----------------------------------------------------------------
        private double GetAngleDegrees()
        {
            var item = cmbAngle.SelectedItem as ComboBoxItem;
            string s = item?.Content?.ToString() ?? "90°";
            if (double.TryParse(s.TrimEnd('°'), out double deg))
                return deg;
            return 90.0;
        }

        private bool IsBypassUp()
        {
            var item = cmbDirection.SelectedItem as ComboBoxItem;
            return item?.Content?.ToString() == "Сверху";
        }

        // ----------------------------------------------------------------
        // Event handlers
        // ----------------------------------------------------------------
        private void CmbAngle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Will be called during InitializeComponent too — guard null
            if (chkAutoClearance == null) return;
            // Notify session to recalculate if pipes are known
            AutoRecalcRequested?.Invoke();
        }

        /// <summary>Raised when angle or auto-checkbox changes so session can refresh preview.</summary>
        public event Action AutoRecalcRequested;

        private void ChkAuto_Changed(object sender, RoutedEventArgs e)
        {
            ApplyAutoState();
            AutoRecalcRequested?.Invoke();
        }

        private void ApplyAutoState()
        {
            if (txtClearance == null || txtSegmentLength == null) return;

            bool isAutoClear = chkAutoClearance?.IsChecked == true;
            txtClearance.IsEnabled = !isAutoClear;
            txtClearance.Opacity   = isAutoClear ? 0.5 : 1.0;

            bool isAutoLen = chkAutoLength?.IsChecked == true;
            txtSegmentLength.IsEnabled = !isAutoLen;
            txtSegmentLength.Opacity   = isAutoLen ? 0.5 : 1.0;
        }

        private void BtnResolve_Click(object sender, RoutedEventArgs e)
        {
            TriggerResolve();
        }

        /// <summary>Programmatically trigger the Resolve action (e.g. from Enter key).</summary>
        public void TriggerResolve()
        {
            bool autoClearance = chkAutoClearance.IsChecked == true;
            bool autoLength    = chkAutoLength.IsChecked == true;
            double angleDeg    = GetAngleDegrees();
            bool bypassUp      = IsBypassUp();
            DebugLogger.Log($"[OPTIONS-BAR] TriggerResolve: angle={angleDeg:F0}° bypassUp={bypassUp} selectedItem={cmbAngle.SelectedItem?.GetType().Name} content={(cmbAngle.SelectedItem as ComboBoxItem)?.Content}");

            double clearanceMm = 0;
            if (!autoClearance)
            {
                if (!double.TryParse(txtClearance.Text.Trim(), out clearanceMm) || clearanceMm < 0)
                {
                    MessageBox.Show("Введите корректное значение величины опуска (мм).",
                        "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            double segmentLengthMm = 0;
            if (!autoLength)
            {
                if (!double.TryParse(txtSegmentLength.Text.Trim(), out segmentLengthMm) || segmentLengthMm <= 0)
                {
                    MessageBox.Show("Введите корректное значение длины сегмента (мм).",
                        "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            ResolveRequested?.Invoke(new ClashResolveOptionsBarParams
            {
                AngleDegrees   = angleDeg,
                ClearanceMm    = clearanceMm,
                AutoClearance  = autoClearance,
                HalfLengthMm   = segmentLengthMm / 2.0,  // halved before passing to ClashPair
                AutoHalfLength = autoLength,
                BypassUp       = bypassUp,
                UseTable       = chkUseTable.IsChecked == true
            });
        }
        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            ExitRequested?.Invoke();
        }

        private void BtnLookupTable_Click(object sender, RoutedEventArgs e)
        {
            var win = ClashLookupWindow.GetOrCreate();
            win.Show();
            win.Activate();
        }

        private void ChkUseTable_Changed(object sender, RoutedEventArgs e)
        {
            bool enabled = chkUseTable.IsChecked == true;
            // Persist globally so MultiClash bar and resolver pick it up too
            ClashLookupService.Instance.GlobalEnabled = enabled;
        }
    }

    /// <summary>Parameters collected from the options bar.</summary>
    public class ClashResolveOptionsBarParams
    {
        public double AngleDegrees   { get; set; } = 90.0;
        public double ClearanceMm    { get; set; }
        public bool   AutoClearance  { get; set; }
        public double HalfLengthMm   { get; set; }
        public bool   AutoHalfLength { get; set; }
        /// <summary>When true, pipe A bypasses pipe B from above.</summary>
        public bool   BypassUp       { get; set; }
        /// <summary>When true, use lookup tables instead of formula-based auto calculation.</summary>
        public bool   UseTable       { get; set; }
    }
}
