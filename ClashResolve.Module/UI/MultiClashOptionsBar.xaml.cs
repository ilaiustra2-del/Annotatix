using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using PluginsManager.Core;

namespace ClashResolve.Module.UI
{
    /// <summary>
    /// Multi-step Options Bar for "Множественное исправление коллизий".
    /// Three steps: (1) select pipes A, (2) select pipes B, (3) configure params and execute.
    /// </summary>
    public partial class MultiClashOptionsBar : UserControl
    {
        // ----------------------------------------------------------------
        // Events — subscribed via reflection from MultiClashSession
        // ----------------------------------------------------------------

        /// <summary>User confirmed selection of pipes A (step 1).</summary>
        public event Action<List<Autodesk.Revit.DB.ElementId>> Step1Done;

        /// <summary>User confirmed selection of pipes B (step 2).</summary>
        public event Action<List<Autodesk.Revit.DB.ElementId>> Step2Done;

        /// <summary>User confirmed parameters (step 3).</summary>
        public event Action<MultiClashOptionsBarParams> Step3Done;

        /// <summary>User clicked Отмена on any step.</summary>
        public event Action CancelRequested;

        /// <summary>Raised when angle or auto-checkbox changes so session can refresh preview.</summary>
        public event Action AutoRecalcRequested;

        // ----------------------------------------------------------------
        // State
        // ----------------------------------------------------------------
        public int CurrentStep { get; private set; } = 1;

        // Ids held by the bar for current-step selection tracking
        private List<Autodesk.Revit.DB.ElementId> _currentIds = new List<Autodesk.Revit.DB.ElementId>();

        // ----------------------------------------------------------------
        // Constructor
        // ----------------------------------------------------------------
        public MultiClashOptionsBar()
        {
            InitializeComponent();
            // Wire checkboxes after InitializeComponent so named elements exist
            chkAutoClearance.Checked   += ChkAuto_Changed;
            chkAutoClearance.Unchecked += ChkAuto_Changed;
            chkAutoLength.Checked      += ChkAuto_Changed;
            chkAutoLength.Unchecked    += ChkAuto_Changed;
            ApplyAutoState();
        }

        // ----------------------------------------------------------------
        // Public API — called by MultiClashSession
        // ----------------------------------------------------------------

        /// <summary>Advance to the specified step (2 or 3), resetting the counter.</summary>
        public void AdvanceToStep(int step)
        {
            // Always called from the UI thread (via session OnStep1Done/OnStep2Done button-click chain)
            // Do NOT use Dispatcher.Invoke — it causes nested re-entrant deadlock
            CurrentStep = step;
            panelStep1.Visibility = step == 1 ? Visibility.Visible : Visibility.Collapsed;
            panelStep2.Visibility = step == 2 ? Visibility.Visible : Visibility.Collapsed;
            panelStep3.Visibility = step == 3 ? Visibility.Visible : Visibility.Collapsed;

            if (step == 2)
                txtCount2.Text = "Выбрано: 0";

            DebugLogger.Log($"[MULTI-BAR] Advanced to step {step}");
        }

        /// <summary>Update the selection counter label shown in step 1 or 2.</summary>
        public void UpdateSelectionCount(int count)
        {
            Dispatcher.Invoke(() =>
            {
                if (CurrentStep == 1)
                    txtCount1.Text = $"Выбрано: {count}";
                else if (CurrentStep == 2)
                    txtCount2.Text = $"Выбрано: {count}";
            });
        }

        /// <summary>
        /// Programmatically trigger "Готово" for the current step (called by Enter key).
        /// </summary>
        public void TriggerDone()
        {
            switch (CurrentStep)
            {
                case 1: FireStep1Done(); break;
                case 2: FireStep2Done(); break;
                case 3: FireStep3Done(); break;
            }
        }

        /// <summary>
        /// Update auto-preview fields when pipes are known (step 3 only).
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
                    double halfMm = Math.Ceiling(rMaxMm * (isAngled ? 3.0 : 2.5));
                    txtSegmentLength.Text = ((int)(halfMm * 2)).ToString();
                }
            });
        }

        // ----------------------------------------------------------------
        // Button handlers
        // ----------------------------------------------------------------
        private void BtnDone1_Click(object sender, RoutedEventArgs e) => FireStep1Done();
        private void BtnDone2_Click(object sender, RoutedEventArgs e) => FireStep2Done();
        private void BtnDone3_Click(object sender, RoutedEventArgs e) => FireStep3Done();
        private void BtnCancel_Click(object sender, RoutedEventArgs e) => CancelRequested?.Invoke();

        // ----------------------------------------------------------------
        // ComboBox / checkbox handlers
        // ----------------------------------------------------------------
        private void CmbAngle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (chkAutoClearance == null) return;
            AutoRecalcRequested?.Invoke();
        }

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

        private bool IsAngle45() => false; // kept for safety — use GetAngleDegrees()

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
        // Internal fire helpers
        // ----------------------------------------------------------------
        private void FireStep1Done()
        {
            Step1Done?.Invoke(new List<Autodesk.Revit.DB.ElementId>(_currentIds));
        }

        private void FireStep2Done()
        {
            if (_currentIds.Count == 0)
            {
                MessageBox.Show("Выберите хотя бы одну трубу B (препятствие).",
                    "Ошибка выбора", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Step2Done?.Invoke(new List<Autodesk.Revit.DB.ElementId>(_currentIds));
        }

        private void FireStep3Done()
        {
            bool autoClearance = chkAutoClearance.IsChecked == true;
            bool autoLength    = chkAutoLength.IsChecked == true;
            double angleDeg    = GetAngleDegrees();
            bool bypassUp      = IsBypassUp();

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

            Step3Done?.Invoke(new MultiClashOptionsBarParams
            {
                AngleDegrees   = angleDeg,
                ClearanceMm    = clearanceMm,
                AutoClearance  = autoClearance,
                HalfLengthMm   = segmentLengthMm / 2.0,
                AutoHalfLength = autoLength,
                BypassUp       = bypassUp
            });
        }

        // ----------------------------------------------------------------
        // Selection state — updated by MultiClashSession
        // ----------------------------------------------------------------

        /// <summary>Replace the current selection list (steps 1 and 2).</summary>
        public void SetCurrentIds(IEnumerable<Autodesk.Revit.DB.ElementId> ids)
        {
            _currentIds = new List<Autodesk.Revit.DB.ElementId>(ids);
        }
    }

    /// <summary>Parameters collected from step 3 of the multi-clash options bar.</summary>
    public class MultiClashOptionsBarParams
    {
        public double AngleDegrees   { get; set; } = 90.0;
        public double ClearanceMm    { get; set; }
        public bool   AutoClearance  { get; set; }
        public double HalfLengthMm   { get; set; }
        public bool   AutoHalfLength { get; set; }
        /// <summary>When true, pipe A bypasses pipe B from above.</summary>
        public bool   BypassUp       { get; set; }
    }
}
