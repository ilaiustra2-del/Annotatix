using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClashResolve.Module.Core;
using PluginsManager.Core;
using MediaColor = System.Windows.Media.Color;
using WpfVisibility = System.Windows.Visibility;

namespace ClashResolve.Module.UI
{
    public partial class ClashResolvePanel : UserControl
    {
        // ----------------------------------------------------------------
        // Fields
        // ----------------------------------------------------------------
        private readonly UIApplication _uiApp;

        private readonly ExternalEvent _clashEvent;
        private readonly IClashExecuteProxy _clashHandler;

        private readonly ExternalEvent _pickEventA;
        private readonly IPickElementProxy _pickHandlerA;

        private readonly ExternalEvent _pickEventB;
        private readonly IPickElementProxy _pickHandlerB;

        private ElementId _pipeAId;
        private ElementId _pipeBId;

        // ----------------------------------------------------------------
        // Constructor
        // ----------------------------------------------------------------
        public ClashResolvePanel(
            UIApplication uiApp,
            ExternalEvent clashEvent, IClashExecuteProxy clashHandler,
            ExternalEvent pickEventA, IPickElementProxy pickHandlerA,
            ExternalEvent pickEventB, IPickElementProxy pickHandlerB)
        {
            _uiApp = uiApp;
            _clashEvent = clashEvent;
            _clashHandler = clashHandler;
            _pickEventA = pickEventA;
            _pickHandlerA = pickHandlerA;
            _pickEventB = pickEventB;
            _pickHandlerB = pickHandlerB;

            InitializeComponent();

            // Wire up pick callbacks
            _pickHandlerA.OnPicked = (id, name) => Dispatcher.Invoke(() => OnPickedA(id, name));
            _pickHandlerB.OnPicked = (id, name) => Dispatcher.Invoke(() => OnPickedB(id, name));
        }

        // ----------------------------------------------------------------
        // Pick A
        // ----------------------------------------------------------------
        private void BtnPickA_Click(object sender, RoutedEventArgs e)
        {
            if (_pickEventA == null)
            {
                ShowStatus("ExternalEvent не инициализирован.");
                return;
            }

            txtPipeA.Text = "Выберите элемент в Revit...";
            txtPipeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(0, 120, 212));
            btnPickA.IsEnabled = false;
            ShowStatus("Выберите трубу A в Revit...");

            _pickHandlerA.PromptMessage = "Выберите трубу/воздуховод A (будет обходить)";
            _pickEventA.Raise();
        }

        private void OnPickedA(ElementId id, string name)
        {
            btnPickA.IsEnabled = true;
            if (id != null)
            {
                _pipeAId = id;
                txtPipeA.Text = name;
                txtPipeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(45, 102, 45));
                ShowStatus("Труба A выбрана.");
            }
            else
            {
                _pipeAId = null;
                txtPipeA.Text = "Выбор отменён";
                txtPipeA.Foreground = new SolidColorBrush(MediaColor.FromRgb(136, 136, 136));
                ShowStatus("Выбор трубы A отменён.");
            }
            UpdateResolveButton();
            TryUpdateAutoDisplay();
        }

        // ----------------------------------------------------------------
        // Pick B
        // ----------------------------------------------------------------
        private void BtnPickB_Click(object sender, RoutedEventArgs e)
        {
            if (_pickEventB == null)
            {
                ShowStatus("ExternalEvent не инициализирован.");
                return;
            }

            txtPipeB.Text = "Выберите элемент в Revit...";
            txtPipeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(0, 120, 212));
            btnPickB.IsEnabled = false;
            ShowStatus("Выберите трубу B в Revit...");

            _pickHandlerB.PromptMessage = "Выберите трубу/воздуховод B (препятствие, остаётся неподвижным)";
            _pickEventB.Raise();
        }

        private void OnPickedB(ElementId id, string name)
        {
            btnPickB.IsEnabled = true;
            if (id != null)
            {
                _pipeBId = id;
                txtPipeB.Text = name;
                txtPipeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(45, 102, 45));
                ShowStatus("Труба B выбрана.");
            }
            else
            {
                _pipeBId = null;
                txtPipeB.Text = "Выбор отменён";
                txtPipeB.Foreground = new SolidColorBrush(MediaColor.FromRgb(136, 136, 136));
                ShowStatus("Выбор трубы B отменён.");
            }
            UpdateResolveButton();
            TryUpdateAutoDisplay();
        }

        // ----------------------------------------------------------------
        // Execute clash resolve
        // ----------------------------------------------------------------
        private void BtnResolve_Click(object sender, RoutedEventArgs e)
        {
            if (_pipeAId == null || _pipeBId == null)
            {
                ShowStatus("Выберите обе трубы перед выполнением.");
                return;
            }

            if (_pipeAId == _pipeBId)
            {
                ShowStatus("Труба A и труба B не могут быть одним элементом.");
                return;
            }

            bool autoClearance  = chkAutoClearance.IsChecked == true;
            bool autoHalfLength = chkAutoHalfLength.IsChecked == true;
            double angleDeg     = rbAngle30.IsChecked == true ? 30.0
                                : rbAngle45.IsChecked == true ? 45.0
                                : rbAngle60.IsChecked == true ? 60.0
                                : 90.0;
            double clearanceMm  = 0;

            if (!autoClearance)
            {
                if (!double.TryParse(txtClearance.Text.Trim(), out clearanceMm) || clearanceMm < 0)
                {
                    ShowStatus("Введите корректное значение отступа (мм).");
                    return;
                }
            }

            double halfLengthMm = 0;
            if (!autoHalfLength)
            {
                if (!double.TryParse(txtHalfLength.Text.Trim(), out double segmentLengthMm) || segmentLengthMm <= 0)
                {
                    ShowStatus("Введите корректное значение длины сегмента (мм).");
                    return;
                }
                halfLengthMm = segmentLengthMm / 2.0; // full segment → half-length
            }

            btnResolve.IsEnabled = false;
            btnPickA.IsEnabled = false;
            btnPickB.IsEnabled = false;
            grpDiagnostics.Visibility = WpfVisibility.Collapsed;
            ShowStatus("Выполняется обход...");

            // Capture values for closure
            var pipeAId = _pipeAId;
            var pipeBId = _pipeBId;

            _clashHandler.PendingAction = (app) =>
            {
                var doc = app.ActiveUIDocument.Document;
                var resolver = new ClashResolver();
                var result = resolver.ResolveClash(doc, new ClashPair
                {
                    PipeAId        = pipeAId,
                    PipeBId        = pipeBId,
                    ClearanceMm    = clearanceMm,
                    HalfLengthMm   = halfLengthMm,
                    AutoClearance  = autoClearance,
                    AutoHalfLength = autoHalfLength,
                    AngleDegrees   = angleDeg
                });

                DebugLogger.Log($"[CLASH-PANEL] Result: Success={result.Success}, Msg={result.Message}");
                Dispatcher.Invoke(() => OnClashCompleted(result));
            };

            _clashEvent.Raise();
        }

        private void OnClashCompleted(ClashResolveResult result)
        {
            btnPickA.IsEnabled = true;
            btnPickB.IsEnabled = true;
            UpdateResolveButton();

            // Show computed auto values in the fields
            if (chkAutoClearance.IsChecked == true && result.UsedClearanceMm > 0)
                txtClearance.Text = ((int)result.UsedClearanceMm).ToString();
            if (chkAutoHalfLength.IsChecked == true && result.UsedHalfLengthMm > 0)
                txtHalfLength.Text = ((int)(result.UsedHalfLengthMm * 2)).ToString(); // display full segment length

            ShowStatus(result.Success
                ? "✓ " + result.Message
                : "✗ " + result.Message);

            ShowDiagnostics(result);

            DebugLogger.Log($"[CLASH-PANEL] Clash resolve completed: {result.Success}, {result.Message}");
        }

        private void ShowDiagnostics(ClashResolveResult result)
        {
            var sb = new System.Text.StringBuilder();

            if (result.PipeAId != null && result.PipeAStart != null)
            {
                sb.AppendLine($"Труба A  ID: {result.PipeAId.Value}");
                sb.AppendLine($"  Начало:  X={result.PipeAStart.X * 304.8:F1}  Y={result.PipeAStart.Y * 304.8:F1}  Z={result.PipeAStart.Z * 304.8:F1} мм");
                sb.AppendLine($"  Конец:   X={result.PipeAEnd.X * 304.8:F1}  Y={result.PipeAEnd.Y * 304.8:F1}  Z={result.PipeAEnd.Z * 304.8:F1} мм");
                sb.AppendLine($"  Радиус:  {result.PipeARadiusMm:F1} мм");
                sb.AppendLine();
            }

            if (result.PipeBId != null && result.PipeBStart != null)
            {
                sb.AppendLine($"Труба B  ID: {result.PipeBId.Value}");
                sb.AppendLine($"  Начало:  X={result.PipeBStart.X * 304.8:F1}  Y={result.PipeBStart.Y * 304.8:F1}  Z={result.PipeBStart.Z * 304.8:F1} мм");
                sb.AppendLine($"  Конец:   X={result.PipeBEnd.X * 304.8:F1}  Y={result.PipeBEnd.Y * 304.8:F1}  Z={result.PipeBEnd.Z * 304.8:F1} мм");
                sb.AppendLine($"  Радиус:  {result.PipeBRadiusMm:F1} мм");
                sb.AppendLine();
            }

            if (result.IntersectionPoint != null)
            {
                sb.AppendLine($"Точка пересечения (на оси A):");
                sb.AppendLine($"  X={result.IntersectionPoint.X * 304.8:F1}  Y={result.IntersectionPoint.Y * 304.8:F1}  Z={result.IntersectionPoint.Z * 304.8:F1} мм");
                sb.AppendLine();
            }

            if (result.DropMm != 0)
            {
                if (result.UsedClearanceMm > 0)
                    sb.AppendLine($"Отступ под трубой B: {result.UsedClearanceMm:F0} мм{(result.UsedClearanceMm > 0 && chkAutoClearance.IsChecked == true ? " (авто)" : "")}");
                sb.AppendLine($"Опускание среднего участка: {result.DropMm:F1} мм");
            }

            if (sb.Length > 0)
            {
                txtDiagnostics.Text = sb.ToString().TrimEnd();
                grpDiagnostics.Visibility = WpfVisibility.Visible;
            }
        }

        // ----------------------------------------------------------------
        // Auto-clearance checkbox
        // ----------------------------------------------------------------
        private void ChkAutoClearance_Changed(object sender, RoutedEventArgs e)
        {
            bool isAuto = chkAutoClearance.IsChecked == true;
            txtClearance.IsEnabled = !isAuto;
            txtClearance.Background = isAuto
                ? new SolidColorBrush(MediaColor.FromRgb(240, 240, 240))
                : System.Windows.Media.Brushes.White;
            TryUpdateAutoDisplay();
        }

        private void ChkAutoHalfLength_Changed(object sender, RoutedEventArgs e)
        {
            bool isAuto = chkAutoHalfLength.IsChecked == true;
            txtHalfLength.IsEnabled = !isAuto;
            txtHalfLength.Background = isAuto
                ? new SolidColorBrush(MediaColor.FromRgb(240, 240, 240))
                : System.Windows.Media.Brushes.White;
            TryUpdateAutoDisplay();
        }

        /// <summary>
        /// When auto checkboxes are on and both pipes are selected,
        /// preview computed values in the (read-only) text fields.
        /// </summary>
        private void TryUpdateAutoDisplay()
        {
            if (_pipeAId == null || _pipeBId == null) return;

            try
            {
                var doc = _uiApp?.ActiveUIDocument?.Document;
                if (doc == null) return;

                Element elemA = doc.GetElement(_pipeAId);
                Element elemB = doc.GetElement(_pipeBId);
                if (elemA == null || elemB == null) return;

                double rA = GetOuterRadius(elemA);
                double rB = GetOuterRadius(elemB);
                if (rA <= 0 || rB <= 0) return;

                double rMax  = Math.Max(rA, rB) * 304.8; // mm
                bool isAngled = rbAngle45?.IsChecked == true
                             || rbAngle60?.IsChecked == true
                             || rbAngle30?.IsChecked == true;

                if (chkAutoClearance.IsChecked == true)
                {
                    double clearanceMm = isAngled ? 50.0 : Math.Ceiling(rMax * 3.5);
                    txtClearance.Text = ((int)clearanceMm).ToString();
                }

                if (chkAutoHalfLength.IsChecked == true)
                {
                    double halfLengthMm = Math.Ceiling(rMax * (isAngled ? 3.0 : 2.5));
                    txtHalfLength.Text = ((int)(halfLengthMm * 2)).ToString(); // display full segment length
                }
            }
            catch { /* silently skip */ }
        }

        private static double GetOuterRadius(Element elem)
        {
            var mep = elem as Autodesk.Revit.DB.MEPCurve;
            if (mep == null) return 0;

            Parameter p = mep.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                       ?? mep.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                       ?? mep.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

            if (p != null && p.HasValue)
                return p.AsDouble() / 2.0;

            return 0;
        }

        private void RbAngle_Changed(object sender, RoutedEventArgs e)
        {
            TryUpdateAutoDisplay();
        }

        // ----------------------------------------------------------------
        // Helpers
        // ----------------------------------------------------------------
        private void UpdateResolveButton()
        {
            btnResolve.IsEnabled = _pipeAId != null && _pipeBId != null;
        }

        private void ShowStatus(string message)
        {
            txtStatus.Text = message;
        }
    }
}
