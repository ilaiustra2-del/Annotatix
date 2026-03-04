using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace AutoNumbering.Module.UI
{
    public partial class AutoNumberingPanel : UserControl
    {
        private readonly UIApplication _uiApp;
        private List<Core.RiserGroup> _riserGroups;
        private List<string> _availableSystemTypes; // All unique system types from analysis
        private List<Core.SystemCompatibilityGroup> _compatibilityGroups = new List<Core.SystemCompatibilityGroup>(); // User-defined groups
        private Core.RiserAnalyzer _lastAnalyzer; // Store last analyzer for statistics
        private ExternalEvent _numberingEvent;
        private object _numberingHandler; // PluginsManager.Commands.NumberRisersHandler

        public AutoNumberingPanel(UIApplication uiApp, ExternalEvent numberingEvent, object numberingHandler)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _numberingEvent = numberingEvent;
            _numberingHandler = numberingHandler;
            
            DebugLogger.Log($"[AUTONUMBERING-PANEL] Panel initialized: UIApp={uiApp != null}, Event={numberingEvent != null}, Handler={numberingHandler != null}");

            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Display active view name
                var doc = _uiApp.ActiveUIDocument?.Document;
                if (doc != null && doc.ActiveView != null)
                {
                    txtActiveView.Text = doc.ActiveView.Name;
                    LoadPipeParameters();
                }
                else
                {
                    txtActiveView.Text = "[Документ не открыт]";
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[AUTONUMBERING] ERROR on load: {ex.Message}");
            }
        }

        /// <summary>
        /// Load available pipe parameters into ComboBox
        /// </summary>
        private void LoadPipeParameters()
        {
            try
            {
                var doc = _uiApp.ActiveUIDocument?.Document;
                if (doc == null) return;

                cmbParameter.Items.Clear();

                // Get first pipe to extract parameters
                var collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .FirstElement() as Pipe;

                if (collector == null)
                {
                    cmbParameter.Items.Add("[Нет труб в проекте]");
                    cmbParameter.SelectedIndex = 0;
                    return;
                }

                // Get all instance text parameters
                var parameters = new List<string>();
                foreach (Parameter param in collector.Parameters)
                {
                    if (!param.IsReadOnly && 
                        param.StorageType == StorageType.String)
                    {
                        parameters.Add(param.Definition.Name);
                    }
                }

                // Sort and add to ComboBox
                parameters = parameters.Distinct().OrderBy(p => p).ToList();
                foreach (var param in parameters)
                {
                    cmbParameter.Items.Add(param);
                }

                // Try to select "ADSK_Номер стояка" by default
                var defaultParam = "ADSK_Номер стояка";
                if (parameters.Contains(defaultParam))
                {
                    cmbParameter.SelectedItem = defaultParam;
                }
                else if (cmbParameter.Items.Count > 0)
                {
                    cmbParameter.SelectedIndex = 0;
                }

                DebugLogger.Log($"[AUTONUMBERING] Loaded {parameters.Count} parameters");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[AUTONUMBERING] ERROR loading parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Analyze button click handler
        /// </summary>
        private void BtnAnalyze_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = _uiApp.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    MessageBox.Show("Документ не открыт", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var activeView = doc.ActiveView;
                if (activeView == null)
                {
                    MessageBox.Show("Активный вид не определен", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                DebugLogger.Log($"[AUTONUMBERING] Analyzing view: {activeView.Name}");
                txtStatus.Text = "Анализ...";
                
                // Parse grouping distance from textbox
                double groupDistanceMm = 400.0; // Default
                if (double.TryParse(txtGroupDistance.Text, out double parsedValue))
                {
                    groupDistanceMm = Math.Max(0, parsedValue); // Ensure non-negative
                }
                
                DebugLogger.Log($"[AUTONUMBERING] Grouping distance: {groupDistanceMm} mm");

                // Find vertical risers
                var analyzer = new Core.RiserAnalyzer(doc, activeView, groupDistanceMm);
                var risers = analyzer.FindRisers();
                _lastAnalyzer = analyzer;

                if (risers.Count == 0)
                {
                    txtStatus.Text = $"⚠ Стояков не найдено (проанализировано труб: {analyzer.TotalPipesAnalyzed})";
                    btnNumber.IsEnabled = false;
                    btnManageGroups.IsEnabled = false;
                    _riserGroups = null;
                    _availableSystemTypes = null;
                    return;
                }
                
                // Get unique system types
                _availableSystemTypes = analyzer.GetUniqueSystemTypes(risers);

                // Group risers (with compatibility groups if defined)
                _riserGroups = analyzer.GroupAndNumberRisers(risers, _compatibilityGroups);

                int totalPipes = analyzer.TotalPipesAnalyzed;
                int totalRisers = analyzer.TotalRisersFound;
                int totalGroups = _riserGroups.Count;

                txtStatus.Text = $"✓ Проанализировано труб: {totalPipes} | Найдено стояков: {totalRisers} | Групп: {totalGroups}";
                btnNumber.IsEnabled = true;
                btnManageGroups.IsEnabled = true; // Always enable if risers found (even if no system types)

                DebugLogger.Log($"[AUTONUMBERING] Analysis complete: {totalPipes} pipes analyzed, {totalRisers} risers found, {totalGroups} groups created");
                DebugLogger.Log($"[AUTONUMBERING] Available system types: {_availableSystemTypes.Count}");
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"✗ Ошибка анализа: {ex.Message}";
                DebugLogger.Log($"[AUTONUMBERING] ERROR during analysis: {ex.Message}");
                MessageBox.Show($"Ошибка при анализе:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Manage system compatibility groups button click handler
        /// </summary>
        private void BtnManageGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[AUTONUMBERING] Opening system groups management window...");
                
                var window = new SystemGroupsWindow(_availableSystemTypes, _compatibilityGroups);
                var result = window.ShowDialog();
                
                if (result == true)
                {
                    // User clicked Apply, update compatibility groups
                    _compatibilityGroups = window.GetCompatibilityGroups();
                    
                    DebugLogger.Log($"[AUTONUMBERING] Updated compatibility groups: {_compatibilityGroups.Count} groups defined");
                    
                    // Re-analyze with new groups
                    if (_lastAnalyzer != null)
                    {
                        var doc = _uiApp.ActiveUIDocument?.Document;
                        if (doc != null)
                        {
                            var activeView = doc.ActiveView;
                            var analyzer = new Core.RiserAnalyzer(doc, activeView, 
                                double.TryParse(txtGroupDistance.Text, out double d) ? Math.Max(0, d) : 400.0);
                            var risers = analyzer.FindRisers();
                            
                            if (risers.Count > 0)
                            {
                                // Re-group with new compatibility settings
                                _riserGroups = analyzer.GroupAndNumberRisers(risers, _compatibilityGroups);
                                
                                int totalPipes = analyzer.TotalPipesAnalyzed;
                                int totalRisers = analyzer.TotalRisersFound;
                                int totalGroups = _riserGroups.Count;
                                
                                txtStatus.Text = $"✓ Проанализировано труб: {totalPipes} | Найдено стояков: {totalRisers} | Групп: {totalGroups} (обновлено)";
                                
                                DebugLogger.Log($"[AUTONUMBERING] Re-grouped with compatibility: {totalGroups} groups created");
                            }
                        }
                    }
                    
                    MessageBox.Show(
                        $"Группы совместимости обновлены!\n\nСоздано групп: {_compatibilityGroups.Count}",
                        "Успех",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[AUTONUMBERING] ERROR in manage groups: {ex.Message}");
                MessageBox.Show(
                    $"Ошибка управления группами:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Number button click handler
        /// </summary>
        private void BtnNumber_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_riserGroups == null || _riserGroups.Count == 0)
                {
                    MessageBox.Show("Сначала выполните анализ", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var selectedParam = cmbParameter.SelectedItem as string;
                if (string.IsNullOrEmpty(selectedParam))
                {
                    MessageBox.Show("Выберите параметр для записи номеров", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string prefix = txtPrefix.Text ?? "";
                string postfix = txtPostfix.Text ?? "";

                DebugLogger.Log($"[AUTONUMBERING] Starting numbering with prefix='{prefix}', postfix='{postfix}', param='{selectedParam}'");
                txtStatus.Text = "Нумерация...";

                // Set handler parameters using reflection (handler is object type)
                var handlerType = _numberingHandler.GetType();
                handlerType.GetProperty("RiserGroups").SetValue(_numberingHandler, _riserGroups);
                handlerType.GetProperty("ParameterName").SetValue(_numberingHandler, selectedParam);
                handlerType.GetProperty("Prefix").SetValue(_numberingHandler, prefix);
                handlerType.GetProperty("Postfix").SetValue(_numberingHandler, postfix);

                // Raise ExternalEvent to execute in Revit API context
                _numberingEvent.Raise();

                // Wait for event to complete
                // MessageBox will be shown from handler when ready
                System.Threading.Thread.Sleep(100);

                // Update status based on results (reuse handlerType from above)
                var errorMessage = handlerType.GetProperty("ErrorMessage").GetValue(_numberingHandler) as string;
                
                if (errorMessage != null)
                {
                    txtStatus.Text = $"✗ Ошибка нумерации: {errorMessage}";
                    DebugLogger.Log($"[AUTONUMBERING] Numbering failed: {errorMessage}");
                }
                else
                {
                    // Status will be updated after handler completes
                    int successCount = (int)handlerType.GetProperty("SuccessCount").GetValue(_numberingHandler);
                    int totalPipes = (int)handlerType.GetProperty("TotalCount").GetValue(_numberingHandler);
                    txtStatus.Text = $"✓ Пронумеровано: {successCount} из {totalPipes} стояков";
                    DebugLogger.Log($"[AUTONUMBERING] Numbering complete: {successCount}/{totalPipes} successful");
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"✗ Ошибка нумерации: {ex.Message}";
                DebugLogger.Log($"[AUTONUMBERING] ERROR during numbering: {ex.Message}");
                MessageBox.Show($"Ошибка при нумерации:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
