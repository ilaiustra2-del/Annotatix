using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

/// <summary>
/// FamilySync Module - Nested Family Parameter Synchronization
/// 
/// IMPLEMENTATION APPROACH (based on Dynamo algorithm):
/// This module implements parameter synchronization for nested families with two strategies:
/// 
/// 1. SHARED NESTED COMPONENTS (Implemented):
///    - These are special nested families with SuperComponent property
///    - Can be modified directly from project context without opening family file
///    - Algorithm: Check element.SuperComponent != null to identify shared nested
///    - Recursively process all levels of nesting using GetSubComponentIds()
///    - Directly set parameter values using LookupParameter().Set()
/// 
/// 2. REGULAR NESTED FAMILIES (Future implementation):
///    - Standard nested families without SuperComponent
///    - Cannot be modified from project - require opening family file
///    - Would need: Document.EditFamily() -> modify -> save -> reload
/// 
/// ALGORITHM FLOW:
/// 1. Get main family instance from user selection
/// 2. Separate nested families into:
///    - Shared nested (SuperComponent != null) - can modify
///    - Regular nested (SuperComponent == null) - cannot modify from project
/// 3. Apply parameter changes to main family + all shared nested (recursive)
/// 4. Report statistics: total/shared/regular nested families
/// 
/// REFERENCE: Based on Dynamo Python script from:
/// https://muratovbim.pro/blog/dynamo-kopiruem-znachenie-parametra-vo-vse-vlozhennye-semejstva
/// </summary>
namespace FamilySync.Module.UI
{
    /// <summary>
    /// External Event Handler for synchronization operations
    /// </summary>
    public class FamilySyncHandler : IExternalEventHandler
    {
        public FamilySyncPanel Panel { get; set; }
        public string ParameterName { get; set; }
        public string SyncValue { get; set; }
        public bool CreateIfNotExists { get; set; }
        public bool CopyFromParent { get; set; }
        
        public int SuccessCount { get; private set; }
        public int ErrorCount { get; private set; }
        public List<string> Errors { get; private set; } = new List<string>();
        public bool ParameterCreated { get; private set; }
        
        // New properties for detailed statistics
        public int TotalNestedCount { get; private set; }
        public int SharedNestedCount { get; private set; }
        public int RegularNestedCount { get; private set; }

        public void Execute(UIApplication app)
        {
            if (Panel == null) return;
            
            SuccessCount = 0;
            ErrorCount = 0;
            Errors.Clear();
            ParameterCreated = false;
            TotalNestedCount = 0;
            SharedNestedCount = 0;
            RegularNestedCount = 0;

            try
            {
                int successCount = 0;
                int errorCount = 0;
                List<string> errors = new List<string>();
                bool parameterCreated = false;
                int totalNested = 0;
                int sharedNested = 0;
                int regularNested = 0;

                // Execute for ALL selected family instances
                Panel.ExecuteSynchronizationForAllInstances(
                    ParameterName, 
                    SyncValue, 
                    CreateIfNotExists,
                    CopyFromParent,
                    out successCount, 
                    out errorCount, 
                    out errors, 
                    out parameterCreated,
                    out totalNested,
                    out sharedNested,
                    out regularNested
                );

                // Store results
                SuccessCount = successCount;
                ErrorCount = errorCount;
                Errors = errors;
                ParameterCreated = parameterCreated;
                TotalNestedCount = totalNested;
                SharedNestedCount = sharedNested;
                RegularNestedCount = regularNested;
            }
            catch (Exception ex)
            {
                Errors.Add($"Критическая ошибка: {ex.Message}");
                ErrorCount++;
            }
        }

        public string GetName()
        {
            return "FamilySync Synchronization Handler";
        }
    }

    public partial class FamilySyncPanel : UserControl
    {
        private UIApplication _uiApp;
        private Document _doc;
        private List<NestedFamilyInfo> _analyzedFamilies = new List<NestedFamilyInfo>();
        private FamilyInstance _mainFamilyInstance; // Keep for backward compatibility
        private List<FamilyInstance> _selectedFamilyInstances = new List<FamilyInstance>(); // Support multiple selection
        private List<string> _availableParameters = new List<string>();
        
        // External Event for synchronization
        private FamilySyncHandler _syncHandler;
        private ExternalEvent _syncEvent;

        /// <summary>
        /// Stores information about a nested family
        /// </summary>
        public class NestedFamilyInfo
        {
            public string FamilyName { get; set; }
            public string TypeName { get; set; }
            public string Category { get; set; }
            public ElementId ElementId { get; set; }
            public List<NestedFamilyInfo> Children { get; set; } = new List<NestedFamilyInfo>();
            public bool IsMainFamily { get; set; }
        }

        public FamilySyncPanel(UIApplication uiApp, FamilySyncHandler syncHandler = null, ExternalEvent syncEvent = null)
        {
            InitializeComponent();
            _uiApp = uiApp;
            
            // Use provided External Event or create new one (if called in API context)
            if (syncEvent != null && syncHandler != null)
            {
                _syncHandler = syncHandler;
                _syncEvent = syncEvent;
                _syncHandler.Panel = this;
                System.Diagnostics.Debug.WriteLine("[FAMILY-SYNC] Using provided ExternalEvent");
            }
            else
            {
                try
                {
                    // Try to create (will fail if not in API context)
                    _syncHandler = new FamilySyncHandler { Panel = this };
                    _syncEvent = ExternalEvent.Create(_syncHandler);
                    System.Diagnostics.Debug.WriteLine("[FAMILY-SYNC] Created new ExternalEvent");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Failed to create ExternalEvent: {ex.Message}");
                    _syncHandler = null;
                    _syncEvent = null;
                }
            }
            
            System.Diagnostics.Debug.WriteLine("[FAMILY-SYNC] Panel instance created");
            PluginsManager.Core.DebugLogger.Log("[FAMILY-SYNC] Panel instance created");

            this.Loaded += (s, e) => InitializePanel();
        }

        private void InitializePanel()
        {
            try
            {
                if (_uiApp == null)
                {
                    txtStatus.Text = "Ошибка: UIApplication не инициализирован.";
                    btnAnalyze.IsEnabled = false;
                    return;
                }

                _doc = _uiApp.ActiveUIDocument?.Document;

                if (_doc == null)
                {
                    txtStatus.Text = "Ошибка: Нет активного документа.";
                    btnAnalyze.IsEnabled = false;
                    return;
                }

                // Initialize parameter combobox with default value
                InitializeParameterComboBox();
                
                // Set initial state for txtSyncValue based on chkCopyFromParent
                UpdateSyncValueState();

                txtStatus.Text = "Готов к работе. Выделите элементы и нажмите 'Анализировать'.";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Ошибка инициализации: {ex.Message}";
                btnAnalyze.IsEnabled = false;
            }
        }

        private void InitializeParameterComboBox()
        {
            cmbParameter.Items.Clear();
            
            // Add common parameters
            var defaultParameters = new List<string>
            {
                "ИмяСемейства",
                "Семейство",
                "Марка",
                "Комментарии",
                "Описание",
                "Код по классификатору",
                "URL",
                "Модель"
            };

            foreach (var param in defaultParameters)
            {
                cmbParameter.Items.Add(param);
            }

            // Select default parameter
            cmbParameter.SelectedIndex = 0;
        }
        
        /// <summary>
        /// Handle checkbox state change for "Copy from parent"
        /// </summary>
        private void ChkCopyFromParent_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateSyncValueState();
        }
        
        /// <summary>
        /// Update txtSyncValue enabled state based on chkCopyFromParent
        /// </summary>
        private void UpdateSyncValueState()
        {
            if (chkCopyFromParent.IsChecked == true)
            {
                txtSyncValue.IsEnabled = false;
                txtSyncValue.Text = "Значение будет скопировано из родительского семейства";
                txtSyncValue.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(248, 249, 250));
            }
            else
            {
                txtSyncValue.IsEnabled = true;
                if (txtSyncValue.Text == "Значение будет скопировано из родительского семейства")
                {
                    txtSyncValue.Text = "";
                }
                txtSyncValue.Background = System.Windows.Media.Brushes.White;
            }
        }

        private void BtnAnalyze_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnAnalyze.IsEnabled = false;
                txtProgress.Text = "Анализ выбранных элементов...";
                
                _analyzedFamilies.Clear();
                treeNestedFamilies.Items.Clear();
                _selectedFamilyInstances.Clear(); // Clear previous selection

                if (_uiApp?.ActiveUIDocument == null)
                {
                    MessageBox.Show("Нет активного документа.", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                _doc = _uiApp.ActiveUIDocument.Document;
                var selection = _uiApp.ActiveUIDocument.Selection;
                var selectedIds = selection.GetElementIds();

                if (selectedIds.Count == 0)
                {
                    txtStatus.Text = "Не выбрано ни одного элемента. Выделите семейство на виде.";
                    txtProgress.Text = "";
                    txtSelectedFamily.Text = "(Семейство не выбрано)";
                    return;
                }

                // Process ALL selected elements (support multiple selection)
                int processedCount = 0;
                foreach (ElementId elemId in selectedIds)
                {
                    Element element = _doc.GetElement(elemId);
                    if (element == null) continue;

                    // Check if it's a FamilyInstance
                    if (element is FamilyInstance familyInstance)
                    {
                        _selectedFamilyInstances.Add(familyInstance);
                        
                        // Keep first as main for backward compatibility
                        if (_mainFamilyInstance == null)
                        {
                            _mainFamilyInstance = familyInstance;
                        }
                        
                        AnalyzeFamilyInstance(familyInstance);
                        processedCount++;
                    }
                }
                
                if (processedCount == 0)
                {
                    txtStatus.Text = "Ни один из выбранных элементов не является семейством.";
                    txtSelectedFamily.Text = "(Нет подходящих элементов)";
                    return;
                }
                
                txtSelectedFamily.Text = processedCount == 1 
                    ? _mainFamilyInstance.Name 
                    : $"{processedCount} семейств";

                txtProgress.Text = "Анализ завершён!";
                btnSynchronize.IsEnabled = _analyzedFamilies.Count > 0;

                // Update parameters combobox with available parameters from the family
                UpdateAvailableParameters();
            }
            catch (Exception ex)
            {
                txtProgress.Text = "Ошибка при анализе.";
                txtStatus.Text = $"Ошибка: {ex.Message}\n{ex.StackTrace}";
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Error: {ex.Message}");
            }
            finally
            {
                btnAnalyze.IsEnabled = true;
            }
        }

        private void AnalyzeFamilyInstance(FamilyInstance familyInstance)
        {
            try
            {
                Family family = familyInstance.Symbol?.Family;
                if (family == null)
                {
                    txtStatus.Text = "Не удалось получить семейство из выбранного элемента.";
                    return;
                }

                string familyName = family.Name;
                string typeName = familyInstance.Symbol?.Name ?? "";
                string categoryName = familyInstance.Category?.Name ?? "";

                txtSelectedFamily.Text = $"{familyName}: {typeName}";

                // Create main family info
                var mainFamilyInfo = new NestedFamilyInfo
                {
                    FamilyName = familyName,
                    TypeName = typeName,
                    Category = categoryName,
                    ElementId = familyInstance.Id,
                    IsMainFamily = true
                };

                // Get nested families from family document
                var nestedFamilies = GetNestedFamilies(family);
                mainFamilyInfo.Children = nestedFamilies;

                _analyzedFamilies.Add(mainFamilyInfo);
                _analyzedFamilies.AddRange(FlattenNestedFamilies(nestedFamilies));

                // Build tree view
                BuildTreeView(mainFamilyInfo);

                // Update count
                int totalNested = CountNestedFamilies(mainFamilyInfo) - 1; // Exclude main family
                txtNestedCount.Text = $"(найдено: {totalNested})";

                txtStatus.Text = $"Проанализировано семейство: {familyName}\n" +
                                $"Тип: {typeName}\n" +
                                $"Категория: {categoryName}\n" +
                                $"Вложенных семейств: {totalNested}";
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Ошибка анализа семейства: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] AnalyzeFamilyInstance error: {ex.Message}");
            }
        }

        private List<NestedFamilyInfo> GetNestedFamilies(Family family)
        {
            var nestedFamilies = new List<NestedFamilyInfo>();

            try
            {
                // Open family document for editing (read-only mode)
                Document familyDoc = _doc.EditFamily(family);
                
                if (familyDoc == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Could not open family document for: {family.Name}");
                    return nestedFamilies;
                }

                try
                {
                    // Collect all FamilyInstance elements in the family document
                    FilteredElementCollector collector = new FilteredElementCollector(familyDoc)
                        .OfClass(typeof(FamilyInstance));

                    var processedFamilies = new HashSet<string>();

                    foreach (FamilyInstance nestedInstance in collector)
                    {
                        if (nestedInstance?.Symbol?.Family == null)
                            continue;

                        Family nestedFamily = nestedInstance.Symbol.Family;
                        string familyKey = $"{nestedFamily.Name}:{nestedInstance.Symbol.Name}";

                        // Avoid duplicates
                        if (processedFamilies.Contains(familyKey))
                            continue;

                        processedFamilies.Add(familyKey);

                        var nestedInfo = new NestedFamilyInfo
                        {
                            FamilyName = nestedFamily.Name,
                            TypeName = nestedInstance.Symbol.Name,
                            Category = nestedInstance.Category?.Name ?? "",
                            ElementId = nestedInstance.Id,
                            IsMainFamily = false
                        };

                        // Recursively get nested families (limited depth to avoid infinite loops)
                        // Note: Deep recursion may cause performance issues
                        // nestedInfo.Children = GetNestedFamilies(nestedFamily);

                        nestedFamilies.Add(nestedInfo);
                    }
                }
                finally
                {
                    // Close the family document without saving
                    familyDoc.Close(false);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] GetNestedFamilies error: {ex.Message}");
            }

            return nestedFamilies;
        }

        private List<NestedFamilyInfo> FlattenNestedFamilies(List<NestedFamilyInfo> families)
        {
            var result = new List<NestedFamilyInfo>();
            
            foreach (var family in families)
            {
                result.Add(family);
                if (family.Children.Count > 0)
                {
                    result.AddRange(FlattenNestedFamilies(family.Children));
                }
            }
            
            return result;
        }

        private int CountNestedFamilies(NestedFamilyInfo root)
        {
            int count = 1; // Count self
            foreach (var child in root.Children)
            {
                count += CountNestedFamilies(child);
            }
            return count;
        }

        private void BuildTreeView(NestedFamilyInfo rootFamily)
        {
            treeNestedFamilies.Items.Clear();

            var rootItem = CreateTreeViewItem(rootFamily, true);
            treeNestedFamilies.Items.Add(rootItem);
        }

        private TreeViewItem CreateTreeViewItem(NestedFamilyInfo familyInfo, bool isRoot)
        {
            var item = new TreeViewItem
            {
                IsExpanded = true
            };

            // Create header with icon and text
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
            
            // Icon based on whether it's main or nested
            var icon = new TextBlock
            {
                Text = isRoot ? "📦" : "📁",
                Margin = new Thickness(0, 0, 5, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            headerPanel.Children.Add(icon);

            // Family name and type
            var nameText = new TextBlock
            {
                Text = string.IsNullOrEmpty(familyInfo.TypeName) 
                    ? familyInfo.FamilyName 
                    : $"{familyInfo.FamilyName}: {familyInfo.TypeName}",
                FontWeight = isRoot ? FontWeights.Bold : FontWeights.Normal,
                VerticalAlignment = VerticalAlignment.Center
            };
            headerPanel.Children.Add(nameText);

            // Category in brackets
            if (!string.IsNullOrEmpty(familyInfo.Category))
            {
                var categoryText = new TextBlock
                {
                    Text = $" [{familyInfo.Category}]",
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(108, 117, 125)),
                    FontSize = 10,
                    VerticalAlignment = VerticalAlignment.Center
                };
                headerPanel.Children.Add(categoryText);
            }

            item.Header = headerPanel;
            item.Tag = familyInfo;

            // Add children
            foreach (var child in familyInfo.Children)
            {
                item.Items.Add(CreateTreeViewItem(child, false));
            }

            return item;
        }

        private void UpdateAvailableParameters()
        {
            try
            {
                if (_mainFamilyInstance == null)
                    return;

                _availableParameters.Clear();
                var existingParams = new HashSet<string>();

                // Get parameters from the family instance
                foreach (Parameter param in _mainFamilyInstance.Parameters)
                {
                    if (param.Definition == null)
                        continue;

                    string paramName = param.Definition.Name;
                    if (!existingParams.Contains(paramName))
                    {
                        existingParams.Add(paramName);
                        _availableParameters.Add(paramName);
                    }
                }

                // Sort and update combobox
                _availableParameters.Sort();

                // Remember current selection
                string currentSelection = cmbParameter.SelectedItem?.ToString() ?? "ИмяСемейства";

                cmbParameter.Items.Clear();

                // Add default first
                cmbParameter.Items.Add("ИмяСемейства");

                foreach (var param in _availableParameters)
                {
                    if (param != "ИмяСемейства")
                    {
                        cmbParameter.Items.Add(param);
                    }
                }

                // Restore selection
                if (cmbParameter.Items.Contains(currentSelection))
                {
                    cmbParameter.SelectedItem = currentSelection;
                }
                else
                {
                    cmbParameter.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] UpdateAvailableParameters error: {ex.Message}");
            }
        }

        private async void BtnSynchronize_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_analyzedFamilies.Count == 0)
                {
                    MessageBox.Show("Сначала выполните анализ семейства.", "Предупреждение", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Get parameter name from text input (not from SelectedItem)
                // This ensures user can manually enter parameter name
                string parameterName = cmbParameter.Text?.Trim();
                string syncValue = txtSyncValue.Text;
                bool createIfNotExists = chkCreateParameter.IsChecked == true;
                bool copyFromParent = chkCopyFromParent.IsChecked == true;

                if (string.IsNullOrWhiteSpace(parameterName))
                {
                    MessageBox.Show("Выберите параметр для синхронизации.", "Предупреждение", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Check if parameter name matches an existing parameter from the list
                bool isExistingParameter = false;
                if (cmbParameter.Items.Count > 0)
                {
                    isExistingParameter = cmbParameter.Items.Cast<string>()
                        .Any(item => item.Equals(parameterName, StringComparison.OrdinalIgnoreCase));
                }

                // Log parameter selection info for debugging
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Parameter name: '{parameterName}'");
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Is existing parameter: {isExistingParameter}");
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Create if not exists: {createIfNotExists}");
                System.Diagnostics.Debug.WriteLine($"[FAMILY-SYNC] Copy from parent: {copyFromParent}");

                // Validate syncValue only if NOT copying from parent
                if (!copyFromParent && string.IsNullOrWhiteSpace(syncValue))
                {
                    var result = MessageBox.Show(
                        "Значение для синхронизации пустое. Продолжить?", 
                        "Подтверждение", 
                        MessageBoxButton.YesNo, 
                        MessageBoxImage.Question);
                    
                    if (result != MessageBoxResult.Yes)
                        return;
                }

                btnSynchronize.IsEnabled = false;
                txtProgress.Text = "Синхронизация параметров...";

                // Check if ExternalEvent is available
                if (_syncHandler == null || _syncEvent == null)
                {
                    MessageBox.Show(
                        "ExternalEvent не инициализирован. Возможно, модуль был загружен некорректно.\n" +
                        "Попробуйте перезапустить Revit и открыть модуль заново.",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                // Set parameters for handler
                _syncHandler.ParameterName = parameterName;
                _syncHandler.SyncValue = syncValue;
                _syncHandler.CreateIfNotExists = createIfNotExists;
                _syncHandler.CopyFromParent = copyFromParent;

                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ===== SYNCHRONIZATION START =====");
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Main family instance: {_mainFamilyInstance?.Name}");
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Analyzed families count: {_analyzedFamilies.Count}");
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Raising ExternalEvent...");

                // Raise external event
                _syncEvent.Raise();

                // Wait for completion
                await Task.Delay(500); // Increased delay to ensure transaction completes

                // Get results
                int successCount = _syncHandler.SuccessCount;
                int errorCount = _syncHandler.ErrorCount;
                var errors = _syncHandler.Errors;
                bool parameterCreated = _syncHandler.ParameterCreated;
                int totalNested = _syncHandler.TotalNestedCount;
                int sharedNested = _syncHandler.SharedNestedCount;
                int regularNested = _syncHandler.RegularNestedCount;

                // Update UI after transaction completes
                txtProgress.Text = "Синхронизация завершена!";
                
                string statusMessage = $"Синхронизация завершена.\n" +
                                      $"Параметр: {parameterName}\n" +
                                      $"Значение: {syncValue}\n";
                
                if (parameterCreated)
                {
                    statusMessage += $"Параметр создан в проекте\n";
                }
                
                statusMessage += $"\nУспешно: {successCount}\n" +
                               $"Ошибок: {errorCount}";
                
                // Add nested families statistics
                if (totalNested > 0)
                {
                    statusMessage += $"\n\nВложенные семейства:\n";
                    statusMessage += $"  Всего: {totalNested}\n";
                    statusMessage += $"  Общих (изменено): {sharedNested}\n";
                    statusMessage += $"  Обычных (требуют редактирования файла): {regularNested}";
                }
                
                txtStatus.Text = statusMessage;

                if (errors.Count > 0)
                {
                    txtStatus.Text += $"\n\nОшибки/Предупреждения:\n{string.Join("\n", errors)}";
                }

                string resultMessage = $"Синхронизация завершена!\n\n" +
                                     $"Параметр: {parameterName}\n";
                
                if (parameterCreated)
                {
                    resultMessage += $"✓ Параметр создан в проекте\n";
                }
                
                resultMessage += $"\nУспешно обновлено: {successCount}\n" +
                               $"Ошибок: {errorCount}";
                
                // Add detailed statistics to result message
                if (totalNested > 0)
                {
                    resultMessage += $"\n\nВложенные семейства:\n";
                    resultMessage += $"  Всего: {totalNested}\n";
                    if (sharedNested > 0)
                    {
                        resultMessage += $"  ✓ Общих (изменено): {sharedNested}\n";
                    }
                    if (regularNested > 0)
                    {
                        resultMessage += $"  ⚠ Обычных (не изменено): {regularNested}\n";
                        resultMessage += $"     (требуется редактирование файла семейства)";
                    }
                }

                MessageBox.Show(
                    resultMessage,
                    "Результат",
                    MessageBoxButton.OK,
                    errorCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                txtProgress.Text = "Ошибка при синхронизации.";
                txtStatus.Text = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Ошибка при синхронизации:\n{ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnSynchronize.IsEnabled = _analyzedFamilies.Count > 0;
            }
        }

        /// <summary>
        /// Execute synchronization in transaction (called from ExternalEvent handler)
        /// Processes ALL selected family instances
        /// </summary>
        public void ExecuteSynchronizationForAllInstances(
            string parameterName, 
            string syncValue, 
            bool createIfNotExists,
            bool copyFromParent,
            out int successCount,
            out int errorCount,
            out List<string> errors,
            out bool parameterCreated,
            out int totalNestedCount,
            out int sharedNestedCount,
            out int regularNestedCount)
        {
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ========================================");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ExecuteSynchronizationForAllInstances CALLED");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Parameter: {parameterName}");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Value: {syncValue}");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Create if not exists: {createIfNotExists}");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Copy from parent: {copyFromParent}");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Selected instances: {_selectedFamilyInstances.Count}");
            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ========================================");

            successCount = 0;
            errorCount = 0;
            errors = new List<string>();
            parameterCreated = false;
            totalNestedCount = 0;
            sharedNestedCount = 0;
            regularNestedCount = 0;

            // Execute synchronization using transaction
            using (Transaction trans = new Transaction(_doc, "Family Sync - Синхронизация параметров"))
            {
                trans.Start();

                try
                {
                    // Create parameter if needed and enabled
                    if (createIfNotExists)
                    {
                        if (CreateProjectParameterIfNotExists(parameterName, out string createError))
                        {
                            parameterCreated = true;
                        }
                        else if (!string.IsNullOrEmpty(createError))
                        {
                            errors.Add($"Ошибка создания параметра: {createError}");
                        }
                    }

                    // Process EACH selected family instance
                    foreach (var familyInstance in _selectedFamilyInstances)
                    {
                        if (familyInstance == null) continue;

                        // Get value to sync - either from parent or from user input
                        string valueToSync = syncValue;
                        
                        if (copyFromParent)
                        {
                            // Read value from THIS parent family instance
                            valueToSync = GetParameterValue(familyInstance, parameterName);
                            
                            if (string.IsNullOrEmpty(valueToSync))
                            {
                                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] WARNING: Parent {familyInstance.Name} has empty parameter '{parameterName}'");
                                errors.Add($"У элемента {familyInstance.Name} пустой параметр '{parameterName}' - пропускаем");
                                continue;
                            }
                            
                            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Copying value '{valueToSync}' from parent {familyInstance.Name}");
                        }

                        // Set parameter for parent family
                        if (SetParameterValue(familyInstance, parameterName, valueToSync))
                        {
                            successCount++;
                            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ✓ Set parameter for parent: {familyInstance.Name}");
                        }
                        else
                        {
                            errorCount++;
                            errors.Add($"Не удалось установить параметр для: {familyInstance.Name}");
                        }

                        // Synchronize SHARED nested components with THIS parent's value
                        try
                        {
                            var sharedNestedInstances = GetSharedNestedComponents(familyInstance);
                            int totalNested = GetNestedFamilyInstances(familyInstance).Count;
                            int sharedCount = sharedNestedInstances.Count;
                            int regularCount = totalNested - sharedCount;

                            // Accumulate statistics
                            totalNestedCount += totalNested;
                            sharedNestedCount += sharedCount;
                            regularNestedCount += regularCount;

                            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Parent: {familyInstance.Name}, Total nested: {totalNested}, Shared: {sharedCount}, Regular: {regularCount}");

                            if (sharedNestedInstances.Count > 0)
                            {
                                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Processing {sharedNestedInstances.Count} shared nested for {familyInstance.Name}...");

                                foreach (var nestedInstance in sharedNestedInstances)
                                {
                                    if (SetParameterValue(nestedInstance, parameterName, valueToSync))
                                    {
                                        successCount++;
                                        PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ✓ Set parameter for nested: {nestedInstance.Name} = {valueToSync}");
                                    }
                                    else
                                    {
                                        errorCount++;
                                        errors.Add($"Не удалось установить параметр для вложенного: {nestedInstance.Name}");
                                        PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] ✗ Failed: {nestedInstance.Name}");
                                    }
                                }
                            }
                            else if (totalNested > 0)
                            {
                                string warningMsg = $"Элемент {familyInstance.Name}: найдено {totalNested} вложенных, но все обычные (не общие)";
                                errors.Add(warningMsg);
                                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] WARNING: {warningMsg}");
                            }
                        }
                        catch (Exception nestedEx)
                        {
                            PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Error processing nested for {familyInstance.Name}: {nestedEx.Message}");
                            errors.Add($"Ошибка обработки вложенных для {familyInstance.Name}: {nestedEx.Message}");
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    errors.Add($"Ошибка транзакции: {ex.Message}");
                    errorCount++;
                }
            }
        }
        
        /// <summary>
        /// Get parameter value from element as string
        /// </summary>
        private string GetParameterValue(Element element, string parameterName)
        {
            try
            {
                Parameter param = element.LookupParameter(parameterName);
                
                if (param == null)
                {
                    // Try to find by definition name
                    foreach (Parameter p in element.Parameters)
                    {
                        if (p.Definition?.Name == parameterName)
                        {
                            param = p;
                            break;
                        }
                    }
                }

                if (param == null || !param.HasValue)
                {
                    return null;
                }

                // Get value based on storage type
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString();

                    case StorageType.Integer:
                        return param.AsInteger().ToString();

                    case StorageType.Double:
                        return param.AsDouble().ToString();

                    case StorageType.ElementId:
                        ElementId elemId = param.AsElementId();
                        if (elemId != null && elemId != ElementId.InvalidElementId)
                        {
                            Element refElement = _doc.GetElement(elemId);
                            return refElement?.Name ?? elemId.ToString();
                        }
                        return null;

                    default:
                        return param.AsValueString();
                }
            }
            catch (Exception ex)
            {
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Error getting parameter value: {ex.Message}");
                return null;
            }
        }

        private bool CreateProjectParameterIfNotExists(string parameterName, out string errorMessage)
        {
            errorMessage = null;

            try
            {
                // Check if parameter already exists
                ParameterElement existingParam = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ParameterElement))
                    .Cast<ParameterElement>()
                    .FirstOrDefault(pe => pe.Name == parameterName);

                if (existingParam != null)
                {
                    // Parameter already exists, no need to create
                    return false;
                }

                // Get all categories used by the analyzed families
                var categoriesToBind = new HashSet<ElementId>();
                
                if (_mainFamilyInstance?.Category != null)
                {
                    categoriesToBind.Add(_mainFamilyInstance.Category.Id);
                }

                foreach (var family in _analyzedFamilies)
                {
                    if (!string.IsNullOrEmpty(family.Category))
                    {
                        Category cat = _doc.Settings.Categories.Cast<Category>()
                            .FirstOrDefault(c => c.Name == family.Category);
                        
                        if (cat != null)
                        {
                            categoriesToBind.Add(cat.Id);
                        }
                    }
                }

                if (categoriesToBind.Count == 0)
                {
                    errorMessage = "Не удалось определить категории для параметра";
                    return false;
                }

                // Create CategorySet for binding
                CategorySet categorySet = _doc.Application.Create.NewCategorySet();
                foreach (var catId in categoriesToBind)
                {
                    Category cat = Category.GetCategory(_doc, catId);
                    if (cat != null)
                    {
                        categorySet.Insert(cat);
                    }
                }

                // Create shared parameter file if needed
                DefinitionFile defFile = _doc.Application.OpenSharedParameterFile();
                
                if (defFile == null)
                {
                    // Create temporary shared parameters file
                    string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "TempSharedParams_FamilySync.txt");
                    if (!System.IO.File.Exists(tempFile))
                    {
                        System.IO.File.WriteAllText(tempFile, "# Temporary shared parameters file for FamilySync\n");
                    }
                    _doc.Application.SharedParametersFilename = tempFile;
                    defFile = _doc.Application.OpenSharedParameterFile();
                }

                if (defFile == null)
                {
                    errorMessage = "Не удалось создать файл общих параметров";
                    return false;
                }

                // Create or get definition group "Механизмы" (Mechanical Systems)
                DefinitionGroup defGroup = defFile.Groups.get_Item("Механизмы");
                if (defGroup == null)
                {
                    defGroup = defFile.Groups.Create("Механизмы");
                }

                // Check if definition already exists in the group
                Definition existingDef = defGroup.Definitions.get_Item(parameterName);
                if (existingDef == null)
                {
                    // Create new external definition (Text type)
                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(parameterName, SpecTypeId.String.Text);
                    options.UserModifiable = true;
                    
                    existingDef = defGroup.Definitions.Create(options);
                }

                if (existingDef == null)
                {
                    errorMessage = "Не удалось создать определение параметра";
                    return false;
                }

                // Create instance binding (parameter appears in instance properties)
                InstanceBinding binding = _doc.Application.Create.NewInstanceBinding(categorySet);

                // Bind parameter to categories in group "Механизмы" (Mechanical)
                BindingMap bindingMap = _doc.ParameterBindings;
                bool bindSuccess = bindingMap.Insert(existingDef, binding, GroupTypeId.Mechanical);

                if (!bindSuccess)
                {
                    // Try to re-insert if already exists
                    bindSuccess = bindingMap.ReInsert(existingDef, binding, GroupTypeId.Mechanical);
                    
                    if (!bindSuccess)
                    {
                        errorMessage = "Не удалось привязать параметр к категориям";
                        return false;
                    }
                }

                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Parameter '{parameterName}' created successfully");
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Bound to {categoriesToBind.Count} categories in group 'Механизмы'");

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Error creating parameter: {ex.Message}");
                return false;
            }
        }

        private bool SetParameterValue(Element element, string parameterName, string value)
        {
            try
            {
                Parameter param = element.LookupParameter(parameterName);
                
                if (param == null)
                {
                    foreach (Parameter p in element.Parameters)
                    {
                        if (p.Definition?.Name == parameterName)
                        {
                            param = p;
                            break;
                        }
                    }
                }

                if (param == null || param.IsReadOnly)
                {
                    return false;
                }

                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        return true;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intValue))
                        {
                            param.Set(intValue);
                            return true;
                        }
                        break;

                    case StorageType.Double:
                        if (double.TryParse(value, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out double doubleValue))
                        {
                            param.Set(doubleValue);
                            return true;
                        }
                        break;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Get all nested family instances from a family instance
        /// Note: In Revit API, nested families are encapsulated and cannot be directly modified
        /// This method attempts to find sub-components that may be accessible
        /// </summary>
        /// <summary>
        /// Get shared nested components that can be modified from project context
        /// Based on Dynamo algorithm: checks SuperComponent property
        /// </summary>
        private List<FamilyInstance> GetSharedNestedComponents(FamilyInstance mainInstance)
        {
            var sharedNested = new List<FamilyInstance>();

            try
            {
                // Get all sub-components (nested family instances)
                var subComponentIds = mainInstance.GetSubComponentIds();
                
                if (subComponentIds != null && subComponentIds.Count > 0)
                {
                    PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Main instance has {subComponentIds.Count} sub-components");
                    
                    foreach (ElementId subId in subComponentIds)
                    {
                        Element subElement = _doc.GetElement(subId);
                        
                        if (subElement is FamilyInstance nestedInstance)
                        {
                            // KEY CHECK from Dynamo algorithm: only Shared Nested Components have SuperComponent
                            // These can be modified from project context without opening family file
                            if (nestedInstance.SuperComponent != null)
                            {
                                sharedNested.Add(nestedInstance);
                                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Found SHARED nested: {nestedInstance.Name} (Category: {nestedInstance.Category?.Name})");
                                
                                // Recursively get deeper shared nested components
                                var deeperNested = GetSharedNestedComponents(nestedInstance);
                                sharedNested.AddRange(deeperNested);
                            }
                            else
                            {
                                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Found REGULAR nested (not shared): {nestedInstance.Name} - cannot modify from project");
                            }
                        }
                    }
                }
                else
                {
                    PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Main instance has no sub-components (nested families are internal)");
                }
            }
            catch (Exception ex)
            {
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Error getting shared nested components: {ex.Message}");
            }

            return sharedNested;
        }

        /// <summary>
        /// Legacy method - gets all nested instances (both shared and regular)
        /// Kept for compatibility and analysis purposes
        /// </summary>
        private List<FamilyInstance> GetNestedFamilyInstances(FamilyInstance mainInstance)
        {
            var nestedInstances = new List<FamilyInstance>();

            try
            {
                var subComponentIds = mainInstance.GetSubComponentIds();
                
                if (subComponentIds != null && subComponentIds.Count > 0)
                {
                    foreach (ElementId subId in subComponentIds)
                    {
                        Element subElement = _doc.GetElement(subId);
                        
                        if (subElement is FamilyInstance nestedInstance)
                        {
                            nestedInstances.Add(nestedInstance);
                            
                            // Recursively get nested instances from this nested instance
                            var deeperNested = GetNestedFamilyInstances(nestedInstance);
                            nestedInstances.AddRange(deeperNested);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PluginsManager.Core.DebugLogger.Log($"[FAMILY-SYNC] Error getting nested instances: {ex.Message}");
            }

            return nestedInstances;
        }

    }
}
