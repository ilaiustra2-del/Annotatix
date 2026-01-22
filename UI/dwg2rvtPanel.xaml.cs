using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace dwg2rvt.UI
{
    public partial class dwg2rvtPanel : UserControl
    {
        private UIApplication _uiApp;
        private Document _doc;
        private List<ImportInstance> _importedFiles;
        private ExternalEvent _externalEvent;
        private ExternalEvent _placeElementsEvent;
        private ExternalEvent _placeSingleBlockTypeEvent;
        private Dictionary<string, BlockTypeInfo> _blockTypesData = new Dictionary<string, BlockTypeInfo>();

        // Class to store block type information with family selection
        public class BlockTypeInfo
        {
            public string BlockTypeName { get; set; }
            public int Count { get; set; }
            public List<BlockInstanceData> Instances { get; set; } = new List<BlockInstanceData>();
            public string SelectedFamily { get; set; }
            public System.Windows.Controls.ComboBox FamilyComboBox { get; set; }
        }

        public class BlockInstanceData
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public double CenterX { get; set; }
            public double CenterY { get; set; }
            public double RotationAngle { get; set; } = 0;  // Rotation angle in degrees
        }

        public dwg2rvtPanel(UIApplication uiApp, ExternalEvent annotateEvent, ExternalEvent placeElementsEvent = null, ExternalEvent placeSingleBlockTypeEvent = null)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _externalEvent = annotateEvent;
            _placeElementsEvent = placeElementsEvent;
            _placeSingleBlockTypeEvent = placeSingleBlockTypeEvent;
            
            // Debug logging
            System.Diagnostics.Debug.WriteLine($"[dwg2rvtPanel] Constructor called");
            System.Diagnostics.Debug.WriteLine($"[dwg2rvtPanel] annotateEvent: {annotateEvent != null}");
            System.Diagnostics.Debug.WriteLine($"[dwg2rvtPanel] placeElementsEvent: {placeElementsEvent != null}");
            System.Diagnostics.Debug.WriteLine($"[dwg2rvtPanel] placeSingleBlockTypeEvent: {placeSingleBlockTypeEvent != null}");
            
            // Set this panel as the active panel for PlaceElementsEventHandler
            PlaceElementsEventHandler.SetActivePanel(this);

            this.Loaded += (s, e) => InitializePanel();
        }

        private void InitializePanel()
        {
            try
            {
                _doc = _uiApp.ActiveUIDocument?.Document;

                if (_doc == null)
                {
                    txtStatus.Text = "Error: No active document found.";
                    if (btnAnalyze != null) btnAnalyze.IsEnabled = false;
                    return;
                }

                // Display active view name
                View activeView = _doc.ActiveView;
                if (activeView != null && txtActiveView != null)
                {
                    txtActiveView.Text = activeView.Name;
                }

                // Load imported DWG files
                LoadImportedFiles();

                if (txtStatus != null)
                    txtStatus.Text = "Ready. Select an imported DWG file and click Analyze.";
            }
            catch (Exception ex)
            {
                if (txtStatus != null)
                    txtStatus.Text = $"Error initializing panel: {ex.Message}";
                if (btnAnalyze != null) btnAnalyze.IsEnabled = false;
            }
        }

        private void LoadImportedFiles()
        {
            try
            {
                _importedFiles = new List<ImportInstance>();
                cmbImportedFiles.Items.Clear();

                // Get all ImportInstance elements (CAD links and imports)
                FilteredElementCollector collector = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                    .OfClass(typeof(ImportInstance))
                    .WhereElementIsNotElementType();

                foreach (ImportInstance importInstance in collector)
                {
                    if (importInstance != null)
                    {
                        _importedFiles.Add(importInstance);
                        
                        // Get the import symbol name
                        string displayName = GetImportDisplayName(importInstance);
                        cmbImportedFiles.Items.Add(displayName);
                    }
                }

                if (cmbImportedFiles.Items.Count > 0)
                {
                    cmbImportedFiles.SelectedIndex = 0;
                    txtStatus.Text = $"Found {cmbImportedFiles.Items.Count} imported file(s) on the active view.";
                }
                else
                {
                    txtStatus.Text = "No imported DWG files found on the active view.";
                    btnAnalyze.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Error loading imported files: {ex.Message}";
                btnAnalyze.IsEnabled = false;
            }
        }

        private string GetImportDisplayName(ImportInstance importInstance)
        {
            try
            {
                // Try to get the name parameter
                Parameter nameParam = importInstance.get_Parameter(BuiltInParameter.IMPORT_SYMBOL_NAME);
                if (nameParam != null && !string.IsNullOrEmpty(nameParam.AsString()))
                {
                    return nameParam.AsString();
                }

                // Fallback to element name
                return importInstance.Name ?? $"Import {importInstance.Id.Value}";
            }
            catch
            {
                return $"Import {importInstance.Id.Value}";
            }
        }

        private void BtnAnalyzeByName_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Document doc = _uiApp.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    MessageBox.Show("No active document found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (cmbImportedFiles.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select an imported DWG file.", "No Selection", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                btnAnalyze.IsEnabled = false;
                txtProgress.Text = "Analyzing Block Names...";
                txtStatus.Text = "Starting Block Name analysis...\n";
                
                // Clear previous data
                _blockTypesData.Clear();
                statusStackPanel.Children.Clear();
                statusStackPanel.Children.Add(txtStatus);

                // Get selected import instance
                ImportInstance selectedImport = _importedFiles[cmbImportedFiles.SelectedIndex];

                // Perform analysis using new AnalyzeByBlockName class
                Core.AnalyzeByBlockName analyzer = new Core.AnalyzeByBlockName(doc);
                var analysisResult = analyzer.Analyze(selectedImport, UpdateStatus);

                if (analysisResult.Success)
                {
                    txtProgress.Text = "Analysis complete!";
                    txtStatus.Text += $"Log file saved to: {analysisResult.LogFilePath}\n";
                    
                    // Parse log file to get block data
                    ParseAnalysisResults(analysisResult.LogFilePath);
                    
                    // Create dynamic UI for family selection
                    CreateFamilySelectionUI();

                    // Success message removed - status is shown in status window
                }
                else
                {
                    txtProgress.Text = "Analysis failed.";
                    txtStatus.Text += $"\nError: {analysisResult.ErrorMessage}";
                    MessageBox.Show($"Analysis failed:\n{analysisResult.ErrorMessage}", 
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                txtProgress.Text = "Error occurred.";
                txtStatus.Text += $"\nException: {ex.Message}\n{ex.StackTrace}";
                MessageBox.Show($"An error occurred:\n{ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnAnalyze.IsEnabled = true;
            }
        }

        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text += message + "\n";
            });
        }
        
        private void ParseAnalysisResults(string logFilePath)
        {
            try
            {
                var lines = System.IO.File.ReadAllLines(logFilePath);
                string currentBlockType = null;
                
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    
                    // Parse block instances from log
                    if (line.Contains("№") && line.Contains("Координаты центра блока:"))
                    {
                        // Extract block name and number
                        int hashIndex = line.IndexOf("№");
                        if (hashIndex > 0)
                        {
                            string blockName = line.Substring(0, hashIndex).Trim();
                            
                            // Extract coordinates section
                            int coordIndex = line.IndexOf("(");
                            int coordEndIndex = line.IndexOf(")");
                            if (coordIndex > 0 && coordEndIndex > coordIndex)
                            {
                                string coordSection = line.Substring(coordIndex + 1, coordEndIndex - coordIndex - 1);
                                
                                // Split by comma and space to separate coordinates
                                // Format: (X, Y) where comma separates coordinates, not decimal separator
                                string[] coordParts = coordSection.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                if (coordParts.Length >= 2)
                                {
                                    // Get first two parts (X and Y coordinates)
                                    string xStr = coordParts[0].Trim();
                                    string yStr = coordParts[1].Trim();
                                    
                                    // Extract number between № and Координаты
                                    int numberEndIndex = line.IndexOf("Координаты");
                                    string numberSection = line.Substring(hashIndex + 1, numberEndIndex - hashIndex - 1).Trim();
                                    
                                    // Extract rotation angle if present
                                    double rotationAngle = 0;
                                    if (line.Contains("Поворот:"))
                                    {
                                        int rotIndex = line.IndexOf("Поворот:");
                                        string rotSection = line.Substring(rotIndex + "Поворот:".Length).Trim();
                                        // Extract number before °
                                        int degIndex = rotSection.IndexOf("°");
                                        if (degIndex > 0)
                                        {
                                            string rotStr = rotSection.Substring(0, degIndex).Trim();
                                            double.TryParse(rotStr, System.Globalization.NumberStyles.Any, 
                                                System.Globalization.CultureInfo.InvariantCulture, out rotationAngle);
                                        }
                                    }
                                    
                                    int number;
                                    double centerX, centerY;
                                    
                                    if (int.TryParse(numberSection.Trim(), out number) &&
                                        double.TryParse(xStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out centerX) &&
                                        double.TryParse(yStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out centerY))
                                    {
                                        // Extract base block type name (without number)
                                        string blockTypeName = blockName;
                                        
                                        if (!_blockTypesData.ContainsKey(blockTypeName))
                                        {
                                            _blockTypesData[blockTypeName] = new BlockTypeInfo
                                            {
                                                BlockTypeName = blockTypeName,
                                                Count = 0
                                            };
                                        }
                                        
                                        _blockTypesData[blockTypeName].Instances.Add(new BlockInstanceData
                                        {
                                            Name = blockName,
                                            Number = number,
                                            CenterX = centerX,
                                            CenterY = centerY,
                                            RotationAngle = rotationAngle
                                        });
                                        
                                        _blockTypesData[blockTypeName].Count++;
                                    }
                                }
                            }
                        }
                    }
                    // Alternative parsing: look for lines with block type and count in summary
                    else if (line.Contains("–") && i > 0 && !line.StartsWith("Общее"))
                    {
                        string[] summaryParts = line.Split('–');
                        if (summaryParts.Length == 2)
                        {
                            string typeName = summaryParts[0].Trim();
                            if (!_blockTypesData.ContainsKey(typeName))
                            {
                                _blockTypesData[typeName] = new BlockTypeInfo
                                {
                                    BlockTypeName = typeName,
                                    Count = 0
                                };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error parsing analysis results: {ex.Message}");
            }
        }
        
        private void CreateFamilySelectionUI()
        {
            try
            {
                // Get available families from document
                List<string> availableFamilies = GetAvailableFamilySymbols();
                
                // Add table header
                var headerGrid = new System.Windows.Controls.Grid
                {
                    Margin = new Thickness(0, 10, 0, 5),
                    Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(248, 249, 250))
                };
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.2, GridUnitType.Star) });
                
                var headerCol1 = new TextBlock
                {
                    Text = "Наименование блока",
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(73, 80, 87)),
                    Padding = new Thickness(8, 6, 8, 6),
                    VerticalAlignment = VerticalAlignment.Center
                };
                System.Windows.Controls.Grid.SetColumn(headerCol1, 0);
                headerGrid.Children.Add(headerCol1);
                
                var headerCol2 = new TextBlock
                {
                    Text = "Кол-во на виде",
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(73, 80, 87)),
                    Padding = new Thickness(8, 6, 8, 6),
                    TextAlignment = TextAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                System.Windows.Controls.Grid.SetColumn(headerCol2, 1);
                headerGrid.Children.Add(headerCol2);
                
                var headerCol3 = new TextBlock
                {
                    Text = "Семейство для замены",
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(73, 80, 87)),
                    Padding = new Thickness(8, 6, 8, 6),
                    VerticalAlignment = VerticalAlignment.Center
                };
                System.Windows.Controls.Grid.SetColumn(headerCol3, 2);
                headerGrid.Children.Add(headerCol3);
                
                var headerCol4 = new TextBlock
                {
                    Text = "Действия",
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(73, 80, 87)),
                    Padding = new Thickness(8, 6, 8, 6),
                    TextAlignment = TextAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                System.Windows.Controls.Grid.SetColumn(headerCol4, 3);
                headerGrid.Children.Add(headerCol4);
                
                statusStackPanel.Children.Add(headerGrid);
                
                // Add header separator
                var headerSeparator = new Separator
                {
                    Margin = new Thickness(0, 0, 0, 0),
                    Height = 2,
                    Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(222, 226, 230))
                };
                statusStackPanel.Children.Add(headerSeparator);
                
                foreach (var blockType in _blockTypesData.Values)
                {
                    // Create table row using Grid
                    var rowGrid = new System.Windows.Controls.Grid
                    {
                        Margin = new Thickness(0, 5, 0, 5)
                    };
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.2, GridUnitType.Star) });
                    
                    // Column 1: Block name
                    var blockNameText = new TextBlock
                    {
                        Text = blockType.BlockTypeName,
                        FontSize = 11,
                        Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(33, 37, 41)),
                        Padding = new Thickness(8, 4, 8, 4),
                        TextWrapping = TextWrapping.Wrap,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    System.Windows.Controls.Grid.SetColumn(blockNameText, 0);
                    rowGrid.Children.Add(blockNameText);
                    
                    // Column 2: Count
                    var countText = new TextBlock
                    {
                        Text = blockType.Count.ToString(),
                        FontSize = 11,
                        Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(33, 37, 41)),
                        Padding = new Thickness(8, 4, 8, 4),
                        TextAlignment = TextAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    System.Windows.Controls.Grid.SetColumn(countText, 1);
                    rowGrid.Children.Add(countText);
                    
                    // Column 3: Family ComboBox container
                    var comboBoxContainer = new StackPanel
                    {
                        Margin = new Thickness(4, 0, 4, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    System.Windows.Controls.Grid.SetColumn(comboBoxContainer, 2);
                    rowGrid.Children.Add(comboBoxContainer);
                    
                    var familyComboBox = new System.Windows.Controls.ComboBox
                    {
                        Height = 28,
                        Margin = new Thickness(0),
                        FontSize = 10,
                        Background = System.Windows.Media.Brushes.White,
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(206, 212, 218)),
                        IsEditable = true,  // Enable text search
                        IsTextSearchEnabled = false,  // Disable default search
                        MaxDropDownHeight = 300  // Allow scrolling in dropdown
                    };
                    
                    // Store all families for substring filtering
                    var allFamiliesForComboBox = new List<string>(availableFamilies);
                    
                    // Add available families to combobox
                    foreach (var family in availableFamilies)
                    {
                        familyComboBox.Items.Add(family);
                    }
                    
                    // Create timer for delayed filtering to avoid UI freeze
                    System.Windows.Threading.DispatcherTimer filterTimer = null;
                    
                    // Add custom text changed handler for substring filtering
                    familyComboBox.AddHandler(System.Windows.Controls.Primitives.TextBoxBase.TextChangedEvent,
                        new System.Windows.Controls.TextChangedEventHandler((s, args) =>
                        {
                            var comboBox = s as System.Windows.Controls.ComboBox;
                            if (comboBox == null || !comboBox.IsEditable) return;
                            
                            // Stop previous timer
                            if (filterTimer != null)
                            {
                                filterTimer.Stop();
                            }
                            
                            // Create new timer for delayed filtering (300ms)
                            filterTimer = new System.Windows.Threading.DispatcherTimer
                            {
                                Interval = TimeSpan.FromMilliseconds(300)
                            };
                            
                            filterTimer.Tick += (ts, te) =>
                            {
                                filterTimer.Stop();
                                
                                string searchText = comboBox.Text;
                                
                                // Don't filter if user just selected an item
                                if (comboBox.SelectedItem != null)
                                {
                                    string selectedText = comboBox.SelectedItem.ToString();
                                    if (selectedText.Equals(searchText, StringComparison.Ordinal))
                                    {
                                        return;
                                    }
                                }
                                
                                object selectedBeforeFilter = comboBox.SelectedItem;
                                bool wasDropDownOpen = comboBox.IsDropDownOpen;
                                
                                comboBox.Items.Clear();
                                
                                if (string.IsNullOrEmpty(searchText))
                                {
                                    // Restore all items
                                    foreach (var family in allFamiliesForComboBox)
                                    {
                                        comboBox.Items.Add(family);
                                    }
                                }
                                else
                                {
                                    // Filter items that contain search text (case insensitive)
                                    foreach (var family in allFamiliesForComboBox)
                                    {
                                        if (family.ToLower().Contains(searchText.ToLower()))
                                        {
                                            comboBox.Items.Add(family);
                                        }
                                    }
                                }
                                
                                // Restore selection if it's still in filtered list
                                if (selectedBeforeFilter != null && comboBox.Items.Contains(selectedBeforeFilter))
                                {
                                    comboBox.SelectedItem = selectedBeforeFilter;
                                }
                                
                                // Open dropdown after filtering
                                if (!string.IsNullOrEmpty(searchText) && comboBox.Items.Count > 0)
                                {
                                    comboBox.IsDropDownOpen = true;
                                }
                            };
                            
                            filterTimer.Start();
                        }));
                    
                    // Auto-select family based on keywords
                    string autoSelectedFamilyName = AutoSelectFamily(blockType.BlockTypeName);
                    string matchedFamily = FindBestFamilyMatch(availableFamilies, autoSelectedFamilyName);
                    
                    if (!string.IsNullOrEmpty(matchedFamily))
                    {
                        familyComboBox.SelectedItem = matchedFamily;
                    }
                    else if (familyComboBox.Items.Count > 0)
                    {
                        familyComboBox.SelectedIndex = 0;
                    }
                    
                    // Store reference to combobox
                    blockType.FamilyComboBox = familyComboBox;
                    
                    comboBoxContainer.Children.Add(familyComboBox);
                    
                    // Column 4: Place Elements button
                    var placeButton = new Button
                    {
                        Content = "Расставить",
                        Height = 28,
                        Margin = new Thickness(4, 0, 0, 0),
                        FontSize = 10,
                        FontWeight = FontWeights.SemiBold,
                        Background = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(255, 193, 7)),
                        Foreground = System.Windows.Media.Brushes.Black,
                        BorderThickness = new Thickness(0),
                        Cursor = System.Windows.Input.Cursors.Hand,
                        Tag = blockType.BlockTypeName,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    
                    placeButton.Click += (s, args) =>
                    {
                        System.Diagnostics.Debug.WriteLine($"[Button.Click] Clicked for block type: {blockType.BlockTypeName}");
                        PlaceSingleBlockType(blockType.BlockTypeName);
                    };
                    
                    System.Windows.Controls.Grid.SetColumn(placeButton, 3);
                    rowGrid.Children.Add(placeButton);
                    
                    statusStackPanel.Children.Add(rowGrid);
                    
                    // Add row separator
                    var rowSeparator = new Separator
                    {
                        Margin = new Thickness(0, 0, 0, 0),
                        Height = 1,
                        Background = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(238, 238, 238))
                    };
                    statusStackPanel.Children.Add(rowSeparator);
                }
                
                // Enable Place Elements button if we have data
                if (_blockTypesData.Count > 0)
                {
                    btnPlaceElements.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error creating family selection UI: {ex.Message}");
            }
        }
        
        private void BtnAnnotate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_externalEvent != null)
                {
                    _externalEvent.Raise();
                }
                else
                {
                    MessageBox.Show("Не удалось инициализировать внешнее событие Revit. Попробуйте перезапустить панель.", 
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при вызове аннотирования: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private List<string> GetAvailableFamilySymbols()
        {
            List<string> familySymbols = new List<string>();
            
            try
            {
                // Collect all FamilySymbol elements in the document
                FilteredElementCollector collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType();
                
                foreach (FamilySymbol symbol in collector)
                {
                    if (symbol != null && symbol.Family != null)
                    {
                        // Get full family name with type
                        string familyName = $"{symbol.Family.Name}: {symbol.Name}";
                        familySymbols.Add(familyName);
                    }
                }
                
                familySymbols.Sort();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error getting family symbols: {ex.Message}");
            }
            
            return familySymbols;
        }
        
        private string AutoSelectFamily(string blockTypeName)
        {
            string lowerBlockName = blockTypeName.ToLower();
            
            // Auto-selection rules based on keywords
            if (lowerBlockName.Contains("розетка"))
            {
                return "TSL_EF_т_СТ_н_IP20_Рзт_1P+N+PE";
            }
            else if (lowerBlockName.Contains("светильник"))
            {
                return "TSL_LF_э_ПЛ_Патрон (нагрузка)";
            }
            
            return null;
        }
        
        private string FindBestFamilyMatch(List<string> availableFamilies, string targetFamilyName)
        {
            if (string.IsNullOrEmpty(targetFamilyName))
                return null;
                
            // Try exact match with family name (before colon)
            foreach (var family in availableFamilies)
            {
                string familyName = family.Split(new[] { ": " }, StringSplitOptions.None)[0];
                if (familyName.Equals(targetFamilyName, StringComparison.OrdinalIgnoreCase))
                {
                    return family;
                }
            }
            
            // Try partial match
            foreach (var family in availableFamilies)
            {
                if (family.Contains(targetFamilyName))
                {
                    return family;
                }
            }
            
            return null;
        }
        
        private void BtnPlaceElements_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_blockTypesData.Count == 0)
                {
                    MessageBox.Show("Нет данных для размещения. Пожалуйста, сначала выполните анализ.", 
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                if (_placeElementsEvent != null)
                {
                    _placeElementsEvent.Raise();
                }
                else
                {
                    MessageBox.Show("Не удалось инициализировать внешнее событие Revit.", 
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void PlaceSingleBlockType(string blockTypeName)
        {
            System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] === METHOD ENTRY === blockTypeName: {blockTypeName}");
            
            try
            {
                System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] Inside try block");
                System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] _placeSingleBlockTypeEvent is null: {_placeSingleBlockTypeEvent == null}");
                
                if (!_blockTypesData.ContainsKey(blockTypeName))
                {
                    MessageBox.Show($"Тип блока '{blockTypeName}' не найден.", 
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                if (_placeSingleBlockTypeEvent == null)
                {
                    MessageBox.Show("Не удалось создать событие для размещения.\n\nПроверьте Output в Visual Studio для подробностей.", 
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // Get handler from OpenHubCommand
                // We need to set BlockTypeName through the static reference
                Commands.OpenHubCommand.SetBlockTypeNameForPlacement(blockTypeName);
                
                _placeSingleBlockTypeEvent.Raise();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] Exception: {ex.Message}");
                MessageBox.Show($"Ошибка при размещении: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public class AnnotateEventHandler : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                if (app.ActiveUIDocument == null) return;
                Document doc = app.ActiveUIDocument.Document;
                if (doc == null) return;

                try
                {
                    string logDirectory = @"C:\Users\Свеж как огурец\Desktop\Эксперимент Annotatix\logs";
                    
                    if (!System.IO.Directory.Exists(logDirectory))
                    {
                        MessageBox.Show("Log directory not found. Please run analysis first.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    var logFiles = System.IO.Directory.GetFiles(logDirectory, "*.txt")
                        .Where(f => System.IO.Path.GetFileName(f).Contains("_NAME.txt") ||
                                    System.IO.Path.GetFileName(f).Contains("DWG_Analysis_"))
                        .OrderByDescending(f => System.IO.File.GetLastWriteTime(f))
                        .ToList();
                    
                    if (logFiles.Count == 0)
                    {
                        MessageBox.Show("No analysis results found. Please run analysis first.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    string latestLogFile = logFiles[0];
                    var blocks = Commands.AnnotateBlocksCommand.ParseLogFileStatic(latestLogFile);
                    
                    if (blocks.Count == 0)
                    {
                        MessageBox.Show("No blocks found in analysis results.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    using (Transaction trans = new Transaction(doc, "Annotate DWG Blocks"))
                    {
                        trans.Start();
                        View activeView = doc.ActiveView;
                        
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        TextNoteType textNoteType = collector
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .FirstOrDefault(t => t.Name.Contains("GOST") || t.Name.Contains("Common"));
                        
                        if (textNoteType == null)
                        {
                            textNoteType = collector.OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();
                        }
                        
                        if (textNoteType == null)
                        {
                            MessageBox.Show("No text note type found in the document.", "Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            trans.RollBack();
                            return;
                        }
                        
                        int annotatedCount = 0;
                        foreach (var block in blocks)
                        {
                            try
                            {
                                double offsetY = -1.5; 
                                XYZ location = new XYZ(block.CenterX, block.CenterY + offsetY, 0);
                                TextNote textNote = TextNote.Create(doc, activeView.Id, location, block.Name, textNoteType.Id);
                                try { textNote.HorizontalAlignment = HorizontalTextAlignment.Center; } catch { }
                                annotatedCount++;
                            }
                            catch { continue; }
                        }
                        
                        trans.Commit();
                        // Success message removed - status is shown in status window
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}", "Annotation Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            public string GetName() => "AnnotateDWGBlocks";
        }
        
        public class PlaceElementsEventHandler : IExternalEventHandler
        {
            private static dwg2rvtPanel _activePanel;
            
            public PlaceElementsEventHandler()
            {
            }
            
            public static void SetActivePanel(dwg2rvtPanel panel)
            {
                _activePanel = panel;
            }
            
            public void Execute(UIApplication app)
            {
                if (_activePanel == null)
                {
                    MessageBox.Show("Панель не инициализирована.", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                if (app.ActiveUIDocument == null) return;
                Document doc = app.ActiveUIDocument.Document;
                if (doc == null) return;
                
                try
                {
                    if (_activePanel._blockTypesData.Count == 0)
                    {
                        MessageBox.Show("Нет данных для размещения.", "Предупреждение", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    _activePanel.Dispatcher.Invoke(() =>
                    {
                        _activePanel.btnPlaceElements.IsEnabled = false;
                        _activePanel.txtProgress.Text = "Размещение элементов...";
                    });
                    
                    int totalPlaced = 0;
                    int totalFailed = 0;
                    
                    using (Transaction trans = new Transaction(doc, "Place Family Instances"))
                    {
                        trans.Start();
                        
                        try
                        {
                            View activeView = doc.ActiveView;
                            Level level = activeView.GenLevel ?? GetFirstLevel(doc);
                            
                            if (level == null)
                            {
                                MessageBox.Show("Не удалось найти уровень для размещения.", "Ошибка", 
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                                trans.RollBack();
                                return;
                            }
                            
                            double offsetFromLevel = 500 / 304.8; // 500mm to feet
                            
                            foreach (var blockType in _activePanel._blockTypesData.Values)
                            {
                                string selectedFamily = null;
                                _activePanel.Dispatcher.Invoke(() =>
                                {
                                    if (blockType.FamilyComboBox != null && blockType.FamilyComboBox.SelectedItem != null)
                                    {
                                        selectedFamily = blockType.FamilyComboBox.SelectedItem.ToString();
                                    }
                                });
                                
                                if (string.IsNullOrEmpty(selectedFamily))
                                {
                                    continue;
                                }
                                
                                // Parse family and type names
                                // Format: "FamilyName: TypeName" where TypeName may contain colons
                                int firstColonIndex = selectedFamily.IndexOf(": ");
                                if (firstColonIndex < 0)
                                {
                                    _activePanel.UpdateStatus($"Неверный формат семейства: {selectedFamily}");
                                    continue;
                                }
                                
                                string familyName = selectedFamily.Substring(0, firstColonIndex);
                                string typeName = selectedFamily.Substring(firstColonIndex + 2);  // Skip ": "
                                
                                // Find the family symbol
                                FamilySymbol familySymbol = FindFamilySymbol(doc, familyName, typeName);
                                
                                if (familySymbol == null)
                                {
                                    _activePanel.UpdateStatus($"Семейство не найдено: {selectedFamily}");
                                    totalFailed += blockType.Instances.Count;
                                    continue;
                                }
                                
                                // Activate the symbol if not active
                                if (!familySymbol.IsActive)
                                {
                                    try
                                    {
                                        familySymbol.Activate();
                                        doc.Regenerate();
                                        _activePanel.UpdateStatus($"Семейство '{selectedFamily}' активировано.");
                                    }
                                    catch (Exception activateEx)
                                    {
                                        _activePanel.UpdateStatus($"Ошибка активации семейства '{selectedFamily}': {activateEx.Message}");
                                        totalFailed += blockType.Instances.Count;
                                        continue;
                                    }
                                }
                                
                                // Place instances at block coordinates
                                foreach (var instance in blockType.Instances)
                                {
                                    try
                                    {
                                        // Coordinates from log file are already in Revit coordinates (feet)
                                        // Use them directly, same as AnnotateBlocksCommand
                                        XYZ revitLocation = new XYZ(instance.CenterX, instance.CenterY, 
                                            level.Elevation + offsetFromLevel);
                                                                        
                                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(
                                            revitLocation, familySymbol, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                                        
                                        // Apply rotation if specified (for sockets)
                                        if (instance.RotationAngle > 0)
                                        {
                                            // Get the location point
                                            LocationPoint locationPoint = familyInstance.Location as LocationPoint;
                                            if (locationPoint != null)
                                            {
                                                // Convert degrees to radians
                                                double angleRadians = instance.RotationAngle * (Math.PI / 180.0);
                                                                                
                                                // Create axis line (vertical through the insertion point)
                                                Line axis = Line.CreateBound(
                                                    new XYZ(revitLocation.X, revitLocation.Y, revitLocation.Z),
                                                    new XYZ(revitLocation.X, revitLocation.Y, revitLocation.Z + 10));
                                                                                
                                                // Rotate the family instance
                                                locationPoint.Rotate(axis, angleRadians);
                                            }
                                        }
                                                                        
                                        totalPlaced++;
                                    }
                                    catch (Exception ex)
                                    {
                                        _activePanel.UpdateStatus($"Ошибка размещения {instance.Name} №{instance.Number}: {ex.Message}");
                                        totalFailed++;
                                    }
                                }
                                                                
                                _activePanel.UpdateStatus($"Размещено {blockType.Instances.Count} элементов типа '{blockType.BlockTypeName}'");
                            }
                            
                            trans.Commit();
                            
                            _activePanel.Dispatcher.Invoke(() =>
                            {
                                _activePanel.txtProgress.Text = $"Размещение завершено: {totalPlaced} успешно, {totalFailed} ошибок";
                            });
                            
                            // Success message removed - status is shown in status window
                        }
                        catch (Exception ex)
                        {
                            trans.RollBack();
                            throw;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _activePanel.Dispatcher.Invoke(() =>
                    {
                        _activePanel.txtProgress.Text = "Ошибка при размещении.";
                    });
                    MessageBox.Show($"Ошибка: {ex.Message}\n{ex.StackTrace}", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    _activePanel.Dispatcher.Invoke(() =>
                    {
                        _activePanel.btnPlaceElements.IsEnabled = true;
                    });
                }
            }
            
            private static FamilySymbol FindFamilySymbol(Document doc, string familyName, string typeName)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType();
                
                foreach (FamilySymbol symbol in collector)
                {
                    if (symbol.Family.Name == familyName && symbol.Name == typeName)
                    {
                        return symbol;
                    }
                }
                
                return null;
            }
            
            private static Level GetFirstLevel(Document doc)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level));
                
                return collector.FirstElement() as Level;
            }
            
            public string GetName() => "PlaceFamilyElements";
        }
        
        // Handler for placing a single block type
        public class PlaceSingleBlockTypeEventHandler : IExternalEventHandler
        {
            private dwg2rvtPanel _activePanel;
            
            public string BlockTypeName { get; set; }
            
            public PlaceSingleBlockTypeEventHandler()
            {
            }
            
            public void SetPanel(dwg2rvtPanel panel)
            {
                _activePanel = panel;
            }
            
            public void Execute(UIApplication uiApp)
            {
                try
                {
                    if (string.IsNullOrEmpty(BlockTypeName))
                    {
                        MessageBox.Show("Тип блока не указан.", 
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    
                    Document doc = uiApp.ActiveUIDocument.Document;
                    UIDocument uidoc = uiApp.ActiveUIDocument;
                    View activeView = doc.ActiveView;
                    
                    if (!_activePanel._blockTypesData.ContainsKey(BlockTypeName))
                    {
                        MessageBox.Show($"Тип блока '{BlockTypeName}' не найден.", 
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    
                    var blockType = _activePanel._blockTypesData[BlockTypeName];
                    
                    // Get selected family from combobox
                    string selectedFamily = null;
                    _activePanel.Dispatcher.Invoke(() =>
                    {
                        if (blockType.FamilyComboBox != null && blockType.FamilyComboBox.SelectedItem != null)
                        {
                            selectedFamily = blockType.FamilyComboBox.SelectedItem.ToString();
                        }
                    });
                    
                    if (string.IsNullOrEmpty(selectedFamily))
                    {
                        MessageBox.Show($"Выберите семейство для '{BlockTypeName}'.", 
                            "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    using (Transaction trans = new Transaction(doc, $"Place {BlockTypeName}"))
                    {
                        try
                        {
                            trans.Start();
                            
                            int totalPlaced = 0;
                            int totalFailed = 0;
                            
                            // Get level
                            Level level = GetFirstLevel(doc);
                            if (level == null)
                            {
                                MessageBox.Show("Не удалось найти уровень в проекте.", 
                                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                            
                            // Parse family name and type name
                            // Format: "FamilyName: TypeName" where TypeName may contain colons
                            int firstColonIndex = selectedFamily.IndexOf(": ");
                            if (firstColonIndex < 0)
                            {
                                _activePanel.UpdateStatus($"Неверный формат имени семейства: {selectedFamily}");
                                totalFailed += blockType.Instances.Count;
                                trans.RollBack();
                                return;
                            }
                            
                            string familyName = selectedFamily.Substring(0, firstColonIndex);
                            string typeName = selectedFamily.Substring(firstColonIndex + 2);  // Skip ": "
                            
                            // Find the family symbol
                            FamilySymbol familySymbol = FindFamilySymbol(doc, familyName, typeName);
                            
                            if (familySymbol == null)
                            {
                                _activePanel.UpdateStatus($"Семейство не найдено: {selectedFamily}");
                                MessageBox.Show($"Семейство не найдено: {selectedFamily}", 
                                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                totalFailed += blockType.Instances.Count;
                                trans.RollBack();
                                return;
                            }
                            
                            // Activate the symbol if not active
                            if (!familySymbol.IsActive)
                            {
                                try
                                {
                                    familySymbol.Activate();
                                    doc.Regenerate();
                                    _activePanel.UpdateStatus($"Семейство '{selectedFamily}' активировано.");
                                }
                                catch (Exception ex)
                                {
                                    _activePanel.UpdateStatus($"Ошибка активации семейства: {ex.Message}");
                                    MessageBox.Show($"Не удалось активировать семейство '{selectedFamily}': {ex.Message}", 
                                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                    trans.RollBack();
                                    return;
                                }
                            }
                            
                            double offsetFromLevel = 500 / 304.8; // 500mm to feet
                                                        
                            // Place instances at block coordinates
                            foreach (var instance in blockType.Instances)
                            {
                                try
                                {
                                    // Coordinates from log file are already in Revit coordinates (feet)
                                    // Use them directly, same as AnnotateBlocksCommand
                                    XYZ revitLocation = new XYZ(instance.CenterX, instance.CenterY, 
                                        level.Elevation + offsetFromLevel);
                                                                
                                    FamilyInstance familyInstance = doc.Create.NewFamilyInstance(
                                        revitLocation, familySymbol, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                                
                                    // Apply rotation if specified (for sockets)
                                    if (instance.RotationAngle > 0)
                                    {
                                        // Get the location point
                                        LocationPoint locationPoint = familyInstance.Location as LocationPoint;
                                        if (locationPoint != null)
                                        {
                                            // Convert degrees to radians
                                            double angleRadians = instance.RotationAngle * (Math.PI / 180.0);
                                                                        
                                            // Create axis line (vertical through the insertion point)
                                            Line axis = Line.CreateBound(
                                                new XYZ(revitLocation.X, revitLocation.Y, revitLocation.Z),
                                                new XYZ(revitLocation.X, revitLocation.Y, revitLocation.Z + 10));
                                                                        
                                            // Rotate the family instance
                                            locationPoint.Rotate(axis, angleRadians);
                                        }
                                    }
                                                                
                                    totalPlaced++;
                                }
                                catch (Exception ex)
                                {
                                    _activePanel.UpdateStatus($"Ошибка размещения {instance.Name} №{instance.Number}: {ex.Message}");
                                    totalFailed++;
                                }
                            }
                                                        
                            trans.Commit();
                                                        
                            _activePanel.Dispatcher.Invoke(() =>
                            {
                                _activePanel.txtProgress.Text = $"Размещено {totalPlaced} элементов '{BlockTypeName}'. Ошибок: {totalFailed}";
                            });
                            
                            // Success message removed - status is shown in status window
                        }
                        catch (Exception ex)
                        {
                            trans.RollBack();
                            MessageBox.Show($"Ошибка при размещении: {ex.Message}\n{ex.StackTrace}", 
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}\n{ex.StackTrace}", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            
            private static FamilySymbol FindFamilySymbol(Document doc, string familyName, string typeName)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType();
                
                foreach (FamilySymbol symbol in collector)
                {
                    if (symbol.Family.Name == familyName && symbol.Name == typeName)
                    {
                        return symbol;
                    }
                }
                
                return null;
            }
            
            private static Level GetFirstLevel(Document doc)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level));
                
                return collector.FirstElement() as Level;
            }
            
            public string GetName() => "PlaceSingleBlockType";
        }
    }
}
