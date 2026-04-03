using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace dwg2rvt.Module.UI
{
    public partial class dwg2rvtPanel : UserControl
    {
        private UIApplication _uiApp;
        private Document _doc;
        private List<ImportInstance> _importedFiles;
        private ExternalEvent _externalEvent;
        private ExternalEvent _placeElementsEvent;
        private ExternalEvent _placeSingleBlockTypeEvent;
        private ExternalEvent _placeAnnotationsEvent;
        private ExternalEvent _placeAnnotationsSingleEvent;
        private object _placeAnnotationsSingleHandler; // for setting block type name
        private Dictionary<string, BlockTypeInfo> _blockTypesData = new Dictionary<string, BlockTypeInfo>();
        private Dictionary<string, AnnotationConfig> _annotationConfigs = new Dictionary<string, AnnotationConfig>();
        private bool _annotationsPanelVisible = false;
        // ViewId captured at button-click time (UI thread) so handlers use the correct active view
        public ElementId PlacementViewId { get; set; }
        // List of floor plan views for the placement view combobox
        private List<View> _planViews = new List<View>();

        // Class to store block type information with family selection
        public class BlockTypeInfo
        {
            public string BlockTypeName { get; set; }
            public int Count { get; set; }
            public List<BlockInstanceData> Instances { get; set; } = new List<BlockInstanceData>();
            public string SelectedFamily { get; set; }
            public System.Windows.Controls.ComboBox FamilyComboBox { get; set; }
            public System.Windows.Controls.CheckBox CheckBox { get; set; }
        }

        public class BlockInstanceData
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public double CenterX { get; set; }
            public double CenterY { get; set; }
            public double RotationAngle { get; set; } = 0;  // Rotation angle in degrees
        }

        public dwg2rvtPanel(UIApplication uiApp, ExternalEvent annotateEvent, ExternalEvent placeElementsEvent = null, ExternalEvent placeSingleBlockTypeEvent = null, ExternalEvent placeAnnotationsEvent = null, ExternalEvent placeAnnotationsSingleEvent = null)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _externalEvent = annotateEvent;
            _placeElementsEvent = placeElementsEvent;
            _placeSingleBlockTypeEvent = placeSingleBlockTypeEvent;
            _placeAnnotationsEvent = placeAnnotationsEvent;
            _placeAnnotationsSingleEvent = placeAnnotationsSingleEvent;
            
            // Debug logging with timestamp to verify dynamic loading
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            PluginsManager.Core.DebugLogger.Log("[dwg2rvtPanel] *** PANEL INSTANCE CREATED ***");
            PluginsManager.Core.DebugLogger.Log("[dwg2rvtPanel] This proves the panel was NOT loaded at Revit startup");
            PluginsManager.Core.DebugLogger.Log("[dwg2rvtPanel] Panel created AFTER authentication");
            PluginsManager.Core.DebugLogger.Log($"[dwg2rvtPanel] annotateEvent: {annotateEvent != null}");
            PluginsManager.Core.DebugLogger.Log($"[dwg2rvtPanel] placeElementsEvent: {placeElementsEvent != null}");
            PluginsManager.Core.DebugLogger.Log($"[dwg2rvtPanel] placeSingleBlockTypeEvent: {placeSingleBlockTypeEvent != null}");
            
            // Set this panel as the active panel for PlaceElementsEventHandler
            PlaceElementsEventHandler.SetActivePanel(this);
            PlaceAnnotationsAllEventHandler.SetActivePanel(this);
            PlaceAnnotationsSingleEventHandler.SetActivePanel(this);

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
                // Load floor plan views for placement view selection
                LoadPlanViews();

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

        private void LoadPlanViews()
        {
            try
            {
                _planViews = new List<View>();
                cmbPlacementView.Items.Clear();

                // Collect all floor plan views that have a GenLevel
                var views = new FilteredElementCollector(_doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate &&
                                (v.ViewType == ViewType.FloorPlan || v.ViewType == ViewType.CeilingPlan) &&
                                v.GenLevel != null)
                    .OrderBy(v => v.Name)
                    .ToList();

                View activeView = _doc.ActiveView;
                int autoSelectIndex = 0;

                for (int i = 0; i < views.Count; i++)
                {
                    _planViews.Add(views[i]);
                    cmbPlacementView.Items.Add(views[i].Name);
                    if (activeView != null && views[i].Id == activeView.Id)
                        autoSelectIndex = i;
                }

                if (cmbPlacementView.Items.Count > 0)
                {
                    cmbPlacementView.SelectedIndex = autoSelectIndex;
                    // Also set directly in case SelectionChanged fires before _planViews is ready
                    PlacementViewId = _planViews[autoSelectIndex].Id;
                }
            }
            catch (Exception ex)
            {
                // Non-critical: placement view selection just won't be available
                System.Diagnostics.Debug.WriteLine($"[LoadPlanViews] Error: {ex.Message}");
            }
        }

        private void CmbPlacementView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int idx = cmbPlacementView.SelectedIndex;
                if (idx >= 0 && idx < _planViews.Count)
                    PlacementViewId = _planViews[idx].Id;
            }
            catch { }
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

                // Get logging preference (disabled by default since checkbox was removed from UI)
                bool enableLogging = false;

                // Perform analysis using new AnalyzeByBlockName class
                dwg2rvt.Module.Core.AnalyzeByBlockName analyzer = new dwg2rvt.Module.Core.AnalyzeByBlockName(doc);
                var analysisResult = analyzer.Analyze(selectedImport, UpdateStatus, enableLogging);

                // Store the active view ID so placement handlers use the correct level
                if (analysisResult.Success)
                    analysisResult.ActiveViewId = doc.ActiveView?.Id?.IntegerValue ?? -1;

                // --- Crop filter ---
                if (analysisResult.Success && chkCropFilter.IsChecked == true)
                {
                    View activeView = doc.ActiveView;
                    if (activeView != null && activeView.CropBoxActive && activeView.CropBox != null)
                    {
                        BoundingBoxXYZ cropBox = activeView.CropBox;
                        // CropBox is in view-local coordinates; transform to world coordinates
                        Transform t = cropBox.Transform;
                        XYZ minWorld = t.OfPoint(cropBox.Min);
                        XYZ maxWorld = t.OfPoint(cropBox.Max);
                        double xMin = Math.Min(minWorld.X, maxWorld.X);
                        double xMax = Math.Max(minWorld.X, maxWorld.X);
                        double yMin = Math.Min(minWorld.Y, maxWorld.Y);
                        double yMax = Math.Max(minWorld.Y, maxWorld.Y);

                        int beforeCount = analysisResult.BlockData.Count;
                        analysisResult.BlockData.RemoveAll(b =>
                            b.CenterX < xMin || b.CenterX > xMax ||
                            b.CenterY < yMin || b.CenterY > yMax);

                        // Rebuild BlocksByType from filtered list
                        analysisResult.BlocksByType.Clear();
                        foreach (var bd in analysisResult.BlockData)
                        {
                            if (!analysisResult.BlocksByType.ContainsKey(bd.Name))
                                analysisResult.BlocksByType[bd.Name] = new List<dwg2rvt.Module.Core.BlockData>();
                            analysisResult.BlocksByType[bd.Name].Add(bd);
                        }

                        int afterCount = analysisResult.BlockData.Count;
                        UpdateStatus($"Crop filter: {beforeCount - afterCount} blocks removed outside crop region ({afterCount} remain).");

                        // Re-store filtered result in cache
                        dwg2rvt.Module.Core.AnalysisDataCache.StoreAnalysisResult(analysisResult);
                    }
                    else
                    {
                        UpdateStatus("Crop filter: view has no active crop box — showing all blocks.");
                    }
                }

                if (analysisResult.Success)
                {
                    txtProgress.Text = "Analysis complete!";
                    txtStatus.Text += $"Analysis data stored in memory cache\n";
                    
                    // Display logging status
                    if (string.IsNullOrEmpty(analysisResult.LogFilePath))
                    {
                        txtStatus.Text += "Logging disabled - data in memory only\n";
                    }
                    else
                    {
                        txtStatus.Text += $"Log file: {analysisResult.LogFilePath}\n";
                    }
                    
                    // Load block data from cache instead of parsing log file
                    LoadAnalysisResultsFromCache();
                    
                    // Create dynamic UI for family selection
                    CreateFamilySelectionUI();

                    // Refresh annotation UI if panel is already visible
                    if (_annotationsPanelVisible)
                        CreateAnnotationUI();

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
        
        private void LoadAnalysisResultsFromCache()
        {
            try
            {
                // Clear previous data
                _blockTypesData.Clear();
                
                // Get analysis result from cache
                var analysisResult = dwg2rvt.Module.Core.AnalysisDataCache.GetLastAnalysisResult();
                
                if (analysisResult == null || !analysisResult.Success || analysisResult.BlockData == null)
                {
                    UpdateStatus("No analysis data found in cache.");
                    return;
                }
                
                // Group blocks by type
                foreach (var blockType in analysisResult.BlocksByType)
                {
                    string blockTypeName = blockType.Key;
                    var instances = blockType.Value;
                    
                    if (!_blockTypesData.ContainsKey(blockTypeName))
                    {
                        _blockTypesData[blockTypeName] = new BlockTypeInfo
                        {
                            BlockTypeName = blockTypeName,
                            Count = 0
                        };
                    }
                    
                    // Add instances
                    foreach (var blockData in instances)
                    {
                        _blockTypesData[blockTypeName].Instances.Add(new BlockInstanceData
                        {
                            Name = blockData.Name,
                            Number = blockData.Number,
                            CenterX = blockData.CenterX,
                            CenterY = blockData.CenterY,
                            RotationAngle = blockData.RotationAngle
                        });
                        
                        _blockTypesData[blockTypeName].Count++;
                    }
                }
                
                UpdateStatus($"Loaded {_blockTypesData.Count} block types from cache.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error loading analysis results from cache: {ex.Message}");
            }
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
                
                // Create rows with new table design (matching XAML)
                foreach (var blockType in _blockTypesData.Values)
                {
                    // Create table row using Grid with proper column definitions
                    var rowGrid = new System.Windows.Controls.Grid
                    {
                        Background = System.Windows.Media.Brushes.White
                    };
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                    
                    // Column 0: Checkbox in Border
                    var checkboxBorder = new Border
                    {
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(224, 224, 224)),
                        BorderThickness = new Thickness(0, 0, 1, 1)
                    };
                    var checkbox = new System.Windows.Controls.CheckBox
                    {
                        IsChecked = true,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center
                    };
                    checkboxBorder.Child = checkbox;
                    blockType.CheckBox = checkbox;
                    System.Windows.Controls.Grid.SetColumn(checkboxBorder, 0);
                    rowGrid.Children.Add(checkboxBorder);
                    
                    // Column 1: Block name in Border
                    var nameBorder = new Border
                    {
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(224, 224, 224)),
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        Padding = new Thickness(6, 4, 6, 4)
                    };
                    var blockNameText = new TextBlock
                    {
                        Text = blockType.BlockTypeName,
                        FontSize = 11,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center
                    };
                    nameBorder.Child = blockNameText;
                    System.Windows.Controls.Grid.SetColumn(nameBorder, 1);
                    rowGrid.Children.Add(nameBorder);
                    
                    // Column 2: Count in Border
                    var countBorder = new Border
                    {
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(224, 224, 224)),
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        Padding = new Thickness(6, 4, 6, 4)
                    };
                    var countText = new TextBlock
                    {
                        Text = blockType.Count.ToString(),
                        FontSize = 11,
                        TextAlignment = TextAlignment.Center,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center
                    };
                    countBorder.Child = countText;
                    System.Windows.Controls.Grid.SetColumn(countBorder, 2);
                    rowGrid.Children.Add(countBorder);
                    
                    // Column 3: Family ComboBox in Border
                    var comboBorder = new Border
                    {
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(224, 224, 224)),
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        Padding = new Thickness(2, 2, 2, 2)
                    };
                    
                    var familyComboBox = new System.Windows.Controls.ComboBox
                    {
                        Height = 22,
                        FontSize = 11,
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(173, 173, 173)),
                        IsEditable = true,
                        IsTextSearchEnabled = false,
                        MaxDropDownHeight = 300
                    };
                    
                    // Store all families for substring filtering
                    var allFamiliesForComboBox = new List<string>(availableFamilies);
                    
                    // Add available families to combobox
                    foreach (var family in availableFamilies)
                    {
                        familyComboBox.Items.Add(family);
                    }
                    
                    // Create timer for delayed filtering
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
                    
                    comboBorder.Child = familyComboBox;
                    System.Windows.Controls.Grid.SetColumn(comboBorder, 3);
                    rowGrid.Children.Add(comboBorder);
                    
                    // Column 4: Action link in Border
                    var actionBorder = new Border
                    {
                        BorderBrush = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(224, 224, 224)),
                        BorderThickness = new Thickness(0, 0, 0, 1),
                        Padding = new Thickness(6, 4, 6, 4)
                    };
                    var actionLink = new TextBlock
                    {
                        Text = "Расставить",
                        FontSize = 11,
                        Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0, 120, 212)),
                        Cursor = System.Windows.Input.Cursors.Hand,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        Tag = blockType.BlockTypeName
                    };
                    
                    actionLink.MouseLeftButtonUp += (s, args) =>
                    {
                        PlaceSingleBlockType(blockType.BlockTypeName);
                    };
                    
                    actionBorder.Child = actionLink;
                    System.Windows.Controls.Grid.SetColumn(actionBorder, 4);
                    rowGrid.Children.Add(actionBorder);
                    
                    statusStackPanel.Children.Add(rowGrid);
                }
                
                // Enable Place Elements button if we have data
                if (_blockTypesData.Count > 0)
                {
                    btnPlaceAll.IsEnabled = true;
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
                    string viewInfo = PlacementViewId != null ? 
                        (_planViews.FirstOrDefault(v => v.Id == PlacementViewId)?.Name ?? "?") : "не выбран";
                    UpdateStatus($"Запуск размещения... Вид: {viewInfo}");
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
                
                // Get handler from OpenHubCommand and set block type name
                // We use reflection because we can't reference PluginsManager directly
                try
                {
                    var pluginsManagerAssembly = System.Reflection.Assembly.Load("PluginsManager");
                    var commandType = pluginsManagerAssembly.GetType("PluginsManager.Commands.OpenDwg2rvtPanelCommand");
                    if (commandType != null)
                    {
                        var method = commandType.GetMethod("SetBlockTypeNameForPlacement", 
                            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                        if (method != null)
                        {
                            method.Invoke(null, new object[] { blockTypeName });
                            System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] BlockTypeName set to: {blockTypeName}");
                        }
                    }
                }
                catch (Exception reflectionEx)
                {
                    System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] Reflection error: {reflectionEx.Message}");
                    MessageBox.Show($"Не удалось установить тип блока: {reflectionEx.Message}", 
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // PlacementViewId is already set from the combobox selection
                _placeSingleBlockTypeEvent.Raise();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PlaceSingleBlockType] Exception: {ex.Message}");
                MessageBox.Show($"Ошибка при размещении: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        // === NEW UI EVENT HANDLERS ===
        
        // Callback to go back to hub
        public Action OnBackToHub { get; set; }
        
        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            OnBackToHub?.Invoke();
        }
        
        // Stub button - select all checkboxes
        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var blockType in _blockTypesData.Values)
            {
                if (blockType.CheckBox != null)
                {
                    blockType.CheckBox.IsChecked = true;
                }
            }
            txtProgress.Text = "Выбраны все элементы";
        }
        
        // Stub button - deselect all checkboxes
        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var blockType in _blockTypesData.Values)
            {
                if (blockType.CheckBox != null)
                {
                    blockType.CheckBox.IsChecked = false;
                }
            }
            txtProgress.Text = "Выделение снято";
        }
        
        // Stub button - place selected elements
        private void BtnPlaceSelected_Click(object sender, RoutedEventArgs e)
        {
            // Get list of selected block types
            var selectedBlockTypes = _blockTypesData.Values
                .Where(bt => bt.CheckBox?.IsChecked == true)
                .Select(bt => bt.BlockTypeName)
                .ToList();
            
            if (selectedBlockTypes.Count == 0)
            {
                MessageBox.Show("Не выбрано ни одного элемента", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            // Place selected block types
            try
            {
                foreach (var blockTypeName in selectedBlockTypes)
                {
                    PlaceSingleBlockType(blockTypeName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при размещении:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("DWG2RVT - Конвертация элементов из DWG\n\n" +
                "1. Выберите импортированный DWG файл\n" +
                "2. Нажмите 'Анализировать'\n" +
                "3. Настройте соответствия семейств\n" +
                "4. Нажмите 'Расставить все элементы'\n\n" +
                "По вопросам: support@annotatix.ai",
                "Помощь", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        private void BtnReference_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://annotatix.ai/docs/dwg2rvt");
            }
            catch
            {
                MessageBox.Show("Откройте ссылку: https://annotatix.ai/docs/dwg2rvt", 
                    "Справка", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // === ANNOTATION CONFIG ===

        public class AnnotationConfig
        {
            public string BlockTypeName { get; set; }
            public string Prefix       { get; set; } = "";
            public string Main         { get; set; } = ""; // empty = auto-numbering
            public string Suffix       { get; set; } = "";
            // UI fields
            public System.Windows.Controls.TextBox PrefixBox { get; set; }
            public System.Windows.Controls.TextBox MainBox   { get; set; }
            public System.Windows.Controls.TextBox SuffixBox { get; set; }
        }

        /// <summary>
        /// Rebuilds the annotation UI table from current _blockTypesData.
        /// Called after analysis and every time the panel is shown.
        /// </summary>
        private void CreateAnnotationUI()
        {
            annotationsStackPanel.Children.Clear();

            if (_blockTypesData.Count == 0)
            {
                var hint = new TextBlock
                {
                    Text = "Сначала выполните анализ блоков.",
                    FontSize = 11,
                    Foreground = System.Windows.Media.Brushes.Gray,
                    Margin = new Thickness(8)
                };
                annotationsStackPanel.Children.Add(hint);
                return;
            }

            // --- header row ---
            var headerGrid = new System.Windows.Controls.Grid { Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 240, 240)) };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });

            string[] headers = { "Наименование блока", "Префикс", "Основная часть", "Суффикс", "Действие" };
            for (int i = 0; i < headers.Length; i++)
            {
                var b = new Border { BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(208, 208, 208)), BorderThickness = new Thickness(0, 0, i < headers.Length - 1 ? 1 : 0, 1), Padding = new Thickness(6, 4, 6, 4) };
                b.Child = new TextBlock { Text = headers[i], FontSize = 11, FontWeight = FontWeights.SemiBold };
                System.Windows.Controls.Grid.SetColumn(b, i);
                headerGrid.Children.Add(b);
            }
            annotationsStackPanel.Children.Add(headerGrid);

            // --- data rows ---
            foreach (var blockType in _blockTypesData.Values)
            {
                if (!_annotationConfigs.ContainsKey(blockType.BlockTypeName))
                    _annotationConfigs[blockType.BlockTypeName] = new AnnotationConfig { BlockTypeName = blockType.BlockTypeName };

                var cfg = _annotationConfigs[blockType.BlockTypeName];

                var rowGrid = new System.Windows.Controls.Grid { Background = System.Windows.Media.Brushes.White };
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });

                // col 0 - name
                var nameBorder = new Border { BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(224, 224, 224)), BorderThickness = new Thickness(0, 0, 1, 1), Padding = new Thickness(6, 4, 6, 4) };
                nameBorder.Child = new TextBlock { Text = blockType.BlockTypeName, FontSize = 11, VerticalAlignment = System.Windows.VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap };
                System.Windows.Controls.Grid.SetColumn(nameBorder, 0);
                rowGrid.Children.Add(nameBorder);

                // col 1 - prefix
                var prefixBox = new System.Windows.Controls.TextBox { Text = cfg.Prefix, FontSize = 11, BorderThickness = new Thickness(0), Padding = new Thickness(4, 2, 4, 2), VerticalContentAlignment = System.Windows.VerticalAlignment.Center };
                prefixBox.TextChanged += (s, e2) => cfg.Prefix = prefixBox.Text;
                cfg.PrefixBox = prefixBox;
                var prefixBorder = new Border { BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(224, 224, 224)), BorderThickness = new Thickness(0, 0, 1, 1) };
                prefixBorder.Child = prefixBox;
                System.Windows.Controls.Grid.SetColumn(prefixBorder, 1);
                rowGrid.Children.Add(prefixBorder);

                // col 2 - main
                var mainBox = new System.Windows.Controls.TextBox { Text = cfg.Main, FontSize = 11, BorderThickness = new Thickness(0), Padding = new Thickness(4, 2, 4, 2), VerticalContentAlignment = System.Windows.VerticalAlignment.Center };
                var mainPlaceholder = new TextBlock { Text = "авто-нумерация", FontSize = 10, Foreground = System.Windows.Media.Brushes.LightGray, IsHitTestVisible = false, Margin = new Thickness(5, 3, 0, 0) };
                mainBox.TextChanged += (s, e2) =>
                {
                    cfg.Main = mainBox.Text;
                    mainPlaceholder.Visibility = string.IsNullOrEmpty(mainBox.Text) ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                };
                cfg.MainBox = mainBox;
                var mainGrid = new System.Windows.Controls.Grid();
                mainGrid.Children.Add(mainBox);
                mainGrid.Children.Add(mainPlaceholder);
                var mainBorder = new Border { BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(224, 224, 224)), BorderThickness = new Thickness(0, 0, 1, 1) };
                mainBorder.Child = mainGrid;
                System.Windows.Controls.Grid.SetColumn(mainBorder, 2);
                rowGrid.Children.Add(mainBorder);

                // col 3 - suffix
                var suffixBox = new System.Windows.Controls.TextBox { Text = cfg.Suffix, FontSize = 11, BorderThickness = new Thickness(0), Padding = new Thickness(4, 2, 4, 2), VerticalContentAlignment = System.Windows.VerticalAlignment.Center };
                suffixBox.TextChanged += (s, e2) => cfg.Suffix = suffixBox.Text;
                cfg.SuffixBox = suffixBox;
                var suffixBorder = new Border { BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(224, 224, 224)), BorderThickness = new Thickness(0, 0, 1, 1) };
                suffixBorder.Child = suffixBox;
                System.Windows.Controls.Grid.SetColumn(suffixBorder, 3);
                rowGrid.Children.Add(suffixBorder);

                // col 4 - action
                var captureName = blockType.BlockTypeName;
                var actionBorder = new Border { BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(224, 224, 224)), BorderThickness = new Thickness(0, 0, 0, 1), Padding = new Thickness(6, 4, 6, 4) };
                var actionLink = new TextBlock { Text = "Расставить", FontSize = 11, Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212)), Cursor = System.Windows.Input.Cursors.Hand, VerticalAlignment = System.Windows.VerticalAlignment.Center };
                actionLink.MouseLeftButtonUp += (s, e2) => PlaceAnnotationsSingleBlockType(captureName);
                actionBorder.Child = actionLink;
                System.Windows.Controls.Grid.SetColumn(actionBorder, 4);
                rowGrid.Children.Add(actionBorder);

                annotationsStackPanel.Children.Add(rowGrid);
            }

            // --- Расставить все button ---
            var placeAllBtn = new Button
            {
                Content = "Расставить все",
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                Margin = new Thickness(6, 6, 6, 6),
                Padding = new Thickness(10, 4, 10, 4),
                FontSize = 11
            };
            placeAllBtn.Click += (s, e2) => PlaceAnnotationsAll();
            annotationsStackPanel.Children.Add(placeAllBtn);
        }

        private void BtnToggleAnnotations_Click(object sender, RoutedEventArgs e)
        {
            _annotationsPanelVisible = !_annotationsPanelVisible;
            if (_annotationsPanelVisible)
            {
                CreateAnnotationUI();
                annotationsPanel.Visibility = System.Windows.Visibility.Visible;
                txtAnnotArrow.Text = " \u2227";
            }
            else
            {
                annotationsPanel.Visibility = System.Windows.Visibility.Collapsed;
                txtAnnotArrow.Text = " ∨";
            }
        }

        private void PlaceAnnotationsSingleBlockType(string blockTypeName)
        {
            if (!_annotationConfigs.ContainsKey(blockTypeName))
            {
                MessageBox.Show($"Конфигурация аннотаций для '{blockTypeName}' не найдена.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (_placeAnnotationsSingleEvent == null)
            {
                MessageBox.Show("Событие для размещения аннотаций не инициализировано.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            // Pass block type name to handler via stored reference
            try
            {
                if (_placeAnnotationsSingleHandler != null)
                {
                    var prop = _placeAnnotationsSingleHandler.GetType().GetProperty("BlockTypeName");
                    prop?.SetValue(_placeAnnotationsSingleHandler, blockTypeName);
                }
            }
            catch { }
            _placeAnnotationsSingleEvent.Raise();
        }

        private void PlaceAnnotationsAll()
        {
            if (_blockTypesData.Count == 0)
            {
                MessageBox.Show("Нет данных блоков. Сначала выполните анализ.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (_placeAnnotationsEvent == null)
            {
                MessageBox.Show("Событие для размещения аннотаций не инициализировано.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            _placeAnnotationsEvent.Raise();
        }

        /// <summary>
        /// Exposes annotation configs to EventHandlers.
        /// </summary>
        public Dictionary<string, AnnotationConfig> AnnotationConfigs => _annotationConfigs;

        public void SetAnnotationsSingleHandler(object handler)
        {
            _placeAnnotationsSingleHandler = handler;
        }

        // === ANNOTATION EVENT HANDLERS ===

        public class PlaceAnnotationsAllEventHandler : IExternalEventHandler
        {
            private static dwg2rvtPanel _activePanel;
            public static void SetActivePanel(dwg2rvtPanel panel) { _activePanel = panel; }

            public void Execute(UIApplication app)
            {
                if (_activePanel == null) return;
                Document doc = app.ActiveUIDocument?.Document;
                if (doc == null) return;

                var blockNames = new List<string>(_activePanel._blockTypesData.Keys);
                PlaceAnnotationsForBlocks(doc, _activePanel, blockNames, app.ActiveUIDocument.ActiveView);
            }
            public string GetName() => "PlaceAnnotationsAll";
        }

        public class PlaceAnnotationsSingleEventHandler : IExternalEventHandler
        {
            private static dwg2rvtPanel _activePanel;
            public static void SetActivePanel(dwg2rvtPanel panel) { _activePanel = panel; }

            public string BlockTypeName { get; set; }

            public void Execute(UIApplication app)
            {
                if (_activePanel == null || string.IsNullOrEmpty(BlockTypeName)) return;
                Document doc = app.ActiveUIDocument?.Document;
                if (doc == null) return;

                PlaceAnnotationsForBlocks(doc, _activePanel, new List<string> { BlockTypeName }, app.ActiveUIDocument.ActiveView);
            }
            public string GetName() => "PlaceAnnotationsSingle";
        }

        /// <summary>
        /// Core logic: places TextNotes for the given block names using their AnnotationConfig.
        /// </summary>
        private static void PlaceAnnotationsForBlocks(Document doc, dwg2rvtPanel panel, List<string> blockNames, View activeView)
        {
            // Pick first available TextNoteType
            TextNoteType textType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .OrderBy(t => t.Name)
                .FirstOrDefault();

            if (textType == null)
            {
                MessageBox.Show("В документе не найден ни один тип текстового примечания.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            int totalPlaced = 0;
            int totalErrors = 0;

            using (Transaction trans = new Transaction(doc, "Place DWG Annotations"))
            {
                trans.Start();
                try
                {
                    foreach (var blockName in blockNames)
                    {
                        if (!panel._blockTypesData.ContainsKey(blockName)) continue;
                        var blockTypeInfo = panel._blockTypesData[blockName];

                        AnnotationConfig cfg;
                        if (!panel.AnnotationConfigs.TryGetValue(blockName, out cfg))
                            cfg = new AnnotationConfig { BlockTypeName = blockName };

                        bool autoNumber = string.IsNullOrEmpty(cfg.Main);

                        // Sort instances left-to-right, top-to-bottom for stable numbering
                        var sorted = blockTypeInfo.Instances
                            .OrderByDescending(i => i.CenterY)
                            .ThenBy(i => i.CenterX)
                            .ToList();

                        for (int idx = 0; idx < sorted.Count; idx++)
                        {
                            var inst = sorted[idx];
                            string mainPart = autoNumber ? (idx + 1).ToString() : cfg.Main;
                            string text = cfg.Prefix + mainPart + cfg.Suffix;
                            if (string.IsNullOrEmpty(text)) text = blockName;

                            try
                            {
                                // Offset slightly to the right of the block center
                                double offsetX = 1.0; // ~300 mm
                                XYZ location = new XYZ(inst.CenterX + offsetX, inst.CenterY, 0);
                                TextNote.Create(doc, activeView.Id, location, text, textType.Id);
                                totalPlaced++;
                            }
                            catch
                            {
                                totalErrors++;
                            }
                        }
                    }
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    MessageBox.Show($"Ошибка при расстановке аннотаций:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            panel.Dispatcher.Invoke(() =>
            {
                panel.txtProgress.Text = $"Аннотации размещены: {totalPlaced}, ошибок: {totalErrors}";
            });
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
                    // Get data from cache instead of log files
                    var analysisResult = dwg2rvt.Module.Core.AnalysisDataCache.GetLastAnalysisResult();
                    
                    if (analysisResult == null || !analysisResult.Success)
                    {
                        MessageBox.Show("No analysis data found in cache. Please run analysis first.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    if (analysisResult.BlockData == null || analysisResult.BlockData.Count == 0)
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
                        foreach (var block in analysisResult.BlockData)
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
                        _activePanel.btnPlaceAll.IsEnabled = false;
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

                            // Get level from the view selected in the combobox
                            Level level = null;
                            if (_activePanel?.PlacementViewId != null)
                            {
                                var selectedView = doc.GetElement(_activePanel.PlacementViewId) as View;
                                level = selectedView?.GenLevel;
                            }
                            if (level == null) level = activeView?.GenLevel;
                            if (level == null) level = GetFirstLevel(doc);

                            if (level == null)
                            {
                                MessageBox.Show("Не удалось определить уровень для размещения.", "Ошибка",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                                trans.RollBack();
                                return;
                            }

                            _activePanel.UpdateStatus($"Размещение на уровень: {level.Name}");
                            // Place directly at level elevation (offset = 0).
                            // Z = level.Elevation → offset-from-level = 0mm.
                            double offsetFromLevel = 0;
                            
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
                                        // Z semantics depend on FamilyPlacementType:
                                        // - OneLevelBased: Z = offset from level (feet)
                                        // - WorkPlaneBased: Z is ignored; set offset via parameter after creation
                                        // - TwoLevelsBased: Z = absolute world Z
                                        var fpt = familySymbol.Family.FamilyPlacementType;
                                        bool zIsAbsolute = fpt == FamilyPlacementType.TwoLevelsBased;
                                        const double offsetMm = 500.0;
                                        double placementZ = zIsAbsolute
                                            ? level.Elevation + offsetMm / 304.8
                                            : offsetMm / 304.8;

                                        XYZ revitLocation = new XYZ(instance.CenterX, instance.CenterY, placementZ);
                                                                        
                                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(
                                            revitLocation, familySymbol, level,
                                            Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                        // For WorkPlaneBased families Z is ignored by NewFamilyInstance.
                                        // Set the elevation offset explicitly via parameter.
                                        if (fpt == FamilyPlacementType.WorkPlaneBased)
                                        {
                                            double offsetFt = offsetMm / 304.8;
                                            // Try standard BIPs first, then fall back to named parameter
                                            var offsetParam =
                                                familyInstance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                                                ?? familyInstance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                                                ?? familyInstance.LookupParameter("АДСК_Размер_Смещение от уровня")
                                                ?? familyInstance.LookupParameter("TSL_Отметка от нуля_СП");
                                            offsetParam?.Set(offsetFt);
                                        }
                                                                        
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
                        _activePanel.btnPlaceAll.IsEnabled = true;
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

            /// <summary>
            /// Resolves the correct Level for family placement.
            /// Priority: (1) ImportInstance base level from cache, (2) saved view's GenLevel, (3) first level.
            /// </summary>
            private static Level GetLevelForPlacement(Document doc, View currentActiveView, dwg2rvtPanel panel)
            {
                // 1. GenLevel of the view selected in the placement combobox (explicit user choice)
                try
                {
                    if (panel?.PlacementViewId != null)
                    {
                        var selectedView = doc.GetElement(panel.PlacementViewId) as View;
                        if (selectedView?.GenLevel != null) return selectedView.GenLevel;
                    }
                }
                catch { }
                // 2. GenLevel of the current active view
                if (currentActiveView?.GenLevel != null) return currentActiveView.GenLevel;
                // 3. Last resort
                return GetFirstLevel(doc);
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
                            
                            // Get level from the view selected in the combobox
                            Level level = null;
                            if (_activePanel?.PlacementViewId != null)
                            {
                                var selectedView = doc.GetElement(_activePanel.PlacementViewId) as View;
                                level = selectedView?.GenLevel;
                            }
                            if (level == null) level = activeView?.GenLevel;
                            if (level == null) level = GetFirstLevel(doc);

                            if (level == null)
                            {
                                MessageBox.Show("Не удалось определить уровень для размещения.", 
                                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                            
                            _activePanel.UpdateStatus($"Размещение на уровень: {level.Name}");
                            
                            // Parse family name and type name
                            // Format: "FamilyName: TypeName" where TypeName may contain colons
                            int firstColonIndex = selectedFamily.IndexOf(": ");
                            if (firstColonIndex < 0)
                            {
                                _activePanel.UpdateStatus($"Неверный формат имени семейства: {selectedFamily}");
                                totalFailed += blockType.Instances.Count;
                                trans.Commit();
                                MessageBox.Show($"Неверный формат имени семейства: {selectedFamily}\n\nОжидается формат: 'ИмяСемейства: ИмяТипа'", 
                                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }
                            
                            string familyName = selectedFamily.Substring(0, firstColonIndex);
                            string typeName = selectedFamily.Substring(firstColonIndex + 2);  // Skip ": "
                            
                            // Find the family symbol
                            FamilySymbol familySymbol = FindFamilySymbol(doc, familyName, typeName);
                            
                            if (familySymbol == null)
                            {
                                _activePanel.UpdateStatus($"Семейство не найдено: {selectedFamily}");
                                _activePanel.UpdateStatus($"Пропущено {blockType.Instances.Count} экземпляров типа '{BlockTypeName}'");
                                totalFailed += blockType.Instances.Count;
                                trans.Commit();
                                
                                // Show warning but allow process to complete
                                MessageBox.Show($"Семейство не найдено: {selectedFamily}\n\nПропущено {blockType.Instances.Count} экземпляров.\n\nПроверьте, что семейство загружено в проект.", 
                                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                            
                            // Place directly at level elevation (offset = 0).
                            // Z = level.Elevation → offset-from-level = 0mm.
                            double offsetFromLevel = 0;
                                                        
                            // Place instances at block coordinates
                            foreach (var instance in blockType.Instances)
                            {
                                try
                                {
                                    // For NewFamilyInstance(XYZ, symbol, level, StructuralType):
                                    // Z component = OFFSET FROM LEVEL in feet (not absolute world Z).
                                    // For WorkPlaneBased: Z is ignored, set offset via parameter after creation.
                                    var fpt2 = familySymbol.Family.FamilyPlacementType;
                                    bool zIsAbsolute2 = fpt2 == FamilyPlacementType.TwoLevelsBased;
                                    const double offsetMm = 500.0;
                                    double placementZ = zIsAbsolute2
                                        ? level.Elevation + offsetMm / 304.8
                                        : offsetMm / 304.8;

                                    XYZ revitLocation = new XYZ(instance.CenterX, instance.CenterY, placementZ);
                                                                
                                    FamilyInstance familyInstance = doc.Create.NewFamilyInstance(
                                        revitLocation, familySymbol, level,
                                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    // For WorkPlaneBased: set offset explicitly via parameter
                                    if (fpt2 == FamilyPlacementType.WorkPlaneBased)
                                    {
                                        var offsetParam =
                                            familyInstance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                                            ?? familyInstance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                                            ?? familyInstance.LookupParameter("АДСК_Размер_Смещение от уровня")
                                            ?? familyInstance.LookupParameter("TSL_Отметка от нуля_СП");
                                        offsetParam?.Set(offsetMm / 304.8);
                                    }
                                                                
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

            private static Level GetLevelForPlacement(Document doc, View currentActiveView, dwg2rvtPanel panel)
            {
                // 1. GenLevel of the view selected in the placement combobox (explicit user choice)
                try
                {
                    if (panel?.PlacementViewId != null)
                    {
                        var selectedView = doc.GetElement(panel.PlacementViewId) as View;
                        if (selectedView?.GenLevel != null) return selectedView.GenLevel;
                    }
                }
                catch { }
                // 2. GenLevel of the current active view
                if (currentActiveView?.GenLevel != null) return currentActiveView.GenLevel;
                // 3. Last resort
                return GetFirstLevel(doc);
            }
            
            public string GetName() => "PlaceSingleBlockType";
        }
    }
}
