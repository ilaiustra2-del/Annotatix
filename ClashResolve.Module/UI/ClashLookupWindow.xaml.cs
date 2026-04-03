using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using ClashResolve.Module.Core;

namespace ClashResolve.Module.UI
{
    public partial class ClashLookupWindow : Window
    {
        // ----------------------------------------------------------------
        // Singleton
        // ----------------------------------------------------------------
        private static ClashLookupWindow _instance;
        public static ClashLookupWindow GetOrCreate()
        {
            if (_instance == null || !_instance.IsLoaded)
                _instance = new ClashLookupWindow();
            return _instance;
        }

        private static readonly string[] AllMepTypes = { "RectDuct", "RoundDuct", "Pipe" };

        // ---- colours (match the Revit options bar palette) ----
        private static readonly SolidColorBrush BrBg      = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x2D));
        private static readonly SolidColorBrush BrCell    = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));
        private static readonly SolidColorBrush BrRow     = new SolidColorBrush(Color.FromRgb(0x25, 0x25, 0x25));
        private static readonly SolidColorBrush BrRowAlt  = new SolidColorBrush(Color.FromRgb(0x2A, 0x2A, 0x2A));
        private static readonly SolidColorBrush BrHdr     = new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C));
        private static readonly SolidColorBrush BrHdrGrp  = new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)); // same as header, no blue tint
        private static readonly SolidColorBrush BrBorder  = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55));
        private static readonly SolidColorBrush BrText    = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8));
        private static readonly SolidColorBrush BrTextDim = new SolidColorBrush(Color.FromRgb(0xAA, 0xAA, 0xAA));
        private static readonly SolidColorBrush BrSel     = new SolidColorBrush(Color.FromRgb(0x4A, 0x4A, 0x70));
        private static readonly SolidColorBrush BrBtn     = new SolidColorBrush(Color.FromRgb(0x4A, 0x4A, 0x4A));
        private static readonly SolidColorBrush BrBtnBdr  = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66));
        private static readonly SolidColorBrush BrBtnRed  = new SolidColorBrush(Color.FromRgb(0x5A, 0x2A, 0x2A));
        private static readonly SolidColorBrush BrBtnRedBdr = new SolidColorBrush(Color.FromRgb(0x88, 0x33, 0x33));
        private static readonly SolidColorBrush BrBtnBlue = new SolidColorBrush(Color.FromRgb(0x3A, 0x4A, 0x5A));

        public ClashLookupWindow()
        {
            InitializeComponent();
            Loaded += (s, e) => RebuildTabs();
            Closed += (s, e) => _instance = null;
        }

        // ----------------------------------------------------------------
        // Tabs
        // ----------------------------------------------------------------
        private void RebuildTabs()
        {
            tabControl.Items.Clear();
            foreach (var table in ClashLookupService.Instance.GetAllTables())
                tabControl.Items.Add(BuildTab(table));
            UpdateEmptyHint();
        }

        private TabItem BuildTab(ClashLookupTable table)
        {
            string mepType     = table.MepType;
            string displayName = ClashLookupService.MepKindToDisplayName(mepType);
            var rows = new ObservableCollection<ClashLookupRow>(table.Rows);

            // ---- Auto-fill checkbox (shared ref so DataGrid handler can read it) ----
            var chkAutoFill = new CheckBox
            {
                Content             = "Автоматическое заполнение таблицы",
                IsChecked           = true,
                Foreground          = BrTextDim,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin              = new Thickness(8, 0, 12, 0),
                Cursor              = System.Windows.Input.Cursors.Hand
            };

            // ---- DataGrid ----
            var grid = BuildDataGrid(rows, mepType, chkAutoFill);

            // ---- "Add row" button ----
            var btnAdd = MakeButton("+ Добавить типоразмер", BrBtn, BrBtnBdr);
            btnAdd.HorizontalAlignment = HorizontalAlignment.Left;
            btnAdd.Margin = new Thickness(0, 6, 0, 0);
            btnAdd.Click += (s, e) =>
            {
                var newRow = new ClashLookupRow();
                rows.Add(newRow);
                SyncAndSave(mepType, rows);
            };

            // ---- Fallback policy ----
            var fallbackLabel = new TextBlock
            {
                Text = "При отсутствии типоразмера:",
                Foreground = BrTextDim,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 6, 0)
            };

            var fallbackCombo = BuildDarkCombo(
                new[] { "Использовать посчитанное значение", "Использовать соседнее значение из таблицы" },
                new[] { "Auto", "Nearest" },
                table.FallbackPolicy ?? "Auto");
            fallbackCombo.Width = 300;
            fallbackCombo.VerticalAlignment = VerticalAlignment.Center;
            fallbackCombo.SelectionChanged += (s, e) =>
            {
                var ci = fallbackCombo.SelectedItem as ComboBoxItem;
                if (ci != null) ClashLookupService.Instance.SetFallbackPolicy(mepType, (string)ci.Tag);
            };

            // ---- Delete button ----
            var btnDel = MakeButton("Удалить таблицу", BrBtnRed, BrBtnRedBdr);
            btnDel.HorizontalAlignment = HorizontalAlignment.Right;
            btnDel.Margin = new Thickness(8, 0, 0, 0);
            btnDel.Click += (s, e) => OnDeleteTable(mepType);

            // ---- Bottom bar: two rows ----
            // Row 1: [+ Add] [AutoFill checkbox]       [Delete table]
            // Row 2: ["При отсутствии типоразмера:"] [combo]

            var row1 = new DockPanel { Margin = new Thickness(0, 6, 0, 2) };
            DockPanel.SetDock(btnDel,      Dock.Right);
            DockPanel.SetDock(btnAdd,      Dock.Left);
            DockPanel.SetDock(chkAutoFill, Dock.Left);
            row1.Children.Add(btnDel);
            row1.Children.Add(btnAdd);
            row1.Children.Add(chkAutoFill);

            var row2 = new DockPanel { Margin = new Thickness(0, 0, 0, 0) };
            DockPanel.SetDock(fallbackLabel, Dock.Left);
            row2.Children.Add(fallbackLabel);
            row2.Children.Add(fallbackCombo);

            var bottomBar = new StackPanel { Margin = new Thickness(0, 4, 0, 0) };
            bottomBar.Children.Add(row1);
            bottomBar.Children.Add(row2);

            var content = new DockPanel { Margin = new Thickness(4) };
            DockPanel.SetDock(bottomBar, Dock.Bottom);
            content.Children.Add(bottomBar);
            content.Children.Add(grid);

            // Use explicit TextBlock as header so foreground is always under our control
            var hdrBlock = new TextBlock
            {
                Text       = displayName,
                Foreground = BrText,
                Margin     = new Thickness(0)
            };
            return new TabItem { Header = hdrBlock, Tag = mepType, Content = content };
        }

        // ----------------------------------------------------------------
        // DataGrid with 2-row header
        // ----------------------------------------------------------------
        private Grid BuildDataGrid(ObservableCollection<ClashLookupRow> rows, string mepType, CheckBox chkAutoFill)
        {
            // Column widths
            double wSize  = 90;
            double wAngle = 68;

            // ---- Two-row header panel ----
            var headerGrid = new System.Windows.Controls.Grid();
            headerGrid.Background = BrHdr;
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(wSize) });
            for (int i = 0; i < 8; i++)
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(wAngle) });
            headerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            headerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(22) });

            headerGrid.Children.Add(MakeHdrCell("Типоразмер, мм", 0, 0, 1, 2));
            headerGrid.Children.Add(MakeHdrCell("Величина опуска, мм", 1, 0, 4, 1));
            headerGrid.Children.Add(MakeHdrCell("Длина сегмента, мм", 5, 0, 4, 1));
            string[] angles = { "30°", "45°", "60°", "90°", "30°", "45°", "60°", "90°" };
            for (int i = 0; i < 8; i++)
                headerGrid.Children.Add(MakeHdrCell(angles[i], i + 1, 1, 1, 1));

            // ---- DataGrid ----
            var dg = new DataGrid
            {
                Background           = BrCell,
                Foreground           = BrText,
                BorderBrush          = BrBorder,
                BorderThickness      = new Thickness(1),
                GridLinesVisibility  = DataGridGridLinesVisibility.All,
                HorizontalGridLinesBrush = BrBorder,
                VerticalGridLinesBrush   = BrBorder,
                RowBackground            = BrRow,
                AlternatingRowBackground = BrRowAlt,
                ColumnHeaderHeight       = 0,
                HeadersVisibility        = DataGridHeadersVisibility.None,
                RowHeight                = 24,
                SelectionMode            = DataGridSelectionMode.Single,
                AutoGenerateColumns      = false,
                CanUserAddRows           = false,
                CanUserDeleteRows        = false,
                CanUserResizeRows        = false,
                ItemsSource              = rows
            };

            // Cell style
            var cellStyle = new Style(typeof(DataGridCell));
            cellStyle.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.Transparent));
            cellStyle.Setters.Add(new Setter(DataGridCell.ForegroundProperty, BrText));
            cellStyle.Setters.Add(new Setter(DataGridCell.BorderThicknessProperty, new Thickness(0)));
            cellStyle.Setters.Add(new Setter(DataGridCell.PaddingProperty, new Thickness(4, 0, 4, 0)));
            var selTrigger = new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true };
            selTrigger.Setters.Add(new Setter(DataGridCell.BackgroundProperty, BrSel));
            cellStyle.Triggers.Add(selTrigger);
            dg.CellStyle = cellStyle;

            var rowStyle = new Style(typeof(DataGridRow));
            rowStyle.Setters.Add(new Setter(DataGridRow.ForegroundProperty, BrText));
            dg.RowStyle = rowStyle;

            // ---- Columns ----
            // SizeMm: plain text column
            dg.Columns.Add(MakeCol("Типоразмер", "SizeMm", wSize));

            // Value columns: DataGridTemplateColumn so we can control foreground per cell
            string[] fields = { "Drop30","Drop45","Drop60","Drop90","Seg30","Seg45","Seg60","Seg90" };
            foreach (var field in fields)
                dg.Columns.Add(MakeValueCol(field, wAngle));

            // Delete-row button column (reuses the extra blank column on the right)
            var delColFactory = new FrameworkElementFactory(typeof(Button));
            delColFactory.SetValue(Button.ContentProperty, "×");
            delColFactory.SetValue(Button.ForegroundProperty, new SolidColorBrush(Color.FromRgb(0xFF, 0x88, 0x88)));
            delColFactory.SetValue(Button.BackgroundProperty, Brushes.Transparent);
            delColFactory.SetValue(Button.BorderThicknessProperty, new Thickness(0));
            delColFactory.SetValue(Button.FontSizeProperty, 13.0);
            delColFactory.SetValue(Button.FontWeightProperty, FontWeights.Bold);
            delColFactory.SetValue(Button.CursorProperty, System.Windows.Input.Cursors.Hand);
            delColFactory.SetValue(Button.ToolTipProperty, "Удалить строку");
            delColFactory.SetValue(Button.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            delColFactory.SetValue(Button.VerticalAlignmentProperty, VerticalAlignment.Center);
            delColFactory.AddHandler(Button.ClickEvent, new RoutedEventHandler((btnSender, btnArgs) =>
            {
                var btn = btnSender as Button;
                var rowItem = btn?.DataContext as ClashLookupRow;
                if (rowItem != null)
                {
                    rows.Remove(rowItem);
                    SyncAndSave(mepType, rows);
                }
            }));
            var delCol = new DataGridTemplateColumn
            {
                Header = "",
                Width  = 60,
                CellTemplate = new DataTemplate { VisualTree = delColFactory },
                CanUserResize = false,
                CanUserSort   = false
            };
            dg.Columns.Add(delCol);

            // CellEditEnding: detect which field changed and handle auto-fill / manual-mark
            dg.CellEditEnding += (s, e) =>
            {
                var row = e.Row.Item as ClashLookupRow;
                if (row == null) return;

                int colIdx = e.Column.DisplayIndex;
                string editedField = colIdx == 0 ? "SizeMm" : fields[colIdx - 1];

                // Capture new text from editing element before commit
                string newText = "";
                if (e.EditingElement is TextBox tb) newText = tb.Text?.Trim() ?? "";

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (editedField == "SizeMm")
                    {
                        if (!double.TryParse(newText,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.CurrentCulture,
                                out double parsedSize) || parsedSize <= 0)
                        {
                            double.TryParse(newText,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out parsedSize);
                        }

                        bool autoFillEnabled = chkAutoFill.IsChecked == true;

                        if (autoFillEnabled && parsedSize > 0)
                        {
                            // Always recalculate the entire row when SizeMm changes
                            var calc = ClashLookupService.Instance.AutoCalculateRow(mepType, parsedSize);
                            row.Drop30 = calc.Drop30; row.Drop45 = calc.Drop45;
                            row.Drop60 = calc.Drop60; row.Drop90 = calc.Drop90;
                            row.Seg30  = calc.Seg30;  row.Seg45  = calc.Seg45;
                            row.Seg60  = calc.Seg60;  row.Seg90  = calc.Seg90;
                            // Reset all flags to auto (white) since values are freshly calculated
                            foreach (var f in fields) row.AutoFlags[f] = true;
                        }

                        // Auto-sort by SizeMm ascending
                        dg.CommitEdit(DataGridEditingUnit.Row, true);
                        var sorted = rows.OrderBy(r => r.SizeMm).ToList();
                        rows.Clear();
                        foreach (var r in sorted) rows.Add(r);
                        dg.Items.Refresh();
                    }
                    else
                    {
                        // Mark as manual ONLY if value actually changed
                        double? oldValue = GetFieldValue(row, editedField);
                        double? newValue = null;
                        if (double.TryParse(newText,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.CurrentCulture,
                                out double parsed))
                            newValue = parsed;
                        else if (double.TryParse(newText,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out double parsedInv))
                            newValue = parsedInv;

                        bool valueChanged = Math.Abs((oldValue ?? 0) - (newValue ?? 0)) > 0.001
                                         || (oldValue == null) != (newValue == null);
                        if (valueChanged)
                            row.AutoFlags[editedField] = false;

                        dg.CommitEdit(DataGridEditingUnit.Row, true);
                        dg.Items.Refresh();
                    }
                    SyncAndSave(mepType, rows);
                }), System.Windows.Threading.DispatcherPriority.Background);
            };

            // ---- Outer container ----
            // headerGrid and DataGrid share the same column widths; both are placed in a
            // Grid so the DataGrid scrolls vertically while the header stays fixed.
            // Fix alignment: make headerGrid exactly as wide as the DataGrid columns.
            double totalWidth = wSize + wAngle * 8 + 60; // +60 for delete column
            headerGrid.Width = totalWidth;
            headerGrid.HorizontalAlignment = HorizontalAlignment.Left;

            var outer = new System.Windows.Controls.Grid();
            outer.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            outer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            outer.Margin = new Thickness(0, 6, 0, 0);
            System.Windows.Controls.Grid.SetRow(headerGrid, 0);
            System.Windows.Controls.Grid.SetRow(dg, 1);
            outer.Children.Add(headerGrid);
            outer.Children.Add(dg);
            return outer;
        }

        // Returns the nullable double value for a named field on a row
        private static double? GetFieldValue(ClashLookupRow row, string field)
        {
            switch (field)
            {
                case "Drop30": return row.Drop30; case "Drop45": return row.Drop45;
                case "Drop60": return row.Drop60; case "Drop90": return row.Drop90;
                case "Seg30":  return row.Seg30;  case "Seg45":  return row.Seg45;
                case "Seg60":  return row.Seg60;  case "Seg90":  return row.Seg90;
                default: return null;
            }
        }

        /// <summary>
        /// Creates a DataGridTemplateColumn for a Drop/Seg value field.
        /// The display TextBlock shows white text for auto-calculated values, blue for manual.
        /// </summary>
        private static DataGridTemplateColumn MakeValueCol(string field, double width)
        {
            var col = new DataGridTemplateColumn { Header = field, Width = width };

            // Display template: TextBlock whose Foreground depends on AutoFlags[field]
            var displayFactory = new FrameworkElementFactory(typeof(TextBlock));
            displayFactory.SetBinding(TextBlock.TextProperty, new Binding(field)
            {
                Converter       = new NullableDoubleConverter(),
                TargetNullValue = ""
            });
            displayFactory.SetBinding(TextBlock.ForegroundProperty, new Binding(".")
            {
                Converter          = new AutoFlagForegroundConverter(field),
                ConverterParameter = field
            });
            displayFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            displayFactory.SetValue(TextBlock.MarginProperty, new Thickness(4, 0, 4, 0));
            col.CellTemplate = new DataTemplate { VisualTree = displayFactory };

            // Edit template: TextBox with dark style
            var editFactory = new FrameworkElementFactory(typeof(TextBox));
            editFactory.SetBinding(TextBox.TextProperty, new Binding(field)
            {
                Mode                = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.LostFocus,
                Converter           = new NullableDoubleConverter(),
                TargetNullValue     = ""
            });
            editFactory.SetValue(TextBox.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0x3A, 0x3A, 0x5A)));
            editFactory.SetValue(TextBox.ForegroundProperty, new SolidColorBrush(Colors.White));
            editFactory.SetValue(TextBox.BorderThicknessProperty, new Thickness(0));
            editFactory.SetValue(TextBox.VerticalAlignmentProperty, VerticalAlignment.Stretch);
            editFactory.SetValue(TextBox.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
            editFactory.SetValue(TextBox.PaddingProperty, new Thickness(4, 0, 4, 0));
            col.CellEditingTemplate = new DataTemplate { VisualTree = editFactory };

            return col;
        }

        private static UIElement MakeHdrCell(string text, int col, int row, int colSpan, int rowSpan,
                                              SolidColorBrush bg = null)
        {
            var border = new Border
            {
                Background   = bg ?? new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)),
                BorderBrush  = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Child = new TextBlock
                {
                    Text = text,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                    FontSize = 11,
                    Padding  = new Thickness(2, 0, 2, 0)
                }
            };
            System.Windows.Controls.Grid.SetColumn(border, col);
            System.Windows.Controls.Grid.SetRow(border, row);
            if (colSpan > 1) System.Windows.Controls.Grid.SetColumnSpan(border, colSpan);
            if (rowSpan > 1) System.Windows.Controls.Grid.SetRowSpan(border, rowSpan);
            return border;
        }

        private static DataGridTextColumn MakeCol(string header, string path, double width)
        {
            return new DataGridTextColumn
            {
                Header  = header,
                Width   = width,
                Binding = new Binding(path)
                {
                    Mode                = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.LostFocus,
                    TargetNullValue     = "",
                    Converter           = new NullableDoubleConverter()
                }
            };
        }

        // ----------------------------------------------------------------
        // Create / Delete
        // ----------------------------------------------------------------
        private void BtnCreateTable_Click(object sender, RoutedEventArgs e)
        {
            var existing = ClashLookupService.Instance.GetAllTables().Select(t => t.MepType).ToHashSet();
            var available = AllMepTypes.Where(t => !existing.Contains(t))
                                       .Select(ClashLookupService.MepKindToDisplayName).ToList();
            if (available.Count == 0)
            {
                MessageBox.Show("Все три таблицы уже созданы.", "Создать таблицу",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var dlg = new TypeSelectDialog(available) { Owner = this };
            if (dlg.ShowDialog() != true) return;

            string mepType = ClashLookupService.DisplayNameToMepType(dlg.SelectedDisplayName);
            ClashLookupService.Instance.CreateTable(mepType);
            RebuildTabs();
            foreach (TabItem ti in tabControl.Items)
                if ((string)ti.Tag == mepType) { tabControl.SelectedItem = ti; break; }
        }

        private void OnDeleteTable(string mepType)
        {
            string dn = ClashLookupService.MepKindToDisplayName(mepType);
            if (MessageBox.Show($"Удалить таблицу «{dn}»?", "Удаление",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            ClashLookupService.Instance.DeleteTable(mepType);
            RebuildTabs();
        }

        private void UpdateEmptyHint()
        {
            bool has = tabControl.Items.Count > 0;
            tabControl.Visibility   = has ? Visibility.Visible   : Visibility.Collapsed;
            txtEmptyHint.Visibility = has ? Visibility.Collapsed : Visibility.Visible;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private static void SyncAndSave(string mepType, ObservableCollection<ClashLookupRow> rows)
            => ClashLookupService.Instance.UpdateRows(mepType, new List<ClashLookupRow>(rows));

        // ----------------------------------------------------------------
        // Helpers
        // ----------------------------------------------------------------
        private static Button MakeButton(string text, SolidColorBrush bg, SolidColorBrush border)
        {
            // Use a proper ControlTemplate so hover/press work correctly on dark background
            var tpl = new ControlTemplate(typeof(Button));
            var bdFactory = new FrameworkElementFactory(typeof(Border));
            bdFactory.Name = "Bd";
            bdFactory.SetValue(Border.BackgroundProperty, bg);
            bdFactory.SetValue(Border.BorderBrushProperty, border);
            bdFactory.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            bdFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));
            bdFactory.SetValue(Border.PaddingProperty, new Thickness(10, 4, 10, 4));
            var cpFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            cpFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            cpFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            bdFactory.AppendChild(cpFactory);
            tpl.VisualTree = bdFactory;

            var hoverTrigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty,
                new SolidColorBrush(Color.FromRgb(
                    (byte)Math.Min(bg.Color.R + 0x18, 0xFF),
                    (byte)Math.Min(bg.Color.G + 0x18, 0xFF),
                    (byte)Math.Min(bg.Color.B + 0x18, 0xFF))), "Bd"));
            tpl.Triggers.Add(hoverTrigger);

            return new Button
            {
                Content         = text,
                Background      = bg,
                Foreground      = BrText,
                BorderBrush     = border,
                BorderThickness = new Thickness(1),
                Padding         = new Thickness(10, 4, 10, 4),
                Cursor          = System.Windows.Input.Cursors.Hand,
                Template        = tpl
            };
        }

        private static ComboBox BuildDarkCombo(string[] labels, string[] tags, string selectedTag)
        {
            // Full dark ControlTemplate for the ComboBox toggle button
            var toggleTemplate = new ControlTemplate(typeof(ToggleButton));
            var toggleBd = new FrameworkElementFactory(typeof(Border));
            toggleBd.Name = "Bd";
            toggleBd.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)));
            toggleBd.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
            toggleBd.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            var toggleGrid = new FrameworkElementFactory(typeof(System.Windows.Controls.Grid));
            var col0 = new FrameworkElementFactory(typeof(ColumnDefinition));
            col0.SetValue(ColumnDefinition.WidthProperty, new GridLength(1, GridUnitType.Star));
            var col1 = new FrameworkElementFactory(typeof(ColumnDefinition));
            col1.SetValue(ColumnDefinition.WidthProperty, new GridLength(18));
            toggleGrid.AppendChild(col0);
            toggleGrid.AppendChild(col1);
            var cpTgl = new FrameworkElementFactory(typeof(ContentPresenter));
            cpTgl.SetValue(System.Windows.Controls.Grid.ColumnProperty, 0);
            cpTgl.SetValue(ContentPresenter.MarginProperty, new Thickness(4, 0, 0, 0));
            cpTgl.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            var arrow = new FrameworkElementFactory(typeof(TextBlock));
            arrow.SetValue(System.Windows.Controls.Grid.ColumnProperty, 1);
            arrow.SetValue(TextBlock.TextProperty, "▾");
            arrow.SetValue(TextBlock.ForegroundProperty, new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)));
            arrow.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            arrow.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            arrow.SetValue(TextBlock.FontSizeProperty, 10.0);
            toggleGrid.AppendChild(cpTgl);
            toggleGrid.AppendChild(arrow);
            toggleBd.AppendChild(toggleGrid);
            toggleTemplate.VisualTree = toggleBd;

            // ComboBox template
            var cbTemplate = new ControlTemplate(typeof(ComboBox));
            var outerBd = new FrameworkElementFactory(typeof(Border));
            outerBd.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)));
            outerBd.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
            outerBd.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            var tglBtn = new FrameworkElementFactory(typeof(ToggleButton));
            tglBtn.Name = "ToggleButton";
            tglBtn.SetValue(ToggleButton.TemplateProperty, toggleTemplate);
            tglBtn.SetBinding(ToggleButton.IsCheckedProperty,
                new Binding("IsDropDownOpen") { Mode = BindingMode.TwoWay,
                    RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            tglBtn.SetValue(ToggleButton.FocusableProperty, false);
            var cpSel = new FrameworkElementFactory(typeof(ContentPresenter));
            cpSel.Name = "ContentSite";
            cpSel.SetValue(ContentPresenter.MarginProperty, new Thickness(4, 2, 24, 2));
            cpSel.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            cpSel.SetBinding(ContentPresenter.ContentProperty,
                new Binding("SelectionBoxItem") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            cpSel.SetBinding(ContentPresenter.ContentTemplateProperty,
                new Binding("SelectionBoxItemTemplate") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            var popup = new FrameworkElementFactory(typeof(Popup));
            popup.Name = "Popup";
            popup.SetValue(Popup.PlacementProperty, PlacementMode.Bottom);
            popup.SetValue(Popup.AllowsTransparencyProperty, true);
            popup.SetBinding(Popup.IsOpenProperty,
                new Binding("IsDropDownOpen") { RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent) });
            popup.SetValue(Popup.StaysOpenProperty, false);
            var dropBd = new FrameworkElementFactory(typeof(Border));
            dropBd.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)));
            dropBd.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)));
            dropBd.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            var scrollViewer = new FrameworkElementFactory(typeof(ScrollViewer));
            scrollViewer.SetValue(ScrollViewer.MarginProperty, new Thickness(2));
            var itemsPresenter = new FrameworkElementFactory(typeof(ItemsPresenter));
            itemsPresenter.SetValue(KeyboardNavigation.DirectionalNavigationProperty,
                KeyboardNavigationMode.Contained);
            scrollViewer.AppendChild(itemsPresenter);
            dropBd.AppendChild(scrollViewer);
            popup.AppendChild(dropBd);
            var cbGrid = new FrameworkElementFactory(typeof(System.Windows.Controls.Grid));
            cbGrid.AppendChild(tglBtn);
            cbGrid.AppendChild(cpSel);
            cbGrid.AppendChild(popup);
            outerBd.AppendChild(cbGrid);
            cbTemplate.VisualTree = outerBd;

            // Item container style
            var itemStyle = new Style(typeof(ComboBoxItem));
            itemStyle.Setters.Add(new Setter(ComboBoxItem.BackgroundProperty,
                new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C))));
            itemStyle.Setters.Add(new Setter(ComboBoxItem.ForegroundProperty, BrText));
            itemStyle.Setters.Add(new Setter(ComboBoxItem.PaddingProperty, new Thickness(6, 3, 6, 3)));
            var hoverTrig = new Trigger { Property = ComboBoxItem.IsMouseOverProperty, Value = true };
            hoverTrig.Setters.Add(new Setter(ComboBoxItem.BackgroundProperty,
                new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55))));
            itemStyle.Triggers.Add(hoverTrig);

            var cb = new ComboBox
            {
                Background  = new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)),
                Foreground  = BrText,
                BorderBrush = BrBorder,
                Template    = cbTemplate,
                ItemContainerStyle = itemStyle
            };

            for (int i = 0; i < labels.Length; i++)
            {
                var ci = new ComboBoxItem { Content = labels[i], Tag = tags[i], Foreground = BrText };
                cb.Items.Add(ci);
                if (tags[i] == selectedTag) cb.SelectedItem = ci;
            }
            if (cb.SelectedItem == null) cb.SelectedIndex = 0;
            return cb;
        }

        /// <summary>Simple dark ComboBox with plain string items.</summary>
        internal static ComboBox BuildDarkComboFromStrings(List<string> items)
        {
            string[] labels = items.ToArray();
            string[] tags   = items.ToArray();
            var cb = BuildDarkCombo(labels, tags, labels.Length > 0 ? labels[0] : null);
            return cb;
        }
    }

    // ====================================================================
    // Type selection dialog
    // ====================================================================
    internal class TypeSelectDialog : Window
    {
        public string SelectedDisplayName { get; private set; }

        public TypeSelectDialog(List<string> names)
        {
            Title = "Выбор типа элементов";
            Width = 340; Height = 160;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background  = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x2D));
            FontFamily  = new FontFamily("Segoe UI");
            FontSize    = 12;
            Foreground  = Brushes.WhiteSmoke;

            var label = new TextBlock
            {
                Text = "Выберите тип элементов:",
                Foreground = Brushes.LightGray,
                Margin = new Thickness(12, 12, 12, 6),
                TextWrapping = TextWrapping.Wrap
            };

            var combo = ClashLookupWindow.BuildDarkComboFromStrings(names);
            combo.Margin = new Thickness(12, 0, 12, 12);

            var btnOk = new Button
            {
                Content = "Создать", Width = 90,
                Background  = new SolidColorBrush(Color.FromRgb(0x4A, 0x4A, 0x4A)),
                Foreground  = Brushes.WhiteSmoke,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
                Margin = new Thickness(12, 0, 6, 12)
            };
            btnOk.Click += (s, e) => { SelectedDisplayName = (combo.SelectedItem as ComboBoxItem)?.Tag?.ToString(); DialogResult = true; };

            var btnCancel = new Button
            {
                Content = "Отмена", Width = 90,
                Background  = new SolidColorBrush(Color.FromRgb(0x5A, 0x3A, 0x3A)),
                Foreground  = Brushes.WhiteSmoke,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x88, 0x33, 0x33)),
                Margin = new Thickness(0, 0, 12, 12)
            };
            btnCancel.Click += (s, e) => DialogResult = false;

            var btnRow = new StackPanel { Orientation = Orientation.Horizontal };
            btnRow.Children.Add(btnOk);
            btnRow.Children.Add(btnCancel);

            var panel = new StackPanel();
            panel.Children.Add(label);
            panel.Children.Add(combo);
            panel.Children.Add(btnRow);
            Content = panel;
        }
    }

    // ====================================================================
    // Converter: auto-flag → foreground colour
    // White = auto-calculated, Blue = manually edited
    // ====================================================================
    internal class AutoFlagForegroundConverter : System.Windows.Data.IValueConverter
    {
        private readonly string _field;
        private static readonly SolidColorBrush BrAuto   = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8)); // white
        private static readonly SolidColorBrush BrManual = new SolidColorBrush(Color.FromRgb(0x66, 0xAA, 0xFF)); // blue

        public AutoFlagForegroundConverter(string field) { _field = field; }

        public object Convert(object value, Type t, object p, System.Globalization.CultureInfo c)
        {
            var row = value as ClashResolve.Module.Core.ClashLookupRow;
            if (row == null) return BrAuto;
            if (row.AutoFlags.TryGetValue(_field, out bool isAuto) && !isAuto)
                return BrManual;
            return BrAuto;
        }

        public object ConvertBack(object value, Type t, object p, System.Globalization.CultureInfo c)
            => System.Windows.Data.Binding.DoNothing;
    }

    // ====================================================================
    // Converter: string ↔ double?
    // ====================================================================
    internal class NullableDoubleConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type t, object p, System.Globalization.CultureInfo c)
            => value == null ? "" : ((double)value).ToString("G", c);

        public object ConvertBack(object value, Type t, object p, System.Globalization.CultureInfo c)
        {
            string s = value as string;
            if (string.IsNullOrWhiteSpace(s)) return null;
            return double.TryParse(s.Trim(), System.Globalization.NumberStyles.Any, c, out double d)
                ? (double?)d : null;
        }
    }
}
