using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PluginsManager.Core;

namespace AutoNumbering.Module.UI
{
    public partial class SystemGroupsWindow : Window
    {
        private List<string> _availableSystemTypes;
        private List<Core.SystemCompatibilityGroup> _compatibilityGroups;
        private int _groupCounter = 1;

        public SystemGroupsWindow(List<string> availableSystemTypes, List<Core.SystemCompatibilityGroup> existingGroups)
        {
            InitializeComponent();
            
            _availableSystemTypes = availableSystemTypes ?? new List<string>();
            _compatibilityGroups = existingGroups != null 
                ? new List<Core.SystemCompatibilityGroup>(existingGroups) 
                : new List<Core.SystemCompatibilityGroup>();
            
            if (existingGroups != null && existingGroups.Count > 0)
            {
                _groupCounter = existingGroups.Count + 1;
            }
            
            LoadSystemTypes();
            RefreshGroupsList();
        }

        public List<Core.SystemCompatibilityGroup> GetCompatibilityGroups()
        {
            return _compatibilityGroups;
        }

        private void LoadSystemTypes()
        {
            if (_availableSystemTypes == null || _availableSystemTypes.Count == 0)
            {
                // No system types available - show message
                lstSystemTypes.ItemsSource = new List<string> 
                { 
                    "(Нет типов систем - параметр 'Тип системы' пустой)" 
                };
                btnCreateGroup.IsEnabled = false;
            }
            else
            {
                lstSystemTypes.ItemsSource = _availableSystemTypes;
                btnCreateGroup.IsEnabled = true;
            }
        }

        private void RefreshGroupsList()
        {
            pnlGroups.Children.Clear();

            if (_compatibilityGroups.Count == 0)
            {
                var emptyText = new TextBlock
                {
                    Text = "Группы не созданы.\nВыберите типы систем слева и нажмите\n'Создать группу из выбранных'",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.FromRgb(153, 153, 153)),
                    TextAlignment = TextAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 20, 0, 0)
                };
                pnlGroups.Children.Add(emptyText);
                return;
            }

            foreach (var group in _compatibilityGroups)
            {
                var groupBorder = CreateGroupControl(group);
                pnlGroups.Children.Add(groupBorder);
            }
        }

        private Border CreateGroupControl(Core.SystemCompatibilityGroup group)
        {
            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(224, 224, 224)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 12),
                Padding = new Thickness(12)
            };

            var stack = new StackPanel();

            // Header with group name and delete button
            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var groupName = new TextBlock
            {
                Text = group.GroupName,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(51, 51, 51))
            };
            Grid.SetColumn(groupName, 0);
            headerGrid.Children.Add(groupName);

            var deleteButton = new Button
            {
                Content = "✕",
                FontSize = 16,
                Width = 28,
                Height = 28,
                Background = new SolidColorBrush(Color.FromRgb(220, 53, 69)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = "Удалить группу"
            };
            deleteButton.Click += (s, e) => DeleteGroup(group);
            Grid.SetColumn(deleteButton, 1);
            headerGrid.Children.Add(deleteButton);

            stack.Children.Add(headerGrid);

            // System types list
            var systemsText = new TextBlock
            {
                Text = string.Join(", ", group.SystemTypes),
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 8, 0, 0)
            };
            stack.Children.Add(systemsText);

            border.Child = stack;
            return border;
        }

        private void DeleteGroup(Core.SystemCompatibilityGroup group)
        {
            var result = MessageBox.Show(
                $"Удалить группу '{group.GroupName}'?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _compatibilityGroups.Remove(group);
                RefreshGroupsList();
                DebugLogger.Log($"[SYSTEM-GROUPS] Deleted group: {group.GroupName}");
            }
        }

        private void BtnCreateGroup_Click(object sender, RoutedEventArgs e)
        {
            // Get selected system types
            var selectedSystems = new List<string>();
            
            foreach (var item in lstSystemTypes.Items)
            {
                var container = lstSystemTypes.ItemContainerGenerator.ContainerFromItem(item);
                if (container != null)
                {
                    var checkBox = FindVisualChild<CheckBox>(container);
                    if (checkBox != null && checkBox.IsChecked == true)
                    {
                        selectedSystems.Add(item.ToString());
                    }
                }
            }

            if (selectedSystems.Count < 2)
            {
                MessageBox.Show(
                    "Выберите минимум 2 типа систем для создания группы совместимости.",
                    "Предупреждение",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Create new group
            var newGroup = new Core.SystemCompatibilityGroup
            {
                GroupName = $"Группа {_groupCounter}",
                SystemTypes = new List<string>(selectedSystems)
            };

            _compatibilityGroups.Add(newGroup);
            _groupCounter++;

            // Clear checkboxes
            foreach (var item in lstSystemTypes.Items)
            {
                var container = lstSystemTypes.ItemContainerGenerator.ContainerFromItem(item);
                if (container != null)
                {
                    var checkBox = FindVisualChild<CheckBox>(container);
                    if (checkBox != null)
                    {
                        checkBox.IsChecked = false;
                    }
                }
            }

            RefreshGroupsList();

            DebugLogger.Log($"[SYSTEM-GROUPS] Created group: {newGroup.GroupName} with {selectedSystems.Count} systems");
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        // Helper to find visual child
        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                    return typedChild;

                var result = FindVisualChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }
    }
}
