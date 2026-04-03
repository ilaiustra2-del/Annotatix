using System;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using PluginsManager.Core;
using Tracer.Module.Core;

namespace Tracer.Module.UI
{
    /// <summary>
    /// Interaction logic for TracerPanel.xaml
    /// </summary>
    public partial class TracerPanel : UserControl
    {
        private readonly UIApplication _uiApp;
        private readonly ExternalEvent _selectMainPipeEvent;
        private readonly ExternalEvent _selectRiserEvent;
        private readonly ExternalEvent _createConnectionEvent;
        private readonly ExternalEvent _createLConnectionEvent;
        
        // Selected elements data
        private MainLineData _selectedMainLine;
        private RiserData _selectedRiser;
        
        // Calculated connection data
        private XYZ _connectionPoint;
        private XYZ _riserConnectionPoint;
        private double _pipeDiameter;
        private double _connectionHeight;

        public TracerPanel(
            UIApplication uiApp, 
            ExternalEvent selectMainPipeEvent,
            ExternalEvent selectRiserEvent,
            ExternalEvent createConnectionEvent,
            ExternalEvent createLConnectionEvent)
        {
            InitializeComponent();
            
            _uiApp = uiApp;
            _selectMainPipeEvent = selectMainPipeEvent;
            _selectRiserEvent = selectRiserEvent;
            _createConnectionEvent = createConnectionEvent;
            _createLConnectionEvent = createLConnectionEvent;
            
            DebugLogger.Log("[TRACER-PANEL] Panel initialized");
        }

        #region Button Click Handlers

        private void SelectMainLineButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[TRACER-PANEL] Select main line button clicked");
                
                var doc = _uiApp.ActiveUIDocument.Document;
                
                try
                {
                    var reference = _uiApp.ActiveUIDocument.Selection.PickObject(
                        ObjectType.Element,
                        new PipeSelectionFilter(),
                        "Выберите магистральную линию канализации");
                    
                    if (reference != null)
                    {
                        _selectedMainLine = RevitPipeUtils.GetMainLineData(doc, reference.ElementId);
                        
                        if (_selectedMainLine != null)
                        {
                            UpdateMainLineUI();
                            SelectRiserButton.IsEnabled = true;
                            DebugLogger.Log($"[TRACER-PANEL] Main line selected: {_selectedMainLine.Name}");
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    DebugLogger.Log("[TRACER-PANEL] Main line selection cancelled");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-PANEL] ERROR selecting main line: {ex.Message}");
                MessageBox.Show($"Ошибка выбора магистрали: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SelectRiserButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[TRACER-PANEL] Select riser button clicked");
                
                var doc = _uiApp.ActiveUIDocument.Document;
                
                try
                {
                    var reference = _uiApp.ActiveUIDocument.Selection.PickObject(
                        ObjectType.Element,
                        new PipeSelectionFilter(),
                        "Выберите стояк канализации");
                    
                    if (reference != null)
                    {
                        _selectedRiser = RevitPipeUtils.GetRiserData(doc, reference.ElementId);
                        
                        if (_selectedRiser != null)
                        {
                            // Auto-calculate connection parameters
                            CalculateConnectionParameters();
                            
                            UpdateRiserUI();
                            CreateConnectionButton.IsEnabled = _connectionPoint != null;
                            CreateLConnectionButton.IsEnabled = _connectionPoint != null;
                            DebugLogger.Log($"[TRACER-PANEL] Riser selected: {_selectedRiser.Name}");
                        }
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    DebugLogger.Log("[TRACER-PANEL] Riser selection cancelled");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-PANEL] ERROR selecting riser: {ex.Message}");
                MessageBox.Show($"Ошибка выбора стояка: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CalculateConnectionParameters()
        {
            if (_selectedMainLine == null || _selectedRiser == null)
                return;

            try
            {
                DebugLogger.Log("[TRACER-PANEL] Calculating connection parameters...");
                
                // Use riser diameter for connection pipe
                _pipeDiameter = _selectedRiser.Diameter;
                
                // Calculate connection height at the riser location (bottom of riser)
                _connectionHeight = _selectedRiser.BottomElevation;
                
                // Calculate connection points according to scheme
                var result = ConnectionCalculator.CalculateConnectionPoints(_selectedMainLine, _selectedRiser);
                _connectionPoint = result.connectionPoint;      // On main line (x5,y5,z5)
                _riserConnectionPoint = result.endPoint;        // At riser (x6,y6,z6)
                
                if (_riserConnectionPoint != null && _connectionPoint != null)
                {
                    // Validate connection
                    bool isValid = ConnectionCalculator.ValidateConnection(
                        _selectedMainLine, _connectionPoint, _riserConnectionPoint);
                    
                    if (!isValid)
                    {
                        DebugLogger.Log("[TRACER-PANEL] Connection validation failed");
                        _connectionPoint = null;
                        _riserConnectionPoint = null;
                    }
                    else
                    {
                        DebugLogger.Log($"[TRACER-PANEL] Connection calculated: " +
                            $"Diameter={_pipeDiameter * 304.8:F0}mm");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-PANEL] ERROR calculating connection: {ex.Message}");
                _connectionPoint = null;
                _riserConnectionPoint = null;
            }
        }

        private void CreateConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[TRACER-PANEL] Create connection button clicked");
                
                if (_selectedMainLine == null || _selectedRiser == null)
                {
                    MessageBox.Show("Необходимо выбрать магистраль и стояк", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                if (_connectionPoint == null || _riserConnectionPoint == null)
                {
                    MessageBox.Show("Не удалось рассчитать точку подключения. Проверьте положение элементов.", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // Store data for ExternalEvent handler - using reflection to avoid circular dependency
                var handlerType = System.Type.GetType("PluginsManager.Commands.TracerCreateConnectionHandler, PluginsManager");
                if (handlerType != null)
                {
                    var method = handlerType.GetMethod("SetConnectionData", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    if (method != null)
                    {
                        method.Invoke(null, new object[] { 
                            _selectedMainLine.ElementId, _selectedRiser.ElementId,
                            _connectionPoint, _riserConnectionPoint, _pipeDiameter,
                            _selectedMainLine.Slope
                        });
                    }
                }
                else
                {
                    DebugLogger.Log("[TRACER-PANEL] ERROR: Could not find TracerCreateConnectionHandler");
                    MessageBox.Show("Ошибка: не найден обработчик создания соединения", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // Raise ExternalEvent to create connection in Revit API context
                _createConnectionEvent.Raise();
                
                DebugLogger.Log("[TRACER-PANEL] ExternalEvent raised for connection creation");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-PANEL] ERROR creating connection: {ex.Message}");
                MessageBox.Show($"Ошибка создания подключения: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateLConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DebugLogger.Log("[TRACER-PANEL] Create L-connection button clicked");
                
                if (_selectedMainLine == null || _selectedRiser == null)
                {
                    MessageBox.Show("Необходимо выбрать магистраль и стояк", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                if (_connectionPoint == null || _riserConnectionPoint == null)
                {
                    MessageBox.Show("Не удалось рассчитать точку подключения. Проверьте положение элементов.", "Ошибка", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // Store data for ExternalEvent handler
                var handlerType = System.Type.GetType("PluginsManager.Commands.TracerCreateLConnectionHandler, PluginsManager");
                if (handlerType != null)
                {
                    var method = handlerType.GetMethod("SetConnectionData", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    if (method != null)
                    {
                        method.Invoke(null, new object[] { 
                            _selectedMainLine.ElementId, _selectedRiser.ElementId,
                            _connectionPoint, _riserConnectionPoint, _pipeDiameter,
                            _selectedMainLine.Slope, _selectedMainLine.StartPoint, _selectedMainLine.EndPoint
                        });
                    }
                }
                else
                {
                    DebugLogger.Log("[TRACER-PANEL] ERROR: Could not find TracerCreateLConnectionHandler");
                    MessageBox.Show("Ошибка: не найден обработчик создания L-образного соединения", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // Raise ExternalEvent to create L-connection in Revit API context
                _createLConnectionEvent.Raise();
                
                DebugLogger.Log("[TRACER-PANEL] ExternalEvent raised for L-connection creation");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-PANEL] ERROR creating L-connection: {ex.Message}");
                MessageBox.Show($"Ошибка создания L-образного подключения: {ex.Message}", "Ошибка", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region UI Update Methods

        private void UpdateMainLineUI()
        {
            if (_selectedMainLine != null)
            {
                MainLineStatusText.Text = $"Магистраль: {_selectedMainLine.Name}";
                MainLineDetailsText.Text = $"Диаметр: {_selectedMainLine.Diameter * 304.8:F0} мм | " +
                    $"Уклон: {_selectedMainLine.Slope:F2}% | " +
                    $"Длина: {(_selectedMainLine.EndPoint - _selectedMainLine.StartPoint).GetLength() * 304.8:F0} мм";
            }
        }

        private void UpdateRiserUI()
        {
            if (_selectedRiser != null)
            {
                RiserStatusText.Text = $"Стояк: {_selectedRiser.Name}";
                RiserDetailsText.Text = $"Диаметр: {_selectedRiser.Diameter * 304.8:F0} мм | " +
                    $"Высота: {(_selectedRiser.TopElevation - _selectedRiser.BottomElevation) * 304.8:F0} мм";
                
                // Display calculated values
                if (_connectionPoint != null)
                {
                    ConnectionHeightTextBlock.Text = $"{_connectionHeight * 304.8:F0}";
                    PipeDiameterTextBlock.Text = $"{_pipeDiameter * 304.8:F0}";
                }
                else
                {
                    ConnectionHeightTextBlock.Text = "Ошибка";
                    PipeDiameterTextBlock.Text = "-";
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// Filter for selecting pipe elements
    /// </summary>
    public class PipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Autodesk.Revit.DB.Plumbing.Pipe;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
}
