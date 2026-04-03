using System;
using System.Linq;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.Commands
{
    /// <summary>
    /// ExternalEvent handler for selecting main pipe in Tracer module
    /// </summary>
    public class TracerSelectMainPipeHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            DebugLogger.Log("[TRACER-COMMAND] Select main pipe event executed");
        }

        public string GetName()
        {
            return "Tracer Select Main Pipe";
        }
    }

    /// <summary>
    /// ExternalEvent handler for selecting riser in Tracer module
    /// </summary>
    public class TracerSelectRiserHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            DebugLogger.Log("[TRACER-COMMAND] Select riser event executed");
        }

        public string GetName()
        {
            return "Tracer Select Riser";
        }
    }

    /// <summary>
    /// ExternalEvent handler for creating 45-degree connection in Tracer module
    /// </summary>
    public class TracerCreateConnectionHandler : IExternalEventHandler
    {
        private static ElementId _mainPipeElementId;
        private static ElementId _riserElementId;
        private static XYZ _connectionPoint;
        private static XYZ _riserConnectionPoint;
        private static double _pipeDiameter;
        private static double _mainLineSlope;

        public string GetName()
        {
            return "Tracer Create 45-degree Connection";
        }

        public static void SetConnectionData(
            ElementId mainPipeId, ElementId riserId,
            XYZ connectionPoint, XYZ riserConnectionPoint, double pipeDiameter,
            double mainLineSlope)
        {
            _mainPipeElementId = mainPipeId;
            _riserElementId = riserId;
            _connectionPoint = connectionPoint;
            _riserConnectionPoint = riserConnectionPoint;
            _pipeDiameter = pipeDiameter;
            _mainLineSlope = mainLineSlope;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Log("[TRACER-COMMAND] Create 45-degree connection event executing...");

                if (_mainPipeElementId == null || _riserElementId == null || _connectionPoint == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Missing connection data");
                    MessageBox.Show("Ошибка: отсутствуют данные для создания подключения", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Document doc = app.ActiveUIDocument.Document;

                Element mainPipeElement = doc.GetElement(_mainPipeElementId);
                if (mainPipeElement == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Main pipe element not found");
                    MessageBox.Show("Ошибка: магистральная труба не найдена", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Pipe mainPipe = mainPipeElement as Pipe;
                if (mainPipe == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Main element is not a pipe");
                    MessageBox.Show("Ошибка: выбранный элемент не является трубой", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                ElementId pipeTypeId = mainPipe.GetTypeId();
                ElementId systemTypeId = mainPipe.MEPSystem?.GetTypeId();
                ElementId levelId = mainPipe.LevelId;
                if (levelId == null || levelId == ElementId.InvalidElementId)
                {
                    Level level = GetLevelByElevation(doc, _connectionPoint.Z);
                    levelId = level?.Id ?? ElementId.InvalidElementId;
                }

                using (Transaction tx = new Transaction(doc, "Создание присоединения канализации (45°)"))
                {
                    tx.Start();

                    try
                    {
                        // Создаём трубу от точки подключения к стояку
                        DebugLogger.Log("[TRACER-COMMAND] Creating connection pipe...");
                        Pipe connectionPipe = Pipe.Create(
                            doc,
                            systemTypeId,
                            pipeTypeId,
                            levelId,
                            _connectionPoint,
                            _riserConnectionPoint);

                        if (connectionPipe == null)
                        {
                            DebugLogger.Log("[TRACER-COMMAND] ERROR: Failed to create connection pipe");
                            tx.RollBack();
                            return;
                        }

                        DebugLogger.Log($"[TRACER-COMMAND] Connection pipe created: {connectionPipe.Id}");

                        // Устанавливаем диаметр
                        Parameter diameterParam = connectionPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diameterParam != null && !diameterParam.IsReadOnly)
                        {
                            diameterParam.Set(_pipeDiameter);
                        }

                        // Подгоняем стояк
                        AdjustRiser(doc, _riserElementId, _riserConnectionPoint);

                        tx.Commit();
                        DebugLogger.Log("[TRACER-COMMAND] Connection created successfully");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[TRACER-COMMAND] ERROR during connection creation: {ex.Message}\n{ex.StackTrace}");
                        tx.RollBack();
                        MessageBox.Show($"Ошибка при создании подключения: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-COMMAND] ERROR in Execute: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AdjustRiser(Document doc, ElementId riserId, XYZ connectionPoint)
        {
            Element riserElement = doc.GetElement(riserId);
            if (riserElement == null || !(riserElement is Pipe riserPipe))
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser not found for adjustment");
                return;
            }

            LocationCurve riserLocation = riserPipe.Location as LocationCurve;
            if (riserLocation == null)
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser has no LocationCurve");
                return;
            }

            Line riserLine = riserLocation.Curve as Line;
            if (riserLine == null)
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser curve is not a Line");
                return;
            }

            XYZ riserEnd0 = riserLine.GetEndPoint(0);
            XYZ riserEnd1 = riserLine.GetEndPoint(1);

            double dist0 = riserEnd0.DistanceTo(connectionPoint);
            double dist1 = riserEnd1.DistanceTo(connectionPoint);

            XYZ riserTop = (dist0 > dist1) ? riserEnd0 : riserEnd1;
            XYZ newRiserBottom = connectionPoint;

            Line newRiserLine = Line.CreateBound(newRiserBottom, riserTop);
            riserLocation.Curve = newRiserLine;

            DebugLogger.Log($"[TRACER-COMMAND] Adjusted riser bottom to ({newRiserBottom.X:F3},{newRiserBottom.Y:F3},{newRiserBottom.Z:F3})");
        }

        private static Level GetLevelByElevation(Document doc, double elevation)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - elevation))
                .ToList();

            return levels.FirstOrDefault();
        }
    }

    /// <summary>
    /// ExternalEvent handler for creating L-shaped connection in Tracer module
    /// Сегмент "a" (1м) перпендикулярно магистрали, затем к стояку
    /// </summary>
    public class TracerCreateLConnectionHandler : IExternalEventHandler
    {
        private static ElementId _mainPipeElementId;
        private static ElementId _riserElementId;
        private static XYZ _connectionPoint;        // Точка 5 (x5,y5,z5) - на магистрали
        private static XYZ _riserConnectionPoint;   // Точка 6 (x6,y6,z6) - у стояка
        private static double _pipeDiameter;
        private static double _mainLineSlope;
        private static XYZ _mainLineStartPoint;
        private static XYZ _mainLineEndPoint;

        public string GetName()
        {
            return "Tracer Create L-shaped Connection";
        }

        public static void SetConnectionData(
            ElementId mainPipeId, ElementId riserId,
            XYZ connectionPoint, XYZ riserConnectionPoint, double pipeDiameter,
            double mainLineSlope, XYZ mainLineStart, XYZ mainLineEnd)
        {
            _mainPipeElementId = mainPipeId;
            _riserElementId = riserId;
            _connectionPoint = connectionPoint;
            _riserConnectionPoint = riserConnectionPoint;
            _pipeDiameter = pipeDiameter;
            _mainLineSlope = mainLineSlope;
            _mainLineStartPoint = mainLineStart;
            _mainLineEndPoint = mainLineEnd;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Log("[TRACER-COMMAND] Create L-shaped connection event executing...");

                if (_mainPipeElementId == null || _riserElementId == null || _connectionPoint == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Missing connection data");
                    MessageBox.Show("Ошибка: отсутствуют данные для создания подключения", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Document doc = app.ActiveUIDocument.Document;

                Element mainPipeElement = doc.GetElement(_mainPipeElementId);
                if (mainPipeElement == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Main pipe element not found");
                    MessageBox.Show("Ошибка: магистральная труба не найдена", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Pipe mainPipe = mainPipeElement as Pipe;
                if (mainPipe == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Main element is not a pipe");
                    MessageBox.Show("Ошибка: выбранный элемент не является трубой", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                ElementId pipeTypeId = mainPipe.GetTypeId();
                ElementId systemTypeId = mainPipe.MEPSystem?.GetTypeId();
                ElementId levelId = mainPipe.LevelId;
                if (levelId == null || levelId == ElementId.InvalidElementId)
                {
                    Level level = GetLevelByElevation(doc, _connectionPoint.Z);
                    levelId = level?.Id ?? ElementId.InvalidElementId;
                }

                using (Transaction tx = new Transaction(doc, "Создание L-образного присоединения"))
                {
                    tx.Start();

                    try
                    {
                        // Рассчитываем точки для L-образного присоединения по схеме:
                        // 1. Точка 4 - проекция оси стояка на магистраль
                        // 2. Точка 5 - на магистрали, ниже точки 4 на sqrt(a/2) по направлению уклона
                        // 3. Сегмент "a" = 1 метр под 45° ВВЕРХ от магистрали (от точки 5 к точке 7)
                        // 4. Сегмент "б" перпендикулярно магистрали от точки 7 к стояку
                        
                        double segmentALength = 1.0; // 1 метр в футах
                        double slopeRatio = _mainLineSlope / 100.0;
                        
                        // === ТОЧКА 4 - проекция оси стояка на магистраль (в плане XY) ===
                        // Направление магистрали в плане (XY)
                        XYZ mainStartXY = new XYZ(_mainLineStartPoint.X, _mainLineStartPoint.Y, 0);
                        XYZ mainEndXY = new XYZ(_mainLineEndPoint.X, _mainLineEndPoint.Y, 0);
                        XYZ mainDirXY = (mainEndXY - mainStartXY).Normalize();
                        
                        // Ось стояка в плане
                        XYZ riserAxisXY = new XYZ(_riserConnectionPoint.X, _riserConnectionPoint.Y, 0);
                        
                        // Проекция оси стояка на магистраль
                        XYZ toRiserXY = riserAxisXY - mainStartXY;
                        double projLength = toRiserXY.DotProduct(mainDirXY);
                        XYZ point4XY = mainStartXY + mainDirXY * projLength;
                        
                        // Координата Z точки 4 по уравнению прямой магистрали
                        double t = projLength / (mainEndXY - mainStartXY).GetLength();
                        double z4 = _mainLineStartPoint.Z + t * (_mainLineEndPoint.Z - _mainLineStartPoint.Z);
                        XYZ point4 = new XYZ(point4XY.X, point4XY.Y, z4);
                        
                        // === НАПРАВЛЕНИЕ УКЛОНА ===
                        // Определяем направление потока (от высокой точки к низкой)
                        bool flowFromStartToEnd = _mainLineStartPoint.Z > _mainLineEndPoint.Z;
                        // Направление ВНИЗ по течению (по потоку)
                        XYZ downstreamDir = flowFromStartToEnd ? mainDirXY : -mainDirXY;
                        
                        // Направление перпендикуляра к магистрали (в плане)
                        XYZ perpDirXY = new XYZ(-mainDirXY.Y, mainDirXY.X, 0);
                        
                        // Определяем знак направления к стояку
                        XYZ toRiser = new XYZ(_riserConnectionPoint.X - point4.X, 
                                              _riserConnectionPoint.Y - point4.Y, 0);
                        double sideDot = toRiser.DotProduct(perpDirXY);
                        int directionSign = sideDot > 0 ? 1 : -1;
                        perpDirXY = perpDirXY * directionSign;
                        
                        // === ТОЧКА 5 - на магистрали, ниже точки 4 на sqrt(a/2) ===
                        // Откладываем от точки 4 по направлению уклона вниз
                        double offsetDown = Math.Sqrt(segmentALength / 2); // sqrt(0.5) = 0.707м
                        double zOffsetDown = offsetDown * slopeRatio;
                        
                        XYZ point5 = new XYZ(
                            point4.X + downstreamDir.X * offsetDown,
                            point4.Y + downstreamDir.Y * offsetDown,
                            point4.Z - zOffsetDown);
                        
                        // === ТОЧКА 7 - конец сегмента "а", начало сегмента "б" ===
                        // Сегмент "a" - равнобедренный прямоугольный треугольник:
                        // - Катет b = √(a/2) = 0.707м перпендикулярно магистрали (perpDirXY)
                        // - Катет b = 0.707м вдоль магистрали ВВЕРХ (upstreamDir)
                        // - Гипотенуза a = 1м
                        // - Уклон сегмента А = уклону магистрали (2%)
                        XYZ upstreamDir = -downstreamDir;
                        double b = Math.Sqrt(segmentALength / 2); // 0.707м
                        
                        // Подъем по Z для сегмента А: длина А * уклон = 1м * slopeRatio
                        double zRiseA = segmentALength * slopeRatio;
                        
                        XYZ point7 = new XYZ(
                            point5.X + upstreamDir.X * b + perpDirXY.X * b,
                            point5.Y + upstreamDir.Y * b + perpDirXY.Y * b,
                            point5.Z + zRiseA);
                        
                        // === ТОЧКА 6 - у стояка (перпендикулярно от точки 7 с уклоном) ===
                        // Сегмент "б" идет перпендикулярно магистрали (90°)
                        // с уклоном как у магистрали
                        // x6 = x3, y6 = y3 (координаты оси стояка)
                        
                        // Расстояние в плане от точки 7 до оси стояка
                        double distB_XY = Math.Sqrt(
                            Math.Pow(_riserConnectionPoint.X - point7.X, 2) + 
                            Math.Pow(_riserConnectionPoint.Y - point7.Y, 2));
                        // Подъем по Z для сегмента Б
                        double zRiseB = distB_XY * slopeRatio;
                        
                        XYZ point6 = new XYZ(
                            _riserConnectionPoint.X,
                            _riserConnectionPoint.Y,
                            point7.Z + zRiseB);
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Point 4 (projection on main): ({point4.X:F3},{point4.Y:F3},{point4.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 5 (start of A on main): ({point5.X:F3},{point5.Y:F3},{point5.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 7 (end of 45° segment A): ({point7.X:F3},{point7.Y:F3},{point7.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 6 (at riser): ({point6.X:F3},{point6.Y:F3},{point6.Z:F3})");
                        
                        // === СЕГМЕНТ "а" (от точки 5 к точке 7) - под 45° вверх от магистрали ===
                        DebugLogger.Log("[TRACER-COMMAND] Creating segment A (45° up from main, 1m)...");
                        Pipe segmentA = Pipe.Create(
                            doc,
                            systemTypeId,
                            pipeTypeId,
                            levelId,
                            point5,
                            point7);
                        
                        if (segmentA == null)
                        {
                            DebugLogger.Log("[TRACER-COMMAND] ERROR: Failed to create segment A");
                            tx.RollBack();
                            return;
                        }
                        
                        Parameter diameterParamA = segmentA.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diameterParamA != null && !diameterParamA.IsReadOnly)
                            diameterParamA.Set(_pipeDiameter);
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Segment A (45°) created: {segmentA.Id}");
                        
                        // === СЕГМЕНТ "б" (от точки 7 к точке 6) - под 90° к магистрали ===
                        DebugLogger.Log("[TRACER-COMMAND] Creating segment B (90° to main, with slope)...");
                        Pipe segmentB = Pipe.Create(
                            doc,
                            systemTypeId,
                            pipeTypeId,
                            levelId,
                            point7,
                            point6);
                        
                        if (segmentB == null)
                        {
                            DebugLogger.Log("[TRACER-COMMAND] ERROR: Failed to create segment B");
                            tx.RollBack();
                            return;
                        }
                        
                        Parameter diameterParamB = segmentB.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diameterParamB != null && !diameterParamB.IsReadOnly)
                            diameterParamB.Set(_pipeDiameter);
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Segment B (90°) created: {segmentB.Id}");
                        
                        // Подгоняем стояк к точке 6 (сохраняем вертикальность)
                        AdjustRiserForLConnection(doc, _riserElementId, point6);

                        tx.Commit();
                        DebugLogger.Log("[TRACER-COMMAND] L-shaped connection created successfully");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[TRACER-COMMAND] ERROR during L-connection creation: {ex.Message}\n{ex.StackTrace}");
                        tx.RollBack();
                        MessageBox.Show($"Ошибка при создании L-образного подключения: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-COMMAND] ERROR in Execute: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AdjustRiser(Document doc, ElementId riserId, XYZ connectionPoint)
        {
            Element riserElement = doc.GetElement(riserId);
            if (riserElement == null || !(riserElement is Pipe riserPipe))
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser not found for adjustment");
                return;
            }

            LocationCurve riserLocation = riserPipe.Location as LocationCurve;
            if (riserLocation == null)
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser has no LocationCurve");
                return;
            }

            Line riserLine = riserLocation.Curve as Line;
            if (riserLine == null)
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser curve is not a Line");
                return;
            }

            XYZ riserEnd0 = riserLine.GetEndPoint(0);
            XYZ riserEnd1 = riserLine.GetEndPoint(1);

            double dist0 = riserEnd0.DistanceTo(connectionPoint);
            double dist1 = riserEnd1.DistanceTo(connectionPoint);

            XYZ riserTop = (dist0 > dist1) ? riserEnd0 : riserEnd1;
            XYZ newRiserBottom = connectionPoint;

            Line newRiserLine = Line.CreateBound(newRiserBottom, riserTop);
            riserLocation.Curve = newRiserLine;

            DebugLogger.Log($"[TRACER-COMMAND] Adjusted riser bottom to ({newRiserBottom.X:F3},{newRiserBottom.Y:F3},{newRiserBottom.Z:F3})");
        }

        private void AdjustRiserForLConnection(Document doc, ElementId riserId, XYZ connectionPoint)
        {
            Element riserElement = doc.GetElement(riserId);
            if (riserElement == null || !(riserElement is Pipe riserPipe))
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser not found for adjustment");
                return;
            }

            LocationCurve riserLocation = riserPipe.Location as LocationCurve;
            if (riserLocation == null)
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser has no LocationCurve");
                return;
            }

            Line riserLine = riserLocation.Curve as Line;
            if (riserLine == null)
            {
                DebugLogger.Log("[TRACER-COMMAND] WARNING: Riser curve is not a Line");
                return;
            }

            XYZ riserEnd0 = riserLine.GetEndPoint(0);
            XYZ riserEnd1 = riserLine.GetEndPoint(1);

            double dist0 = riserEnd0.DistanceTo(connectionPoint);
            double dist1 = riserEnd1.DistanceTo(connectionPoint);

            XYZ riserTop = (dist0 > dist1) ? riserEnd0 : riserEnd1;
            XYZ riserBottom = (dist0 > dist1) ? riserEnd1 : riserEnd0;
            
            // Стояк должен оставаться вертикальным - сохраняем X, Y оси стояка
            // Меняем только Z координату нижней точки
            XYZ newRiserBottom = new XYZ(riserBottom.X, riserBottom.Y, connectionPoint.Z);

            Line newRiserLine = Line.CreateBound(newRiserBottom, riserTop);
            riserLocation.Curve = newRiserLine;

            DebugLogger.Log($"[TRACER-COMMAND] Adjusted riser bottom to Z={newRiserBottom.Z:F3} (X,Y unchanged: {newRiserBottom.X:F3},{newRiserBottom.Y:F3})");
        }

        private static Level GetLevelByElevation(Document doc, double elevation)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - elevation))
                .ToList();

            return levels.FirstOrDefault();
        }
    }
}
