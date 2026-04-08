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
        private static bool _addFittings;

        public string GetName()
        {
            return "Tracer Create 45-degree Connection";
        }

        public static void SetConnectionData(
            ElementId mainPipeId, ElementId riserId,
            XYZ connectionPoint, XYZ riserConnectionPoint, double pipeDiameter,
            double mainLineSlope, bool addFittings = false)
        {
            _mainPipeElementId = mainPipeId;
            _riserElementId = riserId;
            _connectionPoint = connectionPoint;
            _riserConnectionPoint = riserConnectionPoint;
            _pipeDiameter = pipeDiameter;
            _mainLineSlope = mainLineSlope;
            _addFittings = addFittings;
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
                        // Получаем геометрию магистрали
                        LocationCurve mainLoc = mainPipe.Location as LocationCurve;
                        XYZ mainStart = mainLoc.Curve.GetEndPoint(0);
                        XYZ mainEnd = mainLoc.Curve.GetEndPoint(1);
                        
                        // Рассчитываем точки для 45° присоединения по схеме:
                        // 1. Точка 4 - проекция оси стояка на магистраль
                        // 2. Точка 5 - на магистрали, смещена от точки 4 на расстояние = расстоянию до стояка
                        //    (это создает угол 45° в плане)
                        // 3. Труба идет от стояка к точке 5 с заданным уклоном
                        
                        double slopeRatio = _mainLineSlope / 100.0;
                        
                        // === ТОЧКА 4 - проекция оси стояка на магистраль (в плане XY) ===
                        // Направление магистрали в плане (XY)
                        XYZ mainStartXY = new XYZ(mainStart.X, mainStart.Y, 0);
                        XYZ mainEndXY = new XYZ(mainEnd.X, mainEnd.Y, 0);
                        XYZ mainDirXY = (mainEndXY - mainStartXY).Normalize();
                        double mainLengthXY = (mainEndXY - mainStartXY).GetLength();
                        
                        // Ось стояка в плане
                        XYZ riserAxisXY = new XYZ(_riserConnectionPoint.X, _riserConnectionPoint.Y, 0);
                        
                        // Проекция оси стояка на магистраль
                        XYZ toRiserXY = riserAxisXY - mainStartXY;
                        double projLength = toRiserXY.DotProduct(mainDirXY);
                        XYZ point4XY = mainStartXY + mainDirXY * projLength;
                        
                        // Параметр t для точки 4 на магистрали
                        double t4 = projLength / mainLengthXY;
                        
                        // Координата Z точки 4 по уравнению прямой магистрали
                        double z4 = mainStart.Z + t4 * (mainEnd.Z - mainStart.Z);
                        XYZ point4 = new XYZ(point4XY.X, point4XY.Y, z4);
                        
                        // === НАПРАВЛЕНИЕ ПОТОКА ===
                        bool flowFromStartToEnd = mainStart.Z > mainEnd.Z;
                        XYZ downstreamDir = flowFromStartToEnd ? mainDirXY : -mainDirXY;
                        
                        // === РАССТОЯНИЕ ОТ СТОЯКА ДО МАГИСТРАЛИ ===
                        double distanceToRiser = (riserAxisXY - point4XY).GetLength();
                        
                        // === ТОЧКА 5 - на магистрали, смещена от точки 4 ===
                        // Для угла 45° в плане, смещаемся от точки 4 на то же расстояние,
                        // что и от стояка до магистрали
                        // Смещаемся ВНИЗ по потоку (в направлении уклона магистрали)
                        double offsetAlongMain = distanceToRiser;
                        
                        // Параметр t для точки 5
                        double t5 = t4 + (flowFromStartToEnd ? offsetAlongMain / mainLengthXY : -offsetAlongMain / mainLengthXY);
                        t5 = Math.Max(0, Math.Min(1, t5)); // Ограничиваем пределами магистрали
                        
                        // Точка 5 на магистрали (с Z по уравнению магистрали)
                        XYZ point5 = new XYZ(
                            mainStart.X + t5 * (mainEnd.X - mainStart.X),
                            mainStart.Y + t5 * (mainEnd.Y - mainStart.Y),
                            mainStart.Z + t5 * (mainEnd.Z - mainStart.Z));
                        
                        // === ТОЧКА У СТОЯКА ===
                        // Высота точки у стояка = высота точки 5 + падение по уклону
                        double xyDistancePipe = (riserAxisXY - new XYZ(point5.X, point5.Y, 0)).GetLength();
                        double zDrop = xyDistancePipe * slopeRatio;
                        double riserZ = point5.Z + zDrop;
                        
                        XYZ startPoint = new XYZ(
                            _riserConnectionPoint.X,
                            _riserConnectionPoint.Y,
                            riserZ);
                        
                        // Точка на магистрали
                        XYZ endPoint = point5;
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Creating 45° connection pipe with slope {_mainLineSlope}%...");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 4 (perpendicular): ({point4.X:F3},{point4.Y:F3},{point4.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Distance to riser: {distanceToRiser:F3}");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 5 (on main, offset): ({point5.X:F3},{point5.Y:F3},{point5.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Start (at riser): ({startPoint.X:F3},{startPoint.Y:F3},{startPoint.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] End (at main): ({endPoint.X:F3},{endPoint.Y:F3},{endPoint.Z:F3})");
                        
                        Pipe connectionPipe = Pipe.Create(
                            doc,
                            systemTypeId,
                            pipeTypeId,
                            levelId,
                            startPoint,
                            endPoint);

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

                        // Подгоняем стояк (используем startPoint - точка у стояка выше)
                        AdjustRiser(doc, _riserElementId, startPoint);

                        // Create fitting if checkbox is enabled
                        if (_addFittings)
                        {
                            TracerFittingHelper.CreateFittingBetweenRiserAndPipe(doc, _riserElementId, connectionPipe, startPoint);
                        }

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
        private static bool _addFittings;

        public string GetName()
        {
            return "Tracer Create L-shaped Connection";
        }

        public static void SetConnectionData(
            ElementId mainPipeId, ElementId riserId,
            XYZ connectionPoint, XYZ riserConnectionPoint, double pipeDiameter,
            double mainLineSlope, XYZ mainLineStart, XYZ mainLineEnd, bool addFittings = false)
        {
            _mainPipeElementId = mainPipeId;
            _riserElementId = riserId;
            _connectionPoint = connectionPoint;
            _riserConnectionPoint = riserConnectionPoint;
            _pipeDiameter = pipeDiameter;
            _mainLineSlope = mainLineSlope;
            _mainLineStartPoint = mainLineStart;
            _mainLineEndPoint = mainLineEnd;
            _addFittings = addFittings;
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
                        double mainLengthXY = (mainEndXY - mainStartXY).GetLength();
                        
                        // Ось стояка в плане
                        XYZ riserAxisXY = new XYZ(_riserConnectionPoint.X, _riserConnectionPoint.Y, 0);
                        
                        // Проекция оси стояка на магистраль
                        XYZ toRiserXY = riserAxisXY - mainStartXY;
                        double projLength = toRiserXY.DotProduct(mainDirXY);
                        XYZ point4XY = mainStartXY + mainDirXY * projLength;
                        
                        // Координата Z точки 4 по уравнению прямой магистрали
                        double t4 = projLength / mainLengthXY;
                        double z4 = _mainLineStartPoint.Z + t4 * (_mainLineEndPoint.Z - _mainLineStartPoint.Z);
                        XYZ point4 = new XYZ(point4XY.X, point4XY.Y, z4);
                        
                        // === УКЛОН МАГИСТРАЛИ ===
                        // Рассчитываем реальный уклон магистрали
                        double mainLineZDiff = _mainLineEndPoint.Z - _mainLineStartPoint.Z;
                        double mainLineSlopeRatio = mainLineZDiff / mainLengthXY; // может быть отрицательным
                        double mainLineSlopeAbs = Math.Abs(mainLineSlopeRatio);
                        
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
                        // Важно: точка 5 должна быть на линии магистрали, поэтому используем уклон магистрали
                        double offsetDown = Math.Sqrt(segmentALength / 2); // sqrt(0.5) = 0.707м
                        
                        // Параметр t для точки 5 (смещена на offsetDown вниз по потоку)
                        double t5 = t4 + (flowFromStartToEnd ? offsetDown / mainLengthXY : -offsetDown / mainLengthXY);
                        t5 = Math.Max(0, Math.Min(1, t5)); // Ограничиваем пределами магистрали
                        
                        // Точка 5 на магистрали
                        XYZ point5 = new XYZ(
                            _mainLineStartPoint.X + t5 * (_mainLineEndPoint.X - _mainLineStartPoint.X),
                            _mainLineStartPoint.Y + t5 * (_mainLineEndPoint.Y - _mainLineStartPoint.Y),
                            _mainLineStartPoint.Z + t5 * (_mainLineEndPoint.Z - _mainLineStartPoint.Z));
                        
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
                        
                        // Устанавливаем уклон для сегмента Б (горизонтальный, уклон = 0)
                        Parameter slopeParamB = segmentB.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                        if (slopeParamB != null && !slopeParamB.IsReadOnly)
                        {
                            slopeParamB.Set(0);
                        }
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Segment B (90°) created: {segmentB.Id}");
                        
                        // Create fitting between segments A and B if checkbox is enabled
                        if (_addFittings)
                        {
                            TracerFittingHelper.CreateFittingBetweenPipes(doc, segmentA, segmentB, point7);
                        }
                        
                        // Подгоняем стояк к точке 6 (сохраняем вертикальность)
                        AdjustRiserForLConnection(doc, _riserElementId, point6);
                        
                        // Create fitting if checkbox is enabled (between riser and segment B)
                        if (_addFittings)
                        {
                            TracerFittingHelper.CreateFittingBetweenRiserAndPipe(doc, _riserElementId, segmentB, point6);
                        }
                        
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

    /// <summary>
    /// ExternalEvent handler for creating bottom perpendicular connection in Tracer module
    /// </summary>
    public class TracerCreateBottomConnectionHandler : IExternalEventHandler
    {
        private static ElementId _mainPipeElementId;
        private static ElementId _riserElementId;
        private static XYZ _connectionPoint;
        private static XYZ _riserConnectionPoint;
        private static double _pipeDiameter;
        private static double _mainLineSlope;
        private static XYZ _mainLineStartPoint;
        private static XYZ _mainLineEndPoint;
        private static bool _addFittings;

        public string GetName()
        {
            return "Tracer Create Bottom Perpendicular Connection";
        }

        public static void SetConnectionData(
            ElementId mainPipeId, ElementId riserId,
            XYZ connectionPoint, XYZ riserConnectionPoint, double pipeDiameter,
            double mainLineSlope, XYZ mainLineStartPoint, XYZ mainLineEndPoint, bool addFittings = false)
        {
            _mainPipeElementId = mainPipeId;
            _riserElementId = riserId;
            _connectionPoint = connectionPoint;
            _riserConnectionPoint = riserConnectionPoint;
            _pipeDiameter = pipeDiameter;
            _mainLineSlope = mainLineSlope;
            _mainLineStartPoint = mainLineStartPoint;
            _mainLineEndPoint = mainLineEndPoint;
            _addFittings = addFittings;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Log("[TRACER-COMMAND] Create bottom perpendicular connection event executing...");

                if (_mainPipeElementId == null || _riserElementId == null || _connectionPoint == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Missing connection data");
                    MessageBox.Show("Ошибка: отсутствуют данные для создания подключения", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Document doc = app.ActiveUIDocument.Document;

                // Get main pipe system type and pipe type
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
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Main pipe element is not a Pipe");
                    return;
                }

                // Get pipe type and system type from main pipe
                ElementId pipeTypeId = mainPipe.PipeType.Id;
                ElementId systemTypeId = mainPipe.MEPSystem.GetTypeId();
                
                // Get level
                Level level = GetLevelByElevation(doc, _connectionPoint.Z);
                if (level == null)
                {
                    DebugLogger.Log("[TRACER-COMMAND] ERROR: Could not find appropriate level");
                    MessageBox.Show("Ошибка: не удалось определить уровень для создания трубы", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                ElementId levelId = level.Id;

                DebugLogger.Log($"[TRACER-COMMAND] Creating bottom perpendicular connection...");
                DebugLogger.Log($"[TRACER-COMMAND] Main line slope: {_mainLineSlope:F2}%");
                DebugLogger.Log($"[TRACER-COMMAND] Pipe diameter: {_pipeDiameter * 304.8:F0}mm");

                using (Transaction tx = new Transaction(doc, "Create Bottom Perpendicular Connection"))
                {
                    tx.Start();

                    try
                    {
                        // Calculate geometry for bottom perpendicular connection
                        double segmentALength = 1.0; // 1 meter vertical drop
                        double slopeRatio = _mainLineSlope / 100.0;
                        
                        // Направление магистрали в плане (XY)
                        XYZ mainStartXY = new XYZ(_mainLineStartPoint.X, _mainLineStartPoint.Y, 0);
                        XYZ mainEndXY = new XYZ(_mainLineEndPoint.X, _mainLineEndPoint.Y, 0);
                        XYZ mainDirXY = (mainEndXY - mainStartXY).Normalize();
                        
                        // === ТОЧКА 4 - проекция оси стояка на магистраль (перпендикуляр) ===
                        // Рассчитываем заново, т.к. _connectionPoint сдвинут для 45° подключения
                        XYZ riserAxisXY = new XYZ(_riserConnectionPoint.X, _riserConnectionPoint.Y, 0);
                        XYZ toRiserFromStart = riserAxisXY - mainStartXY;
                        double projLength = toRiserFromStart.DotProduct(mainDirXY);
                        XYZ point4XY = mainStartXY + mainDirXY * projLength;
                        
                        // Получаем Z точки 4 на магистрали
                        double z4 = _mainLineStartPoint.Z + 
                            (projLength / (mainEndXY - mainStartXY).GetLength()) * 
                            (_mainLineEndPoint.Z - _mainLineStartPoint.Z);
                        XYZ point4 = new XYZ(point4XY.X, point4XY.Y, z4);
                        
                        // Направление перпендикуляра к магистрали (в плане)
                        XYZ perpDirXY = new XYZ(-mainDirXY.Y, mainDirXY.X, 0);
                        
                        // Определяем знак направления к стояку
                        XYZ toRiser = new XYZ(_riserConnectionPoint.X - point4.X, 
                                              _riserConnectionPoint.Y - point4.Y, 0);
                        double sideDot = toRiser.DotProduct(perpDirXY);
                        int directionSign = sideDot > 0 ? 1 : -1;
                        perpDirXY = perpDirXY * directionSign;
                        
                        // === ТОЧКА 5 - ниже точки 4 на 1 метр (сегмент А - вертикальный спуск) ===
                        // Сегмент А идет строго вертикально вниз от магистрали на 1 метр
                        // Начало сегмента А на оси магистрали (точка 4)
                        // Конец сегмента А внизу (точка 5)
                        // x и y остаются теми же, меняется только z
                        double zDropA = segmentALength * slopeRatio; // Подъем по уклону за 1м
                        
                        XYZ point5 = new XYZ(
                            point4.X,  // x тот же
                            point4.Y,  // y тот же
                            point4.Z - segmentALength + zDropA); // вниз на 1м + уклон
                        
                        // === ТОЧКА 6 - у стояка (сегмент Б - перпендикулярно к магистрали) ===
                        // От точки 5 перпендикулярно к магистрали до оси стояка
                        // 
                        // Точка 4 - проекция стояка на магистраль
                        // Точка 5 - под точкой 4 (вертикально)
                        // Стояк находится на перпендикуляре от точки 4
                        // Сегмент Б должен идти от точки 5 перпендикулярно к магистрали
                        // до той же точки перпендикуляра, где находится стояк
                        
                        // Расстояние от магистрали до стояка (в плане)
                        double distRiserToMain = Math.Sqrt(
                            Math.Pow(_riserConnectionPoint.X - point4.X, 2) + 
                            Math.Pow(_riserConnectionPoint.Y - point4.Y, 2));
                        
                        // Подъем по Z за это расстояние
                        double zRiseB = distRiserToMain * slopeRatio;
                        
                        // Точка 6: от точки 5 перпендикулярно к магистрали на то же расстояние
                        // что и стояк от магистрали
                        XYZ point6 = new XYZ(
                            point5.X + perpDirXY.X * distRiserToMain,
                            point5.Y + perpDirXY.Y * distRiserToMain,
                            point5.Z + zRiseB);
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Point 4 (on main): ({point4.X:F3},{point4.Y:F3},{point4.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 5 (end of A): ({point5.X:F3},{point5.Y:F3},{point5.Z:F3})");
                        DebugLogger.Log($"[TRACER-COMMAND] Point 6 (at riser): ({point6.X:F3},{point6.Y:F3},{point6.Z:F3})");
                        
                        // === СЕГМЕНТ "А" (от точки 4 к точке 5) - перпендикулярно от магистрали ===
                        DebugLogger.Log("[TRACER-COMMAND] Creating segment A (perpendicular from main, 1m)...");
                        Pipe segmentA = Pipe.Create(
                            doc,
                            systemTypeId,
                            pipeTypeId,
                            levelId,
                            point4,
                            point5);
                        
                        if (segmentA == null)
                        {
                            DebugLogger.Log("[TRACER-COMMAND] ERROR: Failed to create segment A");
                            tx.RollBack();
                            return;
                        }
                        
                        Parameter diameterParamA = segmentA.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (diameterParamA != null && !diameterParamA.IsReadOnly)
                            diameterParamA.Set(_pipeDiameter);
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Segment A created: {segmentA.Id}");
                        
                        // === СЕГМЕНТ "Б" (от точки 5 к точке 6) - к стояку ===
                        DebugLogger.Log("[TRACER-COMMAND] Creating segment B (to riser)...");
                        Pipe segmentB = Pipe.Create(
                            doc,
                            systemTypeId,
                            pipeTypeId,
                            levelId,
                            point5,
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
                        
                        // Устанавливаем уклон для сегмента Б (горизонтальный, уклон = 0)
                        Parameter slopeParamB = segmentB.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE);
                        if (slopeParamB != null && !slopeParamB.IsReadOnly)
                        {
                            slopeParamB.Set(0);
                        }
                        
                        DebugLogger.Log($"[TRACER-COMMAND] Segment B created: {segmentB.Id}");
                        
                        // Create fitting between segments A and B if checkbox is enabled
                        if (_addFittings)
                        {
                            TracerFittingHelper.CreateFittingBetweenPipes(doc, segmentA, segmentB, point5);
                        }
                        
                        // Подгоняем стояк к точке 6 (по Z, сохраняя X,Y оси стояка)
                        AdjustRiserForBottomConnection(doc, _riserElementId, point6);

                        // Create fitting if checkbox is enabled (between riser and segment B)
                        if (_addFittings)
                        {
                            TracerFittingHelper.CreateFittingBetweenRiserAndPipe(doc, _riserElementId, segmentB, point6);
                        }

                        tx.Commit();
                        DebugLogger.Log("[TRACER-COMMAND] Bottom perpendicular connection created successfully");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[TRACER-COMMAND] ERROR during bottom connection creation: {ex.Message}\n{ex.StackTrace}");
                        tx.RollBack();
                        MessageBox.Show($"Ошибка при создании нижнего подключения: {ex.Message}", "Ошибка",
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

        private void AdjustRiserForBottomConnection(Document doc, ElementId riserId, XYZ connectionPoint)
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

    }

    /// <summary>
    /// Helper class for creating pipe fittings in Tracer module
    /// </summary>
    public static class TracerFittingHelper
    {
        /// <summary>
        /// Creates an elbow fitting between two pipes using NewElbowFitting
        /// </summary>
        public static Element CreateFittingBetweenPipes(Document doc, Pipe pipe1, Pipe pipe2, XYZ connectionPoint)
        {
            try
            {
                DebugLogger.Log("[TRACER-FITTING] Creating fitting between two pipes...");

                // Find the closest connector on pipe1 to the connection point
                Connector pipe1Connector = GetClosestConnector(pipe1, connectionPoint);
                if (pipe1Connector == null)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: Could not find pipe1 connector");
                    return null;
                }

                // Find the closest connector on pipe2 to the connection point
                Connector pipe2Connector = GetClosestConnector(pipe2, connectionPoint);
                if (pipe2Connector == null)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: Could not find pipe2 connector");
                    return null;
                }

                // Check if connectors are already connected
                if (pipe1Connector.IsConnected || pipe2Connector.IsConnected)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: One of the connectors is already connected");
                    return null;
                }

                // Create the elbow fitting
                Element fitting = doc.Create.NewElbowFitting(pipe1Connector, pipe2Connector);
                if (fitting != null)
                {
                    DebugLogger.Log($"[TRACER-FITTING] Successfully created fitting between pipes: {fitting.Id}");
                }
                else
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: NewElbowFitting returned null");
                }

                return fitting;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-FITTING] ERROR creating fitting between pipes: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Creates an elbow fitting between a riser and a pipe using NewElbowFitting
        /// </summary>
        public static Element CreateFittingBetweenRiserAndPipe(Document doc, ElementId riserId, Pipe pipe, XYZ connectionPoint)
        {
            try
            {
                DebugLogger.Log("[TRACER-FITTING] Creating fitting between riser and pipe...");

                // Get the riser pipe
                Pipe riserPipe = doc.GetElement(riserId) as Pipe;
                if (riserPipe == null)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: Riser not found for fitting creation");
                    return null;
                }

                // Find the closest connector on the riser to the connection point
                Connector riserConnector = GetClosestConnector(riserPipe, connectionPoint);
                if (riserConnector == null)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: Could not find riser connector");
                    return null;
                }

                // Find the closest connector on the pipe to the connection point
                Connector pipeConnector = GetClosestConnector(pipe, connectionPoint);
                if (pipeConnector == null)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: Could not find pipe connector");
                    return null;
                }

                // Check if connectors are already connected
                if (riserConnector.IsConnected || pipeConnector.IsConnected)
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: One of the connectors is already connected");
                    return null;
                }

                // Create the elbow fitting
                Element fitting = doc.Create.NewElbowFitting(riserConnector, pipeConnector);
                if (fitting != null)
                {
                    DebugLogger.Log($"[TRACER-FITTING] Successfully created fitting: {fitting.Id}");
                }
                else
                {
                    DebugLogger.Log("[TRACER-FITTING] WARNING: NewElbowFitting returned null");
                }

                return fitting;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-FITTING] ERROR creating fitting: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets the closest connector on a pipe to a given point
        /// </summary>
        private static Connector GetClosestConnector(Pipe pipe, XYZ point)
        {
            ConnectorManager connectorManager = pipe.ConnectorManager;
            if (connectorManager == null) return null;

            ConnectorSet connectors = connectorManager.Connectors;
            if (connectors == null || connectors.Size == 0) return null;

            Connector closestConnector = null;
            double minDistance = double.MaxValue;

            foreach (Connector connector in connectors)
            {
                double distance = connector.Origin.DistanceTo(point);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestConnector = connector;
                }
            }

            return closestConnector;
        }
    }
}
