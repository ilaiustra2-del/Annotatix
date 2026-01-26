using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;

using System.Reflection;

namespace dwg2rvt.Core
{
    /// <summary>
    /// Analyzer for DWG imports using block name extraction method
    /// </summary>
    public class AnalyzeByBlockName
    {
        private readonly Document _doc;
        private const string LOG_DIRECTORY = @"C:\Users\Свеж как огурец\Desktop\Эксперимент Annotatix\logs";
    
        public AnalyzeByBlockName(Document doc)
        {
            _doc = doc;
        }
    
        private class ComponentInfo
        {
            public string Type { get; set; }
            public XYZ Center { get; set; }
            public double? ArcStartAngle { get; set; }  // For arcs: start angle in radians
            public double? ArcEndAngle { get; set; }    // For arcs: end angle in radians
            public double? ArcRadius { get; set; }      // For arcs: radius
            public XYZ ArcStartPoint { get; set; }      // For arcs: actual start point
            public XYZ ArcEndPoint { get; set; }        // For arcs: actual end point
        }
    
        private class DetailedBlockInfo
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public XYZ ComputedCenter { get; set; }
            public double RotationAngle { get; set; } = 0;  // Rotation angle in degrees (0/90/180/270)
            public List<ComponentInfo> Components { get; set; } = new List<ComponentInfo>();
        }
    
        public AnalysisResult Analyze(ImportInstance importInstance, Action<string> statusCallback = null, bool enableLogging = false)
        {
            AnalysisResult result = new AnalysisResult();
            result.ImportInstanceName = importInstance.Name;
            result.ImportInstanceId = (int)importInstance.Id.Value;
            result.AnalysisTimestamp = DateTime.Now;
            
            List<DetailedBlockInfo> blocks = new List<DetailedBlockInfo>();
            List<string> debugLog = new List<string>();
        
            try
            {
                statusCallback?.Invoke($"Analyzing import: {importInstance.Name}");
        
                Options opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };
                GeometryElement geoElement = importInstance.get_Geometry(opt);
                        
                if (geoElement == null)
                {
                    result.ErrorMessage = "Failed to get geometry.";
                    return result;
                }
        
                // Temporary list to collect raw block instances
                List<DetailedBlockInfo> rawInstances = new List<DetailedBlockInfo>();
                CollectBlockData(geoElement, rawInstances, Transform.Identity, debugLog, null);
        
                // Group and Number blocks
                var groupedByName = rawInstances.GroupBy(b => b.Name);
                foreach (var group in groupedByName)
                {
                    int count = 1;
                    foreach (var instance in group)
                    {
                        instance.Number = count++;
                        blocks.Add(instance);
                    }
                }
        
                result.BlockCount = blocks.Count;
                                
                // Generate Formatted Summary
                StringBuilder summaryBuilder = new StringBuilder();
                summaryBuilder.AppendLine($"Общее количество блоков – {blocks.Count}");
                                
                var summaryCounts = rawInstances.GroupBy(i => i.Name)
                    .OrderByDescending(g => g.Count());
                                    
                foreach (var group in summaryCounts)
                {
                    summaryBuilder.AppendLine($"\t{group.Key} – {group.Count()}");
                }
                                
                result.Summary = summaryBuilder.ToString().TrimEnd();
                                
                statusCallback?.Invoke(result.Summary);
                statusCallback?.Invoke($"Processed {blocks.Count} blocks.");
                
                // Convert to BlockData for in-memory storage
                foreach (var block in blocks)
                {
                    BlockData blockData = new BlockData
                    {
                        Name = block.Name,
                        Number = block.Number,
                        CenterX = block.ComputedCenter.X,
                        CenterY = block.ComputedCenter.Y,
                        RotationAngle = block.RotationAngle
                    };
                    
                    // Convert components
                    foreach (var comp in block.Components)
                    {
                        blockData.Components.Add(new ComponentData
                        {
                            Type = comp.Type,
                            CenterX = comp.Center.X,
                            CenterY = comp.Center.Y,
                            ArcRadius = comp.ArcRadius,
                            ArcStartAngle = comp.ArcStartAngle,
                            ArcEndAngle = comp.ArcEndAngle
                        });
                    }
                    
                    result.BlockData.Add(blockData);
                    
                    // Group by type for easy access
                    if (!result.BlocksByType.ContainsKey(block.Name))
                    {
                        result.BlocksByType[block.Name] = new List<BlockData>();
                    }
                    result.BlocksByType[block.Name].Add(blockData);
                }
                
                // Store in cache
                AnalysisDataCache.StoreAnalysisResult(result);
                statusCallback?.Invoke("Analysis data stored in memory cache.");
                
                // Optional log file generation (controlled by user setting)
                if (enableLogging)
                {
                    string logPath = GenerateDetailedLog(importInstance, blocks, debugLog, result.Summary, statusCallback);
                    result.LogFilePath = logPath;
                    statusCallback?.Invoke($"Log file saved: {logPath}");
                }
                else
                {
                    statusCallback?.Invoke("Logging disabled - data stored ONLY in memory cache.");
                    result.LogFilePath = null;
                }
                
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Error: {ex.Message}";
                statusCallback?.Invoke($"ERROR: {ex.Message}");
            }
        
            return result;
        }
        
        private void CollectBlockData(GeometryElement geoElement, List<DetailedBlockInfo> instances, 
            Transform transform, List<string> debugLog, DetailedBlockInfo currentBlock)
        {
            if (geoElement == null) return;
        
            foreach (GeometryObject geoObj in geoElement)
            {
                if (geoObj is GeometryInstance geoInstance)
                {
                    Element sym = _doc.GetElement(geoInstance.GetSymbolGeometryId().SymbolId);
                    string symbolName = sym?.Name ?? "";
                            
                    string blockName = ExtractBlockNameFromString(symbolName);
                    Transform combinedTransform = transform.Multiply(geoInstance.Transform);
        
                    DetailedBlockInfo nextBlock = currentBlock;
        
                    if (!string.IsNullOrEmpty(blockName))
                    {
                        debugLog.Add($"Found Block: '{blockName}' (Symbol: {symbolName})");
                        nextBlock = new DetailedBlockInfo { Name = blockName };
                        instances.Add(nextBlock);
                    }
                    else
                    {
                        debugLog.Add($"Found Container: '{symbolName}'");
                    }
                            
                    // Recursively search deeper
                    CollectBlockData(geoInstance.GetSymbolGeometry(), instances, combinedTransform, debugLog, nextBlock);
                }
                else if (currentBlock != null)
                {
                    // If we are inside a named block, collect geometric primitives
                    ComponentInfo compInfo = GetElementInfo(geoObj);
                    if (compInfo != null && compInfo.Center != null)
                    {
                        XYZ worldCenter = transform.OfPoint(compInfo.Center);
                        compInfo.Center = worldCenter;
                                
                        // Finer deduplication: Type + World Center
                        string coordKey = $"{compInfo.Type}_{worldCenter.X:F6}_{worldCenter.Y:F6}";
                                
                        // Check if this component is already added to this specific block instance
                        if (!currentBlock.Components.Any(c => $"{c.Type}_{c.Center.X:F6}_{c.Center.Y:F6}" == coordKey))
                        {
                            currentBlock.Components.Add(compInfo);
                                    
                            // Re-calculate Block Center based on block type
                            bool isSocket = currentBlock.Name.ToLower().Contains("розетка");
                            
                            if (isSocket)
                            {
                                // For sockets: use arc center if available
                                var arcComponent = currentBlock.Components.FirstOrDefault(c => c.Type == "Arc" && c.ArcRadius.HasValue);
                                if (arcComponent != null)
                                {
                                    currentBlock.ComputedCenter = arcComponent.Center;
                                    
                                    // Calculate rotation based on arc position relative to lines
                                    currentBlock.RotationAngle = CalculateSocketRotation(currentBlock.Components, arcComponent);
                                }
                                else
                                {
                                    // Fallback to average if no arc found yet
                                    double avgX = currentBlock.Components.Average(c => c.Center.X);
                                    double avgY = currentBlock.Components.Average(c => c.Center.Y);
                                    currentBlock.ComputedCenter = new XYZ(avgX, avgY, 0);
                                }
                            }
                            else
                            {
                                // For non-sockets: use average of all components
                                double avgX = currentBlock.Components.Average(c => c.Center.X);
                                double avgY = currentBlock.Components.Average(c => c.Center.Y);
                                currentBlock.ComputedCenter = new XYZ(avgX, avgY, 0);
                            }
                        }
                    }
                }
            }
        }
    
        private ComponentInfo GetElementInfo(GeometryObject obj)
        {
            if (obj is Line line)
            {
                return new ComponentInfo
                {
                    Type = "Line",
                    Center = (line.GetEndPoint(0) + line.GetEndPoint(1)) / 2.0
                };
            }
            
            if (obj is Arc arc)
            {
                return new ComponentInfo
                {
                    Type = "Arc",
                    Center = arc.Center,
                    ArcStartAngle = arc.GetEndParameter(0),
                    ArcEndAngle = arc.GetEndParameter(1),
                    ArcRadius = arc.Radius,
                    ArcStartPoint = arc.GetEndPoint(0),
                    ArcEndPoint = arc.GetEndPoint(1)
                };
            }
            
            if (obj is PolyLine pline)
            {
                var points = pline.GetCoordinates();
                if (points.Count == 0) return null;
                double minX = points.Min(p => p.X);
                double minY = points.Min(p => p.Y);
                double maxX = points.Max(p => p.X);
                double maxY = points.Max(p => p.Y);
                return new ComponentInfo
                {
                    Type = "PolyLine",
                    Center = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, 0)
                };
            }
            
            return null;
        }
        
        /// <summary>
        /// Calculate socket rotation angle based on arc orientation
        /// Returns 0, 90, 180, or 270 degrees
        /// Logic: Determine where the arc opens (where Y coordinate is lower for bottom arc)
        /// </summary>
        private double CalculateSocketRotation(List<ComponentInfo> components, ComponentInfo arcComponent)
        {
            if (arcComponent == null || arcComponent.ArcStartPoint == null || arcComponent.ArcEndPoint == null)
                return 0;
            
            XYZ arcCenter = arcComponent.Center;
            XYZ arcStart = arcComponent.ArcStartPoint;
            XYZ arcEnd = arcComponent.ArcEndPoint;
            
            // Get the two endpoints of the arc
            double startY = arcStart.Y;
            double endY = arcEnd.Y;
            double startX = arcStart.X;
            double endX = arcEnd.X;
            double centerY = arcCenter.Y;
            double centerX = arcCenter.X;
            
            // Determine which side the arc opens to
            // For a semicircle, the endpoints have the same Y (or X) coordinate
            // The center is offset in the direction opposite to where it opens
            
            const double tolerance = 0.01; // feet tolerance for coordinate comparison
            
            // Check if arc is horizontal (endpoints have similar Y)
            if (Math.Abs(startY - endY) < tolerance)
            {
                // Horizontal arc
                double endpointsY = (startY + endY) / 2.0;
                
                if (centerY > endpointsY)
                {
                    // Arc opens downward (center above endpoints)
                    return 180;
                }
                else
                {
                    // Arc opens upward (center below endpoints)
                    return 0;
                }
            }
            // Check if arc is vertical (endpoints have similar X)
            else if (Math.Abs(startX - endX) < tolerance)
            {
                // Vertical arc
                double endpointsX = (startX + endX) / 2.0;
                
                if (centerX > endpointsX)
                {
                    // Arc opens left (center to the right of endpoints)
                    return 270;
                }
                else
                {
                    // Arc opens right (center to the left of endpoints)
                    return 90;
                }
            }
            else
            {
                // Diagonal arc - use vector from center to midpoint of endpoints
                XYZ endpointMidpoint = (arcStart + arcEnd) / 2.0;
                XYZ direction = (endpointMidpoint - arcCenter).Normalize();
                
                // Calculate angle
                double angleRadians = Math.Atan2(direction.Y, direction.X);
                double angleDegrees = angleRadians * (180.0 / Math.PI);
                
                // Normalize to 0-360
                if (angleDegrees < 0) angleDegrees += 360;
                
                // Round to nearest 90 degrees
                double roundedAngle = Math.Round(angleDegrees / 90.0) * 90.0;
                if (roundedAngle >= 360) roundedAngle = 0;
                
                return roundedAngle;
            }
        }
    
        private string ExtractBlockNameFromString(string fullString)
        {
            if (string.IsNullOrWhiteSpace(fullString)) return null;
            string lowerStr = fullString.ToLower();
                
            if (lowerStr.Contains(".dwg."))
            {
                int dwgIndex = lowerStr.LastIndexOf(".dwg.");
                return fullString.Substring(dwgIndex + 5).Trim();
            }
            return null; // Ignore containers and non-DWG-block-formatted names
        }
    
        private string GenerateDetailedLog(ImportInstance importInstance, List<DetailedBlockInfo> blocks, List<string> debugLog, string summary, Action<string> statusCallback)
        {
            string timestamp = DateTime.Now.ToString("dd.MM.yy_HH-mm");
            string fileName = $"{timestamp}_NAME.txt";
            string logPath = Path.Combine(LOG_DIRECTORY, fileName);
            string version = GetBuildVersion();
    
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("=======================================================");
            sb.AppendLine($"           {version} - DWG2RVT - Analysis Report");
            sb.AppendLine("=======================================================");
            sb.AppendLine($"Analysis Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Анализируемый документ: {importInstance.Name}");
            sb.AppendLine($"Import ID: {importInstance.Id.Value}");
            sb.AppendLine(summary);
            sb.AppendLine("=======================================================");
            sb.AppendLine();
            // DEBUG INFO
            sb.AppendLine("DEBUG INFO:");
            foreach (var log in debugLog)
            {
                sb.AppendLine($"    [SCAN] {log}");
            }
            sb.AppendLine();
            
            sb.AppendLine("Блоки:");
            sb.AppendLine();
    
            foreach (var block in blocks)
            {
                string rotationInfo = block.RotationAngle > 0 ? $" Поворот: {block.RotationAngle}°" : "";
                sb.AppendLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, 
                    "    {0} №{1} Координаты центра блока: ({2:F4}, {3:F4}){4}", 
                    block.Name, block.Number, block.ComputedCenter.X, block.ComputedCenter.Y, rotationInfo));
                foreach (var comp in block.Components)
                {
                    string compDetails = comp.Type;
                    if (comp.ArcRadius.HasValue)
                    {
                        compDetails += $" [Radius: {comp.ArcRadius.Value:F4}]";
                    }
                    sb.AppendLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "        {0} ({1:F4}, {2:F4})", 
                        compDetails, comp.Center.X, comp.Center.Y));
                }
                sb.AppendLine();
            }
    
            sb.AppendLine("=======================================================");
            sb.AppendLine("End of Report");
            sb.AppendLine("=======================================================");
    
            if (!Directory.Exists(LOG_DIRECTORY)) Directory.CreateDirectory(LOG_DIRECTORY);
            File.WriteAllText(logPath, sb.ToString(), new UTF8Encoding(true));
            return logPath;
        }
    
        private string GenerateEmptyLogFile(ImportInstance importInstance, List<string> debugLog, Action<string> statusCallback)
        {
            // Similar logic but with empty message
            return GenerateDetailedLog(importInstance, new List<DetailedBlockInfo>(), debugLog, "Блоков не найдено", statusCallback);
        }
    
        private string GetBuildVersion()
        {
            try
            {
                string buildNumberPath = @"C:\Users\Свеж как огурец\Desktop\Эксперимент Annotatix\dwg2rvt\BuildNumber.txt";
                if (File.Exists(buildNumberPath))
                {
                    return $"2.{File.ReadAllText(buildNumberPath).Trim()}";
                }
            }
            catch { }
            return "1.0";
        }
    }
}
