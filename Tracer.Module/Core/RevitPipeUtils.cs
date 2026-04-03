using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using PluginsManager.Core;

namespace Tracer.Module.Core
{
    /// <summary>
    /// Utility class for working with pipes in Revit
    /// </summary>
    public static class RevitPipeUtils
    {
        /// <summary>
        /// Gets pipe data from a selected pipe element
        /// </summary>
        public static PipeData GetPipeData(Document doc, ElementId pipeId)
        {
            try
            {
                var pipe = doc.GetElement(pipeId) as Pipe;
                if (pipe == null) return null;
                
                // Get pipe geometry
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                if (locationCurve == null) return null;
                
                Curve curve = locationCurve.Curve;
                XYZ startPoint = curve.GetEndPoint(0);
                XYZ endPoint = curve.GetEndPoint(1);
                
                // Get pipe diameter
                double diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() ?? 0.1;
                
                // Check if pipe is sloped
                double slope = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE)?.AsDouble() ?? 0;
                bool isSloped = Math.Abs(slope) > 0.001;
                
                return new PipeData
                {
                    ElementId = pipeId,
                    Name = pipe.Name,
                    StartPoint = startPoint,
                    EndPoint = endPoint,
                    Diameter = diameter,
                    Slope = slope * 100, // Convert to percentage
                    IsSloped = isSloped,
                    SlopeDirection = isSloped ? (endPoint - startPoint).Normalize() : XYZ.BasisZ
                };
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-UTILS] ERROR getting pipe data: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Gets main line data with elevation information
        /// </summary>
        public static MainLineData GetMainLineData(Document doc, ElementId pipeId)
        {
            var pipeData = GetPipeData(doc, pipeId);
            if (pipeData == null) return null;
            
            return new MainLineData
            {
                ElementId = pipeData.ElementId,
                Name = pipeData.Name,
                StartPoint = pipeData.StartPoint,
                EndPoint = pipeData.EndPoint,
                Diameter = pipeData.Diameter,
                Slope = pipeData.Slope,
                IsSloped = pipeData.IsSloped,
                SlopeDirection = pipeData.SlopeDirection,
                StartElevation = pipeData.StartPoint.Z,
                EndElevation = pipeData.EndPoint.Z
            };
        }
        
        /// <summary>
        /// Gets riser data from a vertical pipe
        /// </summary>
        public static RiserData GetRiserData(Document doc, ElementId pipeId)
        {
            var pipeData = GetPipeData(doc, pipeId);
            if (pipeData == null) return null;
            
            // Check if pipe is approximately vertical
            XYZ direction = (pipeData.EndPoint - pipeData.StartPoint).Normalize();
            bool isVertical = Math.Abs(direction.Z) > 0.95;
            
            if (!isVertical)
            {
                DebugLogger.Log("[TRACER-UTILS] WARNING: Selected pipe is not vertical, may not be a riser");
            }
            
            double topZ = Math.Max(pipeData.StartPoint.Z, pipeData.EndPoint.Z);
            double bottomZ = Math.Min(pipeData.StartPoint.Z, pipeData.EndPoint.Z);
            
            return new RiserData
            {
                ElementId = pipeData.ElementId,
                Name = pipeData.Name,
                StartPoint = pipeData.StartPoint,
                EndPoint = pipeData.EndPoint,
                Diameter = pipeData.Diameter,
                Slope = 0,
                IsSloped = false,
                SlopeDirection = XYZ.BasisZ,
                TopElevation = topZ,
                BottomElevation = bottomZ,
                ConnectionPoint = new XYZ(pipeData.StartPoint.X, pipeData.StartPoint.Y, (topZ + bottomZ) / 2)
            };
        }
        
        /// <summary>
        /// Creates a pipe connection between main line and riser at 45° angle
        /// </summary>
        public static bool CreateConnection(
            Document doc,
            MainLineData mainLine,
            RiserData riser,
            XYZ connectionPoint,
            XYZ riserConnectionPoint,
            double pipeDiameter)
        {
            try
            {
                DebugLogger.Log("[TRACER-UTILS] Creating pipe connection...");
                
                // Get pipe type and system type from main line
                var mainPipe = doc.GetElement(mainLine.ElementId) as Pipe;
                if (mainPipe == null)
                {
                    DebugLogger.Log("[TRACER-UTILS] ERROR: Cannot get main pipe element");
                    return false;
                }
                
                ElementId pipeTypeId = mainPipe.GetTypeId();
                ElementId systemTypeId = mainPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsElementId() ?? ElementId.InvalidElementId;
                
                // Get level for pipe creation
                Level level = doc.GetElement(mainPipe.LevelId) as Level;
                if (level == null)
                {
                    // Try to find level by elevation
                    level = FindLevelByElevation(doc, connectionPoint.Z);
                }
                
                if (level == null)
                {
                    DebugLogger.Log("[TRACER-UTILS] ERROR: Cannot determine level for pipe creation");
                    return false;
                }
                
                // Create the connecting pipe
                using (Transaction trans = new Transaction(doc, "Create Tracer Connection"))
                {
                    trans.Start();
                    
                    // Create pipe from main line connection point to riser connection point
                    Pipe connectingPipe = Pipe.Create(
                        doc,
                        systemTypeId,
                        pipeTypeId,
                        level.Id,
                        connectionPoint,
                        riserConnectionPoint
                    );
                    
                    if (connectingPipe == null)
                    {
                        DebugLogger.Log("[TRACER-UTILS] ERROR: Failed to create connecting pipe");
                        trans.RollBack();
                        return false;
                    }
                    
                    // Set pipe diameter
                    connectingPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(pipeDiameter);
                    
                    // Apply slope to connecting pipe (same as main line)
                    // This ensures proper drainage
                    double mainLineSlope = mainLine.Slope / 100.0; // Convert from percentage
                    connectingPipe.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).Set(mainLineSlope);
                    
                    DebugLogger.Log("[TRACER-UTILS] Connecting pipe created successfully");
                    
                    // TODO: Create 45° tee fitting at connection point
                    // This requires more complex fitting creation logic
                    
                    trans.Commit();
                }
                
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-UTILS] ERROR creating connection: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Finds a level by elevation
        /// </summary>
        private static Level FindLevelByElevation(Document doc, double elevation)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - elevation))
                .ToList();
            
            return levels.FirstOrDefault();
        }
        
        /// <summary>
        /// Gets available pipe types in the document
        /// </summary>
        public static List<PipeType> GetPipeTypes(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(PipeType))
                .Cast<PipeType>()
                .ToList();
        }
    }
}
