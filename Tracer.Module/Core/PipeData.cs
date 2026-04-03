using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;

namespace Tracer.Module.Core
{
    /// <summary>
    /// Data class representing a selected pipe (main line or riser)
    /// </summary>
    public class PipeData
    {
        public ElementId ElementId { get; set; }
        public string Name { get; set; }
        public XYZ StartPoint { get; set; }
        public XYZ EndPoint { get; set; }
        public double Diameter { get; set; }
        public double Slope { get; set; } // Slope in percentage (e.g., 2.0 for 2%)
        public bool IsSloped { get; set; }
        
        // For sloped pipes: direction of flow (from higher to lower)
        public XYZ SlopeDirection { get; set; }
        
        public override string ToString()
        {
            return $"{Name} (ID: {ElementId.IntegerValue}, Dia: {Diameter*304.8:F0}mm)";
        }
    }
    
    /// <summary>
    /// Represents a vertical riser pipe
    /// </summary>
    public class RiserData : PipeData
    {
        public double TopElevation { get; set; }
        public double BottomElevation { get; set; }
        public XYZ ConnectionPoint { get; set; } // Point where riser should connect to main
    }
    
    /// <summary>
    /// Represents a sloped main sewage line
    /// </summary>
    public class MainLineData : PipeData
    {
        public double StartElevation { get; set; }
        public double EndElevation { get; set; }
        
        /// <summary>
        /// Gets the elevation at a specific point along the main line
        /// </summary>
        public double GetElevationAtPoint(XYZ point)
        {
            // Project point onto main line
            XYZ lineDirection = (EndPoint - StartPoint).Normalize();
            XYZ toPoint = point - StartPoint;
            double distanceAlongLine = toPoint.DotProduct(lineDirection);
            double totalLength = (EndPoint - StartPoint).GetLength();
            
            if (totalLength < 0.001) return StartElevation;
            
            double ratio = distanceAlongLine / totalLength;
            return StartElevation + (EndElevation - StartElevation) * ratio;
        }
    }
}
