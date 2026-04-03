using System;
using Autodesk.Revit.DB;
using PluginsManager.Core;

namespace Tracer.Module.Core
{
    /// <summary>
    /// Calculates connection geometry for sewage pipe connection
    /// Algorithm based on scheme:
    /// 1. Find perpendicular from riser to main line (point x4,y4,z4)
    /// 2. From perpendicular point, move along main line by distance = riser-to-perpendicular distance 
    ///    to get connection point (x5,y5,z5)
    /// 3. Build pipe from (x5,y5,z5) to (x6=x3, y6=y3, z6 calculated by slope)
    /// 4. The pipe makes 45° angle with main line direction
    /// </summary>
    public static class ConnectionCalculator
    {
        /// <summary>
        /// Calculates connection points according to the scheme
        /// </summary>
        /// <param name="mainLine">Main sloped sewage line</param>
        /// <param name="riser">Vertical riser to connect</param>
        /// <returns>Tuple of (connectionPointOnMain, endPointAtRiser) or null if calculation fails</returns>
        public static (XYZ connectionPoint, XYZ endPoint) CalculateConnectionPoints(
            MainLineData mainLine, 
            RiserData riser)
        {
            try
            {
                // Step 1: Get riser axis position (x3, y3)
                XYZ riserAxis = new XYZ(riser.StartPoint.X, riser.StartPoint.Y, 0);
                
                // Step 2: Find perpendicular projection of riser onto main line (point x4,y4 in scheme)
                XYZ mainStartXY = new XYZ(mainLine.StartPoint.X, mainLine.StartPoint.Y, 0);
                XYZ mainEndXY = new XYZ(mainLine.EndPoint.X, mainLine.EndPoint.Y, 0);
                XYZ mainDirXY = (mainEndXY - mainStartXY).Normalize();
                
                // Project riser onto main line
                XYZ toRiser = riserAxis - mainStartXY;
                double projLength = toRiser.DotProduct(mainDirXY);
                XYZ perpendicularPoint = mainStartXY + mainDirXY * projLength;
                
                // Check if perpendicular point is within main line bounds
                double mainLength = (mainEndXY - mainStartXY).GetLength();
                if (projLength < 0.001 || projLength > mainLength - 0.001)
                {
                    DebugLogger.Log("[TRACER-CALC] ERROR: Perpendicular point is outside main line bounds");
                    return (null, null);
                }
                
                // Step 3: Calculate distance from riser to perpendicular point
                XYZ riserToPerp = perpendicularPoint - riserAxis;
                double distanceToRiser = riserToPerp.GetLength();
                
                if (distanceToRiser < 0.001)
                {
                    DebugLogger.Log("[TRACER-CALC] ERROR: Riser is too close to main line");
                    return (null, null);
                }
                
                // Step 4: Determine flow direction
                bool flowFromStartToEnd = mainLine.StartElevation > mainLine.EndElevation;
                XYZ flowDir = flowFromStartToEnd ? mainDirXY : -mainDirXY;
                
                // Step 5: Calculate connection point on main line (x5,y5,z5)
                // Move from perpendicular point ALONG flow direction (downhill) by distance = distanceToRiser
                // This creates 45° angle between pipe and main line
                XYZ connectionPointXY = perpendicularPoint + flowDir * distanceToRiser;
                double connectionElevation = mainLine.GetElevationAtPoint(connectionPointXY);
                XYZ connectionPoint = new XYZ(connectionPointXY.X, connectionPointXY.Y, connectionElevation);
                
                // Step 7: Calculate end point at riser (x6,y6,z6)
                // x6 = x3, y6 = y3 (riser axis)
                // z6 = z5 + distance * slope (HIGHER than connection point - pipe goes UP from main to riser)
                double slope = mainLine.Slope / 100.0;
                double pipeLengthXY = (connectionPointXY - riserAxis).GetLength();
                double elevationRise = pipeLengthXY * slope;
                double endElevation = connectionElevation + elevationRise;
                XYZ endPoint = new XYZ(riserAxis.X, riserAxis.Y, endElevation);
                
                DebugLogger.Log($"[TRACER-CALC] Connection calculated:");
                DebugLogger.Log($"[TRACER-CALC]   Perpendicular point: ({perpendicularPoint.X:F3}, {perpendicularPoint.Y:F3})");
                DebugLogger.Log($"[TRACER-CALC]   Distance to riser: {distanceToRiser:F3}");
                DebugLogger.Log($"[TRACER-CALC]   Connection point (main): ({connectionPoint.X:F3}, {connectionPoint.Y:F3}, {connectionPoint.Z:F3})");
                DebugLogger.Log($"[TRACER-CALC]   End point (riser): ({endPoint.X:F3}, {endPoint.Y:F3}, {endPoint.Z:F3})");
                DebugLogger.Log($"[TRACER-CALC]   Pipe XY length: {pipeLengthXY:F3}");
                DebugLogger.Log($"[TRACER-CALC]   Elevation rise: {elevationRise:F3}");
                DebugLogger.Log($"[TRACER-CALC]   Slope: {mainLine.Slope:F2}%");
                
                return (connectionPoint, endPoint);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TRACER-CALC] ERROR calculating connection: {ex.Message}");
                return (null, null);
            }
        }
        
        /// <summary>
        /// Validates that the calculated connection is feasible
        /// </summary>
        public static bool ValidateConnection(
            MainLineData mainLine, 
            XYZ connectionPoint, 
            XYZ endPoint)
        {
            if (connectionPoint == null || endPoint == null)
                return false;
                
            // Check that pipe has positive length
            XYZ pipeVector = connectionPoint - endPoint;
            if (pipeVector.GetLength() < 0.001)
            {
                DebugLogger.Log("[TRACER-CALC] Validation failed: pipe has zero length");
                return false;
            }
            
            // Check that end point is below connection point (flow goes down)
            if (endPoint.Z >= connectionPoint.Z)
            {
                DebugLogger.Log("[TRACER-CALC] WARNING: End point should be below connection point for gravity flow");
            }
            
            DebugLogger.Log("[TRACER-CALC] Connection validation passed");
            return true;
        }
    }
}
