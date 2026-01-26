using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace dwg2rvt.Core
{
    /// <summary>
    /// Centralized in-memory cache for storing DWG analysis results
    /// </summary>
    public static class AnalysisDataCache
    {
        private static AnalysisResult _lastAnalysisResult = null;
        private static readonly object _lock = new object();

        /// <summary>
        /// Store analysis result in memory
        /// </summary>
        public static void StoreAnalysisResult(AnalysisResult result)
        {
            lock (_lock)
            {
                _lastAnalysisResult = result;
            }
        }

        /// <summary>
        /// Retrieve the last analysis result from memory
        /// </summary>
        public static AnalysisResult GetLastAnalysisResult()
        {
            lock (_lock)
            {
                return _lastAnalysisResult;
            }
        }

        /// <summary>
        /// Check if analysis data is available in cache
        /// </summary>
        public static bool HasAnalysisData()
        {
            lock (_lock)
            {
                return _lastAnalysisResult != null && 
                       _lastAnalysisResult.Success && 
                       _lastAnalysisResult.BlockData != null && 
                       _lastAnalysisResult.BlockData.Count > 0;
            }
        }

        /// <summary>
        /// Clear all cached data
        /// </summary>
        public static void Clear()
        {
            lock (_lock)
            {
                _lastAnalysisResult = null;
            }
        }
    }

    /// <summary>
    /// Data structure for storing individual block information
    /// </summary>
    public class BlockData
    {
        public string Name { get; set; }
        public int Number { get; set; }
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public double RotationAngle { get; set; } = 0;
        public List<ComponentData> Components { get; set; } = new List<ComponentData>();
    }

    /// <summary>
    /// Data structure for storing component information (Lines, Arcs, etc.)
    /// </summary>
    public class ComponentData
    {
        public string Type { get; set; }
        public double CenterX { get; set; }
        public double CenterY { get; set; }
        public double? ArcRadius { get; set; }
        public double? ArcStartAngle { get; set; }
        public double? ArcEndAngle { get; set; }
    }
}
