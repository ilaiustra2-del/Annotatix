using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using PluginsManager.Core;

namespace AutoNumbering.Module.Core
{
    /// <summary>
    /// Analyzer for finding and grouping vertical pipes (risers)
    /// </summary>
    public class RiserAnalyzer
    {
        private readonly Document _doc;
        private readonly View _activeView;
        private readonly double _groupDistanceFeet;
        private const double TOLERANCE = 0.001; // ~0.3mm tolerance for coordinate comparison
        
        // Analysis statistics
        public int TotalPipesAnalyzed { get; private set; }
        public int TotalRisersFound { get; private set; }

        public RiserAnalyzer(Document doc, View activeView, double groupDistanceMm = 400.0)
        {
            _doc = doc;
            _activeView = activeView;
            _groupDistanceFeet = groupDistanceMm / 304.8; // Convert mm to feet (Revit units)
            
            DebugLogger.Log($"[RISER-ANALYZER] Initialized with grouping distance: {groupDistanceMm} mm ({_groupDistanceFeet:F4} feet)");
        }

        /// <summary>
        /// Find all vertical pipes (risers) in the active view
        /// </summary>
        public List<RiserInfo> FindRisers()
        {
            var risers = new List<RiserInfo>();

            try
            {
                // Get all pipes visible in the active view
                var collector = new FilteredElementCollector(_doc, _activeView.Id)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType();

                int totalPipes = collector.Count();
                TotalPipesAnalyzed = totalPipes;
                DebugLogger.Log($"[RISER-ANALYZER] Total pipes in view: {totalPipes}");

                foreach (Pipe pipe in collector)
                {
                    if (IsVerticalRiser(pipe))
                    {
                        var riserInfo = CreateRiserInfo(pipe);
                        risers.Add(riserInfo);
                        DebugLogger.Log($"[RISER-ANALYZER] ✓ Found riser: {pipe.Id.IntegerValue}, X={riserInfo.X:F3}, Z={riserInfo.Z:F3}, System={riserInfo.SystemType}");
                    }
                }

                TotalRisersFound = risers.Count;
                DebugLogger.Log($"[RISER-ANALYZER] Total risers found: {risers.Count} out of {totalPipes} pipes");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RISER-ANALYZER] ERROR finding risers: {ex.Message}");
            }

            return risers;
        }
        
        /// <summary>
        /// Get unique system types from found risers
        /// </summary>
        public List<string> GetUniqueSystemTypes(List<RiserInfo> risers)
        {
            var systemTypes = risers
                .Select(r => r.SystemType)
                .Where(s => !string.IsNullOrEmpty(s))
                .Distinct()
                .OrderBy(s => s)
                .ToList();
                
            DebugLogger.Log($"[RISER-ANALYZER] Found {systemTypes.Count} unique system types");
            foreach (var sysType in systemTypes)
            {
                int count = risers.Count(r => r.SystemType == sysType);
                DebugLogger.Log($"[RISER-ANALYZER]   - {sysType}: {count} risers");
            }
            
            return systemTypes;
        }

        /// <summary>
        /// Group risers by proximity and assign numbers
        /// Uses expanding cluster algorithm: if riser is close to ANY member of a group, add it to that group
        /// Also checks system compatibility groups
        /// 
        /// NUMBERING LOGIC:
        /// - If compatibility groups defined:
        ///   * Systems in group: numbering starts from 1 per compatibility group
        ///   * Systems not in any group: separate numbering from 1 per system type
        /// - If no compatibility groups: all risers numbered together from 1
        /// </summary>
        public List<RiserGroup> GroupAndNumberRisers(List<RiserInfo> risers, List<SystemCompatibilityGroup> compatibilityGroups = null)
        {
            var groups = new List<RiserGroup>();
            var processed = new HashSet<ElementId>();

            // Sort by X, then Y for consistent ordering (horizontal plane)
            var sortedRisers = risers.OrderBy(r => r.X).ThenBy(r => r.Y).ToList();
            
            // If compatibility groups defined, use per-compatibility-group or per-system-type numbering
            bool useCompatibilityGrouping = compatibilityGroups != null && compatibilityGroups.Count > 0;
            
            if (useCompatibilityGrouping)
            {
                DebugLogger.Log($"[RISER-ANALYZER] Using compatibility group numbering ({compatibilityGroups.Count} groups defined)");
                
                // Group by compatibility group membership
                var groupedByCompatibility = new Dictionary<string, List<RiserInfo>>();
                
                foreach (var riser in sortedRisers)
                {
                    // Find which compatibility group this system belongs to (if any)
                    string compatibilityKey = GetCompatibilityGroupKey(riser.SystemType, compatibilityGroups);
                    
                    if (!groupedByCompatibility.ContainsKey(compatibilityKey))
                    {
                        groupedByCompatibility[compatibilityKey] = new List<RiserInfo>();
                    }
                    groupedByCompatibility[compatibilityKey].Add(riser);
                }
                
                // Process each compatibility group separately
                foreach (var kvp in groupedByCompatibility.OrderBy(x => x.Key))
                {
                    string compatKey = kvp.Key;
                    var compatRisers = kvp.Value;
                    
                    DebugLogger.Log($"[RISER-ANALYZER] Processing compatibility group '{compatKey}' with {compatRisers.Count} risers");
                    
                    // Create proximity groups within this compatibility group
                    var compatGroups = CreateProximityGroups(compatRisers, compatibilityGroups, processed);
                    
                    // IMPORTANT: Number groups starting from 1 for THIS compatibility group
                    int localGroupNumber = 1;
                    foreach (var group in compatGroups)
                    {
                        group.GroupNumber = localGroupNumber++;
                        DebugLogger.Log($"[RISER-ANALYZER]   Assigned number {group.GroupNumber} to group with {group.Risers.Count} risers (compat: {compatKey})");
                    }
                    
                    groups.AddRange(compatGroups);
                }
            }
            else
            {
                // No compatibility groups - use original logic (all risers together)
                DebugLogger.Log($"[RISER-ANALYZER] Using simple proximity grouping (no compatibility groups)");
                var simpleGroups = CreateProximityGroups(sortedRisers, null, processed);
                
                // Number groups starting from 1
                int groupNumber = 1;
                foreach (var group in simpleGroups)
                {
                    group.GroupNumber = groupNumber++;
                }
                
                groups.AddRange(simpleGroups);
            }

            DebugLogger.Log($"[RISER-ANALYZER] Total groups created: {groups.Count}");
            return groups;
        }
        
        /// <summary>
        /// Get compatibility group key for a system type
        /// Returns "COMPAT_GROUP_X" for systems in compatibility groups,
        /// or "SYSTEM_TypeName" for systems not in any group
        /// </summary>
        private string GetCompatibilityGroupKey(string systemType, List<SystemCompatibilityGroup> compatibilityGroups)
        {
            for (int i = 0; i < compatibilityGroups.Count; i++)
            {
                if (compatibilityGroups[i].SystemTypes.Contains(systemType))
                {
                    return $"COMPAT_GROUP_{i + 1}";
                }
            }
            
            // Not in any compatibility group - use system type as key
            return $"SYSTEM_{systemType}";
        }
        
        /// <summary>
        /// Create proximity groups from list of risers using expanding cluster algorithm
        /// </summary>
        private List<RiserGroup> CreateProximityGroups(List<RiserInfo> risers, List<SystemCompatibilityGroup> compatibilityGroups, HashSet<ElementId> processed)
        {
            var groups = new List<RiserGroup>();
            int localGroupNumber = 1;
            
            // IMPORTANT: Sort by X, Y for consistent proximity checking
            var sortedRisers = risers.OrderBy(r => r.X).ThenBy(r => r.Y).ToList();
            
            foreach (var riser in sortedRisers)
            {
                if (processed.Contains(riser.PipeId))
                    continue;

                // Create new group starting with this riser
                var group = new RiserGroup
                {
                    GroupNumber = localGroupNumber++, // Temporary number, will be reassigned
                    Risers = new List<RiserInfo> { riser }
                };

                processed.Add(riser.PipeId);
                DebugLogger.Log($"[RISER-ANALYZER]   New proximity group started with riser {riser.PipeId.IntegerValue} (system: {riser.SystemType})");

                // EXPANDING CLUSTER: Keep checking until no more risers can be added
                bool addedNewRisers = true;
                while (addedNewRisers)
                {
                    addedNewRisers = false;

                    foreach (var otherRiser in risers)
                    {
                        if (processed.Contains(otherRiser.PipeId))
                            continue;

                        // Check if this riser is close to ANY member of the current group
                        bool shouldAddToGroup = false;
                        double minDistance = double.MaxValue;
                        ElementId closestMemberId = null;

                        foreach (var groupMember in group.Risers)
                        {
                            // FIRST: Check system compatibility (if groups defined)
                            if (compatibilityGroups != null && compatibilityGroups.Count > 0)
                            {
                                bool systemsCompatible = AreSystemsCompatible(
                                    groupMember.SystemType, 
                                    otherRiser.SystemType, 
                                    compatibilityGroups);
                                
                                if (!systemsCompatible)
                                {
                                    continue; // Skip if not compatible
                                }
                            }
                            
                            // SECOND: Check if at same XY location (split risers on same vertical line)
                            bool sameLocation = Math.Abs(groupMember.X - otherRiser.X) < TOLERANCE 
                                             && Math.Abs(groupMember.Y - otherRiser.Y) < TOLERANCE;

                            // THIRD: Check horizontal distance in XY plane
                            double distance = Math.Sqrt(
                                Math.Pow(groupMember.X - otherRiser.X, 2) + 
                                Math.Pow(groupMember.Y - otherRiser.Y, 2));

                            if (distance < minDistance)
                            {
                                minDistance = distance;
                                closestMemberId = groupMember.PipeId;
                            }

                            if (sameLocation || distance <= _groupDistanceFeet)
                            {
                                shouldAddToGroup = true;
                                break; // Found at least one close member
                            }
                        }

                        if (shouldAddToGroup)
                        {
                            group.Risers.Add(otherRiser);
                            processed.Add(otherRiser.PipeId);
                            addedNewRisers = true; // Continue expanding
                            DebugLogger.Log($"[RISER-ANALYZER]     → Added riser {otherRiser.PipeId.IntegerValue} (system: {otherRiser.SystemType}), distance={minDistance * 304.8:F1}mm");
                        }
                    }
                }

                groups.Add(group);
                DebugLogger.Log($"[RISER-ANALYZER]   Proximity group complete: {group.Risers.Count} risers");
            }
            
            return groups;
        }
        
        /// <summary>
        /// Check if two system types are compatible (can be grouped together)
        /// </summary>
        private bool AreSystemsCompatible(string systemType1, string systemType2, List<SystemCompatibilityGroup> compatibilityGroups)
        {
            // If no compatibility groups defined, all systems are compatible
            if (compatibilityGroups == null || compatibilityGroups.Count == 0)
                return true;
            
            // Empty systems are always compatible with each other
            if (string.IsNullOrEmpty(systemType1) && string.IsNullOrEmpty(systemType2))
                return true;
                
            // Find if both systems are in same compatibility group
            foreach (var group in compatibilityGroups)
            {
                bool has1 = group.SystemTypes.Contains(systemType1);
                bool has2 = group.SystemTypes.Contains(systemType2);
                
                if (has1 && has2)
                {
                    // Both systems in same compatibility group
                    return true;
                }
            }
            
            // Systems not in same compatibility group
            return false;
        }

        /// <summary>
        /// Check if pipe is vertical (riser)
        /// In Revit: X and Y are horizontal, Z is vertical (height)
        /// </summary>
        private bool IsVerticalRiser(Pipe pipe)
        {
            try
            {
                var locationCurve = pipe.Location as LocationCurve;
                if (locationCurve == null)
                {
                    DebugLogger.Log($"[RISER-ANALYZER]   Pipe {pipe.Id.IntegerValue}: No LocationCurve");
                    return false;
                }

                var curve = locationCurve.Curve;
                var startPoint = curve.GetEndPoint(0);
                var endPoint = curve.GetEndPoint(1);

                // CORRECTED: Check if X and Y coordinates are the same (only Z differs)
                // Z is vertical coordinate in Revit!
                bool sameX = Math.Abs(startPoint.X - endPoint.X) < TOLERANCE;
                bool sameY = Math.Abs(startPoint.Y - endPoint.Y) < TOLERANCE;
                bool diffZ = Math.Abs(startPoint.Z - endPoint.Z) > TOLERANCE;

                bool isVertical = sameX && sameY && diffZ;
                
                if (!isVertical)
                {
                    // Log why pipe is not vertical
                    double deltaX = Math.Abs(startPoint.X - endPoint.X);
                    double deltaY = Math.Abs(startPoint.Y - endPoint.Y);
                    double deltaZ = Math.Abs(startPoint.Z - endPoint.Z);
                    DebugLogger.Log($"[RISER-ANALYZER]   Pipe {pipe.Id.IntegerValue}: NOT vertical - ΔX={deltaX:F4}, ΔY={deltaY:F4}, ΔZ={deltaZ:F4} (need ΔX<{TOLERANCE} AND ΔY<{TOLERANCE} AND ΔZ>{TOLERANCE})");
                }

                return isVertical;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RISER-ANALYZER] ERROR checking pipe {pipe.Id.IntegerValue}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Create RiserInfo from pipe
        /// </summary>
        private RiserInfo CreateRiserInfo(Pipe pipe)
        {
            var locationCurve = pipe.Location as LocationCurve;
            var curve = locationCurve.Curve;
            var startPoint = curve.GetEndPoint(0);
            
            // Get system type parameter
            // Need SYSTEM TYPE, not system name (e.g., "ADSK_Канализация_К1", not "К1 1")
            string systemType = "";
            
            try
            {
                // Try to get SystemType from pipe's MEP System
                if (pipe.MEPSystem != null)
                {
                    var mepSystemType = pipe.MEPSystem.GetTypeId();
                    if (mepSystemType != null && mepSystemType != ElementId.InvalidElementId)
                    {
                        var systemTypeElement = pipe.Document.GetElement(mepSystemType);
                        if (systemTypeElement != null)
                        {
                            systemType = systemTypeElement.Name ?? "";
                        }
                    }
                }
                
                DebugLogger.Log($"[RISER-ANALYZER] Pipe {pipe.Id.IntegerValue}: System type = '{systemType}'");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RISER-ANALYZER] ERROR getting system type for pipe {pipe.Id.IntegerValue}: {ex.Message}");
            }

            return new RiserInfo
            {
                PipeId = pipe.Id,
                Pipe = pipe,
                X = startPoint.X,
                Y = startPoint.Y,
                Z = startPoint.Z,
                SystemType = systemType
            };
        }
    }

    /// <summary>
    /// Information about a single riser pipe
    /// </summary>
    public class RiserInfo
    {
        public ElementId PipeId { get; set; }
        public Pipe Pipe { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public string SystemType { get; set; } // Тип системы
    }

    /// <summary>
    /// Group of risers that should have the same number
    /// </summary>
    public class RiserGroup
    {
        public int GroupNumber { get; set; }
        public List<RiserInfo> Risers { get; set; } = new List<RiserInfo>();
    }
    
    /// <summary>
    /// System compatibility group - defines which system types can be grouped together
    /// </summary>
    public class SystemCompatibilityGroup
    {
        public string GroupName { get; set; }
        public List<string> SystemTypes { get; set; } = new List<string>();
    }
}
