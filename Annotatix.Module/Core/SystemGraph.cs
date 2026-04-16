using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Represents a node in the system graph (an element like duct, pipe, fitting, etc.)
    /// </summary>
    public class SystemGraphNode
    {
        public long ElementId { get; set; }
        public string Category { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        
        /// <summary>Start point in view coordinates</summary>
        public ViewCoordinates2D ViewStart { get; set; }
        
        /// <summary>End point in view coordinates</summary>
        public ViewCoordinates2D ViewEnd { get; set; }
        
        /// <summary>Start point in model coordinates</summary>
        public ModelCoordinates3D ModelStart { get; set; }
        
        /// <summary>End point in model coordinates</summary>
        public ModelCoordinates3D ModelEnd { get; set; }
        
        /// <summary>True if element has two endpoints (linear element)</summary>
        public bool HasEndPoint { get; set; }
        
        /// <summary>Diameter for round ducts/pipes</summary>
        public double? Diameter { get; set; }
        
        /// <summary>Width for rectangular ducts</summary>
        public double? Width { get; set; }
        
        /// <summary>Height for rectangular ducts</summary>
        public double? Height { get; set; }
        
        /// <summary>Size display string</summary>
        public string SizeDisplay { get; set; }
        
        /// <summary>System ID this element belongs to</summary>
        public long? SystemId { get; set; }
        
        /// <summary>System name</summary>
        public string SystemName { get; set; }
        
        /// <summary>Connected nodes (by ElementId)</summary>
        public List<long> ConnectedNodeIds { get; set; } = new List<long>();
        
        /// <summary>Degree in graph (number of connections)</summary>
        public int Degree => ConnectedNodeIds.Count;
        
        /// <summary>True if this is a junction (degree > 2)</summary>
        public bool IsJunction => Degree > 2;
        
        /// <summary>True if this is an endpoint (degree == 1)</summary>
        public bool IsEndpoint => Degree == 1;
        
        /// <summary>True if this is a through element (degree == 2)</summary>
        public bool IsThrough => Degree == 2;
        
        /// <summary>True if this is a duct or pipe (main element to annotate)</summary>
        public bool IsLinearElement => Category == "Воздуховоды" || Category == "Трубы" || 
                                       Category == "Ducts" || Category == "Pipes";
        
        /// <summary>True if this is a fitting (transition, tee, etc.)</summary>
        public bool IsFitting => Category == "Соединительные детали воздуховодов" || 
                                 Category == "Соединительные детали трубопроводов" ||
                                 Category == "Duct Fittings" || Category == "Pipe Fittings";
        
        /// <summary>True if this is equipment</summary>
        public bool IsEquipment => Category == "Оборудование" || Category == "Mechanical Equipment";
        
        /// <summary>True if this is air terminal (diffuser, grille, etc.)</summary>
        public bool IsAirTerminal => Category == "Воздухораспределители" || Category == "Air Terminals";
        
        /// <summary>True if this is duct accessory (damper, fire damper, etc.)</summary>
        public bool IsAccessory => Category == "Арматура воздуховодов" || Category == "Duct Accessories";
    }
    
    /// <summary>
    /// Represents an edge in the system graph (connection between two elements)
    /// </summary>
    public class SystemGraphEdge
    {
        public long FromNodeId { get; set; }
        public long ToNodeId { get; set; }
        
        /// <summary>Connection point in view coordinates</summary>
        public ViewCoordinates2D ConnectionPoint { get; set; }
        
        /// <summary>Connection point in model coordinates</summary>
        public ModelCoordinates3D ModelConnectionPoint { get; set; }
    }
    
    /// <summary>
    /// Represents a system (ductwork or piping system) with its graph
    /// </summary>
    public class SystemGraph
    {
        public long SystemId { get; set; }
        public string SystemName { get; set; }
        public string SystemType { get; set; }
        
        /// <summary>All nodes in the system graph</summary>
        public Dictionary<long, SystemGraphNode> Nodes { get; set; } = new Dictionary<long, SystemGraphNode>();
        
        /// <summary>All edges in the system graph</summary>
        public List<SystemGraphEdge> Edges { get; set; } = new List<SystemGraphEdge>();
        
        /// <summary>Main trunks (paths from equipment to main branches)</summary>
        public List<List<long>> MainTrunks { get; set; } = new List<List<long>>();
        
        /// <summary>Branches (paths that split from main trunk)</summary>
        public List<List<long>> Branches { get; set; } = new List<List<long>>();
        
        /// <summary>Junction nodes (where branches split)</summary>
        public List<SystemGraphNode> Junctions => Nodes.Values.Where(n => n.IsJunction).ToList();
        
        /// <summary>Endpoint nodes (equipment, terminals)</summary>
        public List<SystemGraphNode> Endpoints => Nodes.Values.Where(n => n.IsEndpoint).ToList();
        
        /// <summary>Linear elements (ducts, pipes) only</summary>
        public List<SystemGraphNode> LinearElements => Nodes.Values.Where(n => n.IsLinearElement).ToList();
        
        /// <summary>Fittings only</summary>
        public List<SystemGraphNode> Fittings => Nodes.Values.Where(n => n.IsFitting).ToList();
    }
    
    /// <summary>
    /// Builds system graphs from collected view data
    /// </summary>
    public class SystemGraphBuilder
    {
        // Tolerance for connector matching in view units (increased for reliability)
        private const double CONNECTOR_TOLERANCE = 0.5; // 0.5 view units (~150mm at 1:100)
        
        // Tolerance for model coordinate matching (in feet)
        private const double MODEL_TOLERANCE = 0.15; // ~45mm in model units
        
        /// <summary>
        /// Build system graphs from view snapshot
        /// </summary>
        public List<SystemGraph> BuildFromSnapshot(ViewSnapshot snapshot)
        {
            var systems = new List<SystemGraph>();
            
            // Group elements by system
            var elementsBySystem = snapshot.Elements
                .Where(e => e.SystemId.HasValue)
                .GroupBy(e => e.SystemId.Value)
                .ToDictionary(g => g.Key, g => g.ToList());
            
            // Also collect elements without system ID that should be annotated
            // (equipment, accessories, air terminals)
            var elementsWithoutSystem = snapshot.Elements
                .Where(e => !e.SystemId.HasValue)
                .Where(e => IsAnnotatableElement(e))
                .ToList();
            
            foreach (var systemGroup in elementsBySystem)
            {
                var systemId = systemGroup.Key;
                var elements = systemGroup.Value;
                
                // Get system info
                var systemInfo = snapshot.Systems?.FirstOrDefault(s => s.SystemId == systemId);
                
                var graph = new SystemGraph
                {
                    SystemId = systemId,
                    SystemName = systemInfo?.SystemName ?? "Unknown",
                    SystemType = systemInfo?.SystemType ?? "Unknown"
                };
                
                // Create nodes
                foreach (var elem in elements)
                {
                    var node = CreateNode(elem);
                    graph.Nodes[node.ElementId] = node;
                }
                
                // Build connections by finding matching endpoints
                BuildConnections(graph, elements);
                
                // Identify main trunks and branches
                IdentifyTrunksAndBranches(graph);
                
                systems.Add(graph);
            }
            
            // Create a special graph for elements without system (orphaned equipment, etc.)
            if (elementsWithoutSystem.Count > 0)
            {
                var orphanGraph = new SystemGraph
                {
                    SystemId = -1,
                    SystemName = "Orphaned Elements",
                    SystemType = "Unknown"
                };
                
                foreach (var elem in elementsWithoutSystem)
                {
                    var node = CreateNode(elem);
                    orphanGraph.Nodes[node.ElementId] = node;
                }
                
                // Try to build connections among orphaned elements
                BuildConnections(orphanGraph, elementsWithoutSystem);
                
                systems.Add(orphanGraph);
                
                DebugLogger.Log($"[SYSTEM-GRAPH] Created orphan graph with {elementsWithoutSystem.Count} elements");
            }
            
            return systems;
        }
        
        /// <summary>
        /// Check if an element should be annotated even without system ID
        /// </summary>
        private bool IsAnnotatableElement(ElementData elem)
        {
            if (string.IsNullOrEmpty(elem.Category))
                return false;
            
            // Equipment
            if (elem.Category == "Оборудование" || elem.Category == "Mechanical Equipment")
                return true;
            
            // Air terminals (diffusers, grilles, etc.)
            if (elem.Category == "Воздухораспределители" || elem.Category == "Air Terminals")
                return true;
            
            // Duct accessories (dampers, fire dampers, etc.)
            if (elem.Category == "Арматура воздуховодов" || elem.Category == "Duct Accessories")
                return true;
            
            // Pipe accessories
            if (elem.Category == "Арматура трубопроводов" || elem.Category == "Pipe Accessories")
                return true;
            
            // Ducts and pipes
            if (elem.Category == "Воздуховоды" || elem.Category == "Ducts")
                return true;
            
            if (elem.Category == "Трубы" || elem.Category == "Pipes")
                return true;
            
            return false;
        }
        
        private SystemGraphNode CreateNode(ElementData elem)
        {
            return new SystemGraphNode
            {
                ElementId = elem.ElementId,
                Category = elem.Category,
                FamilyName = elem.FamilyName,
                TypeName = elem.TypeName,
                ViewStart = new ViewCoordinates2D { X = elem.ViewStart.X, Y = elem.ViewStart.Y },
                ViewEnd = elem.HasEndPoint ? new ViewCoordinates2D { X = elem.ViewEnd.X, Y = elem.ViewEnd.Y } : null,
                ModelStart = new ModelCoordinates3D { X = elem.ModelStart.X, Y = elem.ModelStart.Y, Z = elem.ModelStart.Z },
                ModelEnd = elem.HasEndPoint ? new ModelCoordinates3D { X = elem.ModelEnd.X, Y = elem.ModelEnd.Y, Z = elem.ModelEnd.Z } : null,
                HasEndPoint = elem.HasEndPoint,
                Diameter = elem.Diameter,
                Width = elem.Width,
                Height = elem.Height,
                SizeDisplay = elem.SizeDisplay,
                SystemId = elem.SystemId,
                SystemName = elem.SystemName
            };
        }
        
        private void BuildConnections(SystemGraph graph, List<ElementData> elements)
        {
            // First try using Revit Connector API for MEP elements
            var elementDict = elements.ToDictionary(e => e.ElementId, e => e);
            
            foreach (var elemData in elements)
            {
                if (!graph.Nodes.ContainsKey(elemData.ElementId))
                    continue;
                
                var node = graph.Nodes[elemData.ElementId];
                
                // Try to get physical connections via Connector API
                try
                {
                    // Get all connectors for this element
                    var connectors = GetConnectorsForElement(elemData.ElementId, elements.FirstOrDefault(e => e.ElementId == elemData.ElementId));
                    
                    foreach (var connectedId in connectors)
                    {
                        if (connectedId != elemData.ElementId && graph.Nodes.ContainsKey(connectedId))
                        {
                            // Add connection
                            if (!node.ConnectedNodeIds.Contains(connectedId))
                            {
                                node.ConnectedNodeIds.Add(connectedId);
                                
                                // Also add reverse connection
                                var otherNode = graph.Nodes[connectedId];
                                if (!otherNode.ConnectedNodeIds.Contains(elemData.ElementId))
                                {
                                    otherNode.ConnectedNodeIds.Add(elemData.ElementId);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[SYSTEM-GRAPH] Connector API failed for {elemData.ElementId}: {ex.Message}");
                }
            }
            
            // Fallback: also check geometric connections for elements not connected via API
            var linearElements = elements.Where(e => e.HasEndPoint).ToList();
            
            for (int i = 0; i < linearElements.Count; i++)
            {
                for (int j = i + 1; j < linearElements.Count; j++)
                {
                    var elemA = linearElements[i];
                    var elemB = linearElements[j];
                    
                    var nodeA = graph.Nodes.ContainsKey(elemA.ElementId) ? graph.Nodes[elemA.ElementId] : null;
                    var nodeB = graph.Nodes.ContainsKey(elemB.ElementId) ? graph.Nodes[elemB.ElementId] : null;
                    
                    // Skip if already connected
                    if (nodeA != null && nodeB != null && 
                        (nodeA.ConnectedNodeIds.Contains(elemB.ElementId) || nodeB.ConnectedNodeIds.Contains(elemA.ElementId)))
                        continue;
                    
                    // Check if endpoints match
                    var connection = FindConnection(elemA, elemB);
                    if (connection != null)
                    {
                        // Add edge and update node connections
                        graph.Edges.Add(connection);
                        
                        if (nodeA != null && !nodeA.ConnectedNodeIds.Contains(elemB.ElementId))
                            nodeA.ConnectedNodeIds.Add(elemB.ElementId);
                        
                        if (nodeB != null && !nodeB.ConnectedNodeIds.Contains(elemA.ElementId))
                            nodeB.ConnectedNodeIds.Add(elemA.ElementId);
                    }
                }
            }
            
            // Also check connections between linear elements and point elements
            var pointElements = elements.Where(e => !e.HasEndPoint).ToList();
            foreach (var pointElem in pointElements)
            {
                if (!graph.Nodes.ContainsKey(pointElem.ElementId))
                    continue;
                    
                foreach (var linearElem in linearElements)
                {
                    if (!graph.Nodes.ContainsKey(linearElem.ElementId))
                        continue;
                        
                    if (IsPointOnLine(pointElem.ViewStart, linearElem.ViewStart, linearElem.ViewEnd))
                    {
                        var pointNode = graph.Nodes[pointElem.ElementId];
                        var lineNode = graph.Nodes[linearElem.ElementId];
                        
                        if (!pointNode.ConnectedNodeIds.Contains(linearElem.ElementId))
                            pointNode.ConnectedNodeIds.Add(linearElem.ElementId);
                        
                        if (!lineNode.ConnectedNodeIds.Contains(pointElem.ElementId))
                            lineNode.ConnectedNodeIds.Add(pointElem.ElementId);
                    }
                }
            }
            
            // Log connection statistics
            var totalConnections = graph.Nodes.Values.Sum(n => n.ConnectedNodeIds.Count) / 2;
            var junctionCount = graph.Nodes.Values.Count(n => n.IsJunction);
            DebugLogger.Log($"[SYSTEM-GRAPH] Built {totalConnections} connections, {junctionCount} junctions found");
        }
        
        /// <summary>
        /// Get connected element IDs using Revit Connector API
        /// </summary>
        private HashSet<long> GetConnectorsForElement(long elementId, ElementData elemData)
        {
            var connectedIds = new HashSet<long>();
            
            // Note: We don't have access to Document here, so we rely on geometric fallback
            // In a real implementation, we would need to pass the Document or use a different approach
            // For now, return empty and let the geometric method handle it
            
            return connectedIds;
        }
        
        private SystemGraphEdge FindConnection(ElementData elemA, ElementData elemB)
        {
            // Check all endpoint combinations using both view and model coordinates
            var endpoints = new[]
            {
                (elemA.ViewStart, elemB.ViewStart, elemA.ModelStart, elemB.ModelStart),
                (elemA.ViewStart, elemB.ViewEnd, elemA.ModelStart, elemB.ModelEnd),
                (elemA.ViewEnd, elemB.ViewStart, elemA.ModelEnd, elemB.ModelStart),
                (elemA.ViewEnd, elemB.ViewEnd, elemA.ModelEnd, elemB.ModelEnd)
            };
        
            foreach (var (viewA, viewB, modelA, modelB) in endpoints)
            {
                // Check view coordinates
                double viewDist = Math.Sqrt(
                    Math.Pow(viewA.X - viewB.X, 2) +
                    Math.Pow(viewA.Y - viewB.Y, 2));
                        
                // Also check model coordinates for more reliable detection
                double modelDist = Math.Sqrt(
                    Math.Pow(modelA.X - modelB.X, 2) +
                    Math.Pow(modelA.Y - modelB.Y, 2) +
                    Math.Pow(modelA.Z - modelB.Z, 2));
                        
                // Accept connection if either view or model coordinates match
                if (viewDist < CONNECTOR_TOLERANCE || modelDist < MODEL_TOLERANCE)
                {
                    DebugLogger.Log($"[SYSTEM-GRAPH] Found connection between {elemA.ElementId} and {elemB.ElementId}: viewDist={viewDist:F4}, modelDist={modelDist:F4}");
                            
                    return new SystemGraphEdge
                    {
                        FromNodeId = elemA.ElementId,
                        ToNodeId = elemB.ElementId,
                        ConnectionPoint = new ViewCoordinates2D
                        {
                            X = (viewA.X + viewB.X) / 2,
                            Y = (viewA.Y + viewB.Y) / 2
                        },
                        ModelConnectionPoint = new ModelCoordinates3D
                        {
                            X = (modelA.X + modelB.X) / 2,
                            Y = (modelA.Y + modelB.Y) / 2,
                            Z = (modelA.Z + modelB.Z) / 2
                        }
                    };
                }
            }
        
            return null;
        }
        
        private bool IsPointOnLine(Coordinates2D point, Coordinates2D lineStart, Coordinates2D lineEnd)
        {
            // Check if point is approximately on the line segment
            double d1 = Distance2D(point, lineStart);
            double d2 = Distance2D(point, lineEnd);
            double lineLen = Distance2D(lineStart, lineEnd);
            
            // Point is on line if d1 + d2 ≈ lineLen
            return Math.Abs(d1 + d2 - lineLen) < CONNECTOR_TOLERANCE;
        }
        
        private double Distance2D(Coordinates2D a, Coordinates2D b)
        {
            return Math.Sqrt(Math.Pow(a.X - b.X, 2) + Math.Pow(a.Y - b.Y, 2));
        }
        
        private void IdentifyTrunksAndBranches(SystemGraph graph)
        {
            // Find equipment nodes as starting points
            var equipmentNodes = graph.Nodes.Values
                .Where(n => n.IsEquipment)
                .ToList();
            
            // Find endpoint nodes (air terminals, etc.)
            var endpoints = graph.Nodes.Values
                .Where(n => n.IsEndpoint && (n.IsAirTerminal || n.IsEquipment))
                .ToList();
            
            // Simple heuristic: main trunk is the path with most flow
            // Branches are paths that split from the trunk
            
            // For each endpoint, trace back to find the main path
            foreach (var endpoint in endpoints)
            {
                var path = TracePathFromEndpoint(graph, endpoint);
                if (path.Count > 0)
                {
                    // Check if this is a main trunk or branch
                    bool hasJunction = path.Any(id => graph.Nodes.ContainsKey(id) && graph.Nodes[id].IsJunction);
                    
                    if (hasJunction || path.Count > 3)
                    {
                        graph.MainTrunks.Add(path);
                    }
                    else
                    {
                        graph.Branches.Add(path);
                    }
                }
            }
        }
        
        private List<long> TracePathFromEndpoint(SystemGraph graph, SystemGraphNode endpoint)
        {
            var path = new List<long> { endpoint.ElementId };
            var visited = new HashSet<long> { endpoint.ElementId };
            
            var current = endpoint;
            while (current != null && current.ConnectedNodeIds.Count > 0)
            {
                // Find next unvisited connected node
                var nextId = current.ConnectedNodeIds.FirstOrDefault(id => !visited.Contains(id));
                if (nextId == 0) break;
                
                visited.Add(nextId);
                path.Add(nextId);
                
                if (graph.Nodes.ContainsKey(nextId))
                    current = graph.Nodes[nextId];
                else
                    break;
            }
            
            return path;
        }
    }
    
    // Helper classes for coordinates
    public class ViewCoordinates2D
    {
        public double X { get; set; }
        public double Y { get; set; }
    }
    
    public class ModelCoordinates3D
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }
}
