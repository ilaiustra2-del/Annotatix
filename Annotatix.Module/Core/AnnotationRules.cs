using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Types of annotations used in ductwork systems
    /// </summary>
    public enum AnnotationType
    {
        /// <summary>Round duct - size and flow (Круглый воздуховод_Размер и расход)</summary>
        DuctRoundSizeFlow = 1,
        
        /// <summary>Rectangular duct - size and flow (Прямоугольный воздуховод_Размер и расход)</summary>
        DuctRectSizeFlow = 2,
        
        /// <summary>Air terminal - type name / flow (Имя типа / Расход_30)</summary>
        AirTerminalTypeFlow = 3,
        
        /// <summary>Air terminal - short name / flow (Наименование краткое / Расход_20)</summary>
        AirTerminalShortNameFlow = 4,
        
        /// <summary>Duct accessory - (Арматура воздуховодов)</summary>
        DuctAccessory = 5,
        
        /// <summary>Spot dimension - elevation mark (Высотная отметка)</summary>
        SpotDimension = 6,
        
        /// <summary>Equipment mark (Марка_20)</summary>
        EquipmentMark = 7
    }
    
    /// <summary>
    /// Annotation positions relative to the element
    /// </summary>
    public enum AnnotationPosition
    {
        TopLeft = 0,
        TopCenter = 1,
        TopRight = 2,
        BottomLeft = 3,
        BottomCenter = 4,
        BottomRight = 5,
        HorizontalLeft = 6,
        HorizontalRight = 7
    }
    
    /// <summary>
    /// Attachment point on the element (for linear elements)
    /// </summary>
    public enum AttachmentPoint
    {
        Start = 0,      // 20% from start
        Middle = 1,     // 50% (center)
        End = 2         // 80% from start
    }
    
    /// <summary>
    /// Represents an annotation to be placed
    /// </summary>
    public class AnnotationPlan
    {
        /// <summary>Element to annotate</summary>
        public long ElementId { get; set; }
        
        /// <summary>Type of annotation to place</summary>
        public AnnotationType AnnotationType { get; set; }
        
        /// <summary>Family name for the annotation</summary>
        public string FamilyName { get; set; }
        
        /// <summary>Type name for the annotation</summary>
        public string TypeName { get; set; }
        
        /// <summary>Reason why this element needs annotation</summary>
        public string Reason { get; set; }
        
        /// <summary>Possible positions to try (ordered by preference)</summary>
        public List<AnnotationPosition> PreferredPositions { get; set; } = new List<AnnotationPosition>();
        
        /// <summary>Attachment point on the element</summary>
        public AttachmentPoint AttachmentPoint { get; set; } = AttachmentPoint.Middle;
        
        /// <summary>Is this annotation mandatory (equipment, accessory)?</summary>
        public bool IsMandatory { get; set; }
        
        /// <summary>Does this need a leader with elbow?</summary>
        public bool HasElbow { get; set; } = true;
        
        /// <summary>Reference to the node in graph</summary>
        public SystemGraphNode Node { get; set; }
    }
    
    /// <summary>
    /// Determines which elements need annotation based on system graph analysis
    /// </summary>
    public class AnnotationRulesEngine
    {
        // Standardized annotation families
        private const string FAMILY_DUCT_ROUND = "ADSK_Марка_Воздуховодов_ОсновныеОбозначения";
        private const string TYPE_DUCT_ROUND = "Круглый воздуховод_Размер и расход (2)";
        
        private const string FAMILY_DUCT_RECT = "ADSK_Марка_Воздуховодов_ОсновныеОбозначения";
        private const string TYPE_DUCT_RECT = "Прямоугольный воздуховод_Размер и расход (2)";
        
        private const string FAMILY_AIR_TERMINAL_1 = "ADSK_M_Воздухораспределители";
        private const string TYPE_AIR_TERMINAL_1 = "Имя типа / Расход (25)";
        
        private const string FAMILY_AIR_TERMINAL_2 = "ADSK_M_Воздухораспределители";
        private const string TYPE_AIR_TERMINAL_2 = "ADSK_Наименование краткое / Расход (2)";
        
        private const string FAMILY_DUCT_ACCESSORY = "ADSK_M_Арматура воздуховодов";
        private const string TYPE_DUCT_ACCESSORY = "ADSK_Марка (2)";
        
        private const string FAMILY_SPOT_DIMENSION = "Сист. семейство: Высотные отметки";
        private const string TYPE_SPOT_DIMENSION = "ADSK_Стрелка_Относительная_Вниз";
        
        private const string FAMILY_EQUIPMENT = "ADSK_M_Оборудование";
        private const string TYPE_EQUIPMENT = "ADSK_Марка (2)";
        
        /// <summary>
        /// Analyze system graph and determine annotations needed
        /// </summary>
        public List<AnnotationPlan> AnalyzeSystem(SystemGraph graph)
        {
            var plans = new List<AnnotationPlan>();
            
            // Rule 1: Equipment always gets annotated
            foreach (var node in graph.Nodes.Values.Where(n => n.IsEquipment))
            {
                plans.Add(CreateEquipmentPlan(node));
            }
            
            // Rule 2: Duct accessories always get annotated
            foreach (var node in graph.Nodes.Values.Where(n => n.IsAccessory))
            {
                plans.Add(CreateAccessoryPlan(node));
            }
            
            // Rule 3: Air terminals always get annotated
            foreach (var node in graph.Nodes.Values.Where(n => n.IsAirTerminal))
            {
                plans.Add(CreateAirTerminalPlan(node));
            }
            
            // Rule 4: Ducts at junctions - all branches get annotated
            foreach (var junction in graph.Junctions)
            {
                plans.AddRange(CreateJunctionPlans(graph, junction));
            }
            
            // Rule 5: Ducts with same size in sequence - annotate representative
            plans.AddRange(CreateSequencePlans(graph));
            
            // Note: Spot dimensions disabled for now - they require special handling
            // Rule 6: Spot dimensions at vertical transitions
            // plans.AddRange(CreateSpotDimensionPlans(graph));
            
            return plans;
        }
        
        private AnnotationPlan CreateEquipmentPlan(SystemGraphNode node)
        {
            return new AnnotationPlan
            {
                ElementId = node.ElementId,
                AnnotationType = AnnotationType.EquipmentMark,
                FamilyName = FAMILY_EQUIPMENT,
                TypeName = TYPE_EQUIPMENT,
                Reason = "Equipment requires mandatory annotation",
                PreferredPositions = GetAllPositions(),
                AttachmentPoint = AttachmentPoint.Middle,
                IsMandatory = true,
                HasElbow = true,
                Node = node
            };
        }
        
        private AnnotationPlan CreateAccessoryPlan(SystemGraphNode node)
        {
            return new AnnotationPlan
            {
                ElementId = node.ElementId,
                AnnotationType = AnnotationType.DuctAccessory,
                FamilyName = FAMILY_DUCT_ACCESSORY,
                TypeName = TYPE_DUCT_ACCESSORY,
                Reason = "Duct accessory requires mandatory annotation",
                PreferredPositions = GetAllPositions(),
                AttachmentPoint = AttachmentPoint.Middle,
                IsMandatory = true,
                HasElbow = true,
                Node = node
            };
        }
        
        private AnnotationPlan CreateAirTerminalPlan(SystemGraphNode node)
        {
            return new AnnotationPlan
            {
                ElementId = node.ElementId,
                AnnotationType = AnnotationType.AirTerminalShortNameFlow,
                FamilyName = FAMILY_AIR_TERMINAL_2,
                TypeName = TYPE_AIR_TERMINAL_2,
                Reason = "Air terminal requires mandatory annotation",
                PreferredPositions = GetPositionsForPointElement(node),
                AttachmentPoint = AttachmentPoint.Middle,
                IsMandatory = true,
                HasElbow = true,
                Node = node
            };
        }
        
        private List<AnnotationPlan> CreateJunctionPlans(SystemGraph graph, SystemGraphNode junction)
        {
            var plans = new List<AnnotationPlan>();
            
            // Get all connected ducts
            var connectedDucts = junction.ConnectedNodeIds
                .Where(id => graph.Nodes.ContainsKey(id) && graph.Nodes[id].IsLinearElement)
                .Select(id => graph.Nodes[id])
                .ToList();
            
            // Annotate each connected duct at the junction
            foreach (var duct in connectedDucts)
            {
                // Skip if already planned
                if (plans.Any(p => p.ElementId == duct.ElementId))
                    continue;
                
                // Determine attachment point based on which end is near the junction
                AttachmentPoint attachPoint = DetermineAttachmentPoint(duct, junction);
                
                plans.Add(CreateDuctPlan(duct, "At junction", attachPoint));
            }
            
            return plans;
        }
        
        private List<AnnotationPlan> CreateSequencePlans(SystemGraph graph)
        {
            var plans = new List<AnnotationPlan>();
            var annotatedIds = new HashSet<long>();
            
            // Group ducts by size
            var ductsBySize = graph.LinearElements
                .GroupBy(d => d.SizeDisplay ?? "Unknown")
                .ToDictionary(g => g.Key, g => g.ToList());
            
            foreach (var sizeGroup in ductsBySize)
            {
                var ducts = sizeGroup.Value;
                
                // Find sequences of ducts with same size
                foreach (var duct in ducts)
                {
                    if (annotatedIds.Contains(duct.ElementId))
                        continue;
                    
                    // Check if this duct is in a sequence (degree 2, same size neighbors)
                    if (duct.IsThrough)
                    {
                        var neighbors = duct.ConnectedNodeIds
                            .Where(id => graph.Nodes.ContainsKey(id))
                            .Select(id => graph.Nodes[id])
                            .Where(n => n.IsLinearElement && n.SizeDisplay == duct.SizeDisplay)
                            .ToList();
                        
                        if (neighbors.Count > 0)
                        {
                            // This duct is in a sequence - annotate one of them
                            plans.Add(CreateDuctPlan(duct, "Same-size sequence", AttachmentPoint.Middle));
                            
                            // Mark neighbors as annotated (skip them)
                            foreach (var neighbor in neighbors)
                            {
                                annotatedIds.Add(neighbor.ElementId);
                            }
                        }
                    }
                    
                    annotatedIds.Add(duct.ElementId);
                }
            }
            
            return plans;
        }
        
        private List<AnnotationPlan> CreateSpotDimensionPlans(SystemGraph graph)
        {
            var plans = new List<AnnotationPlan>();
            
            // Find vertical transitions (ducts that go up or down)
            var verticalDucts = graph.LinearElements
                .Where(d => d.ModelStart != null && d.ModelEnd != null)
                .Where(d => Math.Abs(d.ModelStart.Z - d.ModelEnd.Z) > 0.1) // Significant Z change
                .ToList();
            
            foreach (var duct in verticalDucts)
            {
                // Add spot dimension at both ends
                plans.Add(new AnnotationPlan
                {
                    ElementId = duct.ElementId,
                    AnnotationType = AnnotationType.SpotDimension,
                    FamilyName = FAMILY_SPOT_DIMENSION,
                    TypeName = TYPE_SPOT_DIMENSION,
                    Reason = "Vertical duct - elevation mark required",
                    PreferredPositions = new List<AnnotationPosition> { 
                        AnnotationPosition.TopLeft, 
                        AnnotationPosition.TopRight 
                    },
                    AttachmentPoint = AttachmentPoint.Start,
                    IsMandatory = true,
                    HasElbow = false,
                    Node = duct
                });
            }
            
            // Also check for horizontal segments before/after vertical transitions
            foreach (var junction in graph.Junctions)
            {
                var connectedDucts = junction.ConnectedNodeIds
                    .Where(id => graph.Nodes.ContainsKey(id))
                    .Select(id => graph.Nodes[id])
                    .Where(n => n.IsLinearElement)
                    .ToList();
                
                // Check if this junction connects elements at different elevations
                if (connectedDucts.Count >= 2)
                {
                    var elevations = connectedDucts
                        .Select(d => d.ModelStart?.Z ?? 0)
                        .Distinct()
                        .ToList();
                    
                    if (elevations.Count > 1)
                    {
                        // Elevation change - add spot dimensions
                        foreach (var duct in connectedDucts)
                        {
                            if (!plans.Any(p => p.ElementId == duct.ElementId && p.AnnotationType == AnnotationType.SpotDimension))
                            {
                                plans.Add(new AnnotationPlan
                                {
                                    ElementId = duct.ElementId,
                                    AnnotationType = AnnotationType.SpotDimension,
                                    FamilyName = FAMILY_SPOT_DIMENSION,
                                    TypeName = TYPE_SPOT_DIMENSION,
                                    Reason = "Before/after elevation change",
                                    PreferredPositions = new List<AnnotationPosition> { 
                                        AnnotationPosition.TopLeft, 
                                        AnnotationPosition.TopRight 
                                    },
                                    AttachmentPoint = AttachmentPoint.Middle,
                                    IsMandatory = true,
                                    HasElbow = false,
                                    Node = duct
                                });
                            }
                        }
                    }
                }
            }
            
            return plans;
        }
        
        private AnnotationPlan CreateDuctPlan(SystemGraphNode duct, string reason, AttachmentPoint attachPoint)
        {
            // Determine if round or rectangular
            bool isRound = duct.Diameter.HasValue && duct.Diameter > 0;
            
            return new AnnotationPlan
            {
                ElementId = duct.ElementId,
                AnnotationType = isRound ? AnnotationType.DuctRoundSizeFlow : AnnotationType.DuctRectSizeFlow,
                FamilyName = isRound ? FAMILY_DUCT_ROUND : FAMILY_DUCT_RECT,
                TypeName = isRound ? TYPE_DUCT_ROUND : TYPE_DUCT_RECT,
                Reason = reason,
                PreferredPositions = GetPositionsForLinearElement(duct),
                AttachmentPoint = attachPoint,
                IsMandatory = false,
                HasElbow = true,
                Node = duct
            };
        }
        
        private AttachmentPoint DetermineAttachmentPoint(SystemGraphNode duct, SystemGraphNode junction)
        {
            // Determine which end of the duct is near the junction
            if (duct.ViewStart != null && duct.ViewEnd != null)
            {
                // Calculate distance from junction to each end
                double distToStart = Distance(
                    duct.ViewStart.X, duct.ViewStart.Y,
                    duct.ViewStart.X, duct.ViewStart.Y); // Simplified - would need junction coords
                
                // For now, use middle
                return AttachmentPoint.Middle;
            }
            
            return AttachmentPoint.Middle;
        }
        
        private List<AnnotationPosition> GetPositionsForLinearElement(SystemGraphNode node)
        {
            // All 6 positions for linear elements
            return new List<AnnotationPosition>
            {
                AnnotationPosition.TopLeft,
                AnnotationPosition.TopRight,
                AnnotationPosition.BottomLeft,
                AnnotationPosition.BottomRight,
                AnnotationPosition.HorizontalLeft,
                AnnotationPosition.HorizontalRight
            };
        }
        
        private List<AnnotationPosition> GetPositionsForPointElement(SystemGraphNode node)
        {
            return new List<AnnotationPosition>
            {
                AnnotationPosition.TopLeft,
                AnnotationPosition.TopRight,
                AnnotationPosition.BottomLeft,
                AnnotationPosition.BottomRight
            };
        }
        
        private List<AnnotationPosition> GetAllPositions()
        {
            return Enum.GetValues(typeof(AnnotationPosition))
                .Cast<AnnotationPosition>()
                .ToList();
        }
        
        private double Distance(double x1, double y1, double x2, double y2)
        {
            return Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));
        }
    }
}
