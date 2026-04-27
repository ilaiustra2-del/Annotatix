using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Represents a group of nearby same-type elements that should share one annotation
    /// </summary>
    public class ElementGroup
    {
        public AnnotationPlan PrimaryPlan { get; set; }
        public List<AnnotationPlan> AdditionalPlans { get; set; } = new List<AnnotationPlan>();
        
        /// <summary>
        /// Whether the primary (extreme) element is on the RIGHT side of the group.
        /// If true, prefer right-facing annotation positions; if false, prefer left-facing.
        /// </summary>
        public bool IsRightExtreme { get; set; } = true;
    }
    
    /// <summary>
    /// Result of a single annotation placement attempt
    /// </summary>
    public class PlacementResult
    {
        public bool Success { get; set; }
        public long ElementId { get; set; }
        public AnnotationPlan Plan { get; set; }
        public AnnotationPosition Position { get; set; }
        public AttachmentPoint AttachmentPoint { get; set; }
        public double ElbowHeight { get; set; }
        public string FailureReason { get; set; }
        
        /// <summary>Created tag element (if successful)</summary>
        public IndependentTag CreatedTag { get; set; }
    }
    
    /// <summary>
    /// Configuration for annotation placement
    /// </summary>
    public class PlacementConfig
    {
        // Base offsets (in model units, typically meters)
        // These will be scaled by view scale and element size
        public double BaseElbowHeight { get; set; } = 1.5;  // Base elbow height in model units
        public double BaseHorizontalOffset { get; set; } = 2.0;  // Base horizontal offset in model units
        public double MinLeaderLength { get; set; } = 1.0;  // Minimum leader length for visibility
        public double ElementSizeMultiplier { get; set; } = 0.75;  // Additional offset based on element size (increased from 0.5)
        
        // Scale factor for view scale (e.g., 1:100 = 100)
        public double ViewScaleMultiplier { get; set; } = 0.01;  // Multiply by view scale
        
        // Z offset for 3D views (how far above/below the element the annotation should be)
        public double BaseZOffset { get; set; } = 1.5;  // Base Z offset for 3D views
        public double ZOffsetSizeMultiplier { get; set; } = 0.5;  // Additional Z offset based on element depth
        
        // Legacy properties for backward compatibility - these now set the base values
        public double InitialElbowHeight 
        { 
            get => BaseElbowHeight; 
            set => BaseElbowHeight = value; 
        }
        public double MaxElbowHeight { get; set; } = 5.0;
        public double ElbowHeightStep { get; set; } = 0.5;
        public double HorizontalOffset 
        { 
            get => BaseHorizontalOffset; 
            set => BaseHorizontalOffset = value; 
        }
    }
    
    /// <summary>
    /// Greedy algorithm for placing annotations with collision avoidance
    /// </summary>
    public class GreedyPlacementService
    {
        private readonly Document _document;
        private readonly View _view;
        private readonly CollisionDetector _collisionDetector;
        private readonly Dictionary<AnnotationPlan, AnnotationSize> _sizes;
        private readonly PlacementConfig _config;
        private readonly TagTypeManager _tagTypeManager;
        
        // CSV export data collectors
        private readonly List<AnnotationPlacementRecord> _placementRecords = new List<AnnotationPlacementRecord>();
        private readonly List<PlacementIterationRecord> _iterationRecords = new List<PlacementIterationRecord>();
        
        /// <summary>Get placement summary records for CSV export</summary>
        public List<AnnotationPlacementRecord> PlacementRecords => _placementRecords;
        /// <summary>Get iteration detail records for CSV export</summary>
        public List<PlacementIterationRecord> IterationRecords => _iterationRecords;
        
        public GreedyPlacementService(
            Document document, 
            View view, 
            CollisionDetector collisionDetector,
            Dictionary<AnnotationPlan, AnnotationSize> sizes,
            PlacementConfig config = null)
        {
            _document = document;
            _view = view;
            _collisionDetector = collisionDetector;
            _sizes = sizes;
            _config = config ?? new PlacementConfig();
            _tagTypeManager = new TagTypeManager(document);
        }
        
        /// <summary>
        /// Convert 3D model coordinates to 2D view coordinates (screen space).
        /// This is critical for 3D views where model coordinates don't match screen coordinates.
        /// </summary>
        private (double X, double Y) ConvertModelToViewCoordinates(XYZ modelPoint)
        {
            if (_view == null)
                return (modelPoint.X, modelPoint.Y);
            
            var viewOrigin = _view.Origin;
            var viewRight = _view.RightDirection;
            var viewUp = _view.UpDirection;
            
            var relative = modelPoint - viewOrigin;
            double viewX = relative.DotProduct(viewRight);
            double viewY = relative.DotProduct(viewUp);
            
            return (viewX, viewY);
        }
        
        /// <summary>
        /// Compute the occupied bounding box for a placed annotation in 3D view.
        /// TagHeadPosition is at the shelf EDGE - bbox extends from there outward.
        /// Uses actual shelf length from tag type name for accurate width.
        /// Uses AnnotationSizer height (based on character height and line count) for height.
        /// </summary>
        private BBox2D ComputeOccupiedBbox3D(XYZ headPos3D, string tagName, AnnotationPosition position, AnnotationPlan plan)
        {
            try
            {
                var headView = ConvertModelToViewCoordinates(headPos3D);
                
                // Get actual shelf length from the tag type name
                double? actualShelfMm = TagTypeManager.GetShelfLengthFromTypeNamePublic(tagName);
                double currentViewScale = _view.Scale;
                if (currentViewScale < 1) currentViewScale = 1;
                
                double shelfFeet, textHeightFeet, paddingFeet;
                if (actualShelfMm.HasValue && actualShelfMm.Value > 0)
                {
                    shelfFeet = actualShelfMm.Value / 304.8 * currentViewScale;
                    textHeightFeet = _sizes.TryGetValue(plan, out var planSz) ? planSz.Height : 0.05;
                    paddingFeet = _sizes.TryGetValue(plan, out var planSzPad) ? planSzPad.Padding : 0.05;
                }
                else
                {
                    shelfFeet = _sizes.TryGetValue(plan, out var planSizeW) ? planSizeW.Width : 0.15;
                    textHeightFeet = _sizes.TryGetValue(plan, out var planSizeH) ? planSizeH.Height : 0.05;
                    paddingFeet = _sizes.TryGetValue(plan, out var planSizeP) ? planSizeP.Padding : 0.05;
                }
                
                // Convert shelf width and text height to view coordinates
                // Paper-space dimensions map directly to view coordinate offsets.
                // The view coordinate system uses the same units as model space (feet).
                // Horizontal dimensions (shelf, padding) → viewX, vertical (text height) → viewY.
                double viewShelfWidth, viewTextHeight, viewPad;
                viewShelfWidth = shelfFeet;       // horizontal on screen (viewX)
                viewTextHeight = textHeightFeet;   // vertical on screen (viewY)
                viewPad = paddingFeet;             // both directions
                if (viewShelfWidth < 0.02) viewShelfWidth = 0.02;
                if (viewTextHeight < 0.02) viewTextHeight = 0.02;
                if (viewPad < 0.005) viewPad = 0.005;
                
                // Build bbox: headView is at shelf TIP (far end from leader connection)
                // Shelf extends from headView TOWARD the element by viewShelfWidth
                bool isRight = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
                bool isCenterOrigin = plan.AnnotationType == AnnotationType.DuctAccessory ||
                                      plan.AnnotationType == AnnotationType.EquipmentMark;
                bool isRightEdgeOrigin = plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                                         plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow;
                double viewShelfHalfWidth = viewShelfWidth / 2.0;
                var bbox = new BBox2D();
                if (isCenterOrigin)
                {
                    bbox.MinX = headView.X - viewShelfHalfWidth - viewPad;
                    bbox.MaxX = headView.X + viewShelfHalfWidth + viewPad;
                }
                else if (isRightEdgeOrigin)
                {
                    // RIGHT-EDGE origin: TagHeadPosition = RIGHT edge of shelf
                    // Shelf always extends LEFT from head by full shelfWidth, regardless of direction
                    bbox.MinX = headView.X - viewShelfWidth - viewPad;
                    bbox.MaxX = headView.X + viewPad;
                }
                else // LEFT-EDGE origin
                {
                    // LEFT-EDGE origin: TagHeadPosition = LEFT edge of shelf
                    // Shelf always extends RIGHT from head by full shelfWidth, regardless of direction
                    bbox.MinX = headView.X - viewPad;
                    bbox.MaxX = headView.X + viewShelfWidth + viewPad;
                }
                
                // Y extent: text is centered on the shelf line (headView.Y)
                // For 2-line tags, text extends both above and below the shelf
                bbox.MinY = headView.Y - viewTextHeight / 2 - viewPad;
                bbox.MaxY = headView.Y + viewTextHeight / 2 + viewPad;
                
                return bbox;
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// Place all annotations greedily, avoiding collisions.
        /// For nearby same-type elements (Air Terminals, Duct Accessories),
        /// creates ONE annotation with multiple leaders (Add Basis feature).
        /// </summary>
        public List<PlacementResult> PlaceAll(List<AnnotationPlan> plans)
        {
            var results = new List<PlacementResult>();
            int annotationIndex = 0; // Global annotation counter for logging
            
            // Find groups of nearby same-type elements that should share one annotation
            var groupedElementIds = new HashSet<long>();
            var groups = FindNearbySameTypeGroups(plans);
            
            // Sort plans: mandatory first, then by category priority (air terminals → duct accessories → equipment → ducts),
            // then by node degree (leaves first)
            var sortedPlans = plans
                .OrderByDescending(p => p.IsMandatory)
                .ThenBy(p => GetCategoryPriority(p.AnnotationType))
                .ThenBy(p => p.Node?.Degree ?? 0)
                .ToList();
            
            // First, process groups (create one annotation per group with AddReferences)
            foreach (var group in groups)
            {
                annotationIndex++;
                
                // Try all elements in the group as primary until one succeeds
                // The extreme element is already set as primary by FindNearbySameTypeGroups
                var allGroupPlans = new List<AnnotationPlan> { group.PrimaryPlan };
                allGroupPlans.AddRange(group.AdditionalPlans);
                
                // Override position preferences based on EACH element's position in the group:
                // - If element is the RIGHTMOST → prefer right-facing configs (TopRight, HorizontalRight, BottomRight)
                // - If element is the LEFTMOST → prefer left-facing configs (TopLeft, HorizontalLeft, BottomLeft)
                // - If element is in the MIDDLE → use the group's IsRightExtreme flag
                // OverridePositionOrder=true ensures these positions are used as-is without
                // being reordered by GetOptimalPositionOrderFromOrientation
                // 
                // Calculate each element's viewX to determine its position in the group
                var elemViewXMap = new Dictionary<long, double>();
                foreach (var p in allGroupPlans)
                {
                    var elem = _document.GetElement(new ElementId(p.ElementId));
                    if (elem != null)
                    {
                        var loc = GetElementLocation(elem);
                        if (loc != null)
                        {
                            var (vx, vy) = ConvertModelToViewCoordinates(loc);
                            elemViewXMap[p.ElementId] = vx;
                        }
                    }
                }
                
                double groupMinViewX = elemViewXMap.Count > 0 ? elemViewXMap.Values.Min() : 0;
                double groupMaxViewX = elemViewXMap.Count > 0 ? elemViewXMap.Values.Max() : 0;
                double viewXTolerance = (groupMaxViewX - groupMinViewX) * 0.1; // 10% tolerance for "middle" elements
                if (viewXTolerance < 0.001) viewXTolerance = 0.001;
                
                foreach (var plan in allGroupPlans)
                {
                    plan.OverridePositionOrder = true;
                    
                    // Determine if this specific element is rightmost, leftmost, or middle
                    bool isRightmost = elemViewXMap.ContainsKey(plan.ElementId) &&
                                       (groupMaxViewX - elemViewXMap[plan.ElementId]) < viewXTolerance;
                    bool isLeftmost = elemViewXMap.ContainsKey(plan.ElementId) &&
                                      (elemViewXMap[plan.ElementId] - groupMinViewX) < viewXTolerance;
                    
                    if (isRightmost)
                    {
                        plan.PreferredPositions = new List<AnnotationPosition>
                        {
                            AnnotationPosition.TopRight,
                            AnnotationPosition.HorizontalRight,
                            AnnotationPosition.BottomRight,
                            AnnotationPosition.TopLeft,
                            AnnotationPosition.HorizontalLeft,
                            AnnotationPosition.BottomLeft
                        };
                    }
                    else if (isLeftmost)
                    {
                        plan.PreferredPositions = new List<AnnotationPosition>
                        {
                            AnnotationPosition.TopLeft,
                            AnnotationPosition.HorizontalLeft,
                            AnnotationPosition.BottomLeft,
                            AnnotationPosition.TopRight,
                            AnnotationPosition.HorizontalRight,
                            AnnotationPosition.BottomRight
                        };
                    }
                    else
                    {
                        // Middle element: use group direction as fallback
                        if (group.IsRightExtreme)
                        {
                            plan.PreferredPositions = new List<AnnotationPosition>
                            {
                                AnnotationPosition.TopRight,
                                AnnotationPosition.HorizontalRight,
                                AnnotationPosition.BottomRight,
                                AnnotationPosition.TopLeft,
                                AnnotationPosition.HorizontalLeft,
                                AnnotationPosition.BottomLeft
                            };
                        }
                        else
                        {
                            plan.PreferredPositions = new List<AnnotationPosition>
                            {
                                AnnotationPosition.TopLeft,
                                AnnotationPosition.HorizontalLeft,
                                AnnotationPosition.BottomLeft,
                                AnnotationPosition.TopRight,
                                AnnotationPosition.HorizontalRight,
                                AnnotationPosition.BottomRight
                            };
                        }
                    }
                }
                
                // Log position assignments for each group element
                foreach (var p in allGroupPlans)
                {
                    string dirLabel = elemViewXMap.ContainsKey(p.ElementId) && (groupMaxViewX - elemViewXMap[p.ElementId]) < viewXTolerance ? "RIGHTMOST" :
                                      elemViewXMap.ContainsKey(p.ElementId) && (elemViewXMap[p.ElementId] - groupMinViewX) < viewXTolerance ? "LEFTMOST" : "MIDDLE";
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Element {p.ElementId}: {dirLabel}, positions={string.Join("→", p.PreferredPositions.Select(pos => PositionToRussianString(pos)))}");
                }
                
                // Mark additional elements as grouped (they won't get individual annotations)
                foreach (var additionalPlan in group.AdditionalPlans)
                {
                    groupedElementIds.Add(additionalPlan.ElementId);
                }
                
                PlacementResult groupResult = null;
                AnnotationPlan successfulPrimary = null;
                
                for (int tryIdx = 0; tryIdx < allGroupPlans.Count; tryIdx++)
                {
                    var primaryPlan = allGroupPlans[tryIdx];
                    var result = PlaceSingle(primaryPlan, annotationIndex, allGroupPlans.Select(p => p.ElementId).ToList());
                    
                    if (result.Success && result.CreatedTag != null)
                    {
                        groupResult = result;
                        successfulPrimary = primaryPlan;
                        break;
                    }
                    
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Primary element {primaryPlan.ElementId} placement failed (attempt {tryIdx + 1}/{allGroupPlans.Count})");
                }
                
                // If all attempts failed, ungroup - let all elements get individual annotations
                if (groupResult == null || !groupResult.Success || groupResult.CreatedTag == null)
                {
                    DebugLogger.Log($"[GREEDY-PLACEMENT] All group elements failed, ungrouping {allGroupPlans.Count} elements");
                    foreach (var plan in allGroupPlans)
                    {
                        groupedElementIds.Remove(plan.ElementId);
                    }
                    if (groupResult != null) results.Add(groupResult);
                    continue;
                }
                
                // Success - add additional references to the same tag
                {
                    var tag = groupResult.CreatedTag;
                    bool addRefsSuccess = false;
                    // Add all OTHER group elements as additional references (excluding the successful primary)
                    var additionalPlansForRefs = allGroupPlans.Where(p => p.ElementId != successfulPrimary.ElementId).ToList();
                    try
                    {
                        var additionalRefs = new List<Reference>();
                        foreach (var additionalPlan in additionalPlansForRefs)
                        {
                            var additionalElement = _document.GetElement(new ElementId(additionalPlan.ElementId));
                            if (additionalElement != null)
                            {
                                var elemRef = GetElementReference(additionalElement);
                                if (elemRef != null)
                                {
                                    additionalRefs.Add(elemRef);
                                }
                            }
                        }
                        
                        if (additionalRefs.Count > 0)
                        {
                            tag.AddReferences(additionalRefs);
                            addRefsSuccess = true;
                            
                            // Set leader end positions for each additional reference
                            var taggedRefs = tag.GetTaggedReferences();
                            foreach (var taggedRef in taggedRefs)
                            {
                                try
                                {
                                    // Find the element for this reference
                                    var refElement = _document.GetElement(taggedRef.ElementId);
                                    if (refElement != null)
                                    {
                                        var refLocation = GetElementLocation(refElement);
                                        if (refLocation != null)
                                        {
                                            tag.SetLeaderEnd(taggedRef, refLocation);
                                            
                                            // Set elbow: for group references, always use L-shaped geometry.
                                            // Even when the primary annotation is Horizontal, reference leaders
                                            // need proper right-angle connections: vertical from element to
                                            // head Y, then horizontal to the tag head.
                                            var headPos = tag.TagHeadPosition;
                                            AnnotationPosition refElbowPos = groupResult.Position;
                                            bool isHorizontalRefPos = refElbowPos == AnnotationPosition.HorizontalLeft || refElbowPos == AnnotationPosition.HorizontalRight;
                                            if (isHorizontalRefPos)
                                            {
                                                // Primary is horizontal → group references must be L-shaped
                                                refElbowPos = AnnotationPosition.TopRight;
                                            }
                                            XYZ elbowPos = CalculateElbowPosition(refLocation, headPos, refElbowPos, groupResult.ElbowHeight);
                                            tag.SetLeaderElbow(taggedRef, elbowPos);
                                        }
                                    }
                                }
                                catch (Exception refEx)
                                {
                                    DebugLogger.Log($"[GREEDY-PLACEMENT] Error setting leader for additional reference: {refEx.Message}");
                                }
                            }
                            
                            // Re-set TagHeadPosition after all leader operations
                            try
                            {
                                tag.TagHeadPosition = groupResult.CreatedTag.TagHeadPosition; // maintain original head position
                            }
                            catch { }
                            
                            DebugLogger.Log($"[GREEDY-PLACEMENT] Added {additionalRefs.Count} references to tag {tag.Id} (group annotation)");
                        }
                    }
                    catch (Exception addRefEx)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] AddReferences failed: {addRefEx.Message}");
                        addRefsSuccess = false;
                    }
                    
                    // If AddReferences failed, ungroup so additional elements get individual annotations
                    if (!addRefsSuccess)
                    {
                        groupedElementIds.ExceptWith(additionalPlansForRefs.Select(p => p.ElementId));
                    }
                    
                    // Register occupied area AND annotation leader lines
                    if (_view is View3D)
                    {
                        // 3D view: register occupied area in VIEW coordinates
                        // TagHeadPosition is at shelf EDGE - bbox extends from there outward
                        var headPos3D = tag.TagHeadPosition;
                        if (headPos3D != null)
                        {
                            var occupiedBbox = ComputeOccupiedBbox3D(headPos3D, tag.Name, groupResult.Position, successfulPrimary);
                            if (occupiedBbox != null)
                            {
                                DebugLogger.Log($"[BBOX] 3D Grouped ElemId={successfulPrimary.ElementId}: bbox Min=({occupiedBbox.MinX:F3}, {occupiedBbox.MinY:F3}) Max=({occupiedBbox.MaxX:F3}, {occupiedBbox.MaxY:F3}) Size=({occupiedBbox.Width:F3}x{occupiedBbox.Height:F3}) HeadModel=({headPos3D.X:F3}, {headPos3D.Y:F3}, {headPos3D.Z:F3})");
                                _collisionDetector.AddOccupiedArea(occupiedBbox, successfulPrimary.ElementId);
                            }
                        }
                    }
                    else
                    {
                        var bbox = GetTagBoundingBox(tag);
                        if (bbox != null)
                        {
                            DebugLogger.Log($"[BBOX] 2D Grouped ElemId={successfulPrimary.ElementId}: bbox Min=({bbox.MinX:F3}, {bbox.MinY:F3}) Max=({bbox.MaxX:F3}, {bbox.MaxY:F3}) Size=({bbox.Width:F3}x{bbox.Height:F3})");
                            _collisionDetector.AddOccupiedArea(bbox, successfulPrimary.ElementId);
                        }
                    }
                    
                    // Register annotation leader lines for future collision checks
                    RegisterAnnotationLinesForCollision(tag, successfulPrimary.ElementId);
                }
                
                results.Add(groupResult);
                
                // Add placeholder results for grouped elements (excluding the successful primary)
                foreach (var additionalPlan in allGroupPlans.Where(p => p.ElementId != successfulPrimary.ElementId))
                {
                    results.Add(new PlacementResult
                    {
                        Success = true,
                        ElementId = additionalPlan.ElementId,
                        Plan = additionalPlan,
                        Position = groupResult.Position,
                        AttachmentPoint = additionalPlan.AttachmentPoint,
                        ElbowHeight = groupResult.ElbowHeight,
                        CreatedTag = groupResult.CreatedTag,
                        FailureReason = "Grouped with primary element " + successfulPrimary.ElementId
                    });
                }
            }
            
            // Then process remaining individual plans
            foreach (var plan in sortedPlans)
            {
                // Skip elements that were already grouped
                if (groupedElementIds.Contains(plan.ElementId))
                    continue;
                
                // Skip elements that were the primary of a group (already processed)
                if (groups.Any(g => g.PrimaryPlan.ElementId == plan.ElementId))
                    continue;
                
                annotationIndex++;
                var result = PlaceSingle(plan, annotationIndex);
                results.Add(result);
                
                // If successful, register occupied area AND annotation leader lines
                if (result.Success && result.CreatedTag != null)
                {
                    if (_view is View3D)
                    {
                        // 3D view: register occupied area in VIEW coordinates
                        // TagHeadPosition is at shelf EDGE - bbox extends from there outward
                        var headPos3D = result.CreatedTag.TagHeadPosition;
                        if (headPos3D != null)
                        {
                            var occupiedBbox = ComputeOccupiedBbox3D(headPos3D, result.CreatedTag.Name, result.Position, plan);
                            if (occupiedBbox != null)
                            {
                                DebugLogger.Log($"[BBOX] 3D ElemId={plan.ElementId}: bbox Min=({occupiedBbox.MinX:F3}, {occupiedBbox.MinY:F3}) Max=({occupiedBbox.MaxX:F3}, {occupiedBbox.MaxY:F3}) Size=({occupiedBbox.Width:F3}x{occupiedBbox.Height:F3}) HeadModel=({headPos3D.X:F3}, {headPos3D.Y:F3}, {headPos3D.Z:F3})");
                                _collisionDetector.AddOccupiedArea(occupiedBbox, plan.ElementId);
                            }
                        }
                    }
                    else
                    {
                        var bbox = GetTagBoundingBox(result.CreatedTag);
                        if (bbox != null)
                        {
                            DebugLogger.Log($"[BBOX] 2D ElemId={plan.ElementId}: bbox Min=({bbox.MinX:F3}, {bbox.MinY:F3}) Max=({bbox.MaxX:F3}, {bbox.MaxY:F3}) Size=({bbox.Width:F3}x{bbox.Height:F3})");
                            _collisionDetector.AddOccupiedArea(bbox, plan.ElementId);
                        }
                    }
                    
                    // Register annotation leader lines for future collision checks
                    RegisterAnnotationLinesForCollision(result.CreatedTag, plan.ElementId);
                }
            }
            
            return results;
        }
        
        private PlacementResult PlaceSingle(AnnotationPlan plan, int annotationIndex, List<long> groupElementIds = null)
        {
            var element = _document.GetElement(new ElementId(plan.ElementId));
            if (element == null)
            {
                DebugLogger.Log($"[Аннотация {annotationIndex}]: Элемент не найден (id {plan.ElementId})");
                return new PlacementResult
                {
                    Success = false,
                    ElementId = plan.ElementId,
                    Plan = plan,
                    FailureReason = "Element not found"
                };
            }
                
            // Get annotation size
            if (!_sizes.TryGetValue(plan, out var size))
            {
                size = new AnnotationSize { Width = 0.15, Height = 0.05 };
            }
                
            // Get element location
            var location = GetElementLocation(element);
            if (location == null)
            {
                DebugLogger.Log($"[Аннотация {annotationIndex}]: Не удалось определить положение элемента (id {plan.ElementId})");
                return new PlacementResult
                {
                    Success = false,
                    ElementId = plan.ElementId,
                    Plan = plan,
                    FailureReason = "Could not get element location"
                };
            }
                
            // ═══════════════════════════════════════════════════════════════
            // DETAILED PER-ANNOTATION LOGGING BLOCK
            // ═══════════════════════════════════════════════════════════════
            
            // Element info
            string elemCategory = element.Category?.Name ?? "?";
            string elemFamily = "?";
            string elemType = "?";
            try
            {
                var elemTypeObj = _document.GetElement(element.GetTypeId());
                if (elemTypeObj != null)
                {
                    elemFamily = (elemTypeObj as FamilySymbol)?.Family?.Name ?? elemTypeObj.Name;
                    elemType = elemTypeObj.Name;
                }
            }
            catch { }
            
            DebugLogger.Log("");
            DebugLogger.Log($"[Аннотация {annotationIndex}]:");
            DebugLogger.Log($"- Элемент аннотирования: (id {plan.ElementId}) {elemCategory} - {elemFamily} - {elemType}");
            
            // Element position in space (model + view coords)
            var locationView = ConvertModelToViewCoordinates(location);
            string positionDesc;
            ElementOrientation orientation;
            if (plan.Node != null && plan.Node.ViewStart != null && plan.Node.ViewEnd != null)
            {
                orientation = DetermineElementOrientationFromViewCoords(
                    plan.Node.ViewStart.X, plan.Node.ViewStart.Y,
                    plan.Node.ViewEnd.X, plan.Node.ViewEnd.Y);
                
                // Build position description like: \ (x1<x2), (y1>y2)
                string arrow;
                string coords;
                double x1 = plan.Node.ViewStart.X, y1 = plan.Node.ViewStart.Y;
                double x2 = plan.Node.ViewEnd.X, y2 = plan.Node.ViewEnd.Y;
                switch (orientation)
                {
                    case ElementOrientation.DiagonalLeftHigher:
                        arrow = "\\";  // DiagonalLeftHigher (slopes down-right)
                        coords = $"(x1<x2), (y1>y2)";
                        break;
                    case ElementOrientation.DiagonalRightHigher:
                        arrow = "//";  // DiagonalRightHigher (slopes up-right)
                        coords = $"(x1<x2), (y1<y2)";
                        break;
                    case ElementOrientation.Horizontal:
                        arrow = "--";  // Horizontal
                        coords = $"(y1~=y2)";
                        break;
                    case ElementOrientation.Vertical:
                        arrow = "||";  // Vertical
                        coords = $"(x1~=x2)";
                        break;
                    default:
                        arrow = "?";
                        coords = "";
                        break;
                }
                positionDesc = $"{arrow} {coords}";
                
                DebugLogger.Log($"  StartModelX\tStartModelY\tStartModelZ\tStartViewX\tStartViewY\tEndModelX\tEndModelY\tEndModelZ\tEndViewX\tEndViewY");
                DebugLogger.Log($"  {location.X:F3}\t{location.Y:F3}\t{location.Z:F3}\t{locationView.X:F3}\t{locationView.Y:F3}\t{location.X:F3}\t{location.Y:F3}\t{location.Z:F3}\t{locationView.X:F3}\t{locationView.Y:F3}");
            }
            else
            {
                // For point elements (accessories, air terminals, equipment),
                // try to inherit orientation from connected ducts in the system.
                // These elements are connected to ducts whose orientation is already known.
                orientation = DetermineOrientationFromConnectedDucts(plan);
                if (orientation == ElementOrientation.Horizontal) // fallback if no connected duct found
                {
                    orientation = DetermineElementOrientation(element.get_BoundingBox(_view));
                }
                positionDesc = orientation.ToString();
                DebugLogger.Log($"  Положение элемента: ({location.X:F3}, {location.Y:F3}, {location.Z:F3}), view=({locationView.X:F3}, {locationView.Y:F3}), orientation={orientation} (inherited from connected duct)");
            }
            
            // For group annotations with OverridePositionOrder, use positions as-is
            // (already ordered by extreme element direction, not element orientation)
            var orderedPositions = plan.OverridePositionOrder
                ? plan.PreferredPositions
                : GetOptimalPositionOrderFromOrientation(plan.PreferredPositions, orientation);
            string chosenConfig = PositionToRussianString(orderedPositions.FirstOrDefault());
            DebugLogger.Log($"- Положение в пространстве: {positionDesc} - выбрана конфигурация аннотации {chosenConfig}");
            
            // Annotation family/type info
            string annFamilyName = plan.FamilyName ?? "?";
            string annTypeName = plan.TypeName ?? "?";
            DebugLogger.Log($"- Выбранное семейство аннотации: {annFamilyName}");
            DebugLogger.Log($"  Тип аннотации: {annTypeName}");
            
            // Calculate optimal offsets based on element size and view scale
            double optimalHorizontalOffset = CalculateOptimalOffset(location, element, plan.AnnotationType);
            double optimalBaseElbowHeight = CalculateOptimalElbowHeight(location, element, 
                plan.PreferredPositions.FirstOrDefault());
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] Element {plan.ElementId}: optimalHorizontalOffset={optimalHorizontalOffset:F2}, optimalBaseElbowHeight={optimalBaseElbowHeight:F2}");
                    
            // Per user's algorithm: for each position, try increasing elbow heights
            // before switching to the next priority position.
            // For Top positions: increase height (move annotation further UP)
            // For Bottom positions: increase height (move annotation further DOWN)
            // For Horizontal positions: no height adjustment (just 1 attempt)
            // Max height: 20mm on paper, step: 1mm on paper
            double viewScaleForHeight = _view.Scale;
            if (viewScaleForHeight < 1) viewScaleForHeight = 1;
            double elbowStepModel = 1.0 / 304.8 * viewScaleForHeight; // 1mm paper step in model feet
            double maxElbowModel = 15.0 / 304.8 * viewScaleForHeight;  // 15mm paper max in model feet
            const int maxHeightIterations = 15; // enough iterations to cover 3mm..15mm range with 1mm steps
                    
            bool is3DView = _view is View3D;
            double viewScaleCollision = _view.Scale;
            if (viewScaleCollision < 1) viewScaleCollision = 1;
            double shelfGapViewCollision = 1.0 / 304.8 * viewScaleCollision; // Fixed 1mm shelf gap for collision detection
            
            // Determine attachment points for the element
            // For ducts: 3 points (center, 25% from start, 25% from end)
            // For other elements: 1 point (center only)
            List<(XYZ Point, string Description)> attachmentPoints = GetAttachmentPoints(element, plan);
            
            // Try each position
            int globalIteration = 0;
            AnnotationPosition? lastShapePosition = null; // Track shape changes
            foreach (var position in orderedPositions)
            {
                bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
                bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
                bool isHorizontalPos = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
                bool isLShaped = isTop || isBottom;
                        
                // Detect shape change and log it
                if (lastShapePosition != null)
                {
                    bool lastWasHorizontal = (lastShapePosition == AnnotationPosition.HorizontalLeft || lastShapePosition == AnnotationPosition.HorizontalRight);
                    bool currentIsHorizontal = isHorizontalPos;
                    if (lastWasHorizontal != currentIsHorizontal)
                    {
                        DebugLogger.Log($"--Смена формы аннотации -> {PositionToRussianString(position)}--");
                    }
                    else if (lastShapePosition != position)
                    {
                        DebugLogger.Log($"--Смена направления -> {PositionToRussianString(position)}--");
                    }
                }
                lastShapePosition = position;
                        
                // For ducts: try each attachment point at each elbow height.
                // Loop order: for each position, try each elbow height, then try each attachment point.
                // This ensures we prefer shorter elbow heights across all attachment points.
                // For other elements: single attachment point (center)
                int iterations = isLShaped ? maxHeightIterations : 1;
                
                for (int iteration = 0; iteration < iterations; iteration++)
                {
                    globalIteration++;
                    double elbowHeight = optimalBaseElbowHeight + iteration * elbowStepModel;
                    if (elbowHeight > maxElbowModel)
                        break;
                    
                    // Try each attachment point at this elbow height
                    foreach (var attachPoint in attachmentPoints)
                    {
                        // Update location to current attachment point
                        location = attachPoint.Point;
                        locationView = ConvertModelToViewCoordinates(location);
                        
                        if (attachmentPoints.Count > 1)
                        {
                            DebugLogger.Log($"--Точка крепления: {attachPoint.Description} ({location.X:F3}, {location.Y:F3}, {location.Z:F3})--");
                        }
                            
                    // Calculate candidate placement bbox with dynamic offset
                    var candidateBbox = CalculatePlacementBboxWithOffset(location, position, elbowHeight, size, optimalHorizontalOffset);
                                                    
                    // Calculate leader line coordinates
                    var headPos = CalculateHeadPosition(location, position, elbowHeight, size, optimalHorizontalOffset);
                    var elbowPos = CalculateElbowPositionFromViewCoords(location.X, location.Y, headPos.X, headPos.Y, position);
                                                    
                    bool headerCollides = false;
                    bool leaderCollides = false;
                    CollisionDetector.CollisionDetails headerDetails = null;
                    CollisionDetector.CollisionDetails leaderDetails = null;
                    
                    // View-coordinate variables (used for 3D views and logging)
                    double leaderEndViewX = 0, leaderEndViewY = 0;
                    double elbowViewX = 0, elbowViewY = 0;
                    double headViewX = 0, headViewY = 0;
                    double currentViewOffsetH = 0; // stores viewOffsetH from 3D view block for CreateAnnotation
                                        
                    if (is3DView)
                    {
                        // ============================================================
                        // 3D VIEW: ALL calculations in 2D view coordinates
                        // Per the user's algorithm: project everything to 2D first,
                        // then compute annotation positions and check collisions
                        // entirely in the 2D projected coordinate system.
                        // This ensures collision detection with model elements works
                        // correctly because all geometry exists in the same 2D plane.
                        // ============================================================
                        var locationViewCol = ConvertModelToViewCoordinates(location);
                                            
                        // Calculate shelf dimensions in view coordinates
                        // The view's scale factor affects how model distances map to view distances.
                        // For an isometric 3D view, a 1-foot model offset may appear as a
                        // smaller offset in view coordinates depending on the viewing angle.
                        // We measure the actual mapping by projecting reference points.
                        double viewShelfWidth, viewTextHeight, viewPad, viewOffsetH, viewOffsetV;
                        bool isRightPos = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
                        bool isLeftPos = position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft || position == AnnotationPosition.HorizontalLeft;
                        bool isPointLikeElement = plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                                                  plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow ||
                                                  plan.AnnotationType == AnnotationType.EquipmentMark;
                        {
                            // Paper-space dimensions map directly to view coordinate offsets.
                            // The view coordinate system (viewX=screen horizontal, viewY=screen vertical)
                            // uses the same units as model space (feet), and 1 foot along viewRight
                            // = 1 viewX unit, 1 foot along viewUp = 1 viewY unit.
                            // So paperMm / 304.8 * viewScale gives the correct view offset in feet.
                            // We do NOT project along model X/Y axes because in isometric views
                            // the model axes are not aligned with screen axes.
                            viewShelfWidth = size.Width;         // horizontal on screen (viewX)
                            viewTextHeight = size.Height;        // vertical on screen (viewY)
                            viewPad = size.Padding;              // both directions
                            // For L-shaped leaders: the horizontal offset must clear the element's bbox.
                            // Point-like elements (air terminals, equipment, accessories) extend far horizontally
                            // from the leader end point (element center). A fixed 1mm gap is not enough.
                            // For horizontal leaders: use full optimalHorizontalOffset (element must be cleared horizontally)
                            double defaultLShapedGap = 1.0 / 304.8 * viewScaleCollision; // 1mm paper gap for L-shaped
                            if (isHorizontalPos)
                            {
                                viewOffsetH = optimalHorizontalOffset; // full offset needed to clear element body
                            }
                            else
                            {
                                viewOffsetH = defaultLShapedGap; // start with 1mm, may increase below
                            }
                            viewOffsetV = elbowHeight;           // vertical offset (viewY)
                                                
                            if (viewShelfWidth < 0.02) viewShelfWidth = 0.02;
                            if (viewTextHeight < 0.02) viewTextHeight = 0.02;
                            if (viewPad < 0.005) viewPad = 0.005;
                            if (viewOffsetH < 0.005) viewOffsetH = 0.005;
                            if (viewOffsetV < 0.005) viewOffsetV = 0.005;
                            
                            
                            // HORIZONTAL OFFSET: Ensure the header clears the element's bbox.
                            // For ALL positions (horizontal AND L-shaped) of air terminals and equipment (isPointLikeElement),
                            // the horizontal offset must be large enough to place the shelf past the element's edge.
                            // For group annotations: use the COMBINED group bbox to ensure the shelf clears
                            // ALL group members (not just the primary element).
                            // For single elements: use the element's own bbox.
                            // Note: we do NOT include viewPad here — the padding area may overlap
                            // with the element; only the actual text area must clear the element.
                            bool isGroupAnnotationForOffset = groupElementIds != null && groupElementIds.Count > 1;
                            if (isPointLikeElement)
                            {
                                // For group annotations: clear against the COMBINED group bbox.
                                // For single elements: use the element's own bbox.
                                var combinedBbox = isGroupAnnotationForOffset
                                    ? GetCombinedElementViewBBox(element, groupElementIds)
                                    : GetElementViewBBox(element);
                                if (combinedBbox != null)
                                {
                                    double minClearOffset = defaultLShapedGap; // at least 1mm
                                    double oneMmView = 1.0 / 304.8 * viewScaleCollision;
                                    if (isRightPos)
                                    {
                                        // Header text area must start past the element's RIGHT edge
                                        double clearOffset = combinedBbox.MaxX - locationViewCol.X + oneMmView;
                                        if (clearOffset > minClearOffset) minClearOffset = clearOffset;
                                    }
                                    else if (isLeftPos)
                                    {
                                        // Header text area near edge must be before the element's LEFT edge
                                        double clearOffset = locationViewCol.X - combinedBbox.MinX + oneMmView;
                                        if (clearOffset > minClearOffset) minClearOffset = clearOffset;
                                    }
                                    if (minClearOffset > viewOffsetH)
                                    {
                                        DebugLogger.Log($"[GREEDY-PLACEMENT] Offset increased: {viewOffsetH:F3}→{minClearOffset:F3}ft to clear element bbox (rightPos={isRightPos}, type={plan.AnnotationType}, horizontal={isHorizontalPos})");
                                        viewOffsetH = minClearOffset;
                                    }
                                }
                            }
                            
                            currentViewOffsetH = viewOffsetH;
                        }
                                            
                        // Compute leader end, elbow, and head positions in 2D view coordinates
                        //
                        // Tag families have different reference point (TagHeadPosition) positions:
                        // - LEFT-EDGE origin (duct marks, air terminals): TagHeadPosition = left edge of tag
                        //   LEFT-facing: head at LEFT TIP = leaderEnd - (shelfGap + shelfWidth)
                        //     Near edge = head + shelfWidth. Gap = shelfGap. ✓
                        //   RIGHT-facing: head at LEFT edge (near) = leaderEnd + shelfGap
                        //     Near edge = head. Gap = shelfGap. ✓
                        //
                        // - CENTER origin (duct accessories, equipment): TagHeadPosition = center of tag
                        //   LEFT-facing: head at CENTER = leaderEnd - (shelfGap + shelfWidth/2)
                        //     Near edge = head + shelfWidth/2. Gap = shelfGap. ✓
                        //   RIGHT-facing: head at CENTER = leaderEnd + (shelfGap + shelfWidth/2)
                        //     Near edge = head - shelfWidth/2. Gap = shelfGap. ✓
                        bool isCenterOriginTag = plan.AnnotationType == AnnotationType.DuctAccessory ||
                                                 plan.AnnotationType == AnnotationType.EquipmentMark;
                        bool isRightEdgeOrigin = plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                                                     plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow;
                        double viewShelfHalfWidth = viewShelfWidth / 2.0;
                                                
                        leaderEndViewX = locationViewCol.X;
                        leaderEndViewY = locationViewCol.Y;
                                                                        
                        if (isHorizontalPos)
                        {
                            // Horizontal: leader goes sideways, no vertical segment
                            elbowViewX = locationViewCol.X;
                            elbowViewY = locationViewCol.Y;
                            if (isLeftPos)
                            {
                                if (isCenterOriginTag)
                                    headViewX = locationViewCol.X - viewOffsetH - viewShelfHalfWidth;
                                else if (isRightEdgeOrigin)
                                    // RIGHT-EDGE origin, left-facing: TagHeadPosition = RIGHT edge (near leader)
                                    headViewX = locationViewCol.X - viewOffsetH;
                                else
                                    // LEFT-EDGE origin, left-facing: TagHeadPosition = LEFT tip (far from leader)
                                    headViewX = locationViewCol.X - viewOffsetH - viewShelfWidth;
                            }
                            else // isRightPos
                            {
                                if (isCenterOriginTag)
                                    headViewX = locationViewCol.X + viewOffsetH + viewShelfHalfWidth;
                                else if (isRightEdgeOrigin)
                                    // RIGHT-EDGE origin, right-facing: TagHeadPosition = RIGHT tip (far from leader)
                                    headViewX = locationViewCol.X + viewOffsetH + viewShelfWidth;
                                else
                                    // LEFT-EDGE origin, right-facing: TagHeadPosition = LEFT edge (near leader)
                                    headViewX = locationViewCol.X + viewOffsetH;
                            }
                            headViewY = locationViewCol.Y;
                        }
                        else
                        {
                            // L-shaped: vertical segment from element to elbow, then horizontal to head
                            elbowViewX = locationViewCol.X;
                            if (isTop)
                                elbowViewY = locationViewCol.Y + viewOffsetV;
                            else
                                elbowViewY = locationViewCol.Y - viewOffsetV;
                                                                        
                            if (isRightPos)
                            {
                                if (isCenterOriginTag)
                                    headViewX = locationViewCol.X + viewOffsetH + viewShelfHalfWidth;
                                else if (isRightEdgeOrigin)
                                    // RIGHT-EDGE origin, right-facing: TagHeadPosition = RIGHT tip (far from leader)
                                    headViewX = locationViewCol.X + viewOffsetH + viewShelfWidth;
                                else
                                    // LEFT-EDGE origin, right-facing: TagHeadPosition = LEFT edge (near leader)
                                    headViewX = locationViewCol.X + viewOffsetH;
                            }
                            else // isLeftPos
                            {
                                if (isCenterOriginTag)
                                    headViewX = locationViewCol.X - viewOffsetH - viewShelfHalfWidth;
                                else if (isRightEdgeOrigin)
                                    // RIGHT-EDGE origin, left-facing: TagHeadPosition = RIGHT edge (near leader)
                                    headViewX = locationViewCol.X - viewOffsetH;
                                else
                                    // LEFT-EDGE origin, left-facing: TagHeadPosition = LEFT tip (far from leader)
                                    headViewX = locationViewCol.X - viewOffsetH - viewShelfWidth;
                            }
                            headViewY = elbowViewY;
                        }
                                                                        
                        // Build collision bbox in 2D view coordinates
                        var candidateBboxView = new BBox2D();
                        if (isCenterOriginTag)
                        {
                            // CENTER origin: tag extends shelfWidth/2 on each side from head
                            candidateBboxView.MinX = headViewX - viewShelfHalfWidth - viewPad;
                            candidateBboxView.MaxX = headViewX + viewShelfHalfWidth + viewPad;
                        }
                        else if (isRightEdgeOrigin)
                        {
                            // RIGHT-EDGE origin: TagHeadPosition = RIGHT edge of shelf
                            // Shelf always extends LEFT from head by full shelfWidth, regardless of direction
                            candidateBboxView.MinX = headViewX - viewShelfWidth - viewPad;
                            candidateBboxView.MaxX = headViewX + viewPad;
                        }
                        else
                        {
                            // LEFT-EDGE origin: TagHeadPosition = LEFT edge of shelf
                            // Shelf always extends RIGHT from head by full shelfWidth, regardless of direction
                            candidateBboxView.MinX = headViewX - viewPad;
                            candidateBboxView.MaxX = headViewX + viewShelfWidth + viewPad;
                        }
                                                                        
                        // Y extent: for single-line annotations, text is above the shelf line;
                        // for multi-line (duct marks, air terminals), text extends both above and below.
                        bool isMultiLineTag = plan.AnnotationType == AnnotationType.DuctRoundSizeFlow ||
                                              plan.AnnotationType == AnnotationType.DuctRectSizeFlow ||
                                              plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                                              plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow;
                        if (isMultiLineTag)
                        {
                            // Multi-line: text centered on shelf line
                            candidateBboxView.MinY = headViewY - viewTextHeight / 2 - viewPad;
                            candidateBboxView.MaxY = headViewY + viewTextHeight / 2 + viewPad;
                        }
                        else
                        {
                            // Single-line: text above shelf, bbox bottom at shelf line
                            candidateBboxView.MinY = headViewY - viewPad;
                            candidateBboxView.MaxY = headViewY + viewTextHeight + viewPad;
                        }
                                            
                        // All collision checks in 2D view coordinates
                        // Model element checks are now enabled because all geometry
                        // (both model line segments and annotation geometry) exists
                        // in the same 2D projected coordinate system.
                        headerDetails = _collisionDetector.GetHeaderCollisionDetails(candidateBboxView);
                        leaderDetails = _collisionDetector.GetLeaderCollisionDetails(
                            leaderEndViewX, leaderEndViewY,
                            elbowViewX, elbowViewY,
                            headViewX, headViewY,
                            plan.ElementId);
                        headerCollides = headerDetails.HasCollision;
                        leaderCollides = leaderDetails.HasCollision;
                        
                        // SELF-COLLISION CHECK: Verify header text area doesn't overlap
                        // the annotated element's bbox.
                        // ONLY for air terminals and equipment — these are point-like elements
                        // with no line segments in the collision detector.
                        // For ducts and duct accessories, skip this check — their centerlines are
                        // already handled by the collision detector.
                        // SKIP for group annotations (multiple elements share one tag) —
                        // the header may naturally overlap with group element bboxes.
                        // Note: we check WITHOUT padding — the padding area may overlap with
                        // the element; only the actual text area must not overlap.
                        bool isGroupAnnotation = groupElementIds != null && groupElementIds.Count > 1;
                        if (!headerCollides && isPointLikeElement && !isGroupAnnotation)
                        {
                            var combinedBbox = GetCombinedElementViewBBox(element, groupElementIds);
                            if (combinedBbox != null)
                            {
                                // Build a reduced bbox without padding for self-collision check.
                                var selfCheckBbox = new BBox2D();
                                if (isCenterOriginTag)
                                {
                                    selfCheckBbox.MinX = headViewX - viewShelfHalfWidth;
                                    selfCheckBbox.MaxX = headViewX + viewShelfHalfWidth;
                                }
                                else if (isRightEdgeOrigin)
                                {
                                    // RIGHT-EDGE origin: shelf always extends LEFT from head
                                    selfCheckBbox.MinX = headViewX - viewShelfWidth;
                                    selfCheckBbox.MaxX = headViewX;
                                }
                                else
                                {
                                    // LEFT-EDGE origin: shelf always extends RIGHT from head
                                    selfCheckBbox.MinX = headViewX;
                                    selfCheckBbox.MaxX = headViewX + viewShelfWidth;
                                }
                                if (isMultiLineTag)
                                {
                                    selfCheckBbox.MinY = headViewY - viewTextHeight / 2;
                                    selfCheckBbox.MaxY = headViewY + viewTextHeight / 2;
                                }
                                else
                                {
                                    selfCheckBbox.MinY = headViewY;
                                    selfCheckBbox.MaxY = headViewY + viewTextHeight;
                                }
                                if (selfCheckBbox.Intersects(combinedBbox))
                                {
                                    headerCollides = true;
                                    if (headerDetails == null) headerDetails = new CollisionDetector.CollisionDetails();
                                    headerDetails.HasCollision = true;
                                    headerDetails.CollisionTypes.Add("self_element_bbox");
                                    headerDetails.CollidingElementIds.Add(plan.ElementId);
                                }
                            }
                        }
                    }
                    else
                    {
                        // 2D view: collision detection with details
                        headerDetails = _collisionDetector.GetHeaderCollisionDetails(candidateBbox);
                        leaderDetails = _collisionDetector.GetLeaderCollisionDetails(
                            location.X, location.Y,
                            elbowPos.X, elbowPos.Y,
                            headPos.X, headPos.Y,
                            plan.ElementId);
                        headerCollides = headerDetails.HasCollision;
                        leaderCollides = leaderDetails.HasCollision;
                    }
                    
                    // Detailed iteration logging
                    string posName = PositionToRussianString(position);
                    double iterLeaderEndViewX, iterLeaderEndViewY, iterElbowViewX, iterElbowViewY, iterHeaderViewX, iterHeaderViewY;
                    if (is3DView)
                    {
                        iterLeaderEndViewX = leaderEndViewX; iterLeaderEndViewY = leaderEndViewY;
                        iterElbowViewX = elbowViewX; iterElbowViewY = elbowViewY;
                        iterHeaderViewX = headViewX; iterHeaderViewY = headViewY;
                        DebugLogger.Log($"[Итерация подбора положения {globalIteration}]");
                        DebugLogger.Log($"Форма аннотации: {posName}; LeaderEndViewX,Y = {leaderEndViewX:F2}, {leaderEndViewY:F2}; ElbowViewX,Y = {elbowViewX:F2}, {elbowViewY:F2}; HeaderViewX,Y = {headViewX:F2}, {headViewY:F2}");
                    }
                    else
                    {
                        iterLeaderEndViewX = location.X; iterLeaderEndViewY = location.Y;
                        iterElbowViewX = elbowPos.X; iterElbowViewY = elbowPos.Y;
                        iterHeaderViewX = headPos.X; iterHeaderViewY = headPos.Y;
                        DebugLogger.Log($"[Итерация подбора положения {globalIteration}]");
                        DebugLogger.Log($"Форма аннотации: {posName}; LeaderEndX,Y = {location.X:F2}, {location.Y:F2}; ElbowX,Y = {elbowPos.X:F2}, {elbowPos.Y:F2}; HeaderX,Y = {headPos.X:F2}, {headPos.Y:F2}");
                    }
                    
                    if (headerCollides || leaderCollides)
                    {
                        var allCollidingIds = new List<long>();
                        if (headerDetails != null) allCollidingIds.AddRange(headerDetails.CollidingElementIds);
                        if (leaderDetails != null) allCollidingIds.AddRange(leaderDetails.CollidingElementIds);
                        var distinctIds = allCollidingIds.Where(id => id != 0).Distinct().ToList();
                        string collisionSummary;
                        if (distinctIds.Count > 0)
                            collisionSummary = string.Join(", ", distinctIds);
                        else
                            collisionSummary = "occupied areas";
                        
                        string collisionTypes = "";
                        if (headerCollides) collisionTypes += "header";
                        if (leaderCollides) collisionTypes += (collisionTypes.Length > 0 ? "+" : "") + "leader";
                        DebugLogger.Log($"  - найдены коллизии ({collisionTypes}): {collisionSummary}");
                        
                        // Record iteration for CSV export
                        double elbowHeightMm = elbowHeight / _view.Scale * 304.8;
                        _iterationRecords.Add(new PlacementIterationRecord
                        {
                            AnnotationIndex = annotationIndex,
                            ElementId = plan.ElementId,
                            GlobalIteration = globalIteration,
                            PositionName = posName,
                            LeaderEndViewX = iterLeaderEndViewX,
                            LeaderEndViewY = iterLeaderEndViewY,
                            ElbowViewX = iterElbowViewX,
                            ElbowViewY = iterElbowViewY,
                            HeaderViewX = iterHeaderViewX,
                            HeaderViewY = iterHeaderViewY,
                            ElbowHeightMm = elbowHeightMm,
                            HeaderCollision = headerCollides,
                            LeaderCollision = leaderCollides,
                            CollidingElementIds = collisionSummary,
                            CollisionTypes = collisionTypes,
                            PlacementSucceeded = false
                        });
                    }
                                                    
                    if (!headerCollides && !leaderCollides)
                    {
                        DebugLogger.Log($"  - коллизий не обнаружено, создаем аннотацию");
                        
                        // Record successful iteration for CSV export
                        double elbowHeightMmSuccess = elbowHeight / _view.Scale * 304.8;
                        _iterationRecords.Add(new PlacementIterationRecord
                        {
                            AnnotationIndex = annotationIndex,
                            ElementId = plan.ElementId,
                            GlobalIteration = globalIteration,
                            PositionName = posName,
                            LeaderEndViewX = iterLeaderEndViewX,
                            LeaderEndViewY = iterLeaderEndViewY,
                            ElbowViewX = iterElbowViewX,
                            ElbowViewY = iterElbowViewY,
                            HeaderViewX = iterHeaderViewX,
                            HeaderViewY = iterHeaderViewY,
                            ElbowHeightMm = elbowHeightMmSuccess,
                            HeaderCollision = false,
                            LeaderCollision = false,
                            CollidingElementIds = "",
                            CollisionTypes = "",
                            PlacementSucceeded = true
                        });
                        
                        // Try to create the actual annotation
                        // For 3D views: pass the ACTUAL offset used by the placement algorithm
                        // (may have been increased to clear element bbox).
                        // For 2D views: always use the base optimalHorizontalOffset.
                        // For L-shaped positions: pass as lShapedOffset for CalculateHeadPositionFor3DView.
                        double effectiveOffset = (is3DView && isHorizontalPos) ? currentViewOffsetH : optimalHorizontalOffset;
                        double lShapedOffset = (is3DView && isLShaped) ? currentViewOffsetH : 0;
                        var tag = CreateAnnotation(plan, element, location, position, elbowHeight, effectiveOffset, lShapedOffset);
                
                        if (tag != null)
                        {
                            // After tag creation, the tag type may have been changed to a wider shelf.
                            // Re-check collisions with the actual (post-type-change) bbox.
                            // If collisions are found, delete the tag and continue iterating.
                            bool postCreateCollision = false;
                            string postCreateCollisionIds = "";
                            if (_view is View3D)
                            {
                                var headPos3D = tag.TagHeadPosition;
                                if (headPos3D != null)
                                {
                                    var actualBbox = ComputeOccupiedBbox3D(headPos3D, tag.Name, position, plan);
                                    if (actualBbox != null)
                                    {
                                        var postHeaderDetails = _collisionDetector.GetHeaderCollisionDetails(actualBbox);
                                        if (postHeaderDetails.HasCollision)
                                        {
                                            // Exclude self-element (and all group members) from collision check
                                            // Group annotations naturally overlap with other group members.
                                            var excludedIds = new HashSet<long> { plan.ElementId };
                                            if (groupElementIds != null)
                                            {
                                                foreach (var gid in groupElementIds)
                                                    excludedIds.Add(gid);
                                            }
                                            var nonSelfCollisions = postHeaderDetails.CollidingElementIds
                                                .Where(id => !excludedIds.Contains(id)).ToList();
                                            if (nonSelfCollisions.Count > 0)
                                            {
                                                postCreateCollision = true;
                                                postCreateCollisionIds = string.Join(", ", nonSelfCollisions);
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if (postCreateCollision)
                            {
                                DebugLogger.Log($"[GREEDY-PLACEMENT] Post-creation collision detected after shelf type change! Colliding with: {postCreateCollisionIds}. Deleting tag and continuing iteration.");
                                try
                                {
                                    using (var subTx = new Autodesk.Revit.DB.SubTransaction(_document))
                                    {
                                        subTx.Start();
                                        _document.Delete(tag.Id);
                                        subTx.Commit();
                                    }
                                }
                                catch (Exception delEx)
                                {
                                    DebugLogger.Log($"[GREEDY-PLACEMENT] Warning: could not delete tag {tag.Id}: {delEx.Message}");
                                }
                                // Continue iterating — do NOT return
                            }
                            else
                            {
                                // No post-creation collision — proceed with successful placement
                                
                            // Log annotation content and shelf length
                            string tagText = "";
                            try { tagText = tag.TagText ?? ""; } catch { }
                            double? shelfMm = TagTypeManager.GetShelfLengthFromTypeNamePublic(tag.Name);
                            string shelfInfo = shelfMm.HasValue ? $"{shelfMm.Value:F0}" : "?";
                            double sizeWidthMm = size.Width / _view.Scale * 304.8;
                            DebugLogger.Log($"- Содержимое аннотации: {tagText}. Длина полки: {shelfInfo} мм (расчетная ширина: {sizeWidthMm:F1} мм). Создан тип: {tag.Name}");
                            
                            // Record placement summary for CSV export
                            // Compute bounding box for CSV export
                            double recBboxMinX = 0, recBboxMinY = 0, recBboxMaxX = 0, recBboxMaxY = 0, recBboxW = 0, recBboxH = 0;
                            string recBboxCoordSys = "";
                            if (_view is View3D)
                            {
                                var headPos3D = tag.TagHeadPosition;
                                if (headPos3D != null)
                                {
                                    var occBbox = ComputeOccupiedBbox3D(headPos3D, tag.Name, position, plan);
                                    if (occBbox != null)
                                    {
                                        recBboxMinX = occBbox.MinX;
                                        recBboxMinY = occBbox.MinY;
                                        recBboxMaxX = occBbox.MaxX;
                                        recBboxMaxY = occBbox.MaxY;
                                        recBboxW = occBbox.Width;
                                        recBboxH = occBbox.Height;
                                        recBboxCoordSys = "view";
                                    }
                                }
                            }
                            else
                            {
                                var tagBbox = GetTagBoundingBox(tag);
                                if (tagBbox != null)
                                {
                                    recBboxMinX = tagBbox.MinX;
                                    recBboxMinY = tagBbox.MinY;
                                    recBboxMaxX = tagBbox.MaxX;
                                    recBboxMaxY = tagBbox.MaxY;
                                    recBboxW = tagBbox.Width;
                                    recBboxH = tagBbox.Height;
                                    recBboxCoordSys = "model";
                                }
                            }
                            
                            _placementRecords.Add(new AnnotationPlacementRecord
                            {
                                AnnotationIndex = annotationIndex,
                                ElementId = plan.ElementId,
                                ElementCategory = elemCategory,
                                ElementFamily = elemFamily,
                                ElementType = elemType,
                                LocationModelX = location.X,
                                LocationModelY = location.Y,
                                LocationModelZ = location.Z,
                                LocationViewX = locationView.X,
                                LocationViewY = locationView.Y,
                                OrientationSymbol = positionDesc.Split(' ')[0],
                                OrientationDescription = positionDesc,
                                AnnotationFamily = annFamilyName,
                                AnnotationType = annTypeName,
                                AnnotationContentType = plan.AnnotationType.ToString(),
                                TagText = tagText,
                                FinalTypeName = tag.Name,
                                ShelfLengthMm = shelfMm ?? 0,
                                CalculatedWidthMm = sizeWidthMm,
                                Success = true,
                                FinalPosition = posName,
                                FinalElbowHeightMm = elbowHeightMmSuccess,
                                TotalIterations = globalIteration,
                                FailureReason = "",
                                BboxMinX = recBboxMinX,
                                BboxMinY = recBboxMinY,
                                BboxMaxX = recBboxMaxX,
                                BboxMaxY = recBboxMaxY,
                                BboxWidth = recBboxW,
                                BboxHeight = recBboxH,
                                BboxCoordSystem = recBboxCoordSys
                            });
                            
                            return new PlacementResult
                            {
                                Success = true,
                                ElementId = plan.ElementId,
                                Plan = plan,
                                Position = position,
                                AttachmentPoint = plan.AttachmentPoint,
                                ElbowHeight = elbowHeight,
                                CreatedTag = tag
                            };
                            } // end else (no post-creation collision)
                            } // end if (tag != null)
                    } // end if (!headerCollides && !leaderCollides)
                    } // end foreach (attachPoint)
                } // end for (iteration)
            } // end foreach (position)
            
            DebugLogger.Log($"- НЕ УДАЛОСЬ разместить аннотацию: все позиции исчерпаны ({globalIteration} итераций)");
            
            // Record failed placement summary for CSV export
            _placementRecords.Add(new AnnotationPlacementRecord
            {
                AnnotationIndex = annotationIndex,
                ElementId = plan.ElementId,
                ElementCategory = elemCategory,
                ElementFamily = elemFamily,
                ElementType = elemType,
                LocationModelX = location.X,
                LocationModelY = location.Y,
                LocationModelZ = location.Z,
                LocationViewX = locationView.X,
                LocationViewY = locationView.Y,
                OrientationSymbol = positionDesc.Split(' ')[0],
                OrientationDescription = positionDesc,
                AnnotationFamily = annFamilyName,
                AnnotationType = annTypeName,
                AnnotationContentType = plan.AnnotationType.ToString(),
                TagText = "",
                FinalTypeName = "",
                ShelfLengthMm = 0,
                CalculatedWidthMm = size.Width / _view.Scale * 304.8,
                Success = false,
                FinalPosition = "",
                FinalElbowHeightMm = 0,
                TotalIterations = globalIteration,
                FailureReason = $"All positions exhausted ({globalIteration} iterations)"
            });
                
            return new PlacementResult
            {
                Success = false,
                ElementId = plan.ElementId,
                Plan = plan,
                FailureReason = $"No valid position found after {globalIteration} iterations"
            };
        }
        
        /// <summary>
        /// Convert AnnotationPosition to Russian string for logging
        /// </summary>
        private static string PositionToRussianString(AnnotationPosition? position)
        {
            if (position == null) return "?";
            switch (position.Value)
            {
                case AnnotationPosition.TopLeft: return "Влево вверх";
                case AnnotationPosition.TopCenter: return "Вверх по центру";
                case AnnotationPosition.TopRight: return "Вправо вверх";
                case AnnotationPosition.BottomLeft: return "Влево вниз";
                case AnnotationPosition.BottomCenter: return "Вниз по центру";
                case AnnotationPosition.BottomRight: return "Вправо вниз";
                case AnnotationPosition.HorizontalLeft: return "Горизонтально влево";
                case AnnotationPosition.HorizontalRight: return "Горизонтально вправо";
                default: return position.Value.ToString();
            }
        }
        
        /// <summary>
        /// Get the element's bounding box in 2D view coordinates.
        /// Used for self-collision checking: ensures the annotation header doesn't overlap
        /// the annotated element's own body (which the line-segment collision detector can't detect
        /// for point-like elements like air terminals, equipment, and accessories).
        /// </summary>
        private BBox2D GetElementViewBBox(Element element)
        {
            try
            {
                var bbox = element.get_BoundingBox(_view);
                if (bbox == null)
                {
                    DebugLogger.Log($"[GET-BBOX] elemId={element.Id.IntegerValue}: get_BoundingBox(_view) returned NULL");
                    return null;
                }
                
                // Project all 8 corners of the 3D bbox to 2D view coordinates
                // and find the min/max extents
                double minX = double.MaxValue, maxX = double.MinValue;
                double minY = double.MaxValue, maxY = double.MinValue;
                
                for (int i = 0; i < 8; i++)
                {
                    double x = (i & 1) == 0 ? bbox.Min.X : bbox.Max.X;
                    double y = (i & 2) == 0 ? bbox.Min.Y : bbox.Max.Y;
                    double z = (i & 4) == 0 ? bbox.Min.Z : bbox.Max.Z;
                    var corner = new XYZ(x, y, z);
                    var viewCoord = ConvertModelToViewCoordinates(corner);
                    if (viewCoord.X < minX) minX = viewCoord.X;
                    if (viewCoord.X > maxX) maxX = viewCoord.X;
                    if (viewCoord.Y < minY) minY = viewCoord.Y;
                    if (viewCoord.Y > maxY) maxY = viewCoord.Y;
                }
                
                DebugLogger.Log($"[GET-BBOX] elemId={element.Id.IntegerValue}: result MinX={minX:F3} MaxX={maxX:F3} (bboxMinX={bbox.Min.X:F5} bboxMaxX={bbox.Max.X:F5})");
                return new BBox2D { MinX = minX, MaxX = maxX, MinY = minY, MaxY = maxY };
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// Get the combined bounding box of the primary element and all group elements
        /// in 2D view coordinates. Used for self-collision checking and L-shaped offset
        /// calculation to ensure the annotation header doesn't overlap any annotated element.
        /// </summary>
        private BBox2D GetCombinedElementViewBBox(Element primaryElement, List<long> groupElementIds)
        {
            var primaryBbox = GetElementViewBBox(primaryElement);
            DebugLogger.Log($"[COMBINED-BBOX] primary={primaryElement.Id.IntegerValue}, primaryBbox={(primaryBbox == null ? "null" : $"MinX={primaryBbox.MinX:F3} MaxX={primaryBbox.MaxX:F3}")}, groupIds=[{string.Join(",", groupElementIds ?? new List<long>())}]");
            
            if (groupElementIds == null || groupElementIds.Count == 0)
                return primaryBbox;
            
            // Start with primary element's bbox
            double? minX = primaryBbox?.MinX, maxX = primaryBbox?.MaxX;
            double? minY = primaryBbox?.MinY, maxY = primaryBbox?.MaxY;
            
            // Expand to include all group elements
            foreach (var elemId in groupElementIds)
            {
                if (elemId == primaryElement.Id.IntegerValue) continue; // skip primary (already included)
                var groupElem = _document.GetElement(new ElementId((int)elemId));
                if (groupElem == null)
                {
                    DebugLogger.Log($"[COMBINED-BBOX] elemId={elemId}: GetElement returned NULL");
                    continue;
                }
                var groupBbox = GetElementViewBBox(groupElem);
                if (groupBbox == null)
                {
                    DebugLogger.Log($"[COMBINED-BBOX] elemId={elemId}: GetElementViewBBox returned NULL");
                    continue;
                }
                DebugLogger.Log($"[COMBINED-BBOX] elemId={elemId}: bbox MinX={groupBbox.MinX:F3} MaxX={groupBbox.MaxX:F3}");
                
                if (minX == null || groupBbox.MinX < minX) minX = groupBbox.MinX;
                if (maxX == null || groupBbox.MaxX > maxX) maxX = groupBbox.MaxX;
                if (minY == null || groupBbox.MinY < minY) minY = groupBbox.MinY;
                if (maxY == null || groupBbox.MaxY > maxY) maxY = groupBbox.MaxY;
            }
            
            if (minX == null) return null;
            DebugLogger.Log($"[COMBINED-BBOX] RESULT: MinX={minX.Value:F3} MaxX={maxX.Value:F3}");
            return new BBox2D { MinX = minX.Value, MaxX = maxX.Value, MinY = minY.Value, MaxY = maxY.Value };
        }

        private XYZ GetElementLocation(Element element)
        {
            if (element.Location is LocationPoint lp)
                return lp.Point;
        
            if (element.Location is LocationCurve lc && lc.Curve != null)
                return lc.Curve.Evaluate(0.5, true);
        
            var bbox = element.get_BoundingBox(_view);
            if (bbox != null)
            {
                return new XYZ(
                    (bbox.Min.X + bbox.Max.X) / 2,
                    (bbox.Min.Y + bbox.Max.Y) / 2,
                    (bbox.Min.Z + bbox.Max.Z) / 2
                );
            }
        
            return null;
        }
        
        /// <summary>
        /// Get attachment points for an element.
        /// For ducts: 3 points (center, 25% from start, 25% from end)
        /// For other elements: 1 point (center only)
        /// </summary>
        private List<(XYZ Point, string Description)> GetAttachmentPoints(Element element, AnnotationPlan plan)
        {
            var points = new List<(XYZ Point, string Description)>();
            
            // Only ducts get multiple attachment points
            bool isDuct = plan.AnnotationType == AnnotationType.DuctRoundSizeFlow ||
                          plan.AnnotationType == AnnotationType.DuctRectSizeFlow;
            
            if (isDuct && element.Location is LocationCurve lc && lc.Curve != null)
            {
                var curve = lc.Curve;
                // Center point (primary)
                points.Add((curve.Evaluate(0.5, true), "Центр (50%)"));
                // 25% from start
                points.Add((curve.Evaluate(0.25, true), "Ближе к началу (25%)"));
                // 25% from end
                points.Add((curve.Evaluate(0.75, true), "Ближе к концу (75%)"));
            }
            else
            {
                // Single point (center)
                var center = GetElementLocation(element);
                if (center != null)
                    points.Add((center, "Центр"));
            }
            
            return points;
        }
        
        /// <summary>
        /// Calculate optimal offset based on element size and view scale.
        /// For ducts and duct accessories: fixed 1mm offset on paper.
        /// For equipment and air terminals: dynamic offset (1-4mm) based on element size.
        /// </summary>
        private double CalculateOptimalOffset(XYZ elementLocation, Element element, AnnotationType annotationType)
        {
            // For ducts and duct accessories: fixed 1mm offset (no size-based extra gap)
            if (annotationType == AnnotationType.DuctRoundSizeFlow ||
                annotationType == AnnotationType.DuctRectSizeFlow ||
                annotationType == AnnotationType.DuctAccessory)
            {
                double fixedOffsetPaperMm = 1.0;
                double viewScale = _view.Scale;
                if (viewScale < 1) viewScale = 1;
                double modelOffsetFeet = fixedOffsetPaperMm / 304.8 * viewScale;
                DebugLogger.Log($"[GREEDY-PLACEMENT] Offset (fixed for ducts): paperMm={fixedOffsetPaperMm:F1}, viewScale={viewScale}, modelFeet={modelOffsetFeet:F2}");
                return modelOffsetFeet;
            }
            
            // For equipment and air terminals: dynamic offset based on element size
            // Base offset in mm on paper (distance from element to NEAREST EDGE of tag)
            double baseOffsetPaperMm = 1.0; // 1mm on paper gap from element edge to tag edge
            
            // Add extra offset for larger elements (they need more clearance)
            var bbox = element.get_BoundingBox(_view);
            if (bbox != null)
            {
                double width = bbox.Max.X - bbox.Min.X;
                double height = bbox.Max.Y - bbox.Min.Y;
                double depth = bbox.Max.Z - bbox.Min.Z;
                double maxDimension = Math.Max(Math.Max(width, height), depth);
                
                // For large equipment (>0.5m), add a small extra gap on paper
                if (maxDimension > 0.5)
                {
                    baseOffsetPaperMm += Math.Min(maxDimension * 0.5, 3.0); // Up to 3mm extra on paper
                }
            }
            
            // Convert paper mm to model space feet
            double vs = _view.Scale;
            if (vs < 1) vs = 1;
            double modelOffset = baseOffsetPaperMm / 304.8 * vs;
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] Offset (dynamic for equipment/terminals): paperMm={baseOffsetPaperMm:F1}, viewScale={vs}, modelFeet={modelOffset:F2}");
            
            return Math.Min(modelOffset, 3.0);
        }
                
        /// <summary>
        /// Calculate optimal elbow height based on element position and view.
        /// Like the offset, this is calculated in paper space (mm) then converted to model space.
        /// Typical elbow height: 3-10mm on paper.
        /// </summary>
        private double CalculateOptimalElbowHeight(XYZ elementLocation, Element element, AnnotationPosition position)
        {
            // Fixed base elbow height in mm on paper (distance from element to shelf on printed sheet)
            // All elements start at 3mm; the elbow height increases through iteration if collisions are found.
            // Previously, large elements got up to 5mm extra based on bounding box size,
            // but this caused thin long ducts to start at 8mm unnecessarily.
            double baseHeightPaperMm = 3.0;
            
            // Convert paper mm to model space feet
            double viewScale = _view.Scale;
            if (viewScale < 1) viewScale = 1;
            double modelHeightFeet = baseHeightPaperMm / 304.8 * viewScale;
            
            return modelHeightFeet;
        }
        
        /// <summary>
        /// Calculate optimal Z offset for 3D views.
        /// Like offset/elbow height, this uses paper-space-based calculation.
        /// Typical Z offset: 5-10mm on paper from element.
        /// </summary>
        private double CalculateOptimalZOffset(XYZ elementLocation, Element element, AnnotationPosition position)
        {
            // For 3D views, we need to offset Z for visibility
            if (_view is View3D)
            {
                // Base Z offset in mm on paper
                double zOffsetPaperMm = 2.0; // Small Z offset (was 5mm - too large)
                
                // Get element bounding box
                var bbox = element.get_BoundingBox(_view);
                if (bbox != null)
                {
                    double depth = bbox.Max.Z - bbox.Min.Z;
                    double width = bbox.Max.X - bbox.Min.X;
                    double height = bbox.Max.Y - bbox.Min.Y;
                    double maxDimension = Math.Max(Math.Max(depth, width), height);
                    
                    // For large elements, add more paper-space Z offset
                    if (maxDimension > 0.5)
                    {
                        zOffsetPaperMm += Math.Min(maxDimension * 1.0, 5.0); // Up to 5mm extra on paper
                    }
                }
                
                // Convert paper mm to model space feet
                double viewScale = _view.Scale;
                if (viewScale < 1) viewScale = 1;
                double zOffset = zOffsetPaperMm / 304.8 * viewScale;
                
                // Adjust sign based on position
                bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
                if (isBottom)
                {
                    zOffset = -zOffset;
                }
                
                return zOffset;
            }
            
            // For 2D views, no Z offset needed
            return 0;
        }
        
        private BBox2D CalculatePlacementBbox(
            XYZ location,
            AnnotationPosition position,
            double elbowHeight,
            AnnotationSize size)
        {
            // Calculate head position (CENTER of tag)
            var headPos = CalculateHeadPosition(location, position, elbowHeight, size);
            
            // BBox is centered on head position (TagHeadPosition = CENTER of tag)
            return new BBox2D
            {
                MinX = headPos.X - size.TotalWidth / 2,
                MaxX = headPos.X + size.TotalWidth / 2,
                MinY = headPos.Y - size.TotalHeight / 2,
                MaxY = headPos.Y + size.TotalHeight / 2
            };
        }
        
        /// <summary>
        /// Calculate head position for L-shaped leader geometry.
        /// 
        /// IMPORTANT: TagHeadPosition in Revit is the END of the leader line, where
        /// the tag text is displayed. It is NOT the shelf edge where the leader connects.
        /// The leader line goes: LeaderEnd -> Elbow -> Head(TagHeadPosition).
        /// For Left positions: Head is to the LEFT of Elbow by (shelfWidth + shelfGap).
        /// For Right positions: Head is to the RIGHT of Elbow by (shelfWidth + shelfGap).
        /// 
        /// Visual layout (TopRight example):
        ///   Element -> LeaderEnd -> [up to Elbow] -> [right to ShelfLeftEdge] -> [shelfWidth] -> Head(Tip)
        ///                                                                                          ↑ TagHeadPosition
        /// 
        /// The leader horizontal segment goes from Elbow to TagHeadPosition (shelf edge).
        /// Leader horizontal length = shelfWidth + shelfGap (shelf extends from leader edge to tip)
        /// </summary>
        private (double X, double Y) CalculateHeadPosition(
            XYZ location,
            AnnotationPosition position,
            double elbowHeight,
            AnnotationSize size,
            double horizontalOffset)
        {
            // TagHeadPosition = point where leader meets the tag.
            // CRITICAL: Revit's tag shelf always extends to the RIGHT from TagHeadPosition.
            // - For LEFT-facing: head at LEFT TIP = leaderEnd - (shelfGap + shelfWidth)
            //   Shelf extends RIGHT from tip, covering full width. Near edge = head + shelfWidth.
            //   Gap = leaderEndViewX - nearEdge = shelfGap. ✓
            // - For RIGHT-facing: head at NEAR EDGE = leaderEnd + shelfGap
            //   Shelf extends RIGHT from near edge. Gap = headViewX - leaderEndViewX = shelfGap. ✓
            //
            // IMPORTANT: shelfGap is always 1mm on paper (the gap from leader to shelf edge).
            // For L-shaped leaders: the vertical segment clears the element, so only 1mm horizontal gap is needed.
            // For horizontal leaders: the full optimalHorizontalOffset is needed to clear the element body.
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
            bool isRight = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
            bool isLeft = position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft || position == AnnotationPosition.HorizontalLeft;
            
            // For L-shaped leaders: use fixed 1mm shelf gap (element is cleared by vertical segment)
            // For horizontal leaders: use full optimalHorizontalOffset (element must be cleared horizontally)
            double shelfGapModel;
            if (isHorizontal)
            {
                shelfGapModel = horizontalOffset; // full offset needed to clear element body
            }
            else
            {
                // Fixed 1mm shelf gap on paper, converted to model units
                double vs = _view.Scale;
                if (vs < 1) vs = 1;
                shelfGapModel = 1.0 / 304.8 * vs; // 1mm paper gap
            }
            double shelfWidthModel = size.Width;
            
            double headOffset;
            if (isLeft)
            {
                headOffset = shelfGapModel + shelfWidthModel; // tip position
            }
            else // isRight
            {
                headOffset = shelfGapModel; // near edge position
            }
            
            if (isHorizontal)
            {
                double horizontalDist = headOffset;
                return position switch
                {
                    AnnotationPosition.HorizontalLeft => (
                        location.X - horizontalDist,
                        location.Y
                    ),
                    AnnotationPosition.HorizontalRight => (
                        location.X + horizontalDist,
                        location.Y
                    ),
                    _ => (location.X, location.Y)
                };
            }
            else
            {
                // L-shaped (Top/Bottom): leader goes UP/DOWN then SIDEWAYS
                double headX;
                if (isRight)
                {
                    headX = location.X + headOffset;
                }
                else // isLeft
                {
                    headX = location.X - headOffset;
                }
                
                double headY = isTop ? location.Y + elbowHeight : location.Y - elbowHeight;
                return (headX, headY);
            }
        }
        
        // Overload for backward compatibility
        private (double X, double Y) CalculateHeadPosition(
            XYZ location,
            AnnotationPosition position,
            double elbowHeight,
            AnnotationSize size)
        {
            return CalculateHeadPosition(location, position, elbowHeight, size, _config.HorizontalOffset);
        }
        
        private BBox2D CalculatePlacementBboxWithOffset(
            XYZ location,
            AnnotationPosition position,
            double elbowHeight,
            AnnotationSize size,
            double horizontalOffset)
        {
            var headPos = CalculateHeadPosition(location, position, elbowHeight, size, horizontalOffset);
            
            // TagHeadPosition is now the shelf TIP (far end from leader connection)
            // The shelf extends from headPos TOWARD the element by shelfWidth
            // The leader connection is at headPos + shelfWidth (for Left) or headPos - shelfWidth (for Right)
            double shelfWidth = size.Width;  // shelf length in model feet
            double textHeight = size.Height; // text height in model feet
            double pad = size.Padding;
            
            bool isRight = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
            bool isLeft = position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft || position == AnnotationPosition.HorizontalLeft;
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
            
            var bbox = new BBox2D();
            
            // X extent: depends on direction because Revit's shelf always extends RIGHT from TagHeadPosition.
            // - For LEFT-facing: head is at LEFT TIP, shelf extends RIGHT toward element.
            //   BBox: from (headPos.X - pad) to (headPos.X + shelfWidth + pad)
            // - For RIGHT-facing: head is at NEAR EDGE (left side of shelf), shelf extends RIGHT away from element.
            //   BBox: from (headPos.X - pad) to (headPos.X + shelfWidth + pad)
            // Both cases have the same bbox formula since shelf extends RIGHT from head in both cases!
            bbox.MinX = headPos.X - pad;               // left edge (head position - padding)
            bbox.MaxX = headPos.X + shelfWidth + pad;   // right edge (head + shelf + padding)
            
            // Y extent: text is centered on the shelf line (headPos.Y)
            // For 2-line tags, text extends both above and below the shelf
            bbox.MinY = headPos.Y - textHeight / 2 - pad;
            bbox.MaxY = headPos.Y + textHeight / 2 + pad;
            
            return bbox;
        }
                
        /// <summary>
        /// Calculate elbow position in view coordinates for leader collision detection
        /// </summary>
        private (double X, double Y) CalculateElbowPositionFromViewCoords(
            double leaderEndX, double leaderEndY,
            double headX, double headY,
            AnnotationPosition position)
        {
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
        
            if (isTop || isBottom)
            {
                // For top/bottom: vertical segment from leader end, then horizontal to head
                // Elbow is at (leaderEndX, headY)
                return (leaderEndX, headY);
            }
            else if (isHorizontal)
            {
                // For horizontal positions: create a straight horizontal line
                // Elbow is at midpoint between leaderEnd and head (both have same Y)
                // This creates a "degenerate" L-shape that appears as a straight line
                double midX = (leaderEndX + headX) / 2.0;
                return (midX, leaderEndY);  // Y same as leaderEnd (straight line)
            }
        
            // Default
            return (leaderEndX, headY);
        }
                
        /// <summary>
        /// Element orientation classification based on projected geometry on view.
        /// </summary>
        private enum ElementOrientation
        {
            /// <summary>\ Diagonal: left edge is higher than right edge (slopes down to the right)</summary>
            DiagonalLeftHigher,
            /// <summary>// Diagonal: right edge is higher than left edge (slopes down to the left)</summary>
            DiagonalRightHigher,
            /// <summary>-- Horizontal: element is wider than tall</summary>
            Horizontal,
            /// <summary>|| Vertical: element is taller than wide</summary>
            Vertical
        }
        
        /// <summary>
        /// Determine element orientation from view coordinates (ViewStart/ViewEnd).
        /// This is more accurate than using bounding box projection because it uses
        /// the actual element direction as seen on the view.
        /// User's algorithm:
        /// \ (x1<x2, y1>y2): Left edge higher -> slopes down to the right
        /// // (x1<x2, y1<y2): Right edge higher -> slopes up to the right
        /// -- (x1<x2, y1~=y2): Horizontal
        /// || (x1~=x2, y1<y2): Vertical
        /// </summary>
        private ElementOrientation DetermineElementOrientationFromViewCoords(
            double x1, double y1, double x2, double y2)
        {
            // Ensure x1 < x2 (swap if needed)
            if (x1 > x2)
            {
                double tmp = x1; x1 = x2; x2 = tmp;
                tmp = y1; y1 = y2; y2 = tmp;
            }
            
            double dx = x2 - x1;
            double dy = y2 - y1;
            
            // Threshold for horizontal/vertical classification
            const double threshold = 0.1;
            
            // -- Horizontal: y1 ~= y2
            if (Math.Abs(dy) < threshold * Math.Max(dx, 0.01))
                return ElementOrientation.Horizontal;
            
            // || Vertical: x1 ~= x2
            if (Math.Abs(dx) < threshold * Math.Max(Math.Abs(dy), 0.01))
                return ElementOrientation.Vertical;
            
            // \ Diagonal left higher: x1<x2 and y1>y2 (slopes down to the right)
            if (y1 > y2)
                return ElementOrientation.DiagonalLeftHigher;
            
            // // Diagonal right higher: x1<x2 and y1<y2 (slopes up to the right)
            return ElementOrientation.DiagonalRightHigher;
        }
        
        /// <summary>
        /// Determine orientation for a point element (accessory, air terminal, equipment)
        /// by finding connected ducts in the system graph and inheriting their orientation.
        /// If the element is not part of a system or has no connected ducts, returns Horizontal (fallback).
        /// </summary>
        private ElementOrientation DetermineOrientationFromConnectedDucts(AnnotationPlan plan)
        {
            try
            {
                if (plan.Node == null || plan.Node.ConnectedNodeIds == null || plan.Node.ConnectedNodeIds.Count == 0)
                    return ElementOrientation.Horizontal;
                
                // Look through connected nodes for ducts (linear elements with endpoints)
                foreach (var connectedId in plan.Node.ConnectedNodeIds)
                {
                    try
                    {
                        var connectedElement = _document.GetElement(new ElementId(connectedId));
                        if (connectedElement == null) continue;
                        
                        // Check if it's a duct or pipe (linear element)
                        if (connectedElement is Duct || connectedElement is Pipe)
                        {
                            // Get the duct's direction from its LocationCurve
                            if (connectedElement.Location is LocationCurve lc && lc.Curve != null)
                            {
                                var start = lc.Curve.GetEndPoint(0);
                                var end = lc.Curve.GetEndPoint(1);
                                
                                // Convert to view coordinates
                                var (viewX1, viewY1) = ConvertModelToViewCoordinates(start);
                                var (viewX2, viewY2) = ConvertModelToViewCoordinates(end);
                                
                                // Determine orientation using same logic as for ducts
                                var ductOrientation = DetermineElementOrientationFromViewCoords(
                                    viewX1, viewY1, viewX2, viewY2);
                                
                                DebugLogger.Log($"[GREEDY-PLACEMENT] Element {plan.ElementId}: inherited orientation {ductOrientation} from connected duct {connectedId}");
                                return ductOrientation;
                            }
                        }
                    }
                    catch { continue; }
                }
            }
            catch { }
            
            return ElementOrientation.Horizontal; // No connected duct found
        }
        
        /// <summary>
        /// Get position order from a pre-determined orientation.
        /// </summary>
        private List<AnnotationPosition> GetOptimalPositionOrderFromOrientation(
            List<AnnotationPosition> preferredPositions,
            ElementOrientation orientation)
        {
            if (preferredPositions == null || preferredPositions.Count == 0)
                return new List<AnnotationPosition> { AnnotationPosition.TopRight };

            // Get position priority based on orientation
            var positionPriority = GetPositionPriorityForOrientation(orientation);
        
            // Filter to only include positions from preferred list
            var result = positionPriority
                .Where(p => preferredPositions.Contains(p))
                .ToList();
                    
            // Add any remaining preferred positions not in our priority list
            foreach (var p in preferredPositions)
            {
                if (!result.Contains(p))
                    result.Add(p);
            }
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] Position order: orientation={orientation}, order={string.Join(", ", result)}");
        
            return result;
        }
        
        /// <summary>
        /// Determine the element's orientation based on its bounding box projection on the view.
        /// IMPORTANT: For 3D views, we must project the model-space bounding box onto the view plane
        /// using the view's RightDirection and UpDirection vectors, because model X/Y don't correspond
        /// to screen X/Y in a 3D view.
        /// </summary>
        private ElementOrientation DetermineElementOrientation(BoundingBoxXYZ bbox)
        {
            if (bbox == null)
                return ElementOrientation.Horizontal; // Default
            
            // Get view projection vectors
            var viewRight = _view.RightDirection;
            var viewUp = _view.UpDirection;
            
            // Project bounding box corners onto view plane to get screen-space extents
            // The 8 corners of the model-space bounding box
            var corners = new[]
            {
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
            };
            
            // Project each corner to view coordinates (screen X, Y)
            double minViewX = double.MaxValue, maxViewX = double.MinValue;
            double minViewY = double.MaxValue, maxViewY = double.MinValue;
            
            foreach (var corner in corners)
            {
                double viewX = corner.DotProduct(viewRight);
                double viewY = corner.DotProduct(viewUp);
                minViewX = Math.Min(minViewX, viewX);
                maxViewX = Math.Max(maxViewX, viewX);
                minViewY = Math.Min(minViewY, viewY);
                maxViewY = Math.Max(maxViewY, viewY);
            }
            
            // Width and height in VIEW coordinates (what you see on screen)
            double width = maxViewX - minViewX;
            double height = maxViewY - minViewY;
            
            // Threshold for aspect ratio classification
            const double aspectRatioThreshold = 1.5;
            
            // Check if horizontal (width >> height on screen)
            if (width > height * aspectRatioThreshold)
            {
                return ElementOrientation.Horizontal;
            }
            
            // Check if vertical (height >> width on screen)
            if (height > width * aspectRatioThreshold)
            {
                return ElementOrientation.Vertical;
            }
            
            // For diagonal elements, determine which edge is higher on screen
            // Compare the top-left vs top-right corner projections
            // "Left edge higher" means the left side of the element (in screen coords) is at a higher Y
            double topLeftY = corners.Max(c => c.DotProduct(viewUp) - 0.1 * Math.Abs(c.DotProduct(viewRight) - minViewX)); // heuristic
            
            // Simple approach: check which side of the screen-projected bbox is higher
            // The min-X side vs max-X side projected Y
            double leftSideMaxY = corners.Where(c => c.DotProduct(viewRight) < (minViewX + maxViewX) / 2)
                                         .Select(c => c.DotProduct(viewUp))
                                         .DefaultIfEmpty(0).Max();
            double rightSideMaxY = corners.Where(c => c.DotProduct(viewRight) >= (minViewX + maxViewX) / 2)
                                          .Select(c => c.DotProduct(viewUp))
                                          .DefaultIfEmpty(0).Max();
            
            // Also check min Y on each side for the bottom edge
            double leftSideMinY = corners.Where(c => c.DotProduct(viewRight) < (minViewX + maxViewX) / 2)
                                         .Select(c => c.DotProduct(viewUp))
                                         .DefaultIfEmpty(0).Min();
            double rightSideMinY = corners.Where(c => c.DotProduct(viewRight) >= (minViewX + maxViewX) / 2)
                                          .Select(c => c.DotProduct(viewUp))
                                          .DefaultIfEmpty(0).Min();
            
            // \ Left edge higher: left side has higher max Y OR lower min Y than right side
            // This means the element slopes down to the right
            if (leftSideMaxY > rightSideMaxY + 0.01 || leftSideMinY > rightSideMinY + 0.01)
            {
                return ElementOrientation.DiagonalLeftHigher;
            }
            
            // // Right edge higher: right side is higher
            if (rightSideMaxY > leftSideMaxY + 0.01 || rightSideMinY > leftSideMinY + 0.01)
            {
                return ElementOrientation.DiagonalRightHigher;
            }
            
            // Default for ambiguous cases
            return ElementOrientation.Horizontal;
        }
        
        /// <summary>
        /// Get the position priority list for a given element orientation.
        /// Based on the user's priority algorithm specification.
        /// </summary>
        private List<AnnotationPosition> GetPositionPriorityForOrientation(ElementOrientation orientation)
        {
            var priority = new List<AnnotationPosition>();
            
            switch (orientation)
            {
                case ElementOrientation.DiagonalLeftHigher:
                    // \ Object: left edge higher than right edge
                    // Priority: TopRight > HorizontalRight > HorizontalLeft > BottomLeft > TopLeft > BottomRight
                    // TopLeft and BottomRight only as last resort
                    priority.Add(AnnotationPosition.TopRight);
                    priority.Add(AnnotationPosition.HorizontalRight);
                    priority.Add(AnnotationPosition.HorizontalLeft);
                    priority.Add(AnnotationPosition.BottomLeft);
                    priority.Add(AnnotationPosition.TopLeft);       // Last resort
                    priority.Add(AnnotationPosition.BottomRight);    // Last resort
                    break;
                    
                case ElementOrientation.DiagonalRightHigher:
                    // // Object: right edge higher than left edge
                    // Priority: TopLeft > HorizontalLeft > HorizontalRight > BottomRight > TopRight > BottomLeft
                    // TopRight and BottomLeft only as last resort
                    priority.Add(AnnotationPosition.TopLeft);
                    priority.Add(AnnotationPosition.HorizontalLeft);
                    priority.Add(AnnotationPosition.HorizontalRight);
                    priority.Add(AnnotationPosition.BottomRight);
                    priority.Add(AnnotationPosition.TopRight);      // Last resort
                    priority.Add(AnnotationPosition.BottomLeft);    // Last resort
                    break;
                    
                case ElementOrientation.Horizontal:
                    // -- Object: horizontal orientation
                    // Priority: TopLeft > TopRight > BottomLeft > BottomRight > HorizontalLeft > HorizontalRight
                    // HorizontalLeft/HorizontalRight only when no other elements on that side
                    priority.Add(AnnotationPosition.TopLeft);
                    priority.Add(AnnotationPosition.TopRight);
                    priority.Add(AnnotationPosition.BottomLeft);
                    priority.Add(AnnotationPosition.BottomRight);
                    priority.Add(AnnotationPosition.HorizontalLeft);   // Only if clear on left side
                    priority.Add(AnnotationPosition.HorizontalRight);  // Only if clear on right side
                    break;
                    
                case ElementOrientation.Vertical:
                    // || Object: vertical orientation
                    // Priority: HorizontalLeft > HorizontalRight > TopLeft > BottomLeft > TopRight > BottomRight
                    priority.Add(AnnotationPosition.HorizontalLeft);
                    priority.Add(AnnotationPosition.HorizontalRight);
                    priority.Add(AnnotationPosition.TopLeft);
                    priority.Add(AnnotationPosition.BottomLeft);
                    priority.Add(AnnotationPosition.TopRight);
                    priority.Add(AnnotationPosition.BottomRight);
                    break;
            }
            
            // Add center positions as fallbacks
            priority.Add(AnnotationPosition.TopCenter);
            priority.Add(AnnotationPosition.BottomCenter);
            
            return priority;
        }
                
        /// <summary>
        /// Get view extents (min/max X and Y) in model coordinates
        /// </summary>
        private Tuple<double, double, double, double> GetViewExtents()
        {
            try
            {
                var bbox = _view.get_BoundingBox(_view);
                if (bbox != null)
                {
                    return Tuple.Create(bbox.Min.X, bbox.Max.X, bbox.Min.Y, bbox.Max.Y);
                }
            }
            catch { }
                            
            return null;
        }
                
        /// <summary>
        /// Calculate adjusted head position based on actual tag bounding box.
        /// TagHeadPosition is the CENTER of the tag, so we adjust based on
        /// the actual tag width to ensure the nearest edge is at the correct offset.
        /// </summary>
        private XYZ CalculateAdjustedHeadPosition(
            XYZ leaderEndPos,
            XYZ originalHeadPos,
            AnnotationPosition position,
            double actualWidth,
            double actualHeight,
            double actualCenterX,
            double actualCenterY)
        {
            try
            {
                double adjustedX = originalHeadPos.X;
                double adjustedY = originalHeadPos.Y;
                double adjustedZ = originalHeadPos.Z;
                
                bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
                bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
                bool isLeft = position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft || position == AnnotationPosition.HorizontalLeft;
                bool isRight = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
                bool isLShaped = isTop || isBottom;
                
                if (isLShaped)
                {
                    // L-shaped: shelf edge aligns near leader end X (with horizontalOffset gap)
                    // Head center = leader end X ± (horizontalOffset + actualWidth/2)
                    // Note: horizontalOffset is now used consistently with CalculateHeadPosition
                    double edgeGap = 1.0 / 304.8 * _view.Scale; // 1mm paper gap (consistent with base offset)
                    if (isRight)
                    {
                        adjustedX = leaderEndPos.X + edgeGap + actualWidth / 2;
                    }
                    else // isLeft
                    {
                        adjustedX = leaderEndPos.X - edgeGap - actualWidth / 2;
                    }
                }
                else
                {
                    // Horizontal: use small gap (1mm) between element and annotation edge
                    if (isLeft)
                    {
                        adjustedX = leaderEndPos.X - 1.0 / 304.8 * _view.Scale - actualWidth / 2;
                    }
                    else if (isRight)
                    {
                        adjustedX = leaderEndPos.X + 1.0 / 304.8 * _view.Scale + actualWidth / 2;
                    }
                }
                
                return new XYZ(adjustedX, adjustedY, adjustedZ);
            }
            catch
            {
                return null;
            }
        }
                
        /// <summary>
        /// Calculate adjusted head position for 3D views to ensure proper L-shape on screen.
        /// In 3D views, the leader must form an L-shape in VIEW space (on screen), not in model space.
        /// 
        /// For Top/Bottom positions:
        /// - Screen X(elbow) = Screen X(leaderEnd) for vertical leader on screen
        /// - Screen Y(elbow) = Screen Y(head) for horizontal segment to head
        /// </summary>
        private XYZ CalculateAdjustedHeadPositionFor3DView(
            XYZ leaderEndPos,
            XYZ originalHeadPos,
            AnnotationPosition position,
            double annotationWidth,
            double annotationHeight)
        {
            try
            {
                // Get view projection vectors
                var viewOrigin = _view.Origin;
                var viewRight = _view.RightDirection;
                var viewUp = _view.UpDirection;
                var viewDir = _view.ViewDirection;
                        
                // Project leader end to screen coordinates
                var leaderEndRel = leaderEndPos - viewOrigin;
                double leaderEndViewX = leaderEndRel.DotProduct(viewRight);
                double leaderEndViewY = leaderEndRel.DotProduct(viewUp);
                double leaderEndDepth = leaderEndRel.DotProduct(viewDir);
                        
                // Project original head to screen coordinates
                var headRel = originalHeadPos - viewOrigin;
                double headViewX = headRel.DotProduct(viewRight);
                double headViewY = headRel.DotProduct(viewUp);
                double headDepth = headRel.DotProduct(viewDir);
                        
                // Calculate target head position in SCREEN space
                // For Top/Bottom positions, the leader should form an L-shape on screen
                // The elbow should be at (leaderEndViewX, headViewY) on screen
                        
                bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
                bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
                        
                double targetHeadViewX = headViewX;
                double targetHeadViewY = headViewY;
                double targetHeadDepth = headDepth;
                        
                if (isTop || isBottom)
                {
                    // For L-shaped leaders:
                    // - The horizontal segment goes from elbow to head
                    // - The elbow screen X = leader end screen X
                    // - The elbow screen Y = head screen Y
                    // - So we need to adjust head position to ensure proper L-shape
                            
                    // annotationWidth is already in model feet (paper-space mm * viewScale / 304.8)
                    // In view coordinates, model feet project directly, so use annotationWidth as offset
                    double offsetViewX = annotationWidth / 2; // Half-width offset from leader to head center
                            
                    if (position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft)
                    {
                        // Head should be to the left of leader on screen
                        targetHeadViewX = leaderEndViewX - offsetViewX;
                    }
                    else // TopRight, BottomRight
                    {
                        // Head should be to the right of leader on screen
                        targetHeadViewX = leaderEndViewX + offsetViewX;
                    }
                            
                    // Keep the same screen Y as original (maintains vertical position)
                    targetHeadViewY = headViewY;
                            
                    DebugLogger.Log($"[GREEDY-PLACEMENT] 3D L-shape adjustment: leaderViewX={leaderEndViewX:F3}, targetHeadViewX={targetHeadViewX:F3}");
                }
                        
                // Convert target screen position back to model space
                XYZ adjustedHeadModel = viewOrigin +
                    targetHeadViewX * viewRight +
                    targetHeadViewY * viewUp +
                    targetHeadDepth * viewDir;
                        
                return adjustedHeadModel;
            }
            catch
            {
                return null;
            }
        }
                
        private IndependentTag CreateAnnotation(
            AnnotationPlan plan,
            Element element,
            XYZ location,
            AnnotationPosition position,
            double elbowHeight,
            double horizontalOffset,
            double lShapedOffset = 0)
        {
            try
            {
                // Find the tag type
                var tagType = FindTagType(plan);
                if (tagType == null)
                {
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Tag type not found for plan: {plan.AnnotationType}");
                    return null;
                }
                
                // Activate the tag type if not already active
                if (!tagType.IsActive)
                {
                    using (var activateTx = new Transaction(_document, "Activate Tag Type"))
                    {
                        activateTx.Start();
                        tagType.Activate();
                        activateTx.Commit();
                    }
                }
                
                // Get reference for the element
                Reference elemRef = new Reference(element);
                
                // Calculate head position with dynamic offset
                AnnotationSize size;
                if (!_sizes.TryGetValue(plan, out size))
                    size = new AnnotationSize();
                var headPos = CalculateHeadPosition(location, position, elbowHeight, size, horizontalOffset);
                
                // Calculate Z offset for 3D views - this is critical for proper placement
                double zOffset = CalculateOptimalZOffset(location, element, position);
                double tagZ = location.Z + zOffset;
                
                // For 3D views with HORIZONTAL annotations, we need special handling:
                // The head must project to the same screen Y as the leader end for a straight line.
                // This requires adjusting the head position based on view projection.
                bool is3DView = _view is View3D;
                bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
                                
                DebugLogger.Log($"[GREEDY-PLACEMENT] CreateAnnotation: position={position}, is3DView={is3DView}, isHorizontal={isHorizontal}");
                                
                XYZ tagHeadPosition;
                if (is3DView)
                {
                    // For ALL annotations in 3D views, calculate head position in VIEW space
                    // then convert back to model space. This ensures correct positioning on screen
                    // regardless of the 3D view rotation/tilt.
                                        bool isCenterOriginTagFor3D = plan.AnnotationType == AnnotationType.DuctAccessory ||
                                                         plan.AnnotationType == AnnotationType.EquipmentMark;
                                        bool isRightEdgeOriginFor3D = plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                                                         plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow;
tagHeadPosition = CalculateHeadPositionFor3DView(
                        location, tagZ, horizontalOffset, size, position, elbowHeight, lShapedOffset, isCenterOriginTagFor3D, isRightEdgeOriginFor3D);
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Using CalculateHeadPositionFor3DView: headModel=({tagHeadPosition.X:F3}, {tagHeadPosition.Y:F3}, {tagHeadPosition.Z:F3})");
                }
                else
                {
                    // For 2D views: model coordinates = view coordinates
                    tagHeadPosition = new XYZ(headPos.X, headPos.Y, tagZ);
                }
                
                // Create tag with proper leader
                IndependentTag tag = IndependentTag.Create(
                    _document,
                    tagType.Id,
                    _view.Id,
                    elemRef,
                    true,  // Add leader
                    TagOrientation.Horizontal,
                    tagHeadPosition
                );
                
                if (tag != null)
                {
                    // Get actual tag bounding box to determine text width
                    BoundingBoxXYZ actualBbox = null;
                    try
                    {
                        actualBbox = tag.get_BoundingBox(_view);
                    }
                    catch { }
                                    
                    // CRITICAL: In 3D views, bounding box returns MODEL coordinates, not annotation paper space.
                    // For 3D views, we estimate text width from the tag text content.
                    // is3DView is already defined above
                    
                    // Try to get a tag type with appropriate shelf length
                    if (tagType != null)
                    {
                        double effectiveWidth;
                        
                        if (is3DView)
                        {
                            // 3D view: estimate width from tag text content
                            // IMPORTANT: TagText returns ALL lines concatenated WITHOUT separators.
                            // For multi-line tags, we must estimate the longest line width separately.
                            // Use AnnotationType to determine if the tag is multi-line.
                            string tagText = tag.TagText ?? "";
                            
                            // For round/rect duct tags, text after the 2nd dot is not rendered
                            // and should be excluded from width estimation.
                            // Example: "ø100.0L0.0000" → "ø100.0L0" (keep up to 2nd dot)
                            if (plan.AnnotationType == AnnotationType.DuctRoundSizeFlow ||
                                plan.AnnotationType == AnnotationType.DuctRectSizeFlow)
                            {
                                int dotCount = 0;
                                int truncateAt = tagText.Length;
                                for (int ci = 0; ci < tagText.Length; ci++)
                                {
                                    if (tagText[ci] == '.')
                                    {
                                        dotCount++;
                                        if (dotCount == 2) { truncateAt = ci; break; }
                                    }
                                }
                                if (truncateAt < tagText.Length)
                                {
                                    string originalText = tagText;
                                    tagText = tagText.Substring(0, truncateAt);
                                    DebugLogger.Log($"[GREEDY-PLACEMENT] Truncated tag text from \"{originalText}\" to \"{tagText}\" (removed part after 2nd dot)");
                                }
                            }
                            int maxLineLength;
                            bool isMultiLine = plan.AnnotationType == AnnotationType.DuctRoundSizeFlow ||
                                               plan.AnnotationType == AnnotationType.DuctRectSizeFlow ||
                                               plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                                               plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow;
                            
                            if (isMultiLine && tagText.Length > 0)
                            {
                                // Multi-line tag: TagText is concatenated without separators.
                                // We need to estimate the width of each line separately.
                                // For duct tags: format is like "ø100.0L35.0000" where:
                                //   Line 1 (size): starts with ø or digits, ends before "L" (flow indicator)
                                //   Line 2 (flow): starts with "L" followed by digits
                                // For air terminal tags: format is like "Ecoline 2  100" where:
                                //   Line 1 (name): text part
                                //   Line 2 (flow): numeric part
                                // Strategy: split at common boundary patterns
                                int line1Len = EstimateFirstLineLength(tagText, plan.AnnotationType);
                                int line2Len = tagText.Length - line1Len;
                                
                                // Truncate trailing zeros after decimal point for width estimation.
                                // Revit displays "100.0" as "100" and "0.0000" as "0" on the tag,
                                // so the rendered text is shorter than the raw TagText.
                                // Example: "ø100.0" → "ø100" (4 chars vs 6), "L0.0000" → "L0" (2 chars vs 7)
                                string line1Text = tagText.Substring(0, line1Len);
                                string line2Text = tagText.Substring(line1Len);
                                string line1Truncated = TruncateTrailingZeros(line1Text);
                                string line2Truncated = TruncateTrailingZeros(line2Text);
                                maxLineLength = Math.Max(line1Truncated.Length, line2Truncated.Length);
                                DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view: multi-line tag \"{tagText}\" split: line1=\"{line1Text}\"→\"{line1Truncated}\"({line1Truncated.Length}), line2=\"{line2Text}\"→\"{line2Truncated}\"({line2Truncated.Length}), maxLine={maxLineLength}");
                            }
                            else
                            {
                                // Single-line tag or TagText has line breaks
                                string[] lines = tagText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                // Apply trailing zero truncation to each line
                                var truncatedLines = lines.Select(l => TruncateTrailingZeros(l)).ToList();
                                maxLineLength = truncatedLines.Count > 0 ? truncatedLines.Max(l => l.Length) : tagText.Length;
                            }
                            
                            // Estimate width: 1.5mm per character for Cyrillic/bold text
                            // (1.0mm was too small - caused shelf to be shorter than actual text)
                            double estimatedWidthMm = maxLineLength * 1.5;
                            // Convert to model feet INCLUDING view scale
                            // This is consistent with how 2D views provide model-space bounding box width
                            effectiveWidth = estimatedWidthMm / 304.8 * _view.Scale;
                            DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view: estimated text width from \"{tagText}\" (longest line: {maxLineLength} chars): {estimatedWidthMm:F1}mm, effectiveWidth={effectiveWidth:F4}ft");
                        }
                        else if (actualBbox != null)
                        {
                            // 2D view: use bounding box width
                            effectiveWidth = actualBbox.Max.X - actualBbox.Min.X;
                        }
                        else
                        {
                            effectiveWidth = 0.15; // Default fallback
                        }
                                        
                        // Get or create a type with appropriate shelf length
                        var optimizedType = _tagTypeManager.GetOrCreateTagTypeWithShelfLength(
                            tagType, effectiveWidth, _view.Scale);
                                        
                        // If we got a different type, apply it
                        if (optimizedType != null && optimizedType.Id != tagType.Id)
                        {
                            try
                            {
                                tag.ChangeTypeId(optimizedType.Id);
                                DebugLogger.Log($"[GREEDY-PLACEMENT] Changed tag type from '{tagType.Name}' to '{optimizedType.Name}' for better shelf length");
                                                
                                // Re-get bounding box after type change
                                try
                                {
                                    actualBbox = tag.get_BoundingBox(_view);
                                }
                                catch { }
                            }
                            catch (Exception typeEx)
                            {
                                DebugLogger.Log($"[GREEDY-PLACEMENT] Could not change tag type: {typeEx.Message}");
                            }
                        }
                    }
                                    
                    // Set leader end condition to free for custom positioning
                    tag.LeaderEndCondition = LeaderEndCondition.Free;
                    
                    // Get the tagged reference for leader configuration
                    var taggedRefs = tag.GetTaggedReferences();
                    Reference tagRef = (taggedRefs != null && taggedRefs.Count > 0) ? taggedRefs[0] : elemRef;
                                    
                    // Calculate adjusted head position based on actual tag size
                    // For 3D views, use paper-space-based annotation sizes
                    // Use SMALLER defaults that reflect actual text width, not the full annotation box
                    // Actual annotation width is typically the text width + 2mm padding
                    const double defaultAnnotationWidthPaperMm = 15.0;   // 15mm default annotation width on paper
                    const double defaultAnnotationHeightPaperMm = 6.0;    // 6mm default annotation height on paper (2 lines of text)
                    double currentViewScale = _view.Scale;
                    if (currentViewScale < 1) currentViewScale = 1;
                    double defaultAnnotationWidthFeet = defaultAnnotationWidthPaperMm / 304.8 * currentViewScale;
                    double defaultAnnotationHeightFeet = defaultAnnotationHeightPaperMm / 304.8 * currentViewScale;
                                    
                    if (actualBbox != null && !is3DView)
                    {
                        // 2D view: use actual bounding box (annotation paper space coordinates)
                        double actualWidth = actualBbox.Max.X - actualBbox.Min.X;
                        double actualHeight = actualBbox.Max.Y - actualBbox.Min.Y;
                        double actualCenterX = (actualBbox.Min.X + actualBbox.Max.X) / 2;
                        double actualCenterY = (actualBbox.Min.Y + actualBbox.Max.Y) / 2;
                                        
                        // Adjust head position so leader connects at the edge
                        XYZ adjustedHeadPos = CalculateAdjustedHeadPosition(
                            location, tagHeadPosition, position, 
                            actualWidth, actualHeight, actualCenterX, actualCenterY);
                                        
                        if (adjustedHeadPos != null)
                        {
                            tagHeadPosition = adjustedHeadPos;
                            // Don't set TagHeadPosition here - set it LAST after leader positions
                            // to avoid Revit overriding our position
                            DebugLogger.Log($"[GREEDY-PLACEMENT] Adjusted head position: actualWidth={actualWidth:F4}, actualHeight={actualHeight:F4}");
                        }
                    }
                    else if (is3DView)
                    {
                        // 3D view: After tag type change, log the actual shelf length.
                        // We do NOT recalculate head position here — the head was correctly
                        // positioned BEFORE the type change, and the shelf length difference
                        // is handled by the annotation's inherent design. Recalculating from
                        // scratch would use a wrong gap (1mm default instead of bbox-cleared offset).
                        double? actualShelfMm = TagTypeManager.GetShelfLengthFromTypeNamePublic(tag.Name);
                        if (actualShelfMm.HasValue && actualShelfMm.Value > 0)
                        {
                            DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view: head position unchanged after type change, actualShelf={actualShelfMm.Value:F0}mm");
                        }
                    }
                    
                    // Set leader end position on the element FIRST
                    // Per user's algorithm: LeaderEnd -> Header -> Elbow
                    try
                    {
                        tag.SetLeaderEnd(tagRef, location);
                    }
                    catch (Exception endEx)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Could not set leader end: {endEx.Message}");
                    }
                    
                    // STEP 2: Set head position (TagHeadPosition) BEFORE setting elbow
                    // This is how Revit expects it: set where the shelf should be, then connect the leader
                    try
                    {
                        tag.TagHeadPosition = tagHeadPosition;
                    }
                    catch { }
                    
                    // STEP 3: Set elbow position to form a proper L-shape on screen
                    // The elbow must be at (leaderEndViewX, headViewY) in screen space
                    if (plan.HasElbow)
                    {
                        try
                        {
                            // Read actual head position after our set (Revit may have adjusted it)
                            var actualHeadPos = tag.TagHeadPosition;
                            if (actualHeadPos != null)
                            {
                                tagHeadPosition = actualHeadPos;
                            }
                            
                            XYZ elbowPos = CalculateElbowPosition(location, tagHeadPosition, position, elbowHeight, size);
                            tag.SetLeaderElbow(tagRef, elbowPos);
                            
                            // After setting elbow, Revit may adjust head position again.
                            // Re-read and correct if needed.
                            var postElbowHead = tag.TagHeadPosition;
                            if (postElbowHead != null)
                            {
                                double postDrift = Math.Sqrt(
                                    Math.Pow(postElbowHead.X - tagHeadPosition.X, 2) +
                                    Math.Pow(postElbowHead.Y - tagHeadPosition.Y, 2));
                                
                                if (postDrift > 0.05) // More than ~15mm drift
                                {
                                    DebugLogger.Log($"[GREEDY-PLACEMENT] Head drift after elbow: {postDrift:F3}ft, re-setting");
                                    try { tag.TagHeadPosition = tagHeadPosition; } catch { }
                                    
                                    // Re-calculate elbow for the corrected head position
                                    try
                                    {
                                        elbowPos = CalculateElbowPosition(location, tagHeadPosition, position, elbowHeight, size);
                                        tag.SetLeaderElbow(tagRef, elbowPos);
                                    }
                                    catch { }
                                }
                            }
                            
                            // Log the L-shaped geometry for debugging
                            DebugLogger.Log("[GREEDY-PLACEMENT] L-shape geometry for element " + plan.ElementId + ":");
                            DebugLogger.Log("[GREEDY-PLACEMENT]   LeaderEnd: (" + location.X.ToString("F3") + ", " + location.Y.ToString("F3") + ")");
                            DebugLogger.Log("[GREEDY-PLACEMENT]   ElbowPosition: (" + elbowPos.X.ToString("F3") + ", " + elbowPos.Y.ToString("F3") + ")");
                            DebugLogger.Log("[GREEDY-PLACEMENT]   HeadPosition: (" + tagHeadPosition.X.ToString("F3") + ", " + tagHeadPosition.Y.ToString("F3") + ")");
                        }
                        catch (Exception elbowEx)
                        {
                            DebugLogger.Log("[GREEDY-PLACEMENT] Could not set elbow: " + elbowEx.Message);
                        }
                    }
                                        
                    // Final head position verification and logging
                    var finalHeadPos = tag.TagHeadPosition;
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Set final TagHeadPosition: ({tagHeadPosition.X:F3}, {tagHeadPosition.Y:F3}, {tagHeadPosition.Z:F3})");
                    if (finalHeadPos != null)
                    {
                        double finalError = Math.Sqrt(Math.Pow(finalHeadPos.X - tagHeadPosition.X, 2) + Math.Pow(finalHeadPos.Y - tagHeadPosition.Y, 2));
                        if (finalError > 0.05)
                        {
                            DebugLogger.Log($"[GREEDY-PLACEMENT] WARNING: Final head drift = {finalError:F3}ft, actual=({finalHeadPos.X:F3},{finalHeadPos.Y:F3})");
                        }
                    }
                    
                    DebugLogger.Log("[GREEDY-PLACEMENT] Created tag " + tag.Id + " for element " + plan.ElementId);
                }
                
                return tag;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GREEDY-PLACEMENT] Error creating annotation: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Calculate head position for ALL annotations in 3D views.
        /// Works entirely in VIEW space (screen coordinates) then converts back to model space.
        /// This ensures correct positioning regardless of 3D view rotation/tilt.
        /// 
        /// For horizontal annotations: headViewY = leaderEndViewY (straight line on screen)
        /// For L-shaped annotations: headViewY = leaderEndViewY +/- elbowHeight (in view coords)
        /// </summary>
        private XYZ CalculateHeadPositionFor3DView(
            XYZ leaderEndPosition,
            double headZ,
            double horizontalOffset,
            AnnotationSize size,
            AnnotationPosition position,
            double elbowHeight,
            double lShapedOffset = 0,
            bool isCenterOriginTag = false,
            bool isRightEdgeOrigin = false)
        {
            // Get view projection vectors
            var viewOrigin = _view.Origin;
            var viewRight = _view.RightDirection;
            var viewUp = _view.UpDirection;
            var viewDir = _view.ViewDirection;
            
            // Project leader end to view coordinates
            var leaderEndRel = leaderEndPosition - viewOrigin;
            double leaderEndViewX = leaderEndRel.DotProduct(viewRight);
            double leaderEndViewY = leaderEndRel.DotProduct(viewUp);
            double leaderEndDepth = leaderEndRel.DotProduct(viewDir);
            
            // Calculate target head position in VIEW space
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
            
            double headViewX;
            double headViewY;
            
            // Project shelf dimensions from model space to view coordinates.
            // Paper-space dimensions map directly to view coordinate offsets.
            // The view coordinate system uses the same units as model space (feet).
            // Horizontal dimensions (shelf width, gap) → viewX, vertical (height) → viewY.
            //
            // IMPORTANT: shelfGap depends on the leader type:
            // For L-shaped leaders: the horizontal offset must clear the element's bbox.
            //   If the element is wide, 1mm gap is not enough - use horizontalOffset instead.
            // For horizontal leaders: the full horizontalOffset is needed to clear the element body.
            double viewShelfWidth, viewShelfGap, viewShelfHeight;
            viewShelfWidth = size.Width;           // horizontal on screen (viewX)
            viewShelfHeight = size.Height;         // vertical on screen (viewY)
            
            double vs = _view.Scale;
            if (vs < 1) vs = 1;
            double oneMmView = 1.0 / 304.8 * vs;   // 1mm in view coordinates
            
            if (isHorizontal)
            {
                viewShelfGap = horizontalOffset;   // full offset needed to clear element body
            }
            else
            {
                // For L-shaped leaders: use the lShapedOffset if provided (when the element bbox
                // requires a larger offset to clear), otherwise use the default 1mm gap.
                if (lShapedOffset > oneMmView + 0.001) // lShapedOffset was increased to clear element bbox
                {
                    viewShelfGap = lShapedOffset;
                }
                else
                {
                    viewShelfGap = Math.Max(oneMmView, horizontalOffset);
                }
            }
            
            if (viewShelfWidth < 0.02) viewShelfWidth = 0.02;
            if (viewShelfGap < 0.005) viewShelfGap = 0.005;
            if (viewShelfHeight < 0.02) viewShelfHeight = 0.02;
            
            bool isRight = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
            bool isLeft = position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft || position == AnnotationPosition.HorizontalLeft;
            
            // Tag families have different reference point (TagHeadPosition) positions:
            // - RIGHT-EDGE origin (air terminals): TagHeadPosition = RIGHT edge of shelf
            //   RIGHT-facing: headOffset = gap + shelfWidth (head at RIGHT TIP = far from element)
            //   LEFT-facing: headOffset = gap (head at RIGHT EDGE = near element)
            // - LEFT-EDGE origin (duct marks): TagHeadPosition = LEFT edge of shelf
            //   LEFT-facing: headOffset = gap + shelfWidth (head at LEFT TIP = far from element)
            //   RIGHT-facing: headOffset = gap (head at LEFT EDGE = near element)
            // - CENTER origin (duct accessories, equipment): TagHeadPosition = CENTER of tag
            //   Both directions: headOffset = gap + shelfWidth/2 (head at CENTER)
            double viewShelfHalfWidth = viewShelfWidth / 2.0;
            
            double headOffsetView;
            if (isCenterOriginTag)
            {
                // CENTER origin: head at CENTER for both directions
                headOffsetView = viewShelfGap + viewShelfHalfWidth;
            }
            else
            {
                // Non-center: head at FAR TIP when facing the origin edge, NEAR EDGE otherwise.
                // RIGHT-EDGE RIGHT-facing OR LEFT-EDGE LEFT-facing → head at FAR TIP (gap + shelfWidth)
                // RIGHT-EDGE LEFT-facing OR LEFT-EDGE RIGHT-facing → head at NEAR EDGE (gap only)
                headOffsetView = viewShelfGap + ((isRight == isRightEdgeOrigin) ? viewShelfWidth : 0);
            }
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view offsets (paper-space direct): viewShelfWidth={viewShelfWidth:F3}ft, viewShelfGap={viewShelfGap:F3}ft, viewShelfHeight={viewShelfHeight:F3}ft, direction={((isLeft)?"LEFT":"RIGHT")}, headOffset={headOffsetView:F3}ft");
            
            if (isHorizontal)
            {
                // Horizontal: straight line on screen, headViewY = leaderEndViewY
                headViewY = leaderEndViewY;
                
                if (isLeft)
                {
                    headViewX = leaderEndViewX - headOffsetView;
                }
                else // isRight
                {
                    headViewX = leaderEndViewX + headOffsetView;
                }
            }
            else
            {
                // L-shaped (Top/Bottom): leader goes UP/DOWN then SIDEWAYS
                if (isTop)
                {
                    headViewY = leaderEndViewY + elbowHeight;
                }
                else // Bottom
                {
                    headViewY = leaderEndViewY - elbowHeight;
                }
                
                if (isRight)
                {
                    headViewX = leaderEndViewX + headOffsetView;
                }
                else // isLeft
                {
                    headViewX = leaderEndViewX - headOffsetView;
                }
            }
            
            // Use the same depth as leader end to maintain projection consistency
            double headDepth = leaderEndDepth;
            
            // Convert back to model space
            XYZ headModelPos = viewOrigin +
                headViewX * viewRight +
                headViewY * viewUp +
                headDepth * viewDir;
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] 3D head calculation (position={position}):" );
            DebugLogger.Log($"[GREEDY-PLACEMENT]   LeaderEnd screen: ({leaderEndViewX:F3}, {leaderEndViewY:F3}), depth={leaderEndDepth:F3}");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Target head screen: ({headViewX:F3}, {headViewY:F3}), depth={headDepth:F3}");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Head model calculated: ({headModelPos.X:F3}, {headModelPos.Y:F3}, {headModelPos.Z:F3})");
            
            return headModelPos;
        }
        
        /// <summary>
        /// Estimate the character length of the first line in a multi-line tag.
        /// TagText returns all lines concatenated without separators.
        /// We need to determine where the first line ends to estimate shelf width correctly.
        /// </summary>
        private int EstimateFirstLineLength(string tagText, AnnotationType annotationType)
        {
            if (string.IsNullOrEmpty(tagText))
                return 0;
            
            // For duct tags (DuctRoundSizeFlow, DuctRectSizeFlow):
            // Format: "ø100.0L35.0000" or "200x100L35.0000"
            // Line 1 is the size (e.g., "ø100.0" or "200x100")
            // Line 2 is the flow (e.g., "L35.0000")
            // The boundary is where "L" (flow indicator) starts after the size part
            if (annotationType == AnnotationType.DuctRoundSizeFlow ||
                annotationType == AnnotationType.DuctRectSizeFlow)
            {
                // Look for 'L' preceded by a digit (not the first character)
                // This is the flow line indicator: e.g., "...0L35" means Line 2 starts at 'L'
                for (int i = 1; i < tagText.Length; i++)
                {
                    if (tagText[i] == 'L' && i > 0 && char.IsDigit(tagText[i - 1]))
                    {
                        return i; // Line 1 is everything before this 'L'
                    }
                }
                // Fallback: if no 'L' found, assume single line
                return tagText.Length;
            }
            
            // For air terminal tags (AirTerminalTypeFlow, AirTerminalShortNameFlow):
            // Format: "Ecoline 2  100" or "VS 100M  50" or "L35.0000VS 100M"
            // Line 1 is the type name, Line 2 is the flow value
            // The boundary is typically after the name - look for multiple spaces or a numeric-only suffix
            if (annotationType == AnnotationType.AirTerminalTypeFlow ||
                annotationType == AnnotationType.AirTerminalShortNameFlow)
            {
                // Special case: TagText starts with "L" followed by digits (flow line comes first)
                // Example: "L35.0000VS 100M" → Line 1 = "L35.0000", Line 2 = "VS 100M"
                if (tagText.Length > 1 && tagText[0] == 'L' && char.IsDigit(tagText[1]))
                {
                    // Find where the flow value ends: after digits and optional decimal part
                    int flowEnd = 1;
                    bool seenDot = false;
                    while (flowEnd < tagText.Length)
                    {
                        if (char.IsDigit(tagText[flowEnd]))
                        {
                            flowEnd++;
                        }
                        else if (tagText[flowEnd] == '.' && !seenDot)
                        {
                            seenDot = true;
                            flowEnd++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    // flowEnd is now the start of the name part
                    // But we want line1 to be the FIRST line (above shelf)
                    // and line2 to be the SECOND line (below shelf).
                    // In Revit tag families, the order in TagText is: line1 + line2.
                    // So if flow comes first in TagText, it means flow=line1, name=line2.
                    // We return flowEnd as the line1 length.
                    return flowEnd;
                }
                
                // Look for double space (common separator in Revit tag text)
                int doubleSpace = tagText.IndexOf("  ");
                if (doubleSpace > 0)
                    return doubleSpace;
                
                // Fallback: look for transition from letters to digits
                // E.g., "Ecoline 2100" - the "100" at the end is the flow
                // Find the last letter/dot before a trailing numeric sequence
                for (int i = tagText.Length - 2; i >= 1; i--)
                {
                    if (char.IsLetter(tagText[i]) && char.IsDigit(tagText[i + 1]))
                    {
                        // Check if the rest is mostly digits (flow value)
                        bool restIsNumeric = true;
                        for (int j = i + 1; j < tagText.Length; j++)
                        {
                            if (!char.IsDigit(tagText[j]) && tagText[j] != '.' && tagText[j] != ',' && tagText[j] != '-')
                            {
                                restIsNumeric = false;
                                break;
                            }
                        }
                        if (restIsNumeric)
                            return i + 1;
                    }
                }
                
                return tagText.Length;
            }
            
            // Default: assume single line
            return tagText.Length;
        }
        
        /// <summary>
        /// Truncate trailing zeros after a decimal point in a text segment.
        /// Revit displays numeric values without unnecessary trailing zeros,
        /// so "100.0" renders as "100" and "0.0000" renders as "0".
        /// We must account for this when estimating shelf length from TagText.
        /// 
        /// Examples:
        ///   "ø100.0" → "ø100"  (trailing ".0" removed)
        ///   "L0.0000" → "L0"   (trailing ".0000" removed)
        ///   "L35.0000" → "L35" (trailing ".0000" removed)
        ///   "VS 100M" → "VS 100M" (no decimal point, unchanged)
        /// </summary>
        private string TruncateTrailingZeros(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            
            // Process each numeric segment that contains a decimal point
            // We use regex to find patterns like digits.digits and truncate trailing zeros
            var result = new System.Text.StringBuilder(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                // Check if we're at the start of a number (digit, or dot preceded by digit)
                if (char.IsDigit(text[i]))
                {
                    // Find the end of the numeric segment: digits, optional dot, more digits
                    int numStart = i;
                    bool hasDot = false;
                    int dotPos = -1;
                    while (i < text.Length && (char.IsDigit(text[i]) || (text[i] == '.' && !hasDot)))
                    {
                        if (text[i] == '.')
                        {
                            hasDot = true;
                            dotPos = i;
                        }
                        i++;
                    }
                    
                    if (hasDot && dotPos >= 0)
                    {
                        // Extract the part after the dot and truncate trailing zeros
                        string beforeDot = text.Substring(numStart, dotPos - numStart);
                        string afterDot = text.Substring(dotPos + 1, i - dotPos - 1);
                        string afterDotTrimmed = afterDot.TrimEnd('0');
                        
                        if (afterDotTrimmed.Length == 0)
                        {
                            // All digits after dot were zeros: "100.0" → "100"
                            result.Append(beforeDot);
                        }
                        else
                        {
                            // Some non-zero digits after dot: "100.50" → "100.5"
                            result.Append(beforeDot);
                            result.Append('.');
                            result.Append(afterDotTrimmed);
                        }
                    }
                    else
                    {
                        // No decimal point, keep as-is
                        result.Append(text, numStart, i - numStart);
                    }
                }
                else
                {
                    result.Append(text[i]);
                    i++;
                }
            }
            
            return result.ToString();
        }
        
        /// <summary>
        /// Calculate elbow position for proper 90-degree leader angle (L-shaped leader)
        /// 
        /// IMPORTANT FOR 3D VIEWS:
        /// The leader must form an L-shape in VIEW space (on screen), not in MODEL space.
        /// In 3D views, the projection from model to view coordinates distorts angles.
        /// 
        /// For a proper right angle in the VIEW:
        /// 1. Convert LeaderEnd and HeadPosition to view coordinates
        /// 2. Calculate elbow position in VIEW space (same viewX as LeaderEnd, same viewY as Head)
        /// 3. Project back to MODEL space using the view's projection plane
        /// 
        /// Geometry in VIEW space:
        ///     HeadPosition ──────── ElbowPosition
        ///                               │
        ///                               │ (vertical in view)
        ///                               │
        ///                          LeaderEnd (on element)
        /// </summary>
        private XYZ CalculateElbowPosition(XYZ leaderEndPosition, XYZ headPosition, AnnotationPosition position, double elbowHeight, AnnotationSize size = null)
        {
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
        
            // For 3D views, calculate elbow position based on VIEW projection
            if (_view is View3D view3D)
            {
                var effectiveSize = size ?? new AnnotationSize();
                return CalculateElbowPositionFor3DView(leaderEndPosition, headPosition, position, view3D, effectiveSize);
            }
            
            // For 2D views, use model coordinates directly (they match view coordinates)
            double elbowZ = leaderEndPosition.Z;
        
            if (isTop || isBottom)
            {
                return new XYZ(
                    leaderEndPosition.X,  // Elbow X matches LeaderEnd X (vertical leader)
                    headPosition.Y,       // Elbow Y matches HeadPosition Y (horizontal to head)
                    elbowZ
                );
            }
            else if (isHorizontal)
            {
                // For horizontal positions: create a straight horizontal line
                // Elbow is at midpoint between leaderEnd and head (both have same Y)
                // This creates a "degenerate" L-shape that appears as a straight line
                double midX = (leaderEndPosition.X + headPosition.X) / 2.0;
                return new XYZ(
                    midX,                 // Elbow at midpoint X
                    leaderEndPosition.Y,  // Elbow Y matches LeaderEnd Y (straight line)
                    elbowZ
                );
            }
        
            // Default: L-shape with vertical leader
            return new XYZ(
                leaderEndPosition.X,
                headPosition.Y,
                elbowZ
            );
        }
        
        /// <summary>
        /// Calculate elbow position for 3D views using VIEW space projection.
        /// In 3D views, the leader must form a right angle on SCREEN, not in model space.
        /// 
        /// CRITICAL: We use the SAME projection method as ViewDataCollector.ConvertToViewCoordinates
        /// to ensure consistency between how positions are recorded and calculated.
        /// 
        /// KEY INSIGHT: The elbow must project to (leaderEndViewX, headViewY) in screen space.
        /// This ensures the leader forms a perfect L-shape when viewed.
        /// The depth value determines where the elbow sits in 3D space, but doesn't affect
        /// the screen projection of the right angle.
        /// </summary>
        private XYZ CalculateElbowPositionFor3DView(XYZ leaderEndPosition, XYZ headPosition, AnnotationPosition position, View3D view3D, AnnotationSize size)
        {
            // Get view properties for projection (same method as ConvertToViewCoordinates)
            var viewOrigin = _view.Origin;
            var viewRight = _view.RightDirection;
            var viewUp = _view.UpDirection;
            var viewDir = _view.ViewDirection;
            
            // Project LeaderEnd and HeadPosition to view coordinates (screen space)
            // This uses the SAME projection as ConvertToViewCoordinates for consistency
            var leaderEndRel = leaderEndPosition - viewOrigin;
            var headRel = headPosition - viewOrigin;
            
            // Screen coordinates (viewX = horizontal on screen, viewY = vertical on screen)
            double leaderEndViewX = leaderEndRel.DotProduct(viewRight);
            double leaderEndViewY = leaderEndRel.DotProduct(viewUp);
            double leaderEndDepth = leaderEndRel.DotProduct(viewDir);  // Depth from view plane
            
            double headViewX = headRel.DotProduct(viewRight);
            double headViewY = headRel.DotProduct(viewUp);
            double headDepth = headRel.DotProduct(viewDir);
            
            // Calculate elbow position in VIEW space (screen coordinates)
            // For L-shaped annotations: elbowViewX = leaderEndViewX (vertical leader on screen)
            // The shelf edge automatically aligns with the elbow because:
            //   HeadViewX = leaderEndViewX + shelfHalfWidth (right-attached)
            //   Left shelf edge = HeadViewX - shelfHalfWidth = leaderEndViewX = elbowViewX ✓
            // So elbow is naturally at the shelf edge - no clamping needed.
            
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
            
            double elbowViewX, elbowViewY, elbowDepth;
            
            if (isTop || isBottom)
            {
                // For L-shaped: elbow at leader end X (vertical leader), head Y (horizontal to head)
                elbowViewX = leaderEndViewX;  // Same screen X as leader end (vertical leader on screen)
                elbowViewY = headViewY;        // Same screen Y as head (horizontal segment on screen)
                elbowDepth = headDepth;         // Same depth as head for proper horizontal segment
            }
            else if (isHorizontal)
            {
                // For horizontal positions: create a straight horizontal line on screen
                elbowViewX = (leaderEndViewX + headViewX) / 2.0;  // Midpoint X
                elbowViewY = leaderEndViewY;                        // Same Y as leader end
                elbowDepth = headDepth;                              // Same depth as head
                
                DebugLogger.Log($"[GREEDY-PLACEMENT] Horizontal annotation: leaderEndViewY={leaderEndViewY:F3}, headViewY={headViewY:F3}");
                if (Math.Abs(leaderEndViewY - headViewY) > 0.5)
                {
                    DebugLogger.Log($"[GREEDY-PLACEMENT] WARNING: Horizontal line mismatch! Y difference={Math.Abs(leaderEndViewY - headViewY):F3}");
                }
            }
            else
            {
                // Default: L-shape with vertical leader on screen
                elbowViewX = leaderEndViewX;
                elbowViewY = headViewY;
                elbowDepth = headDepth;
            }
            
            // Convert elbow position from VIEW space back to MODEL space
            // We need to find a 3D point that projects to (elbowViewX, elbowViewY, elbowDepth)
            // The formula is: modelPoint = viewOrigin + viewX*viewRight + viewY*viewUp + depth*viewDir
            XYZ elbowModelPos = viewOrigin + 
                elbowViewX * viewRight + 
                elbowViewY * viewUp + 
                elbowDepth * viewDir;
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view elbow calculation (using head depth for proper L-shape):");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   LeaderEnd model: ({leaderEndPosition.X:F3}, {leaderEndPosition.Y:F3}, {leaderEndPosition.Z:F3})");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   LeaderEnd screen: ({leaderEndViewX:F3}, {leaderEndViewY:F3}), depth={leaderEndDepth:F3}");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Head model: ({headPosition.X:F3}, {headPosition.Y:F3}, {headPosition.Z:F3})");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Head screen: ({headViewX:F3}, {headViewY:F3}), depth={headDepth:F3}");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Elbow screen target: ({elbowViewX:F3}, {elbowViewY:F3}), depth={elbowDepth:F3}");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Elbow model calculated: ({elbowModelPos.X:F3}, {elbowModelPos.Y:F3}, {elbowModelPos.Z:F3})");
            
            // Verify the projection (for debugging)
            var elbowRel = elbowModelPos - viewOrigin;
            double verifyViewX = elbowRel.DotProduct(viewRight);
            double verifyViewY = elbowRel.DotProduct(viewUp);
            double verifyDepth = elbowRel.DotProduct(viewDir);
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Verification - Elbow projects to screen: ({verifyViewX:F3}, {verifyViewY:F3}), depth={verifyDepth:F3}");
            DebugLogger.Log($"[GREEDY-PLACEMENT]   Expected screen match: X={elbowViewX:F3} (should be {leaderEndViewX:F3}), Y={elbowViewY:F3} (should be {headViewY:F3})");
            
            return elbowModelPos;
        }
        
        private FamilySymbol FindTagType(AnnotationPlan plan)
        {
            // Map annotation types to categories
            // Note: OST_AirTerminalTags doesn't exist in Revit API, so we search by family name
            var categories = GetCategoriesForAnnotationType(plan.AnnotationType);
            
            foreach (var category in categories)
            {
                var collector = new FilteredElementCollector(_document)
                    .OfCategory(category)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>();
                
                // PRIORITY 1: Find exact match by BOTH FamilyName AND TypeName
                if (!string.IsNullOrEmpty(plan.FamilyName) && !string.IsNullOrEmpty(plan.TypeName))
                {
                    // Try exact FamilyName + TypeName match first
                    var exactMatch = collector.FirstOrDefault(s =>
                        s.FamilyName.Equals(plan.FamilyName, StringComparison.OrdinalIgnoreCase) &&
                        s.Name.Equals(plan.TypeName, StringComparison.OrdinalIgnoreCase));
                    
                    if (exactMatch != null)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found exact Family+Type match: {exactMatch.FamilyName} - {exactMatch.Name}");
                        return exactMatch;
                    }
                    
                    // Try with flexible TypeName (allow different numeric suffix)
                    var baseTypeName = GetBaseTypeName(plan.TypeName);
                    if (!string.IsNullOrEmpty(baseTypeName))
                    {
                        var flexTypeMatch = collector.FirstOrDefault(s =>
                            s.FamilyName.Equals(plan.FamilyName, StringComparison.OrdinalIgnoreCase) &&
                            GetBaseTypeName(s.Name).Equals(baseTypeName, StringComparison.OrdinalIgnoreCase));
                        
                        if (flexTypeMatch != null)
                        {
                            DebugLogger.Log($"[GREEDY-PLACEMENT] Found flexible Type match: {flexTypeMatch.FamilyName} - {flexTypeMatch.Name} (requested: {plan.TypeName})");
                            return flexTypeMatch;
                        }
                    }
                }
                
                // PRIORITY 2: Find by FamilyName only (fallback within same family)
                if (!string.IsNullOrEmpty(plan.FamilyName))
                {
                    // Try exact family name match
                    var familyMatch = collector.FirstOrDefault(s =>
                        s.FamilyName.Equals(plan.FamilyName, StringComparison.OrdinalIgnoreCase));
                    
                    if (familyMatch != null)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found family match (fallback): {familyMatch.FamilyName} - {familyMatch.Name}");
                        return familyMatch;
                    }
                    
                    // Then try matching with ADSK_ prefix
                    var adskeMatch = collector.FirstOrDefault(s =>
                        s.FamilyName.Equals("ADSK_" + plan.FamilyName, StringComparison.OrdinalIgnoreCase) ||
                        plan.FamilyName.Equals("ADSK_" + s.FamilyName, StringComparison.OrdinalIgnoreCase));
                    
                    if (adskeMatch != null)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found ADSK family match: {adskeMatch.FamilyName} - {adskeMatch.Name}");
                        return adskeMatch;
                    }
                    
                    // Then try partial match (family name contains or is contained by)
                    var partialMatch = collector.FirstOrDefault(s =>
                        s.FamilyName.Contains(plan.FamilyName) ||
                        plan.FamilyName.Contains(s.FamilyName));
                    
                    if (partialMatch != null)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found partial family match: {partialMatch.FamilyName} - {partialMatch.Name}");
                        return partialMatch;
                    }
                }
                
                // PRIORITY 3: Find by TypeName only (if FamilyName not specified)
                if (!string.IsNullOrEmpty(plan.TypeName) && string.IsNullOrEmpty(plan.FamilyName))
                {
                    // Try exact match first
                    var typeMatch = collector.FirstOrDefault(s =>
                        s.Name.Equals(plan.TypeName, StringComparison.OrdinalIgnoreCase));
                    
                    if (typeMatch != null)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found exact type match: {typeMatch.FamilyName} - {typeMatch.Name}");
                        return typeMatch;
                    }
                    
                    // Try matching base name (without number suffix)
                    var baseTypeName = GetBaseTypeName(plan.TypeName);
                    if (!string.IsNullOrEmpty(baseTypeName))
                    {
                        typeMatch = collector.FirstOrDefault(s =>
                        {
                            var symbolBaseName = GetBaseTypeName(s.Name);
                            return !string.IsNullOrEmpty(symbolBaseName) &&
                                   symbolBaseName.Equals(baseTypeName, StringComparison.OrdinalIgnoreCase);
                        });
                        
                        if (typeMatch != null)
                        {
                            DebugLogger.Log($"[GREEDY-PLACEMENT] Found base type match: {typeMatch.FamilyName} - {typeMatch.Name} (looking for {plan.TypeName})");
                            return typeMatch;
                        }
                    }
                    
                    // Try contains match
                    typeMatch = collector.FirstOrDefault(s => s.Name.Contains(plan.TypeName));
                    if (typeMatch != null)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found contains type match: {typeMatch.FamilyName} - {typeMatch.Name}");
                        return typeMatch;
                    }
                }
            }
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] No match found for annotation type {plan.AnnotationType}");
            
            // Try to find default tag type using TagTypeManager
            var defaultType = _tagTypeManager.FindDefaultTagType(plan, _view);
            if (defaultType != null)
            {
                return defaultType;
            }
            
            // LAST RESORT: For air terminals, try to find ANY suitable tag
            if (plan.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                plan.AnnotationType == AnnotationType.AirTerminalShortNameFlow)
            {
                DebugLogger.Log($"[GREEDY-PLACEMENT] Searching for air terminal tag fallback...");
                
                var allTags = new List<FamilySymbol>();
                
                // Try multiple tag categories that might work for air terminals
                var categoriesToTry = new[] 
                { 
                    BuiltInCategory.OST_MultiCategoryTags,      // Multi-category tags can tag air terminals
                    BuiltInCategory.OST_GenericModelTags,        // Generic model tags
                    BuiltInCategory.OST_MechanicalEquipmentTags, // Mechanical equipment tags
                    BuiltInCategory.OST_DuctTags                 // Duct tags as last resort
                };
                
                foreach (var cat in categoriesToTry)
                {
                    try
                    {
                        var tagsInCategory = new FilteredElementCollector(_document)
                            .OfCategory(cat)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .ToList();
                        allTags.AddRange(tagsInCategory);
                        DebugLogger.Log($"[GREEDY-PLACEMENT]   Found {tagsInCategory.Count} tags in category {cat}");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT]   Could not search category {cat}: {ex.Message}");
                    }
                }
                
                DebugLogger.Log($"[GREEDY-PLACEMENT]   Total tags to search: {allTags.Count}");
                
                // Look for tags with keywords related to air terminals
                // Include both Russian and English keywords for ADSK families
                var airTerminalKeywords = new[] { 
                    "воздухораспределитель", "airterminal", "air terminal", "diffuser", "диффузор", 
                    "решетка", "grille", "наименование", "краткое", "расход", "flow",
                    "приток", "вытяжка", "supply", "exhaust"
                };
                
                foreach (var tag in allTags)
                {
                    var familyNameLower = tag.FamilyName.ToLowerInvariant();
                    var typeNameLower = tag.Name.ToLowerInvariant();
                    
                    if (airTerminalKeywords.Any(k => familyNameLower.Contains(k) || typeNameLower.Contains(k)))
                    {
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Found air terminal tag by keyword: {tag.FamilyName} - {tag.Name}");
                        return tag;
                    }
                }
                
                // Try to find a multi-category tag with generic annotation
                var multiCatTag = allTags.FirstOrDefault(t => 
                    t.FamilyName.Contains("Несколько категорий") || 
                    t.FamilyName.ToLowerInvariant().Contains("multi") ||
                    t.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_MultiCategoryTags);
                
                if (multiCatTag != null)
                {
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Using multi-category tag for air terminal: {multiCatTag.FamilyName} - {multiCatTag.Name}");
                    return multiCatTag;
                }
                
                // Ultimate fallback: use any available equipment tag
                var equipmentTag = allTags.FirstOrDefault(t => 
                    t.FamilyName.Contains("Оборудование") || 
                    t.FamilyName.ToLowerInvariant().Contains("equipment"));
                
                if (equipmentTag != null)
                {
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Using equipment tag as fallback for air terminal: {equipmentTag.FamilyName} - {equipmentTag.Name}");
                    return equipmentTag;
                }
                
                // Last resort: any tag
                var fallbackTag = allTags.FirstOrDefault();
                if (fallbackTag != null)
                {
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Using any available tag as last resort for air terminal: {fallbackTag.FamilyName} - {fallbackTag.Name}");
                    return fallbackTag;
                }
                
                DebugLogger.Log($"[GREEDY-PLACEMENT] No fallback tag found for air terminal");
            }
            
            return null;
        }
        
        /// <summary>
        /// Get base type name without numeric suffix or shelf length in parentheses.
        /// e.g. "ADSK_Марка_15" -> "ADSK_Марка"
        /// e.g. "Круглый воздуховод_Размер и расход (14)" -> "Круглый воздуховод_Размер и расход"
        /// </summary>
        private string GetBaseTypeName(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return typeName;
                    
            // First try to remove shelf length in parentheses at the end
            // Pattern: "Name (number)" or "Name(number)"
            var parenMatch = System.Text.RegularExpressions.Regex.Match(typeName, @"^(.+?)\s*\(\d+(?:\.\d+)?\)\s*$");
            if (parenMatch.Success)
            {
                return parenMatch.Groups[1].Value.Trim();
            }
                    
            // Then try to remove trailing underscore followed by numbers
            var lastUnderscore = typeName.LastIndexOf('_');
            if (lastUnderscore > 0)
            {
                var suffix = typeName.Substring(lastUnderscore + 1);
                if (int.TryParse(suffix, out _))
                {
                    return typeName.Substring(0, lastUnderscore);
                }
            }
                    
            return typeName;
        }
        
        /// <summary>
        /// Get categories to search for each annotation type
        /// </summary>
        private List<BuiltInCategory> GetCategoriesForAnnotationType(AnnotationType type)
        {
            return type switch
            {
                AnnotationType.DuctRoundSizeFlow => new List<BuiltInCategory> { BuiltInCategory.OST_DuctTags },
                AnnotationType.DuctRectSizeFlow => new List<BuiltInCategory> { BuiltInCategory.OST_DuctTags },
                // Air terminal tags - search multiple categories since OST_AirTerminalTags doesn't exist
                // Air terminals can be tagged with GenericModel tags, MechanicalEquipment tags,
                // MultiCategory tags, DuctAccessory tags, or even Duct tags
                AnnotationType.AirTerminalTypeFlow => new List<BuiltInCategory> 
                { 
                    BuiltInCategory.OST_DuctTerminalTags,         // Марки воздухораспределителей (если существует)
                    BuiltInCategory.OST_DuctAccessoryTags,        // Марки арматуры воздуховодов
                    BuiltInCategory.OST_MultiCategoryTags,
                    BuiltInCategory.OST_GenericModelTags,
                    BuiltInCategory.OST_MechanicalEquipmentTags,
                    BuiltInCategory.OST_DuctTags
                },
                AnnotationType.AirTerminalShortNameFlow => new List<BuiltInCategory> 
                { 
                    BuiltInCategory.OST_DuctTerminalTags,         // Марки воздухораспределителей (если существует)
                    BuiltInCategory.OST_DuctAccessoryTags,        // Марки арматуры воздуховодов
                    BuiltInCategory.OST_MultiCategoryTags,
                    BuiltInCategory.OST_GenericModelTags,
                    BuiltInCategory.OST_MechanicalEquipmentTags,
                    BuiltInCategory.OST_DuctTags
                },
                AnnotationType.DuctAccessory => new List<BuiltInCategory> { BuiltInCategory.OST_DuctAccessoryTags },
                AnnotationType.EquipmentMark => new List<BuiltInCategory> { BuiltInCategory.OST_MechanicalEquipmentTags },
                _ => new List<BuiltInCategory> { BuiltInCategory.OST_DuctTags }
            };
        }
        
        private BBox2D GetTagBoundingBox(IndependentTag tag)
        {
            try
            {
                var bbox = tag.get_BoundingBox(_view);
                if (bbox != null)
                {
                    return new BBox2D
                    {
                        MinX = bbox.Min.X,
                        MaxX = bbox.Max.X,
                        MinY = bbox.Min.Y,
                        MaxY = bbox.Max.Y
                    };
                }
            }
            catch { }
            
            return null;
        }
        
        /// <summary>
        /// Register annotation leader lines for collision detection.
        /// Per the user's algorithm: after an annotation is placed successfully,
        /// its leader lines must be added to the collision check list so that
        /// subsequent annotations can avoid overlapping with this one.
        /// Lines are registered in VIEW coordinates (same as project element segments).
        /// </summary>
        private void RegisterAnnotationLinesForCollision(IndependentTag tag, long elementId)
        {
            try
            {
                var taggedRefs = tag.GetTaggedReferences();
                if (taggedRefs == null || taggedRefs.Count == 0) return;
                
                var tagRef = taggedRefs[0];
                
                // Get actual positions from the tag
                var leaderEndPos = tag.GetLeaderEnd(tagRef);
                var elbowPos = tag.GetLeaderElbow(tagRef);
                var headPos = tag.TagHeadPosition;
                
                if (leaderEndPos == null || elbowPos == null || headPos == null) return;
                
                if (_view is View3D)
                {
                    // 3D view: convert all positions to view coordinates
                    var leaderEndView = ConvertModelToViewCoordinates(leaderEndPos);
                    var elbowView = ConvertModelToViewCoordinates(elbowPos);
                    var headView = ConvertModelToViewCoordinates(headPos);
                    
                    _collisionDetector.RegisterAnnotationLines(
                        leaderEndView.X, leaderEndView.Y,
                        elbowView.X, elbowView.Y,
                        headView.X, headView.Y,
                        elementId);
                    
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Registered annotation lines (3D): leaderEnd=({leaderEndView.X:F2}, {leaderEndView.Y:F2}), elbow=({elbowView.X:F2}, {elbowView.Y:F2}), head=({headView.X:F2}, {headView.Y:F2})");
                }
                else
                {
                    // 2D view: model coordinates = view coordinates
                    _collisionDetector.RegisterAnnotationLines(
                        leaderEndPos.X, leaderEndPos.Y,
                        elbowPos.X, elbowPos.Y,
                        headPos.X, headPos.Y,
                        elementId);
                    
                    DebugLogger.Log($"[GREEDY-PLACEMENT] Registered annotation lines (2D): leaderEnd=({leaderEndPos.X:F2}, {leaderEndPos.Y:F2}), elbow=({elbowPos.X:F2}, {elbowPos.Y:F2}), head=({headPos.X:F2}, {headPos.Y:F2})");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GREEDY-PLACEMENT] Failed to register annotation lines: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Find groups of nearby same-type elements that should share one annotation.
        /// Currently applies to Air Terminals and Duct Accessories.
        /// Elements are considered "nearby" if they are within a certain radius on the projected view.
        /// </summary>
        private List<ElementGroup> FindNearbySameTypeGroups(List<AnnotationPlan> plans)
        {
            var groups = new List<ElementGroup>();
            
            // Only group certain annotation types
            var groupableTypes = new[]
            {
                AnnotationType.AirTerminalTypeFlow,
                AnnotationType.AirTerminalShortNameFlow,
                AnnotationType.DuctAccessory
            };
            
            var groupablePlans = plans
                .Where(p => groupableTypes.Contains(p.AnnotationType))
                .ToList();
            
            if (groupablePlans.Count < 2)
                return groups;
            
            // Collect element data for grouping: view coordinates, model Z, Расход воздуха
            var elemDataMap = new Dictionary<long, (double viewX, double viewY, double modelZ, string airflow)>();
            foreach (var p in groupablePlans)
            {
                var elem = _document.GetElement(new ElementId(p.ElementId));
                if (elem == null) continue;
                var loc = GetElementLocation(elem);
                if (loc == null) continue;
                var (vx, vy) = ConvertModelToViewCoordinates(loc);
                double modelZ = loc.Z;
                
                // Get Расход воздуха parameter for air terminals
                string airflow = "";
                if (p.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                    p.AnnotationType == AnnotationType.AirTerminalShortNameFlow)
                {
                    var airflowParam = elem.get_Parameter(new Guid("a5c6c5f4-43c8-4c7b-9b8d-6e5f4d3c2b1a"));
                    if (airflowParam == null)
                    {
                        // Try by name
                        airflowParam = elem.LookupParameter("Расход воздуха");
                    }
                    if (airflowParam != null)
                    {
                        airflow = airflowParam.AsString() ?? airflowParam.AsDouble().ToString("F4");
                    }
                }
                
                elemDataMap[p.ElementId] = (vx, vy, modelZ, airflow);
            }
            
            // Group by annotation type and appropriate key
            // Air terminals: FamilyName + TypeName + Расход воздуха
            // Duct accessories: FamilyName only
            var typeGroups = groupablePlans
                .GroupBy(p =>
                {
                    if (p.AnnotationType == AnnotationType.AirTerminalTypeFlow ||
                        p.AnnotationType == AnnotationType.AirTerminalShortNameFlow)
                    {
                        string airflow = elemDataMap.ContainsKey(p.ElementId) ? elemDataMap[p.ElementId].airflow : "";
                        return $"{p.AnnotationType}|{p.FamilyName}|{p.TypeName}|{airflow}";
                    }
                    else
                    {
                        return $"{p.AnnotationType}|{p.FamilyName}";
                    }
                })
                .Where(g => g.Count() >= 2);
            
            foreach (var typeGroup in typeGroups)
            {
                var planList = typeGroup.ToList();
                var processed = new HashSet<long>();
                
                for (int i = 0; i < planList.Count; i++)
                {
                    if (processed.Contains(planList[i].ElementId))
                        continue;
                    
                    var primaryPlan = planList[i];
                    if (!elemDataMap.ContainsKey(primaryPlan.ElementId)) continue;
                    var (primaryViewX, primaryViewY, primaryModelZ, _) = elemDataMap[primaryPlan.ElementId];
                    
                    // Proximity threshold in mm on paper
                    double proximityMm = 300.0;
                    double proximityModelFeet = proximityMm / 304.8 * _view.Scale;
                    double proximityModelFeetSq = proximityModelFeet * proximityModelFeet;
                    
                    // Z-height tolerance: elements must be within ~100mm of each other
                    double zToleranceFeet = 0.33; // ~100mm in feet
                    
                    var nearbyPlans = new List<AnnotationPlan>();
                    processed.Add(primaryPlan.ElementId);
                    
                    for (int j = i + 1; j < planList.Count; j++)
                    {
                        if (processed.Contains(planList[j].ElementId))
                            continue;
                        
                        if (!elemDataMap.ContainsKey(planList[j].ElementId)) continue;
                        var (otherViewX, otherViewY, otherModelZ, _) = elemDataMap[planList[j].ElementId];
                        double distViewX = otherViewX - primaryViewX;
                        double distViewY = otherViewY - primaryViewY;
                        double distViewSq = distViewX * distViewX + distViewY * distViewY;
                        double distZ = Math.Abs(otherModelZ - primaryModelZ);
                        
                        // Must be nearby on 2D view AND at approximately same Z height
                        if (distViewSq <= proximityModelFeetSq && distZ <= zToleranceFeet)
                        {
                            nearbyPlans.Add(planList[j]);
                            processed.Add(planList[j].ElementId);
                        }
                    }
                    
                    if (nearbyPlans.Count > 0)
                    {
                        // Select the extreme element as primary based on 2D view X coordinate
                        var allGroupPlans = new List<AnnotationPlan> { primaryPlan };
                        allGroupPlans.AddRange(nearbyPlans);
                        
                        // Find rightmost and leftmost elements by view X
                        double maxViewX = double.MinValue;
                        double minViewX = double.MaxValue;
                        AnnotationPlan rightmostPlan = null;
                        AnnotationPlan leftmostPlan = null;
                        
                        foreach (var p in allGroupPlans)
                        {
                            if (!elemDataMap.ContainsKey(p.ElementId)) continue;
                            var (vx, vy, _, _) = elemDataMap[p.ElementId];
                            if (vx > maxViewX)
                            {
                                maxViewX = vx;
                                rightmostPlan = p;
                            }
                            if (vx < minViewX)
                            {
                                minViewX = vx;
                                leftmostPlan = p;
                            }
                        }
                        
                        // Determine which extreme element to use as primary.
                        // Strategy: prefer the element that is farther from the group center,
                        // so that the annotation can point AWAY from the group.
                        // - Rightmost element → annotation points RIGHT (away from group)
                        // - Leftmost element → annotation points LEFT (away from group)
                        double centerViewX = 0;
                        int count = 0;
                        foreach (var p in allGroupPlans)
                        {
                            if (!elemDataMap.ContainsKey(p.ElementId)) continue;
                            centerViewX += elemDataMap[p.ElementId].viewX;
                            count++;
                        }
                        if (count > 0) centerViewX /= count;
                        
                        // Use rightmost as primary if it's farther from center than leftmost,
                        // otherwise use leftmost. This ensures the annotation points away from the group.
                        bool isRightExtreme = (maxViewX - centerViewX) >= (centerViewX - minViewX);
                        AnnotationPlan extremePlan = isRightExtreme ? rightmostPlan : leftmostPlan;
                        if (extremePlan == null) extremePlan = primaryPlan;
                        
                        // Rebuild group with extreme element as primary
                        var group = new ElementGroup
                        {
                            PrimaryPlan = extremePlan,
                            IsRightExtreme = isRightExtreme
                        };
                        foreach (var p in allGroupPlans.Where(p => p.ElementId != extremePlan.ElementId))
                        {
                            group.AdditionalPlans.Add(p);
                        }
                        
                        groups.Add(group);
                        string direction = isRightExtreme ? "RIGHT" : "LEFT";
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Created group: primary={extremePlan.ElementId} (extreme element, {direction}), additional={string.Join(",", group.AdditionalPlans.Select(p => p.ElementId))}");
                    }
                }
            }
            
            return groups;
        }
        
        /// <summary>
        /// Returns a priority number for annotation category.
        /// Lower number = higher priority (placed first).
        /// Order: air terminals (1) → duct accessories (2) → equipment (3) → ducts (4) → other (5)
        /// </summary>
        private static int GetCategoryPriority(AnnotationType type)
        {
            switch (type)
            {
                case AnnotationType.AirTerminalTypeFlow:
                case AnnotationType.AirTerminalShortNameFlow:
                    return 1; // Air terminals - highest priority
                case AnnotationType.DuctAccessory:
                    return 2; // Duct accessories
                case AnnotationType.EquipmentMark:
                    return 3; // Equipment
                case AnnotationType.DuctRoundSizeFlow:
                case AnnotationType.DuctRectSizeFlow:
                    return 4; // Ducts - lowest priority
                default:
                    return 5; // Other (spot dimensions, etc.)
            }
        }

        /// <summary>
        /// Get a reference for an element suitable for tagging.
        /// For FamilyInstance elements, uses the element reference directly.
        /// For MEPCurves, gets the curve reference from geometry.
        /// </summary>
        private Reference GetElementReference(Element element)
        {
            try
            {
                if (element is MEPCurve mepCurve)
                {
                    var geoElem = mepCurve.get_Geometry(new Options());
                    if (geoElem != null)
                    {
                        foreach (var geoObj in geoElem)
                        {
                            if (geoObj is Curve curve && curve.Reference != null)
                                return curve.Reference;
                        }
                    }
                }
                
                // For FamilyInstance and other elements, use direct reference
                return new Reference(element);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[GREEDY-PLACEMENT] GetElementReference failed for {element.Id}: {ex.Message}");
                return null;
            }
        }
    }
}
