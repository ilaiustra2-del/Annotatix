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
                double viewShelfWidth, viewTextHeight, viewPad;
                try
                {
                    var headEdgeView = ConvertModelToViewCoordinates(headPos3D);
                    var headRightView = ConvertModelToViewCoordinates(new XYZ(headPos3D.X + shelfFeet, headPos3D.Y, headPos3D.Z));
                    var headUpView = ConvertModelToViewCoordinates(new XYZ(headPos3D.X, headPos3D.Y + textHeightFeet, headPos3D.Z));
                    var padView = ConvertModelToViewCoordinates(new XYZ(headPos3D.X + paddingFeet, headPos3D.Y, headPos3D.Z));
                    viewShelfWidth = Math.Abs(headRightView.X - headEdgeView.X);
                    viewTextHeight = Math.Abs(headUpView.Y - headEdgeView.Y);
                    viewPad = Math.Abs(padView.X - headEdgeView.X);
                    if (viewShelfWidth < 0.02) viewShelfWidth = 0.02;
                    if (viewTextHeight < 0.02) viewTextHeight = 0.02;
                    if (viewPad < 0.005) viewPad = 0.005;
                }
                catch
                {
                    viewShelfWidth = 0.02;
                    viewTextHeight = 0.02;
                    viewPad = 0.005;
                }
                
                // Build bbox: headView is at shelf TIP (far end from leader connection)
                // Shelf extends from headView TOWARD the element by viewShelfWidth
                bool isRight = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
                bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
                bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
                bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
                
                var bbox = new BBox2D();
                
                if (isRight)
                {
                    // Head is at right tip, shelf extends LEFT toward element
                    bbox.MinX = headView.X - viewShelfWidth - viewPad;
                    bbox.MaxX = headView.X + viewPad;
                }
                else // Left
                {
                    // Head is at left tip, shelf extends RIGHT toward element
                    bbox.MinX = headView.X - viewPad;
                    bbox.MaxX = headView.X + viewShelfWidth + viewPad;
                }
                
                if (isTop)
                {
                    bbox.MinY = headView.Y - viewPad;
                    bbox.MaxY = headView.Y + viewTextHeight + viewPad;
                }
                else if (isBottom)
                {
                    bbox.MinY = headView.Y - viewTextHeight - viewPad;
                    bbox.MaxY = headView.Y + viewPad;
                }
                else // Horizontal
                {
                    bbox.MinY = headView.Y - viewTextHeight / 2 - viewPad;
                    bbox.MaxY = headView.Y + viewTextHeight / 2 + viewPad;
                }
                
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
            
            // Sort plans: mandatory first
            var sortedPlans = plans
                .OrderByDescending(p => p.IsMandatory)
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
                    var result = PlaceSingle(primaryPlan, annotationIndex);
                    
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
                                            
                                            // Set elbow at same Y as head but X at leader end
                                            var headPos = tag.TagHeadPosition;
                                            XYZ elbowPos = CalculateElbowPosition(refLocation, headPos, groupResult.Position, groupResult.ElbowHeight);
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
                                _collisionDetector.AddOccupiedArea(occupiedBbox);
                        }
                    }
                    else
                    {
                        var bbox = GetTagBoundingBox(tag);
                        if (bbox != null)
                        {
                            _collisionDetector.AddOccupiedArea(bbox);
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
                                _collisionDetector.AddOccupiedArea(occupiedBbox);
                        }
                    }
                    else
                    {
                        var bbox = GetTagBoundingBox(result.CreatedTag);
                        if (bbox != null)
                        {
                            _collisionDetector.AddOccupiedArea(bbox);
                        }
                    }
                    
                    // Register annotation leader lines for future collision checks
                    RegisterAnnotationLinesForCollision(result.CreatedTag, plan.ElementId);
                }
            }
            
            return results;
        }
        
        private PlacementResult PlaceSingle(AnnotationPlan plan, int annotationIndex)
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
            
            var orderedPositions = GetOptimalPositionOrderFromOrientation(plan.PreferredPositions, orientation);
            string chosenConfig = PositionToRussianString(orderedPositions.FirstOrDefault());
            DebugLogger.Log($"- Положение в пространстве: {positionDesc} - выбрана конфигурация аннотации {chosenConfig}");
            
            // Annotation family/type info
            string annFamilyName = plan.FamilyName ?? "?";
            string annTypeName = plan.TypeName ?? "?";
            DebugLogger.Log($"- Выбранное семейство аннотации: {annFamilyName}");
            DebugLogger.Log($"  Тип аннотации: {annTypeName}");
            
            // Calculate optimal offsets based on element size and view scale
            double optimalHorizontalOffset = CalculateOptimalOffset(location, element);
            double optimalBaseElbowHeight = CalculateOptimalElbowHeight(location, element, 
                plan.PreferredPositions.FirstOrDefault());
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] Element {plan.ElementId}: optimalHorizontalOffset={optimalHorizontalOffset:F2}, optimalBaseElbowHeight={optimalBaseElbowHeight:F2}");
                    
            // Per user's algorithm: for each position, try up to 5 height iterations
            // before switching to the next priority position.
            // For Top positions: increase height (move annotation further UP)
            // For Bottom positions: increase height (move annotation further DOWN)
            // For Horizontal positions: no height adjustment (just 1 attempt)
            const int maxHeightIterations = 8;
                    
            bool is3DView = _view is View3D;
            double viewScaleCollision = _view.Scale;
            if (viewScaleCollision < 1) viewScaleCollision = 1;
            double shelfGapViewCollision = optimalHorizontalOffset; // Use dynamic offset (not hardcoded 1mm)
                
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
                        
                // Number of height iterations: 5 for L-shaped, 1 for horizontal
                int iterations = isLShaped ? maxHeightIterations : 1;
                        
                for (int iteration = 0; iteration < iterations; iteration++)
                {
                    globalIteration++;
                    double elbowHeight = optimalBaseElbowHeight + iteration * _config.ElbowHeightStep;
                    if (elbowHeight > _config.MaxElbowHeight)
                        break;
                            
                    // Calculate candidate placement bbox with dynamic offset
                    var candidateBbox = CalculatePlacementBboxWithOffset(location, position, elbowHeight, size, optimalHorizontalOffset);
                                                    
                    // Calculate leader line coordinates
                    var headPos = CalculateHeadPosition(location, position, elbowHeight, size, optimalHorizontalOffset);
                    var elbowPos = CalculateElbowPositionFromViewCoords(location.X, location.Y, headPos.X, headPos.Y, position);
                                                    
                    bool headerCollides = false;
                    bool leaderCollides = false;
                    CollisionDetector.CollisionDetails headerDetails = null;
                    CollisionDetector.CollisionDetails leaderDetails = null;
                                        
                    if (is3DView)
                    {
                        // 3D view: convert all coordinates to view space for collision detection
                        var locationViewCol = ConvertModelToViewCoordinates(location);
                                                        
                        (double headViewX, double headViewY) headPosView;
                        double elbowViewX, elbowViewY;
                                                        
                        if (isHorizontalPos)
                        {
                            double headViewXPos;
                            if (position == AnnotationPosition.HorizontalLeft)
                                headViewXPos = locationViewCol.X - shelfGapViewCollision;
                            else
                                headViewXPos = locationViewCol.X + shelfGapViewCollision;
                            headPosView = (headViewXPos, locationViewCol.Y);
                            elbowViewX = (locationViewCol.X + headPosView.headViewX) / 2.0;
                            elbowViewY = locationViewCol.Y;
                        }
                        else
                        {
                            headPosView = ConvertModelToViewCoordinates(new XYZ(headPos.X, headPos.Y, location.Z));
                            elbowViewX = locationViewCol.X;
                            elbowViewY = headPosView.headViewY;
                        }
                                                        
                        // Create bbox in view coordinates
                        double viewShelfWidth, viewTextHeight;
                        {
                            var headEdgeView = ConvertModelToViewCoordinates(new XYZ(headPos.X, headPos.Y, location.Z));
                            var headRightView = ConvertModelToViewCoordinates(new XYZ(headPos.X + size.Width, headPos.Y, location.Z));
                            var headUpView = ConvertModelToViewCoordinates(new XYZ(headPos.X, headPos.Y + size.Height, location.Z));
                            viewShelfWidth = Math.Abs(headRightView.X - headEdgeView.X);
                            viewTextHeight = Math.Abs(headUpView.Y - headEdgeView.Y);
                            if (viewShelfWidth < 0.02) viewShelfWidth = 0.02;
                            if (viewTextHeight < 0.02) viewTextHeight = 0.02;
                        }
                        
                        double viewPad;
                        {
                            var padView = ConvertModelToViewCoordinates(new XYZ(headPos.X + size.Padding, headPos.Y, location.Z));
                            viewPad = Math.Abs(padView.X - ConvertModelToViewCoordinates(new XYZ(headPos.X, headPos.Y, location.Z)).X);
                            if (viewPad < 0.005) viewPad = 0.005;
                        }
                        
                        var candidateBboxView = new BBox2D();
                        bool isRightPos = position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight || position == AnnotationPosition.HorizontalRight;
                        bool isLeftPos = position == AnnotationPosition.TopLeft || position == AnnotationPosition.BottomLeft || position == AnnotationPosition.HorizontalLeft;
                        
                        if (isRightPos)
                        {
                            candidateBboxView.MinX = headPosView.headViewX - viewPad;
                            candidateBboxView.MaxX = headPosView.headViewX + viewShelfWidth + viewPad;
                        }
                        else
                        {
                            candidateBboxView.MinX = headPosView.headViewX - viewShelfWidth - viewPad;
                            candidateBboxView.MaxX = headPosView.headViewX + viewPad;
                        }
                        
                        if (isTop)
                        {
                            candidateBboxView.MinY = headPosView.headViewY - viewPad;
                            candidateBboxView.MaxY = headPosView.headViewY + viewTextHeight + viewPad;
                        }
                        else if (isBottom)
                        {
                            candidateBboxView.MinY = headPosView.headViewY - viewTextHeight - viewPad;
                            candidateBboxView.MaxY = headPosView.headViewY + viewPad;
                        }
                        else
                        {
                            candidateBboxView.MinY = headPosView.headViewY - viewTextHeight / 2 - viewPad;
                            candidateBboxView.MaxY = headPosView.headViewY + viewTextHeight / 2 + viewPad;
                        }
                                                        
                        // Use detailed collision detection for logging
                        // 3D view: don't check model elements (ducts/pipes) because leader/header
                        // lines naturally cross their projections in 3D - only check annotation-to-annotation
                        headerDetails = _collisionDetector.GetHeaderCollisionDetails(candidateBboxView, checkModelElements: false);
                        leaderDetails = _collisionDetector.GetLeaderCollisionDetails(
                            locationViewCol.X, locationViewCol.Y,
                            elbowViewX, elbowViewY,
                            headPosView.headViewX, headPosView.headViewY,
                            plan.ElementId, checkModelElements: false);
                        headerCollides = headerDetails.HasCollision;
                        leaderCollides = leaderDetails.HasCollision;
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
                        var lv = ConvertModelToViewCoordinates(location);
                        var hv = ConvertModelToViewCoordinates(new XYZ(headPos.X, headPos.Y, location.Z));
                        iterLeaderEndViewX = lv.X; iterLeaderEndViewY = lv.Y;
                        iterHeaderViewX = hv.X; iterHeaderViewY = hv.Y;
                        iterElbowViewX = lv.X; // elbow view X
                        iterElbowViewY = hv.Y; // elbow view Y
                        DebugLogger.Log($"[Итерация подбора положения {globalIteration}]");
                        DebugLogger.Log($"Форма аннотации: {posName}; LeaderEndViewX,Y = {lv.X:F2}, {lv.Y:F2}; ElbowViewX,Y = {iterElbowViewX:F2}, {iterElbowViewY:F2}; HeaderViewX,Y = {hv.X:F2}, {hv.Y:F2}");
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
                        var distinctIds = allCollidingIds.Where(id => id > 0).Distinct().ToList();
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
                        var tag = CreateAnnotation(plan, element, location, position, elbowHeight, optimalHorizontalOffset);
                
                        if (tag != null)
                        {
                            // Log annotation content and shelf length
                            string tagText = "";
                            try { tagText = tag.TagText ?? ""; } catch { }
                            double? shelfMm = TagTypeManager.GetShelfLengthFromTypeNamePublic(tag.Name);
                            string shelfInfo = shelfMm.HasValue ? $"{shelfMm.Value:F0}" : "?";
                            double sizeWidthMm = size.Width / _view.Scale * 304.8;
                            DebugLogger.Log($"- Содержимое аннотации: {tagText}. Длина полки: {shelfInfo} мм (расчетная ширина: {sizeWidthMm:F1} мм). Создан тип: {tag.Name}");
                            
                            // Record placement summary for CSV export
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
                                FailureReason = ""
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
                        }
                    }
                }
            }
            
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
        /// Calculate optimal offset based on element size and view scale.
        /// For annotations, the offset should be in annotation paper space (mm on paper),
        /// then converted to model space using view scale.
        /// Typical offset on paper: 5-15mm from element edge to annotation.
        /// </summary>
        private double CalculateOptimalOffset(XYZ elementLocation, Element element)
        {
            // Base offset in mm on paper (distance from element to NEAREST EDGE of tag)
            // This is just a small gap; the TotalWidth/2 is added separately for head center position
            double baseOffsetPaperMm = 1.0; // 1mm on paper gap from element edge to tag edge
            
            // Add extra offset for larger elements (they need more clearance)
            var bbox = element.get_BoundingBox(_view);
            if (bbox != null)
            {
                // Calculate element dimensions in model space
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
            // Formula: modelFeet = paperMm / 304.8 * viewScale
            // At 1:100 scale, 5mm on paper = 5/304.8 * 100 = 1.64 feet in model
            double viewScale = _view.Scale;
            if (viewScale < 1) viewScale = 1;
            double modelOffsetFeet = baseOffsetPaperMm / 304.8 * viewScale;
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] Offset: paperMm={baseOffsetPaperMm:F1}, viewScale={viewScale}, modelFeet={modelOffsetFeet:F2}");
            
            // Cap to reasonable maximum (3 meters in model space)
            return Math.Min(modelOffsetFeet, 3.0);
        }
                
        /// <summary>
        /// Calculate optimal elbow height based on element position and view.
        /// Like the offset, this is calculated in paper space (mm) then converted to model space.
        /// Typical elbow height: 3-10mm on paper.
        /// </summary>
        private double CalculateOptimalElbowHeight(XYZ elementLocation, Element element, AnnotationPosition position)
        {
            // Base elbow height in mm on paper (distance from element to shelf on printed sheet)
            // Small value because the leader should be short and elegant
            double baseHeightPaperMm = 3.0; // 3mm on paper (was 5mm - too large)
            
            // Get element bounding box
            var bbox = element.get_BoundingBox(_view);
            if (bbox != null)
            {
                double height = bbox.Max.Y - bbox.Min.Y;
                double depth = bbox.Max.Z - bbox.Min.Z;
                double maxDim = Math.Max(height, depth);
                
                // For large elements, add more elbow height on paper
                if (maxDim > 0.5)
                {
                    baseHeightPaperMm += Math.Min(maxDim * 1.5, 5.0); // Up to 5mm extra on paper
                }
            }
            
            // Convert paper mm to model space feet
            double viewScale = _view.Scale;
            if (viewScale < 1) viewScale = 1;
            double modelHeightFeet = baseHeightPaperMm / 304.8 * viewScale;
            
            // Cap to reasonable maximum
            return Math.Min(modelHeightFeet, 3.0);
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
            // TagHeadPosition = END of the leader line (where tag text is displayed)
            // The leader goes: LeaderEnd -> Elbow -> Head(TagHeadPosition)
            // For Left positions: Head is LEFT of Elbow by (shelfWidth + shelfGap)
            // For Right positions: Head is RIGHT of Elbow by (shelfWidth + shelfGap)
            // For Horizontal: Head is directly left/right of LeaderEnd by (shelfWidth + shelfGap)
            // horizontalOffset = distance from element edge to nearest annotation edge (in model feet)
            // This is calculated by CalculateOptimalOffset based on element size (1-4mm on paper)
            double shelfGapModel = horizontalOffset; // Use the dynamic offset for consistency with collision detection
            double shelfWidthModel = size.Width; // shelf width in model feet
            
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
            
            if (isHorizontal)
            {
                // Horizontal: leader goes SIDEWAYS from element
                // Head is at shelf tip (far end from element)
                double horizontalDist = shelfGapModel + shelfWidthModel;
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
                // Head is at shelf tip, (shelfWidth + shelfGap) from leader end X
                double headX;
                double sideOffset = shelfGapModel + shelfWidthModel;
                if (position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight)
                {
                    headX = location.X + sideOffset;
                }
                else // TopLeft, BottomLeft
                {
                    headX = location.X - sideOffset;
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
            
            // X extent: headPos is at the shelf tip
            // Shelf extends from headPos TOWARD the element
            if (isRight)
            {
                // Head is at right tip, shelf extends LEFT toward element
                bbox.MinX = headPos.X - shelfWidth - pad; // left edge (near element)
                bbox.MaxX = headPos.X + pad;              // right edge (tip + padding)
            }
            else // Left
            {
                // Head is at left tip, shelf extends RIGHT toward element
                bbox.MinX = headPos.X - pad;               // left edge (tip - padding)
                bbox.MaxX = headPos.X + shelfWidth + pad;   // right edge (near element)
            }
            
            // Y extent: text extends from shelf upward (for top) or downward (for bottom)
            if (isTop)
            {
                bbox.MinY = headPos.Y - pad;              // small padding below shelf
                bbox.MaxY = headPos.Y + textHeight + pad; // text extends up + padding
            }
            else if (isBottom)
            {
                bbox.MinY = headPos.Y - textHeight - pad; // text extends down + padding
                bbox.MaxY = headPos.Y + pad;              // small padding above shelf
            }
            else // Horizontal
            {
                bbox.MinY = headPos.Y - textHeight / 2 - pad;
                bbox.MaxY = headPos.Y + textHeight / 2 + pad;
            }
            
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
            double horizontalOffset)
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
                    tagHeadPosition = CalculateHeadPositionFor3DView(
                        location, tagZ, horizontalOffset, size, position, elbowHeight);
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
                                maxLineLength = Math.Max(line1Len, line2Len);
                                DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view: multi-line tag \"{tagText}\" split: line1~={line1Len}, line2~={line2Len}, maxLine={maxLineLength}");
                            }
                            else
                            {
                                // Single-line tag or TagText has line breaks
                                string[] lines = tagText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                maxLineLength = lines.Length > 0 ? lines.Max(l => l.Length) : tagText.Length;
                            }
                            
                            // Estimate width: 1.0mm per character (no padding here - added in TagTypeManager)
                            double estimatedWidthMm = maxLineLength * 1.0;
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
                        // 3D view: After tag type change, the actual shelf length may differ.
                        // Since TagHeadPosition = shelf edge at 1mm from leader, head position
                        // doesn't depend on shelf width (only shelfGap). So no head recalculation needed.
                        // The actual shelf length is used by ComputeOccupiedBbox3D for collision detection.
                        double? actualShelfMm = TagTypeManager.GetShelfLengthFromTypeNamePublic(tag.Name);
                        if (actualShelfMm.HasValue && actualShelfMm.Value > 0)
                        {
                            DebugLogger.Log($"[GREEDY-PLACEMENT] 3D view: tag type changed, actual shelf={actualShelfMm.Value:F0}mm (head position unchanged, shelf gap=1mm)");
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
            double elbowHeight)
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
            
            // TagHeadPosition = END of the leader line (where tag text is displayed)
            // The leader goes: LeaderEnd -> Elbow -> Head(TagHeadPosition)
            // Head is at (shelfWidth + shelfGap) from leader end in the horizontal direction
            // horizontalOffset = distance from element edge to nearest annotation edge (in model feet)
            // This is calculated by CalculateOptimalOffset based on element size (1-4mm on paper)
            double shelfGapView = horizontalOffset; // Use the dynamic offset for consistency with collision detection
            double shelfWidthView = size.Width; // shelf width in model feet
            double sideOffsetView = shelfGapView + shelfWidthView; // total horizontal distance from leader end to head
            
            if (isHorizontal)
            {
                // Horizontal: straight line on screen, headViewY = leaderEndViewY
                // Head is at (shelfWidth + shelfGap) from leader end horizontally
                headViewY = leaderEndViewY;
                
                if (position == AnnotationPosition.HorizontalLeft)
                {
                    headViewX = leaderEndViewX - sideOffsetView;
                }
                else // HorizontalRight
                {
                    headViewX = leaderEndViewX + sideOffsetView;
                }
            }
            else
            {
                // L-shaped (Top/Bottom): leader goes UP/DOWN then SIDEWAYS
                // Head is at shelf tip, (shelfWidth + shelfGap) from leader end X
                if (isTop)
                {
                    headViewY = leaderEndViewY + elbowHeight;
                }
                else // Bottom
                {
                    headViewY = leaderEndViewY - elbowHeight;
                }
                
                if (position == AnnotationPosition.TopRight || position == AnnotationPosition.BottomRight)
                {
                    headViewX = leaderEndViewX + sideOffsetView;
                }
                else // TopLeft, BottomLeft
                {
                    headViewX = leaderEndViewX - sideOffsetView;
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
            // Format: "Ecoline 2  100" or "VS 100M  50"
            // Line 1 is the type name, Line 2 is the flow value
            // The boundary is typically after the name - look for multiple spaces or a numeric-only suffix
            if (annotationType == AnnotationType.AirTerminalTypeFlow ||
                annotationType == AnnotationType.AirTerminalShortNameFlow)
            {
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
            
            // Group by annotation type and family name
            var typeGroups = groupablePlans
                .GroupBy(p => new { p.AnnotationType, p.FamilyName })
                .Where(g => g.Count() >= 2);
            
            foreach (var typeGroup in typeGroups)
            {
                var planList = typeGroup.ToList();
                var processed = new HashSet<long>();
                
                // Collect view coordinates for all elements to determine extreme positions
                var planViewCoords = new Dictionary<long, (double viewX, double viewY)>();
                foreach (var p in planList)
                {
                    var elem = _document.GetElement(new ElementId(p.ElementId));
                    if (elem == null) continue;
                    var loc = GetElementLocation(elem);
                    if (loc == null) continue;
                    var (vx, vy) = ConvertModelToViewCoordinates(loc);
                    planViewCoords[p.ElementId] = (vx, vy);
                }
                
                for (int i = 0; i < planList.Count; i++)
                {
                    if (processed.Contains(planList[i].ElementId))
                        continue;
                    
                    var primaryPlan = planList[i];
                    if (!planViewCoords.ContainsKey(primaryPlan.ElementId)) continue;
                    var (primaryViewX, primaryViewY) = planViewCoords[primaryPlan.ElementId];
                    
                    // Proximity threshold in mm on paper
                    // Two elements are "nearby" if they are within this distance on the printed sheet
                    double proximityMm = 300.0; // 300mm on paper = ~30cm at typical print scale
                    double proximityModelFeet = proximityMm / 304.8 * _view.Scale;
                    double proximityModelFeetSq = proximityModelFeet * proximityModelFeet;
                    
                    var nearbyPlans = new List<AnnotationPlan>();
                    processed.Add(primaryPlan.ElementId);
                    
                    for (int j = i + 1; j < planList.Count; j++)
                    {
                        if (processed.Contains(planList[j].ElementId))
                            continue;
                        
                        if (!planViewCoords.ContainsKey(planList[j].ElementId)) continue;
                        var (otherViewX, otherViewY) = planViewCoords[planList[j].ElementId];
                        double distViewX = otherViewX - primaryViewX;
                        double distViewY = otherViewY - primaryViewY;
                        double distViewSq = distViewX * distViewX + distViewY * distViewY;
                        
                        if (distViewSq <= proximityModelFeetSq)
                        {
                            nearbyPlans.Add(planList[j]);
                            processed.Add(planList[j].ElementId);
                        }
                    }
                    
                    if (nearbyPlans.Count > 0)
                    {
                        // Select the extreme element as primary:
                        // The element with the largest or smallest X or Y coordinate on the view
                        var allGroupPlans = new List<AnnotationPlan> { primaryPlan };
                        allGroupPlans.AddRange(nearbyPlans);
                        
                        AnnotationPlan extremePlan = primaryPlan;
                        double maxAbsCoord = 0;
                        
                        foreach (var p in allGroupPlans)
                        {
                            if (!planViewCoords.ContainsKey(p.ElementId)) continue;
                            var (vx, vy) = planViewCoords[p.ElementId];
                            // Use max absolute coordinate to find the element at the edge of the group
                            double maxCoord = Math.Max(Math.Abs(vx), Math.Abs(vy));
                            if (maxCoord > maxAbsCoord)
                            {
                                maxAbsCoord = maxCoord;
                                extremePlan = p;
                            }
                        }
                        
                        // Rebuild group with extreme element as primary
                        var group = new ElementGroup { PrimaryPlan = extremePlan };
                        foreach (var p in allGroupPlans.Where(p => p.ElementId != extremePlan.ElementId))
                        {
                            group.AdditionalPlans.Add(p);
                        }
                        
                        groups.Add(group);
                        DebugLogger.Log($"[GREEDY-PLACEMENT] Created group: primary={extremePlan.ElementId} (extreme element), additional={string.Join(",", group.AdditionalPlans.Select(p => p.ElementId))}");
                    }
                }
            }
            
            return groups;
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
