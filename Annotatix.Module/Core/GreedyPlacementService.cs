using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
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
        }
        
        /// <summary>
        /// Place all annotations greedily, avoiding collisions
        /// </summary>
        public List<PlacementResult> PlaceAll(List<AnnotationPlan> plans)
        {
            var results = new List<PlacementResult>();
            
            // Sort plans: mandatory first
            var sortedPlans = plans
                .OrderByDescending(p => p.IsMandatory)
                .ThenBy(p => p.Node?.Degree ?? 0)
                .ToList();
            
            foreach (var plan in sortedPlans)
            {
                var result = PlaceSingle(plan);
                results.Add(result);
                
                // If successful, register occupied area
                if (result.Success && result.CreatedTag != null)
                {
                    var bbox = GetTagBoundingBox(result.CreatedTag);
                    if (bbox != null)
                    {
                        _collisionDetector.AddOccupiedArea(bbox);
                    }
                }
            }
            
            return results;
        }
        
        private PlacementResult PlaceSingle(AnnotationPlan plan)
        {
            var element = _document.GetElement(new ElementId(plan.ElementId));
            if (element == null)
            {
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
                return new PlacementResult
                {
                    Success = false,
                    ElementId = plan.ElementId,
                    Plan = plan,
                    FailureReason = "Could not get element location"
                };
            }
            
            // Calculate optimal offsets based on element size and view scale
            double optimalHorizontalOffset = CalculateOptimalOffset(location, element);
            double optimalBaseElbowHeight = CalculateOptimalElbowHeight(location, element, 
                plan.PreferredPositions.FirstOrDefault());
            
            DebugLogger.Log($"[GREEDY-PLACEMENT] Element {plan.ElementId}: optimalHorizontalOffset={optimalHorizontalOffset:F2}, optimalBaseElbowHeight={optimalBaseElbowHeight:F2}");
            
            // Try each position
            foreach (var position in plan.PreferredPositions)
            {
                // Use calculated optimal values with some variation for collision avoidance
                for (double elbowHeight = optimalBaseElbowHeight; 
                     elbowHeight <= _config.MaxElbowHeight; 
                     elbowHeight += _config.ElbowHeightStep)
                {
                    // Calculate candidate placement bbox with dynamic offset
                    var candidateBbox = CalculatePlacementBboxWithOffset(location, position, elbowHeight, size, optimalHorizontalOffset);
                    
                    // Check for collisions
                    if (!_collisionDetector.HasCollision(candidateBbox, plan.ElementId))
                    {
                        // Try to create the actual annotation
                        var tag = CreateAnnotation(plan, element, location, position, elbowHeight, optimalHorizontalOffset);
                        
                        if (tag != null)
                        {
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
            
            return new PlacementResult
            {
                Success = false,
                ElementId = plan.ElementId,
                Plan = plan,
                FailureReason = "No valid position found after trying all options"
            };
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
        /// Calculate optimal offset based on element size and view scale
        /// </summary>
        private double CalculateOptimalOffset(XYZ elementLocation, Element element)
        {
            double baseOffset = _config.BaseHorizontalOffset;
                    
            // Get element bounding box to determine size
            var bbox = element.get_BoundingBox(_view);
            if (bbox != null)
            {
                // Calculate element dimensions
                double width = bbox.Max.X - bbox.Min.X;
                double height = bbox.Max.Y - bbox.Min.Y;
                double depth = bbox.Max.Z - bbox.Min.Z;
                        
                // Use the largest dimension to scale offset
                double maxDimension = Math.Max(Math.Max(width, height), depth);
                
                // Cap the max dimension contribution to avoid excessive offsets for large equipment
                // Maximum additional offset from size is 5 meters
                double sizeContribution = Math.Min(maxDimension * _config.ElementSizeMultiplier, 5.0);
                baseOffset += sizeContribution;
            }
                    
            // Scale by view scale (larger scale = larger offset needed)
            double viewScale = _view.Scale;
            if (viewScale > 1)
            {
                // For scales like 1:100, viewScale = 100
                // We need larger offsets for larger scales
                // Cap the scale multiplier to avoid excessive offsets
                double scaleMultiplier = Math.Min(1 + Math.Log10(viewScale) * _config.ViewScaleMultiplier * 10, 3.0);
                baseOffset *= scaleMultiplier;
            }
                    
            // Cap the final offset to a reasonable maximum (10 meters)
            return Math.Min(Math.Max(baseOffset, _config.MinLeaderLength), 10.0);
        }
                
        /// <summary>
        /// Calculate optimal elbow height based on element position and view
        /// </summary>
        private double CalculateOptimalElbowHeight(XYZ elementLocation, Element element, AnnotationPosition position)
        {
            double baseHeight = _config.BaseElbowHeight;
                    
            // Get element bounding box
            var bbox = element.get_BoundingBox(_view);
            if (bbox != null)
            {
                // Add extra height based on element size
                double height = bbox.Max.Y - bbox.Min.Y;
                double depth = bbox.Max.Z - bbox.Min.Z;
                
                // Cap the size contribution to avoid excessive heights
                double sizeContribution = Math.Min(Math.Max(height, depth) * _config.ElementSizeMultiplier, 5.0);
                baseHeight += sizeContribution;
            }
                    
            // Scale by view scale
            double viewScale = _view.Scale;
            if (viewScale > 1)
            {
                // Cap the scale multiplier to avoid excessive heights
                double scaleMultiplier = Math.Min(1 + Math.Log10(viewScale) * _config.ViewScaleMultiplier * 5, 3.0);
                baseHeight *= scaleMultiplier;
            }
                    
            // Cap to maximum elbow height from config or 5 meters
            return Math.Min(Math.Max(baseHeight, _config.MinLeaderLength), Math.Min(_config.MaxElbowHeight, 5.0));
        }
        
        /// <summary>
        /// Calculate optimal Z offset for 3D views
        /// This ensures annotations are placed at a visible elevation
        /// </summary>
        private double CalculateOptimalZOffset(XYZ elementLocation, Element element, AnnotationPosition position)
        {
            // For 3D views, we need to offset Z for visibility
            if (_view is View3D)
            {
                // Start with base Z offset
                double zOffset = _config.BaseZOffset;
                
                // Get element bounding box
                var bbox = element.get_BoundingBox(_view);
                if (bbox != null)
                {
                    // Calculate element depth (Z dimension)
                    double depth = bbox.Max.Z - bbox.Min.Z;
                    // Also consider width and height for diagonal positioning
                    double width = bbox.Max.X - bbox.Min.X;
                    double height = bbox.Max.Y - bbox.Min.Y;
                    
                    // Use the maximum dimension to ensure we clear the element
                    double maxDimension = Math.Max(Math.Max(depth, width), height);
                    zOffset += maxDimension * _config.ZOffsetSizeMultiplier;
                }
                
                // Scale by view scale - larger scale needs larger offsets
                double viewScale = _view.Scale;
                if (viewScale > 1)
                {
                    // For 1:100 scale, this adds about 50% more
                    zOffset *= (1 + Math.Log10(viewScale) * _config.ViewScaleMultiplier * 5);
                }
                
                // Adjust sign based on position
                // For Top positions, move annotation above element
                // For Bottom positions, move annotation below element
                // For Horizontal positions, move to the side
                bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
                bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
                
                if (isBottom)
                {
                    // Move below the element
                    zOffset = -zOffset;
                }
                // For top and horizontal positions, keep positive (above element)
                
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
            // Calculate head position based on annotation position
            var headPos = CalculateHeadPosition(location, position, elbowHeight, size);
            
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
        /// The leader forms an L-shape:
        /// - LeaderEnd is at the element (location parameter)
        /// - Vertical segment goes from LeaderEnd to ElbowPosition (height = elbowHeight)
        /// - Horizontal segment goes from ElbowPosition to HeadPosition (length = horizontalOffset)
        /// 
        /// Key constraint: HeadPosition.Y = ElbowPosition.Y (horizontal connection)
        /// </summary>
        private (double X, double Y) CalculateHeadPosition(
            XYZ location,
            AnnotationPosition position,
            double elbowHeight,
            AnnotationSize size,
            double horizontalOffset)
        {
            double offsetX = horizontalOffset;
            
            // For L-shaped leader:
            // - HeadPosition.Y must equal ElbowPosition.Y (horizontal connection)
            // - ElbowPosition.Y = location.Y + elbowHeight (for Top positions)
            // - So HeadPosition.Y = location.Y + elbowHeight (NOT adding annotation size!)
            
            return position switch
            {
                // Top positions: leader goes UP from element, head is to the side
                // HeadPosition.Y = ElbowPosition.Y = location.Y + elbowHeight
                AnnotationPosition.TopLeft => (
                    location.X - offsetX - size.TotalWidth / 2,  // Head X is to the left
                    location.Y + elbowHeight                      // Head Y = Elbow Y
                ),
                AnnotationPosition.TopRight => (
                    location.X + offsetX + size.TotalWidth / 2,  // Head X is to the right
                    location.Y + elbowHeight                      // Head Y = Elbow Y
                ),
                // Bottom positions: leader goes DOWN from element, head is to the side
                // HeadPosition.Y = ElbowPosition.Y = location.Y - elbowHeight
                AnnotationPosition.BottomLeft => (
                    location.X - offsetX - size.TotalWidth / 2,
                    location.Y - elbowHeight                      // Head Y = Elbow Y
                ),
                AnnotationPosition.BottomRight => (
                    location.X + offsetX + size.TotalWidth / 2,
                    location.Y - elbowHeight                      // Head Y = Elbow Y
                ),
                // Horizontal positions: leader goes SIDEWAYS, then vertical to head
                // Different L-shape orientation
                AnnotationPosition.HorizontalLeft => (
                    location.X - offsetX - size.TotalWidth,  // Head X is to the left
                    location.Y                               // Head Y at element Y level
                ),
                AnnotationPosition.HorizontalRight => (
                    location.X + offsetX + size.TotalWidth,   // Head X is to the right
                    location.Y                                // Head Y at element Y level
                ),
                _ => (location.X, location.Y + elbowHeight)
            };
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
            
            return new BBox2D
            {
                MinX = headPos.X - size.TotalWidth / 2,
                MaxX = headPos.X + size.TotalWidth / 2,
                MinY = headPos.Y - size.TotalHeight / 2,
                MaxY = headPos.Y + size.TotalHeight / 2
            };
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
                
                // For 3D views, we need to position the annotation at a different elevation
                // than the element to avoid overlap
                var tagHeadPosition = new XYZ(headPos.X, headPos.Y, tagZ);
                
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
                    // Set leader end condition to free for custom positioning
                    tag.LeaderEndCondition = LeaderEndCondition.Free;
                    tag.TagHeadPosition = tagHeadPosition;
                    
                    // Get the tagged reference for leader configuration
                    var taggedRefs = tag.GetTaggedReferences();
                    Reference tagRef = (taggedRefs != null && taggedRefs.Count > 0) ? taggedRefs[0] : elemRef;
                    
                    // Set elbow position for proper 90-degree angle (L-shaped leader)
                    // The leader forms an L-shape:
                    // - LeaderEnd at element location (vertical leader start)
                    // - ElbowPosition: X = LeaderEnd.X, Y = HeadPosition.Y (corner of L)
                    // - HeadPosition: the annotation head
                    if (plan.HasElbow)
                    {
                        try
                        {
                            XYZ elbowPos = CalculateElbowPosition(location, tagHeadPosition, position, elbowHeight);
                            tag.SetLeaderElbow(tagRef, elbowPos);
                            
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
                    
                    // Set leader end position on the element
                    try
                    {
                        tag.SetLeaderEnd(tagRef, location);
                    }
                    catch (Exception endEx)
                    {
                        DebugLogger.Log("[GREEDY-PLACEMENT] Could not set leader end: " + endEx.Message);
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
        private XYZ CalculateElbowPosition(XYZ leaderEndPosition, XYZ headPosition, AnnotationPosition position, double elbowHeight)
        {
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
        
            // For 3D views, calculate elbow position based on VIEW projection
            if (_view is View3D view3D)
            {
                return CalculateElbowPositionFor3DView(leaderEndPosition, headPosition, position, view3D);
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
                return new XYZ(
                    headPosition.X,       // Elbow X matches HeadPosition X
                    leaderEndPosition.Y,  // Elbow Y matches LeaderEnd Y
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
        private XYZ CalculateElbowPositionFor3DView(XYZ leaderEndPosition, XYZ headPosition, AnnotationPosition position, View3D view3D)
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
            
            // Calculate desired elbow position in VIEW space (screen coordinates)
            // For an L-shape on screen, the elbow must satisfy:
            // - Elbow viewX = LeaderEnd viewX (vertical segment on screen)
            // - Elbow viewY = Head viewY (horizontal segment on screen)
            // 
            // CRITICAL: Use headDepth for the elbow depth so that the horizontal segment
            // (elbow to head) lies on the same depth plane. This ensures that when projected,
            // the horizontal segment appears truly horizontal on screen.
            
            bool isTop = position == AnnotationPosition.TopLeft || position == AnnotationPosition.TopRight;
            bool isBottom = position == AnnotationPosition.BottomLeft || position == AnnotationPosition.BottomRight;
            bool isHorizontal = position == AnnotationPosition.HorizontalLeft || position == AnnotationPosition.HorizontalRight;
            
            double elbowViewX, elbowViewY, elbowDepth;
            
            if (isTop || isBottom)
            {
                // For top/bottom positions: vertical leader from element on screen, horizontal to head
                // Elbow screen X = LeaderEnd screen X (vertical on screen)
                // Elbow screen Y = Head screen Y (horizontal to head)
                // Elbow depth = Head depth (so horizontal segment is in same depth plane)
                elbowViewX = leaderEndViewX;  // Same screen X as leader end (for vertical leader on screen)
                elbowViewY = headViewY;        // Same screen Y as head (for horizontal segment on screen)
                elbowDepth = headDepth;         // Same depth as head for proper horizontal segment
            }
            else if (isHorizontal)
            {
                // For horizontal positions: horizontal leader from element, then vertical to head
                elbowViewX = headViewX;        // Same screen X as head
                elbowViewY = leaderEndViewY;   // Same screen Y as leader end
                elbowDepth = headDepth;         // Same depth as head
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
        /// Get base type name without numeric suffix
        /// e.g. "ADSK_Марка_15" -> "ADSK_Марка"
        /// </summary>
        private string GetBaseTypeName(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return typeName;
            
            // Remove trailing underscore followed by numbers
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
    }
}
