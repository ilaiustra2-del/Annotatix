using System;
using System.Collections.Generic;
using System.Linq;
using PluginsManager.Core;
using Autodesk.Revit.DB;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Represents a 2D bounding box for collision detection
    /// </summary>
    public class BBox2D
    {
        public double MinX { get; set; }
        public double MinY { get; set; }
        public double MaxX { get; set; }
        public double MaxY { get; set; }
        
        public double Width => MaxX - MinX;
        public double Height => MaxY - MinY;
        
        public double CenterX => (MinX + MaxX) / 2;
        public double CenterY => (MinY + MaxY) / 2;
        
        /// <summary>
        /// Check if this bbox intersects with another
        /// </summary>
        public bool Intersects(BBox2D other)
        {
            return !(MaxX < other.MinX || MinX > other.MaxX ||
                     MaxY < other.MinY || MinY > other.MaxY);
        }
        
        /// <summary>
        /// Check if this bbox contains a point
        /// </summary>
        public bool Contains(double x, double y)
        {
            return x >= MinX && x <= MaxX && y >= MinY && y <= MaxY;
        }
        
        /// <summary>
        /// Expand bbox by a margin
        /// </summary>
        public BBox2D Expand(double margin)
        {
            return new BBox2D
            {
                MinX = MinX - margin,
                MinY = MinY - margin,
                MaxX = MaxX + margin,
                MaxY = MaxY + margin
            };
        }
    }
    
    /// <summary>
    /// Represents a 2D line segment
    /// </summary>
    public class LineSegment2D
    {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        
        public long ElementId { get; set; }
        
        /// <summary>
        /// Get bounding box of this segment
        /// </summary>
        public BBox2D GetBoundingBox()
        {
            return new BBox2D
            {
                MinX = Math.Min(X1, X2),
                MaxX = Math.Max(X1, X2),
                MinY = Math.Min(Y1, Y2),
                MaxY = Math.Max(Y1, Y2)
            };
        }
        
        /// <summary>
        /// Get point at parameter t (0 to 1)
        /// </summary>
        public (double X, double Y) PointAt(double t)
        {
            return (
                X1 + t * (X2 - X1),
                Y1 + t * (Y2 - Y1)
            );
        }
        
        /// <summary>
        /// Get length of segment
        /// </summary>
        public double Length()
        {
            return Math.Sqrt(Math.Pow(X2 - X1, 2) + Math.Pow(Y2 - Y1, 2));
        }
    }
    
    /// <summary>
    /// Represents a visual intersection point
    /// </summary>
    public class VisualIntersection
    {
        public double X { get; set; }
        public double Y { get; set; }
        
        /// <summary>The two elements that visually intersect</summary>
        public long ElementId1 { get; set; }
        public long ElementId2 { get; set; }
    }
    
    /// <summary>
    /// Detects visual collisions between elements and annotation placements
    /// </summary>
    public class CollisionDetector
    {
        private readonly List<LineSegment2D> _lineSegments = new List<LineSegment2D>();
        private readonly List<BBox2D> _occupiedAreas = new List<BBox2D>();
        private readonly List<VisualIntersection> _intersections = new List<VisualIntersection>();
        
        /// <summary>
        /// Add elements from system graph for collision detection
        /// </summary>
        public void AddElementsFromGraph(SystemGraph graph)
        {
            foreach (var node in graph.Nodes.Values)
            {
                if (node.HasEndPoint && node.ViewStart != null && node.ViewEnd != null)
                {
                    _lineSegments.Add(new LineSegment2D
                    {
                        X1 = node.ViewStart.X,
                        Y1 = node.ViewStart.Y,
                        X2 = node.ViewEnd.X,
                        Y2 = node.ViewEnd.Y,
                        ElementId = node.ElementId
                    });
                }
            }
            
            // Find all visual intersections
            FindAllIntersections();
        }
        
        /// <summary>
        /// Add elements from view snapshot for collision detection
        /// </summary>
        public void AddElementsFromSnapshot(ViewSnapshot snapshot)
        {
            foreach (var elem in snapshot.Elements)
            {
                if (elem.HasEndPoint)
                {
                    _lineSegments.Add(new LineSegment2D
                    {
                        X1 = elem.ViewStart.X,
                        Y1 = elem.ViewStart.Y,
                        X2 = elem.ViewEnd.X,
                        Y2 = elem.ViewEnd.Y,
                        ElementId = elem.ElementId
                    });
                }
            }
            
            // Add existing annotations as occupied areas
            foreach (var annot in snapshot.Annotations)
            {
                if (annot.HeadPosition != null)
                {
                    // Estimate annotation bbox (simplified)
                    _occupiedAreas.Add(new BBox2D
                    {
                        MinX = annot.HeadPosition.X - 0.1,
                        MaxX = annot.HeadPosition.X + 0.1,
                        MinY = annot.HeadPosition.Y - 0.05,
                        MaxY = annot.HeadPosition.Y + 0.05
                    });
                }
            }
            
            FindAllIntersections();
        }
        
        /// <summary>
        /// Find all pairwise intersections between line segments
        /// Uses naive O(n²) algorithm - could be optimized with sweep line for large datasets
        /// </summary>
        private void FindAllIntersections()
        {
            _intersections.Clear();
            
            for (int i = 0; i < _lineSegments.Count; i++)
            {
                for (int j = i + 1; j < _lineSegments.Count; j++)
                {
                    var segA = _lineSegments[i];
                    var segB = _lineSegments[j];
                    
                    // Skip if same element
                    if (segA.ElementId == segB.ElementId)
                        continue;
                    
                    // Check if bounding boxes overlap (quick rejection)
                    var bboxA = segA.GetBoundingBox();
                    var bboxB = segB.GetBoundingBox();
                    
                    if (!bboxA.Intersects(bboxB))
                        continue;
                    
                    // Find exact intersection
                    var intersection = LineSegmentIntersection(segA, segB);
                    if (intersection != null)
                    {
                        intersection.ElementId1 = segA.ElementId;
                        intersection.ElementId2 = segB.ElementId;
                        _intersections.Add(intersection);
                    }
                }
            }
        }
        
        /// <summary>
        /// Calculate intersection point of two line segments
        /// Returns null if no intersection
        /// </summary>
        private VisualIntersection LineSegmentIntersection(LineSegment2D a, LineSegment2D b)
        {
            // Line A: P1 + t * (P2 - P1)
            // Line B: P3 + s * (P4 - P3)
            
            double x1 = a.X1, y1 = a.Y1;
            double x2 = a.X2, y2 = a.Y2;
            double x3 = b.X1, y3 = b.Y1;
            double x4 = b.X2, y4 = b.Y2;
            
            double denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
            
            // Parallel lines
            if (Math.Abs(denom) < 1e-10)
                return null;
            
            double t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom;
            double s = ((x1 - x3) * (y1 - y2) - (y1 - y3) * (x1 - x2)) / denom;
            
            // Check if intersection is within both segments
            if (t >= 0 && t <= 1 && s >= 0 && s <= 1)
            {
                return new VisualIntersection
                {
                    X = x1 + t * (x2 - x1),
                    Y = y1 + t * (y2 - y1)
                };
            }
            
            return null;
        }
        
        /// <summary>
        /// Check if a bounding box collides with OTHER ANNOTATIONS (occupied areas) only.
        /// The header (shelf) of an annotation can visually overlap model elements in 3D views,
        /// but must not overlap other annotation headers.
        /// Model element intersections are handled by HasLeaderCollision (leader line check).
        /// </summary>
        public bool HasHeaderCollision(BBox2D bbox)
        {
            // Check against occupied areas (other annotations) ONLY
            foreach (var occupied in _occupiedAreas)
            {
                if (bbox.Intersects(occupied))
                    return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// Check if an L-shaped leader line collides with any elements
        /// Leader line has: vertical segment from element to elbow, then horizontal to head
        /// </summary>
        public bool HasLeaderCollision(double leaderEndX, double leaderEndY, 
                                        double elbowX, double elbowY,
                                        double headX, double headY,
                                        long? excludeElementId = null)
        {
            return HasLeaderCollision(leaderEndX, leaderEndY, elbowX, elbowY, headX, headY, 
                                      excludeElementId, checkModelElements: true);
        }
        
        /// <summary>
        /// Check if an L-shaped leader line collides with any elements
        /// Leader line has: vertical segment from element to elbow, then horizontal to head
        /// For 3D views: set checkModelElements=false because leader lines naturally cross
        /// model element projections in 3D - only check annotation-to-annotation collisions.
        /// </summary>
        public bool HasLeaderCollision(double leaderEndX, double leaderEndY, 
                                        double elbowX, double elbowY,
                                        double headX, double headY,
                                        long? excludeElementId, bool checkModelElements)
        {
            double elementMargin = 0.02; // 6mm margin around elements
            double annotationMargin = 0.01; // 3mm margin around annotations
            
            // Only check model element line segments in 2D views.
            // In 3D views, leader lines naturally cross model element projections
            // (ducts, pipes, etc.) - this is expected and not a collision.
            if (checkModelElements)
            {
                foreach (var seg in _lineSegments)
                {
                    if (excludeElementId.HasValue && seg.ElementId == excludeElementId.Value)
                        continue;
                            
                    // Check vertical segment (leader end to elbow)
                    if (LineSegmentsIntersectWithMargin(leaderEndX, leaderEndY, elbowX, elbowY,
                                                         seg.X1, seg.Y1, seg.X2, seg.Y2, elementMargin))
                        return true;
                            
                    // Check horizontal segment (elbow to head)
                    if (LineSegmentsIntersectWithMargin(elbowX, elbowY, headX, headY,
                                                         seg.X1, seg.Y1, seg.X2, seg.Y2, elementMargin))
                        return true;
                }
            }
                    
            // Check against occupied areas (other annotations)
            // Create line segments for the leader
            var verticalSegment = new LineSegment2D
            {
                X1 = leaderEndX, Y1 = leaderEndY,
                X2 = elbowX, Y2 = elbowY
            };
            var horizontalSegment = new LineSegment2D
            {
                X1 = elbowX, Y1 = elbowY,
                X2 = headX, Y2 = headY
            };
                    
            foreach (var occupied in _occupiedAreas)
            {
                // Expand the occupied area by margin
                var expanded = occupied.Expand(annotationMargin);
                        
                // Check if leader segments intersect the occupied area
                if (SegmentIntersectsBBox(verticalSegment, expanded) ||
                    SegmentIntersectsBBox(horizontalSegment, expanded))
                    return true;
            }
        
            return false;
        }
                
        /// <summary>
        /// Check if a line segment intersects a bounding box
        /// </summary>
        private bool SegmentIntersectsBBox(LineSegment2D seg, BBox2D bbox)
        {
            // Check if segment is completely outside bbox
            if (seg.X1 < bbox.MinX && seg.X2 < bbox.MinX) return false;
            if (seg.X1 > bbox.MaxX && seg.X2 > bbox.MaxX) return false;
            if (seg.Y1 < bbox.MinY && seg.Y2 < bbox.MinY) return false;
            if (seg.Y1 > bbox.MaxY && seg.Y2 > bbox.MaxY) return false;
                    
            // Check if segment endpoints are inside bbox
            if (bbox.Contains(seg.X1, seg.Y1) || bbox.Contains(seg.X2, seg.Y2))
                return true;
                    
            // Check if segment intersects any edge of bbox
            // Top edge
            if (LineSegmentsIntersect(seg.X1, seg.Y1, seg.X2, seg.Y2,
                                      bbox.MinX, bbox.MaxY, bbox.MaxX, bbox.MaxY))
                return true;
            // Bottom edge
            if (LineSegmentsIntersect(seg.X1, seg.Y1, seg.X2, seg.Y2,
                                      bbox.MinX, bbox.MinY, bbox.MaxX, bbox.MinY))
                return true;
            // Left edge
            if (LineSegmentsIntersect(seg.X1, seg.Y1, seg.X2, seg.Y2,
                                      bbox.MinX, bbox.MinY, bbox.MinX, bbox.MaxY))
                return true;
            // Right edge
            if (LineSegmentsIntersect(seg.X1, seg.Y1, seg.X2, seg.Y2,
                                      bbox.MaxX, bbox.MinY, bbox.MaxX, bbox.MaxY))
                return true;
                    
            return false;
        }
        
        /// <summary>
        /// Check if two line segments intersect with a margin around the first segment
        /// </summary>
        private bool LineSegmentsIntersectWithMargin(double x1, double y1, double x2, double y2,
                                                      double x3, double y3, double x4, double y4,
                                                      double margin)
        {
            // First check actual intersection
            if (LineSegmentsIntersect(x1, y1, x2, y2, x3, y3, x4, y4))
                return true;
            
            // Check if any endpoint of the element segment is within margin of the leader segment
            double dist1 = PointToSegmentDistance(x3, y3, x1, y1, x2, y2);
            double dist2 = PointToSegmentDistance(x4, y4, x1, y1, x2, y2);
            
            return dist1 < margin || dist2 < margin;
        }
        
        /// <summary>
        /// Calculate distance from point to line segment
        /// </summary>
        private double PointToSegmentDistance(double px, double py, double x1, double y1, double x2, double y2)
        {
            double dx = x2 - x1;
            double dy = y2 - y1;
            double lengthSq = dx * dx + dy * dy;
            
            if (lengthSq < 1e-10)
                return Math.Sqrt((px - x1) * (px - x1) + (py - y1) * (py - y1));
            
            double t = Math.Max(0, Math.Min(1, ((px - x1) * dx + (py - y1) * dy) / lengthSq));
            double projX = x1 + t * dx;
            double projY = y1 + t * dy;
            
            return Math.Sqrt((px - projX) * (px - projX) + (py - projY) * (py - projY));
        }
        
        /// <summary>
        /// Check if a bounding box intersects a line segment
        /// </summary>
        private bool BBoxIntersectsLine(BBox2D bbox, LineSegment2D line, double margin)
        {
            // Expand bbox by margin
            var expandedBbox = bbox.Expand(margin);
            
            // Check both endpoints
            bool p1Inside = expandedBbox.Contains(line.X1, line.Y1);
            bool p2Inside = expandedBbox.Contains(line.X2, line.Y2);
            
            if (p1Inside || p2Inside)
                return true;
            
            // Check if line crosses any bbox edge
            // Top edge
            if (LineSegmentsIntersect(line.X1, line.Y1, line.X2, line.Y2,
                                      expandedBbox.MinX, expandedBbox.MaxY,
                                      expandedBbox.MaxX, expandedBbox.MaxY))
                return true;
            
            // Bottom edge
            if (LineSegmentsIntersect(line.X1, line.Y1, line.X2, line.Y2,
                                      expandedBbox.MinX, expandedBbox.MinY,
                                      expandedBbox.MaxX, expandedBbox.MinY))
                return true;
            
            // Left edge
            if (LineSegmentsIntersect(line.X1, line.Y1, line.X2, line.Y2,
                                      expandedBbox.MinX, expandedBbox.MinY,
                                      expandedBbox.MinX, expandedBbox.MaxY))
                return true;
            
            // Right edge
            if (LineSegmentsIntersect(line.X1, line.Y1, line.X2, line.Y2,
                                      expandedBbox.MaxX, expandedBbox.MinY,
                                      expandedBbox.MaxX, expandedBbox.MaxY))
                return true;
            
            return false;
        }
        
        private bool LineSegmentsIntersect(double x1, double y1, double x2, double y2,
                                           double x3, double y3, double x4, double y4)
        {
            double denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
            if (Math.Abs(denom) < 1e-10) return false;
            
            double t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom;
            double s = ((x1 - x3) * (y1 - y2) - (y1 - y3) * (x1 - x2)) / denom;
            
            return t >= 0 && t <= 1 && s >= 0 && s <= 1;
        }
        
        /// <summary>
        /// Register an occupied area (placed annotation)
        /// </summary>
        public void AddOccupiedArea(BBox2D bbox)
        {
            _occupiedAreas.Add(bbox);
        }
        
        /// <summary>
        /// Register annotation leader lines for collision detection.
        /// Per the user's algorithm: after an annotation is successfully placed,
        /// its leader lines (vertical + horizontal segments) and bbox edges must be
        /// added to the collision check list so subsequent annotations avoid overlapping.
        /// </summary>
        public void RegisterAnnotationLines(
            double leaderEndX, double leaderEndY,
            double elbowX, double elbowY,
            double headX, double headY,
            long elementId)
        {
            // Vertical segment: LeaderEnd → Elbow
            _lineSegments.Add(new LineSegment2D
            {
                X1 = leaderEndX, Y1 = leaderEndY,
                X2 = elbowX, Y2 = elbowY,
                ElementId = elementId
            });
            
            // Horizontal segment: Elbow → Head (the shelf portion covered by the leader)
            _lineSegments.Add(new LineSegment2D
            {
                X1 = elbowX, Y1 = elbowY,
                X2 = headX, Y2 = headY,
                ElementId = elementId
            });
        }
        
        /// <summary>
        /// Get all visual intersections
        /// </summary>
        public List<VisualIntersection> GetIntersections()
        {
            return _intersections.ToList();
        }
        
        /// <summary>
        /// Get intersections near a point
        /// </summary>
        public List<VisualIntersection> GetIntersectionsNear(double x, double y, double radius)
        {
            return _intersections
                .Where(i => Math.Sqrt(Math.Pow(i.X - x, 2) + Math.Pow(i.Y - y, 2)) <= radius)
                .ToList();
        }
        
        /// <summary>
        /// Clear all data
        /// </summary>
        public void Clear()
        {
            _lineSegments.Clear();
            _occupiedAreas.Clear();
            _intersections.Clear();
        }
    }
}
