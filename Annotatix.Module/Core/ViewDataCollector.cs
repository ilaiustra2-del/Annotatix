using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Collects data from a Revit view for snapshot export
    /// </summary>
    public class ViewDataCollector
    {
        private readonly Document _document;
        private readonly View _view;
        private readonly UIView _uiView;

        public ViewDataCollector(Document document, View view, UIView uiView)
        {
            _document = document;
            _view = view;
            _uiView = uiView;
        }

        /// <summary>
        /// Collect a complete snapshot of the view
        /// </summary>
        public ViewSnapshot CollectSnapshot(string sessionId, string snapshotType)
        {
            // Log view details for debugging
            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] View type: {_view.ViewType}, Name: {_view.Name}, Id: {_view.Id.Value}");
            
            var snapshot = new ViewSnapshot
            {
                SessionId = sessionId,
                DocumentName = _document.Title,
                ViewId = _view.Id.Value,
                ViewName = _view.Name,
                Timestamp = DateTime.Now,
                SnapshotType = snapshotType
            };

            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Collecting {snapshotType} snapshot for view: {_view.Name}");

            // Collect elements
            var elements = GetElementsOnView();
            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] GetElementsOnView returned {elements.Count} elements");
            
            foreach (var element in elements)
            {
                var elementData = CreateElementData(element);
                if (elementData != null)
                {
                    snapshot.Elements.Add(elementData);
                }
            }

            // Collect annotations
            var annotations = GetAnnotationsOnView();
            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] GetAnnotationsOnView returned {annotations.Count} annotations");
            
            foreach (var annotation in annotations)
            {
                var annotationData = CreateAnnotationData(annotation);
                if (annotationData != null)
                {
                    snapshot.Annotations.Add(annotationData);
                }
            }

            // Collect systems
            var systems = GetSystemsOnView(elements);
            foreach (var system in systems)
            {
                snapshot.Systems.Add(system);
            }

            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Collected {snapshot.Elements.Count} elements, {snapshot.Annotations.Count} annotations, {snapshot.Systems.Count} systems");

            return snapshot;
        }

        /// <summary>
        /// Get all model elements visible on the view
        /// </summary>
        private List<Element> GetElementsOnView()
        {
            var elements = new List<Element>();
            
            try
            {
                // Method 1: Use basic collector for the view
                var collector = new FilteredElementCollector(_document, _view.Id)
                    .WhereElementIsNotElementType();

                int count = 0;
                int skippedAnn = 0;
                int skippedNull = 0;
                var skippedCategories = new System.Collections.Generic.List<string>();
                foreach (var element in collector)
                {
                    // Skip annotation elements (handled separately) - but only skip if they are tags/notes/dims
                    if (element is IndependentTag || element is SpatialElementTag || element is TextNote || element is Dimension)
                    {
                        skippedAnn++;
                        skippedCategories.Add($"{element.Category?.Name ?? "null"}({element.Id.Value})");
                        continue;
                    }
                    
                    // Skip view-based detail items (also annotations)
                    if (element is FamilyInstance fiCheck && fiCheck.Symbol?.Family?.FamilyPlacementType == FamilyPlacementType.ViewBased)
                    {
                        skippedAnn++;
                        skippedCategories.Add($"{element.Category?.Name ?? "null"}({element.Id.Value})");
                        continue;
                    }
                    
                    // Skip null categories
                    if (element.Category == null)
                    {
                        skippedNull++;
                        continue;
                    }

                    elements.Add(element);
                    count++;
                }
                
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] GetElementsOnView: {count} elements, skipped {skippedAnn} annotations, {skippedNull} null categories");
                if (skippedCategories.Count > 0)
                {
                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Skipped annotation categories: {string.Join(", ", skippedCategories)}");
                }
                
                // If no elements found, try document-wide collector as fallback
                if (count == 0)
                {
                    DebugLogger.Log("[ANNOTATIX-COLLECTOR] No elements on view, trying document-wide collector...");
                    var docCollector = new FilteredElementCollector(_document)
                        .WhereElementIsNotElementType();
                    
                    int docCount = 0;
                    foreach (var element in docCollector)
                    {
                        if (element.Category != null && !IsAnnotationElement(element))
                        {
                            // Check if element is not hidden in this view
                            try
                            {
                                if (!element.IsHidden(_view))
                                {
                                    elements.Add(element);
                                    docCount++;
                                }
                            }
                            catch
                            {
                                // Element might not support IsHidden, skip it
                            }
                        }
                    }
                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Document-wide found {docCount} visible elements");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting elements: {ex.Message}");
            }

            return elements;
        }

        /// <summary>
        /// Get all annotation elements on the view
        /// </summary>
        private List<Element> GetAnnotationsOnView()
        {
            var annotations = new List<Element>();

            try
            {
                // Collect ALL elements on view first to see what's there
                var allOnView = new FilteredElementCollector(_document, _view.Id)
                    .WhereElementIsNotElementType();
                
                int totalOnView = 0;
                foreach (var elem in allOnView) totalOnView++;
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Total elements on view: {totalOnView}");

                // Collect ALL elements and classify by type - don't rely on category IDs
                var allCollector = new FilteredElementCollector(_document, _view.Id)
                    .WhereElementIsNotElementType();
                
                int tagCount = 0;
                int textCount = 0;
                int dimCount = 0;
                int detailCount = 0;
                var foundCategories = new System.Collections.Generic.HashSet<string>();
                
                foreach (var elem in allCollector)
                {
                    if (elem.Category != null)
                    {
                        foundCategories.Add($"{elem.Category.Name}({elem.Category.Id.Value})");
                    }
                    
                    // Check by element type first (most reliable)
                    if (elem is IndependentTag || elem is SpatialElementTag)
                    {
                        annotations.Add(elem);
                        tagCount++;
                        continue;
                    }
                    
                    // Text notes
                    if (elem is TextNote)
                    {
                        annotations.Add(elem);
                        textCount++;
                        continue;
                    }
                    
                    // Dimensions
                    if (elem is Dimension)
                    {
                        annotations.Add(elem);
                        dimCount++;
                        continue;
                    }
                    
                    // Detail items (view-based families)
                    if (elem is FamilyInstance fi && fi.Symbol?.Family?.FamilyPlacementType == FamilyPlacementType.ViewBased)
                    {
                        annotations.Add(elem);
                        detailCount++;
                        continue;
                    }
                }
                
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Categories found: {string.Join(", ", foundCategories)}");
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Found {tagCount} tags, {textCount} text notes, {dimCount} dimensions, {detailCount} detail items");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting annotations: {ex.Message}");
            }


            return annotations;
        }

        /// <summary>
        /// Check if an element is an annotation element
        /// </summary>
        private bool IsAnnotationElement(Element element)
        {
            var category = element.Category;
            if (category == null) return false;
        
            var categoryId = category.Id.Value;
        
            // Annotation categories - only these specific categories should be skipped
            // Tags, text notes, dimensions are annotations and handled separately
            // Note: In Revit, almost all categories have negative IDs, so we can't use categoryId < 0
                    
            // Check if this is a tag element (various tag categories)
            if (element is IndependentTag || element is SpatialElementTag)
                return true;
                    
            // Check specific annotation category IDs
            return categoryId == -1010500 || // OST_DuTags (generic tags)
                   categoryId == -1005500 || // OST_TextNotes
                   categoryId == -1005300 || // OST_Dimensions
                   categoryId == -2001000 || // OST_PipeTags
                   categoryId == -2000950 || // OST_DuctTags
                   categoryId == -2001200 || // OST_MechanicalEquipmentTags
                   categoryId == -2001100 || // OST_CableTrayTags
                   categoryId == -2001150 || // OST_ConduitTags
                   categoryId == -133204;    // OST_MaterialTags
        }

        /// <summary>
        /// Create ElementData from a Revit element
        /// </summary>
        private ElementData CreateElementData(Element element)
        {
            try
            {
                // Filter out Grids and Center lines
                string categoryName = element.Category?.Name ?? "";
                if (categoryName == "Оси" || categoryName == "Осевая линия" || 
                    categoryName == "Grids" || categoryName == "Center Line")
                {
                    return null;
                }
                
                var data = new ElementData
                {
                    ElementId = element.Id.Value,
                    Category = categoryName
                };

                // Family and type info
                if (element is FamilyInstance instance)
                {
                    data.FamilyName = instance.Symbol?.FamilyName ?? instance.Symbol?.Family?.Name ?? "";
                    data.TypeName = instance.Symbol?.Name ?? "";
                }
                else
                {
                    data.FamilyName = element.GetType().Name;
                    data.TypeName = element.Name;
                }

                // Get start/end coordinates for linear elements (pipes, ducts, etc.)
                var lineCoords = GetLinearElementCoordinates(element);
                if (lineCoords != null)
                {
                    data.ModelStart = lineCoords.Item1;
                    data.ModelEnd = lineCoords.Item2;
                    data.HasEndPoint = lineCoords.Item3;
                    
                    // Convert to view coordinates
                    data.ViewStart = ConvertToViewCoordinates(lineCoords.Item1);
                    if (data.HasEndPoint)
                    {
                        data.ViewEnd = ConvertToViewCoordinates(lineCoords.Item2);
                    }
                }
                else
                {
                    // Single point element
                    var modelCoords = GetModelCoordinates(element);
                    if (modelCoords != null)
                    {
                        data.ModelStart = modelCoords;
                        data.ViewStart = GetViewCoordinates(element) ?? new Coordinates2D();
                        data.HasEndPoint = false;
                    }
                }

                // System information for MEP elements
                var systemInfo = GetElementSystemInfo(element);
                if (systemInfo != null)
                {
                    data.SystemId = systemInfo.Item1;
                    data.SystemName = systemInfo.Item2;
                    data.BelongTo = $"system:{systemInfo.Item2}";
                }
                                
                // Size dimensions for pipes and ducts
                ExtractElementDimensions(element, data);
                
                return data;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error creating element data: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create AnnotationData from a Revit annotation element
        /// </summary>
        private AnnotationData CreateAnnotationData(Element annotation)
        {
            try
            {
                var data = new AnnotationData
                {
                    ElementId = annotation.Id.Value,
                    Category = annotation.Category?.Name ?? "Unknown"
                };

                // Family and type info
                if (annotation is FamilyInstance instance)
                {
                    data.FamilyName = instance.Symbol?.FamilyName ?? instance.Symbol?.Family?.Name ?? "";
                    data.TypeName = instance.Symbol?.Name ?? "";
                }
                else if (annotation is IndependentTag tag)
                {
                    // For IndependentTag, get the tag type name from the document
                    ElementId typeId = tag.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        Element tagType = _document.GetElement(typeId);
                        if (tagType != null)
                        {
                            data.TypeName = tagType.Name;
                            // FamilyName for tags is typically the category name
                            data.FamilyName = tagType.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsString() ?? "";
                            if (string.IsNullOrEmpty(data.FamilyName))
                            {
                                // Try to get from FamilySymbol property
                                if (tagType is FamilySymbol fs)
                                {
                                    data.FamilyName = fs.FamilyName;
                                }
                            }
                        }
                    }
                    // Store the tag text for reference
                    data.TagText = tag.TagText;
                }
                else
                {
                    data.FamilyName = annotation.GetType().Name;
                    data.TypeName = annotation.Name;
                }

                // Get head and leader positions for tags
                if (annotation is IndependentTag indTag)
                {
                    // Head position (where the tag text is) - in MODEL coordinates
                    XYZ headPoint = indTag.TagHeadPosition;
                    
                    // Save model coordinates (full 3D)
                    data.HeadModelPosition = new Coordinates3D 
                    { 
                        X = headPoint.X, 
                        Y = headPoint.Y, 
                        Z = headPoint.Z 
                    };
                    
                    // Convert to view coordinates (2D projection)
                    data.HeadViewPosition = ConvertToViewCoordinates(new Coordinates3D 
                    { 
                        X = headPoint.X, 
                        Y = headPoint.Y, 
                        Z = headPoint.Z 
                    });
                    
                    // Legacy compatibility
                    data.HeadPosition = new Coordinates2D { X = data.HeadViewPosition.X, Y = data.HeadViewPosition.Y };
                                    
                    // Log coordinates for debugging
                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Tag {annotation.Id}: HeadModel=({headPoint.X:F2}, {headPoint.Y:F2}, {headPoint.Z:F2}), HeadView=({data.HeadViewPosition.X:F2}, {data.HeadViewPosition.Y:F2})");
                                    
                    // Get leader information
                    data.HasLeader = indTag.HasLeader;
                                                        
                    // Leader type (Free vs Attached)
                    var leaderEndCondition = indTag.LeaderEndCondition;
                    data.LeaderType = leaderEndCondition.ToString(); // "Free" or "Attached"
                                                        
                    // Orientation
                    data.Orientation = indTag.TagOrientation.ToString(); // "Horizontal" or "Vertical"
                                        
                    // Get tagged elements for leader position
                    var taggedIds = indTag.GetTaggedLocalElementIds();
                    if (taggedIds != null && taggedIds.Count > 0)
                    {
                        var taggedElem = _document.GetElement(taggedIds.First());
                        if (taggedElem != null)
                        {
                            XYZ leaderEndModelPoint = null;
                            Reference elemRef = null;
                            
                            // CRITICAL: Get the ACTUAL reference that the tag uses
                            // GetTaggedReferences returns the references the tag is attached to
                            var taggedRefs = indTag.GetTaggedReferences();
                            if (taggedRefs != null && taggedRefs.Count > 0)
                            {
                                elemRef = taggedRefs[0]; // Use the first tagged reference
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Got tagged reference: {elemRef.ElementId}, ElementReferenceType={elemRef.ElementReferenceType}");
                            }
                            
                            // Fallback: try to get reference from element
                            if (elemRef == null)
                            {
                                try
                                {
                                    if (taggedElem is MEPCurve mepCurve)
                                    {
                                        LocationCurve lc = mepCurve.Location as LocationCurve;
                                        if (lc != null && lc.Curve != null)
                                        {
                                            elemRef = lc.Curve.Reference;
                                        }
                                    }
                                    if (elemRef == null)
                                    {
                                        elemRef = new Reference(taggedElem);
                                    }
                                }
                                catch (Exception refEx)
                                {
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get element reference: {refEx.Message}");
                                }
                            }
                            
                            // Get leader end position
                            // For Free leaders: GetLeaderEnd returns where the arrow points
                            // For Attached leaders: the end is attached to the element
                            if (leaderEndCondition == LeaderEndCondition.Free && elemRef != null)
                            {
                                try
                                {
                                    leaderEndModelPoint = indTag.GetLeaderEnd(elemRef);
                                    if (leaderEndModelPoint != null)
                                    {
                                        data.LeaderEndModel = new Coordinates3D
                                        {
                                            X = leaderEndModelPoint.X,
                                            Y = leaderEndModelPoint.Y,
                                            Z = leaderEndModelPoint.Z
                                        };
                                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Free leader end from GetLeaderEnd: ({leaderEndModelPoint.X:F2}, {leaderEndModelPoint.Y:F2}, {leaderEndModelPoint.Z:F2})");
                                    }
                                }
                                catch (Exception leaderEx)
                                {
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get Free leader end: {leaderEx.Message}");
                                }
                            }
                            else if (leaderEndCondition == LeaderEndCondition.Attached)
                            {
                                // For attached leader, get the attachment point
                                try
                                {
                                    if (elemRef != null)
                                    {
                                        leaderEndModelPoint = indTag.GetLeaderEnd(elemRef);
                                    }
                                }
                                catch
                                {
                                    // Attached leaders may not have GetLeaderEnd, use element location
                                }
                                
                                // Fallback: use element location
                                if (leaderEndModelPoint == null)
                                {
                                    Location elemLoc = taggedElem.Location;
                                    if (elemLoc is LocationPoint locPoint)
                                    {
                                        leaderEndModelPoint = locPoint.Point;
                                    }
                                    else if (elemLoc is LocationCurve locCurve)
                                    {
                                        leaderEndModelPoint = locCurve.Curve.Evaluate(0.5, true);
                                    }
                                }
                                
                                if (leaderEndModelPoint != null)
                                {
                                    data.LeaderEndModel = new Coordinates3D
                                    {
                                        X = leaderEndModelPoint.X,
                                        Y = leaderEndModelPoint.Y,
                                        Z = leaderEndModelPoint.Z
                                    };
                                }
                            }
                            
                            // If still no leader end, compute it from element geometry
                            if (data.LeaderEndModel == null)
                            {
                                // For Free leader without elbow, calculate the closest point on element
                                // This is where the leader line should point to
                                XYZ computedLeaderEnd = null;
                                
                                try
                                {
                                    // Get element geometry to find closest point
                                    if (taggedElem is MEPCurve mepCurve)
                                    {
                                        // For pipes/ducts, get closest point on curve
                                        LocationCurve lc = mepCurve.Location as LocationCurve;
                                        if (lc != null && lc.Curve != null)
                                        {
                                            computedLeaderEnd = lc.Curve.Project(headPoint).XYZPoint;
                                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Computed leader end from MEPCurve projection: ({computedLeaderEnd.X:F2}, {computedLeaderEnd.Y:F2}, {computedLeaderEnd.Z:F2})");
                                        }
                                    }
                                    else if (taggedElem.Location is LocationCurve lc2 && lc2.Curve != null)
                                    {
                                        // For other curve-based elements
                                        computedLeaderEnd = lc2.Curve.Project(headPoint).XYZPoint;
                                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Computed leader end from LocationCurve projection: ({computedLeaderEnd.X:F2}, {computedLeaderEnd.Y:F2}, {computedLeaderEnd.Z:F2})");
                                    }
                                    else if (taggedElem.Location is LocationPoint lp)
                                    {
                                        // For point-based elements, use the location point
                                        computedLeaderEnd = lp.Point;
                                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Using element location point as leader end: ({computedLeaderEnd.X:F2}, {computedLeaderEnd.Y:F2}, {computedLeaderEnd.Z:F2})");
                                    }
                                    else
                                    {
                                        // Try to get geometry and find closest point
                                        var geomElem = taggedElem.get_Geometry(new Options());
                                        if (geomElem != null)
                                        {
                                            double minDist = double.MaxValue;
                                            foreach (var geomObj in geomElem)
                                            {
                                                if (geomObj is Curve curve)
                                                {
                                                    var proj = curve.Project(headPoint);
                                                    if (proj != null && proj.Distance < minDist)
                                                    {
                                                        minDist = proj.Distance;
                                                        computedLeaderEnd = proj.XYZPoint;
                                                    }
                                                }
                                                else if (geomObj is Solid solid && solid.Volume > 0)
                                                {
                                                    // Find closest point on solid faces
                                                    foreach (Face face in solid.Faces)
                                                    {
                                                        var proj = face.Project(headPoint);
                                                        if (proj != null && proj.Distance < minDist)
                                                        {
                                                            minDist = proj.Distance;
                                                            computedLeaderEnd = proj.XYZPoint;
                                                        }
                                                    }
                                                }
                                            }
                                            if (computedLeaderEnd != null)
                                            {
                                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Computed leader end from geometry projection: ({computedLeaderEnd.X:F2}, {computedLeaderEnd.Y:F2}, {computedLeaderEnd.Z:F2})");
                                            }
                                        }
                                    }
                                }
                                catch (Exception geomEx)
                                {
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not compute leader end from geometry: {geomEx.Message}");
                                }
                                
                                // Use computed point or fallback to head position
                                if (computedLeaderEnd != null)
                                {
                                    data.LeaderEndModel = new Coordinates3D
                                    {
                                        X = computedLeaderEnd.X,
                                        Y = computedLeaderEnd.Y,
                                        Z = computedLeaderEnd.Z
                                    };
                                }
                                else
                                {
                                    data.LeaderEndModel = new Coordinates3D
                                    {
                                        X = headPoint.X,
                                        Y = headPoint.Y,
                                        Z = headPoint.Z
                                    };
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Using head as leader end fallback");
                                }
                            }
                            
                            // Convert leader end to view coordinates
                            data.LeaderEndView = ConvertToViewCoordinates(data.LeaderEndModel);
                            data.LeaderEnd = new Coordinates2D 
                            { 
                                X = data.LeaderEndView.X, 
                                Y = data.LeaderEndView.Y 
                            };
                            
                            // Get elbow position (break point)
                            if (elemRef != null)
                            {
                                try
                                {
                                    // First check if there IS an elbow using HasLeaderElbow
                                    bool hasElbow = indTag.HasLeaderElbow(elemRef);
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] HasLeaderElbow check: {hasElbow}");
                                    
                                    if (hasElbow)
                                    {
                                        XYZ elbowPoint = indTag.GetLeaderElbow(elemRef);
                                        if (elbowPoint != null)
                                        {
                                            data.ElbowModelPosition = new Coordinates3D
                                            {
                                                X = elbowPoint.X,
                                                Y = elbowPoint.Y,
                                                Z = elbowPoint.Z
                                            };
                                            data.ElbowViewPosition = ConvertToViewCoordinates(data.ElbowModelPosition);
                                            data.ElbowPosition = new Coordinates2D 
                                            { 
                                                X = data.ElbowViewPosition.X, 
                                                Y = data.ElbowViewPosition.Y 
                                            };
                                            data.HasElbow = true;
                                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Elbow model: ({elbowPoint.X:F2}, {elbowPoint.Y:F2}, {elbowPoint.Z:F2})");
                                        }
                                    }
                                }
                                catch (Exception elbowEx)
                                {
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get elbow: {elbowEx.Message}");
                                }
                            }
                        }
                    }
                    
                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Tag: HeadModel=({data.HeadModelPosition?.X:F2}, {data.HeadModelPosition?.Y:F2}, {data.HeadModelPosition?.Z:F2}), LeaderEndModel=({data.LeaderEndModel?.X:F2}, {data.LeaderEndModel?.Y:F2}, {data.LeaderEndModel?.Z:F2}), Type={data.LeaderType}, HasElbow={data.HasElbow}");
                }
                // Handle SpotDimension (elevation marks / высотные отметки)
                else if (annotation is SpotDimension spotDim)
                {
                    data.FamilyName = "SpotDimension";
                    data.TypeName = spotDim.SpotDimensionType?.Name ?? "SpotDimension";
                    
                    // SpotDimension may not have a valid Location, use References instead
                    XYZ headPoint = null;
                    XYZ leaderEndPoint = null;
                    
                    // Debug: log what SpotDimension properties are available
                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension analysis: HasLeader={spotDim.HasLeader}, View={spotDim.View?.Name}");
                    
                    // Log ALL SpotDimension properties for analysis
                    try
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] === SPOTDIMENSION FULL PROPERTY DUMP ===");
                        
                        // Location property
                        if (spotDim.Location is LocationPoint locPt)
                        {
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Location (LocationPoint): ({locPt.Point.X:F4}, {locPt.Point.Y:F4}, {locPt.Point.Z:F4})");
                        }
                        else if (spotDim.Location != null)
                        {
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Location type: {spotDim.Location.GetType().Name}");
                        }
                        else
                        {
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Location: null");
                        }
                        
                        // Origin property
                        try
                        {
                            XYZ originPt = spotDim.Origin;
                            if (originPt != null)
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Origin: ({originPt.X:F4}, {originPt.Y:F4}, {originPt.Z:F4})");
                            else
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Origin: null");
                        }
                        catch (Exception ex) { DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Origin error: {ex.Message}"); }
                        
                        // LeaderEndPosition
                        try
                        {
                            XYZ leaderEnd = spotDim.LeaderEndPosition;
                            if (leaderEnd != null)
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderEndPosition: ({leaderEnd.X:F4}, {leaderEnd.Y:F4}, {leaderEnd.Z:F4})");
                            else
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderEndPosition: null");
                        }
                        catch (Exception ex) { DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderEndPosition error: {ex.Message}"); }
                        
                        // LeaderShoulderPosition
                        try
                        {
                            if (spotDim.LeaderHasShoulder)
                            {
                                XYZ shoulder = spotDim.LeaderShoulderPosition;
                                if (shoulder != null)
                                {
                                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderShoulderPosition: ({shoulder.X:F4}, {shoulder.Y:F4}, {shoulder.Z:F4})");
                                    // Save LeaderShoulderPosition for proper placement
                                    data.LeaderShoulderModel = new Coordinates3D
                                    {
                                        X = shoulder.X,
                                        Y = shoulder.Y,
                                        Z = shoulder.Z
                                    };
                                }
                            }
                            else
                            {
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderHasShoulder: False");
                            }
                        }
                        catch (Exception ex) { DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderShoulderPosition error: {ex.Message}"); }
                        
                        // TextPosition
                        try
                        {
                            XYZ textPos = spotDim.TextPosition;
                            if (textPos != null)
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   TextPosition: ({textPos.X:F4}, {textPos.Y:F4}, {textPos.Z:F4})");
                            else
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   TextPosition: null");
                        }
                        catch (Exception ex) { DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   TextPosition error: {ex.Message}"); }
                        
                        // Curve (dimension line)
                        try
                        {
                            Curve dimCurve = spotDim.Curve;
                            if (dimCurve != null)
                            {
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Curve: Start=({dimCurve.GetEndPoint(0).X:F4}, {dimCurve.GetEndPoint(0).Y:F4}, {dimCurve.GetEndPoint(0).Z:F4}), End=({dimCurve.GetEndPoint(1).X:F4}, {dimCurve.GetEndPoint(1).Y:F4}, {dimCurve.GetEndPoint(1).Z:F4})");
                            }
                        }
                        catch (Exception ex) { DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Curve error: {ex.Message}"); }
                        
                        // Value
                        try
                        {
                            double? val = spotDim.Value;
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Value: {val}");
                        }
                        catch (Exception ex) { DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   Value error: {ex.Message}"); }
                        
                        // DimensionShape
                        try
                        {
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   DimensionShape: {spotDim.DimensionShape}");
                        }
                        catch { }
                        
                        // BoundingBox
                        try
                        {
                            var bbox = spotDim.get_BoundingBox(spotDim.View);
                            if (bbox != null)
                            {
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   BoundingBox Min: ({bbox.Min.X:F4}, {bbox.Min.Y:F4}, {bbox.Min.Z:F4})");
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   BoundingBox Max: ({bbox.Max.X:F4}, {bbox.Max.Y:F4}, {bbox.Max.Z:F4})");
                            }
                        }
                        catch { }
                        
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] === END PROPERTY DUMP ===");
                    }
                    catch (Exception dumpEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Property dump error: {dumpEx.Message}");
                    }
                    
                    // Try to get position from References first (most reliable for SpotDimension)
                    try
                    {
                        var refs = spotDim.References;
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension References: Size={refs?.Size ?? -1}");
                        if (refs != null && refs.Size > 0)
                        {
                            var firstRef = refs.get_Item(0);
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] FirstRef: ElementId={firstRef?.ElementId}, ElementReferenceType={firstRef?.ElementReferenceType}, GlobalPoint={firstRef?.GlobalPoint}");
                            if (firstRef != null && firstRef.GlobalPoint != null)
                            {
                                leaderEndPoint = firstRef.GlobalPoint;
                                data.LeaderEndModel = new Coordinates3D
                                {
                                    X = leaderEndPoint.X,
                                    Y = leaderEndPoint.Y,
                                    Z = leaderEndPoint.Z
                                };
                                data.LeaderEndView = ConvertToViewCoordinates(data.LeaderEndModel);
                                data.LeaderEnd = new Coordinates2D 
                                { 
                                    X = data.LeaderEndView.X, 
                                    Y = data.LeaderEndView.Y 
                                };
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension LeaderEnd from GlobalPoint: ({leaderEndPoint.X:F2}, {leaderEndPoint.Y:F2}, {leaderEndPoint.Z:F2})");
                            }
                            else if (firstRef != null)
                            {
                                // Reference exists but GlobalPoint is null - get element and use its location
                                var refElem = _document.GetElement(firstRef.ElementId);
                                if (refElem != null)
                                {
                                    if (refElem.Location is LocationPoint lp)
                                    {
                                        leaderEndPoint = lp.Point;
                                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension LeaderEnd from referenced element location: ({leaderEndPoint.X:F2}, {leaderEndPoint.Y:F2}, {leaderEndPoint.Z:F2})");
                                    }
                                    else if (refElem.Location is LocationCurve lc && lc.Curve != null)
                                    {
                                        leaderEndPoint = lc.Curve.Evaluate(0.5, true);
                                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension LeaderEnd from referenced curve midpoint: ({leaderEndPoint.X:F2}, {leaderEndPoint.Y:F2}, {leaderEndPoint.Z:F2})");
                                    }
                                }
                                
                                // Store the computed leader end
                                if (leaderEndPoint != null)
                                {
                                    data.LeaderEndModel = new Coordinates3D
                                    {
                                        X = leaderEndPoint.X,
                                        Y = leaderEndPoint.Y,
                                        Z = leaderEndPoint.Z
                                    };
                                    data.LeaderEndView = ConvertToViewCoordinates(data.LeaderEndModel);
                                    data.LeaderEnd = new Coordinates2D 
                                    { 
                                        X = data.LeaderEndView.X, 
                                        Y = data.LeaderEndView.Y 
                                    };
                                }
                            }
                        }
                    }
                    catch (Exception refEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get SpotDimension references: {refEx.Message}");
                    }
                    
                    // Get Head position from TextPosition (the actual text location for SpotDimension)
                    // TextPosition is the most accurate source for where the text appears
                    try
                    {
                        XYZ textPos = spotDim.TextPosition;
                        if (textPos != null)
                        {
                            headPoint = textPos;
                            data.HeadModelPosition = new Coordinates3D
                            {
                                X = textPos.X,
                                Y = textPos.Y,
                                Z = textPos.Z
                            };
                            data.HeadViewPosition = ConvertToViewCoordinates(data.HeadModelPosition);
                            data.HeadPosition = new Coordinates2D
                            {
                                X = data.HeadViewPosition.X,
                                Y = data.HeadViewPosition.Y
                            };
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension TextPosition (text): ({textPos.X:F2}, {textPos.Y:F2}, {textPos.Z:F2})");
                        }
                    }
                    catch (Exception textPosEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get TextPosition: {textPosEx.Message}");
                    }
                    
                    // Fallback: Get Head position from Location if TextPosition was not available
                    if (headPoint == null && spotDim.Location is LocationPoint locPoint && locPoint.Point != null)
                    {
                        headPoint = locPoint.Point;
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension Location (text fallback): ({headPoint.X:F2}, {headPoint.Y:F2}, {headPoint.Z:F2})");
                    }
                    
                    // Get Origin (arrow position) - this is where the dimension line starts
                    // The arrow position for SpotDimension is crucial for proper placement
                    try
                    {
                        XYZ origin = spotDim.Origin;
                        if (origin != null)
                        {
                            // Origin is the arrow position for SpotDimension
                            data.SpotOriginModel = new Coordinates3D
                            {
                                X = origin.X,
                                Y = origin.Y,
                                Z = origin.Z
                            };
                            data.SpotOriginView = ConvertToViewCoordinates(data.SpotOriginModel);
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension Origin (arrow): ({origin.X:F2}, {origin.Y:F2}, {origin.Z:F2})");
                        }
                    }
                    catch (Exception originEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get SpotDimension Origin: {originEx.Message}");
                    }
                                        
                    // Get LeaderEndPosition (attachment point to element)
                    // This is the point where the arrow attaches to the element (critical for placement)
                    try
                    {
                        XYZ leaderEndPos = spotDim.LeaderEndPosition;
                        if (leaderEndPos != null)
                        {
                            // LeaderEndPosition is where it attaches to the element
                            // Override LeaderEndModel with this more accurate value
                            data.LeaderEndModel = new Coordinates3D
                            {
                                X = leaderEndPos.X,
                                Y = leaderEndPos.Y,
                                Z = leaderEndPos.Z
                            };
                            data.LeaderEndView = ConvertToViewCoordinates(data.LeaderEndModel);
                            data.LeaderEnd = new Coordinates2D 
                            { 
                                X = data.LeaderEndView.X, 
                                Y = data.LeaderEndView.Y 
                            };
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension LeaderEndPosition (attachment): ({leaderEndPos.X:F2}, {leaderEndPos.Y:F2}, {leaderEndPos.Z:F2})");
                                                
                            // If SpotOriginModel is still empty, use LeaderEndPosition as origin (arrow position)
                            if (data.SpotOriginModel.X == 0 && data.SpotOriginModel.Y == 0 && data.SpotOriginModel.Z == 0)
                            {
                                data.SpotOriginModel = new Coordinates3D
                                {
                                    X = leaderEndPos.X,
                                    Y = leaderEndPos.Y,
                                    Z = leaderEndPos.Z
                                };
                                data.SpotOriginView = ConvertToViewCoordinates(data.SpotOriginModel);
                                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension Origin (from LeaderEndPosition): ({leaderEndPos.X:F2}, {leaderEndPos.Y:F2}, {leaderEndPos.Z:F2})");
                            }
                        }
                    }
                    catch (Exception leaderEndEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Could not get SpotDimension LeaderEndPosition: {leaderEndEx.Message}");
                    }
                                        
                    // If SpotOriginModel is still empty, use LeaderEndModel as fallback for arrow position
                    if (data.SpotOriginModel.X == 0 && data.SpotOriginModel.Y == 0 && data.SpotOriginModel.Z == 0 
                        && leaderEndPoint != null)
                    {
                        data.SpotOriginModel = new Coordinates3D
                        {
                            X = leaderEndPoint.X,
                            Y = leaderEndPoint.Y,
                            Z = leaderEndPoint.Z
                        };
                        data.SpotOriginView = ConvertToViewCoordinates(data.SpotOriginModel);
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension Origin (fallback from leaderEndPoint): ({leaderEndPoint.X:F2}, {leaderEndPoint.Y:F2}, {leaderEndPoint.Z:F2})");
                    }
                    
                    // Fallback for head position if Location was null
                    if (headPoint == null)
                    {
                        // Fallback: use the leader end point as head (common for elevation marks)
                        // Or get from the spot dimension's text position
                        try
                        {
                            // Try to get the text position - SpotDimension doesn't have direct API for this
                            // Use element geometry as fallback
                            var geomElem = spotDim.get_Geometry(new Options());
                            if (geomElem != null)
                            {
                                foreach (var geomObj in geomElem)
                                {
                                    if (geomObj is GeometryInstance geomInst)
                                    {
                                        // Get symbol geometry
                                        var symbolGeom = geomInst.GetSymbolGeometry();
                                        foreach (var obj in symbolGeom)
                                        {
                                            if (obj is Curve curve)
                                            {
                                                // Use endpoint of a curve as approximation
                                                headPoint = curve.GetEndPoint(1);
                                                break;
                                            }
                                        }
                                        if (headPoint != null) break;
                                    }
                                    else if (geomObj is Curve c)
                                    {
                                        // Use endpoint of leader curve
                                        headPoint = c.GetEndPoint(1);
                                        break;
                                    }
                                }
                            }
                        }
                        catch { }
                        
                        // Final fallback: offset from leader end
                        if (headPoint == null && leaderEndPoint != null)
                        {
                            // Offset by a small amount in view direction
                            headPoint = new XYZ(
                                leaderEndPoint.X + 0.5,
                                leaderEndPoint.Y,
                                leaderEndPoint.Z
                            );
                        }
                    }
                    
                    if (headPoint != null)
                    {
                        // Save full 3D model coordinates
                        data.HeadModelPosition = new Coordinates3D
                        {
                            X = headPoint.X,
                            Y = headPoint.Y,
                            Z = headPoint.Z
                        };
                        
                        // Convert to view coordinates
                        data.HeadViewPosition = ConvertToViewCoordinates(data.HeadModelPosition);
                        
                        // Legacy compatibility
                        data.HeadPosition = new Coordinates2D 
                        { 
                            X = data.HeadViewPosition.X, 
                            Y = data.HeadViewPosition.Y 
                        };
                    }
                                                    
                    // Get elevation value
                    try
                    {
                        var spotDimType = spotDim.SpotDimensionType;
                        data.Orientation = "Elevation"; // Mark as elevation
                        
                        // Get the elevation value from SpotDimension
                        // SpotDimension.Value gives the elevation value
                        try
                        {
                            // The elevation is stored in the Value property
                            double? elevationValue = spotDim.Value;
                            if (elevationValue.HasValue)
                            {
                                // Convert from feet to meters (or keep in project units)
                                data.TagText = $"{elevationValue.Value:F3}"; // Store elevation value
                            }
                            else
                            {
                                data.TagText = "";
                            }
                        }
                        catch
                        {
                            // SpotDimension doesn't have TagText property
                            data.TagText = "";
                        }
                                                        
                        // Store leader info
                        if (spotDim.HasLeader)
                        {
                            data.HasLeader = true;
                            
                            // Try to get leader end position
                            try
                            {
                                var refs = spotDim.References;
                                if (refs != null && refs.Size > 0)
                                {
                                    var firstRef = refs.get_Item(0);
                                    if (firstRef != null)
                                    {
                                        // Leader end is where the arrow points to
                                        data.LeaderEndModel = new Coordinates3D
                                        {
                                            X = firstRef.GlobalPoint.X,
                                            Y = firstRef.GlobalPoint.Y,
                                            Z = firstRef.GlobalPoint.Z
                                        };
                                        data.LeaderEndView = ConvertToViewCoordinates(data.LeaderEndModel);
                                        data.LeaderEnd = new Coordinates2D 
                                        { 
                                            X = data.LeaderEndView.X, 
                                            Y = data.LeaderEndView.Y 
                                        };
                                    }
                                }
                            }
                            catch { }
                        }
                                                        
                        // Try to get the reference element
                        ElementId refElemId = null;
                        try
                        {
                            // SpotDimension has References property
                            var refs2 = spotDim.References;
                            if (refs2 != null && refs2.Size > 0)
                            {
                                var firstRef = refs2.get_Item(0);
                                if (firstRef != null)
                                {
                                    refElemId = firstRef.ElementId;
                                    data.TaggedElementId = refElemId.Value;
                                    data.BelongTo = $"element:{refElemId.Value}";
                                }
                            }
                        }
                        catch { }
                                                        
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] SpotDimension: Type={data.TypeName}, HasLeader={data.HasLeader}, RefElement={refElemId}");
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   SpotOriginModel (arrow): ({data.SpotOriginModel?.X:F2}, {data.SpotOriginModel?.Y:F2}, {data.SpotOriginModel?.Z:F2})");
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   HeadModelPosition (text): ({data.HeadModelPosition?.X:F2}, {data.HeadModelPosition?.Y:F2}, {data.HeadModelPosition?.Z:F2})");
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR]   LeaderEndModel (attachment): ({data.LeaderEndModel?.X:F2}, {data.LeaderEndModel?.Y:F2}, {data.LeaderEndModel?.Z:F2})");
                    }
                    catch (Exception spotEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error processing SpotDimension: {spotEx.Message}");
                    }
                }
                else
                {
                    // For other annotations, use general view coordinates
                    var viewCoords = GetViewCoordinates(annotation);
                    if (viewCoords != null)
                    {
                        data.HeadPosition = viewCoords;
                    }
                    data.HasLeader = false;
                }
                
                // Tagged element (for tags)
                var taggedElement = GetTaggedElement(annotation);
                if (taggedElement != null)
                {
                    data.TaggedElementId = taggedElement.Id.Value;
                    data.BelongTo = $"element:{taggedElement.Id.Value}";
                }

                return data;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error creating annotation data: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get 3D model coordinates of an element
        /// </summary>
        private Coordinates3D GetModelCoordinates(Element element)
        {
            try
            {
                var location = element.Location;
                if (location is LocationPoint point)
                {
                    return new Coordinates3D
                    {
                        X = Math.Round(point.Point.X, 4),
                        Y = Math.Round(point.Point.Y, 4),
                        Z = Math.Round(point.Point.Z, 4)
                    };
                }
                else if (location is LocationCurve curve)
                {
                    var start = curve.Curve.GetEndPoint(0);
                    return new Coordinates3D
                    {
                        X = Math.Round(start.X, 4),
                        Y = Math.Round(start.Y, 4),
                        Z = Math.Round(start.Z, 4)
                    };
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting model coordinates: {ex.Message}");
            }

            return new Coordinates3D();
        }
                
        /// <summary>
        /// Extract diameter/size dimensions for pipes and ducts
        /// </summary>
        private void ExtractElementDimensions(Element element, ElementData data)
        {
            try
            {
                // Pipes - get diameter
                if (element is Pipe pipe)
                {
                    data.Diameter = pipe.Diameter;
                    // Convert to mm for display (1 foot = 304.8 mm)
                    double diameterMm = pipe.Diameter * 304.8;
                    data.SizeDisplay = $"⌀{diameterMm:F0}";
                    DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Pipe {element.Id}: Diameter={diameterMm:F1}mm");
                }
                // Ducts - get width/height or diameter
                else if (element is Duct duct)
                {
                    // Check if rectangular or round
                    Parameter widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                    Parameter heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                            
                    if (widthParam != null && heightParam != null && 
                        widthParam.AsDouble() > 0 && heightParam.AsDouble() > 0)
                    {
                        // Rectangular duct
                        data.Width = widthParam.AsDouble();
                        data.Height = heightParam.AsDouble();
                        double widthMm = data.Width.Value * 304.8;
                        double heightMm = data.Height.Value * 304.8;
                        data.SizeDisplay = $"{widthMm:F0}x{heightMm:F0}";
                        DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Duct {element.Id}: {widthMm:F0}x{heightMm:F0}mm");
                    }
                    else
                    {
                        // Round duct - try diameter parameter
                        Parameter diameterParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                        if (diameterParam != null && diameterParam.AsDouble() > 0)
                        {
                            data.Diameter = diameterParam.AsDouble();
                            double diameterMm = data.Diameter.Value * 304.8;
                            data.SizeDisplay = $"⌀{diameterMm:F0}";
                            DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Round Duct {element.Id}: ⌀{diameterMm:F0}mm");
                        }
                        else
                        {
                            // Fallback - try to get from MEPCurve
                            data.Diameter = duct.Diameter;
                            if (data.Diameter > 0)
                            {
                                double diameterMm = data.Diameter.Value * 304.8;
                                data.SizeDisplay = $"⌀{diameterMm:F0}";
                            }
                        }
                    }
                }
                // Cable trays - get width
                else if (element is CableTray cableTray)
                {
                    Parameter widthParam = cableTray.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                    if (widthParam != null)
                    {
                        data.Width = widthParam.AsDouble();
                        double widthMm = data.Width.Value * 304.8;
                        data.SizeDisplay = $"{widthMm:F0}";
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error extracting dimensions: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get 2D view coordinates of an element
        /// </summary>
        private Coordinates2D GetViewCoordinates(Element element)
        {
            try
            {
                var boundingBox = element.get_BoundingBox(_view);
                if (boundingBox != null)
                {
                    // Calculate center of bounding box in model space
                    var center = new XYZ(
                        (boundingBox.Min.X + boundingBox.Max.X) / 2,
                        (boundingBox.Min.Y + boundingBox.Max.Y) / 2,
                        (boundingBox.Min.Z + boundingBox.Max.Z) / 2
                    );

                    // Transform to view coordinates
                    // For plan views and locked 3D views, we need to project to 2D
                    var viewOrigin = _view.Origin;
                    var viewRight = _view.RightDirection;
                    var viewUp = _view.UpDirection;

                    // Calculate relative position in view coordinates
                    var relative = center - viewOrigin;
                    var viewX = relative.DotProduct(viewRight);
                    var viewY = relative.DotProduct(viewUp);

                    return new Coordinates2D
                    {
                        X = Math.Round(viewX, 4),
                        Y = Math.Round(viewY, 4)
                    };
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting view coordinates: {ex.Message}");
            }

            return new Coordinates2D();
        }

        /// <summary>
        /// Get start and end coordinates for linear elements (pipes, ducts, cable trays)
        /// </summary>
        private Tuple<Coordinates3D, Coordinates3D, bool> GetLinearElementCoordinates(Element element)
        {
            try
            {
                // Handle pipes
                if (element is Pipe pipe)
                {
                    var curve = pipe.Location as LocationCurve;
                    if (curve != null && curve.Curve != null)
                    {
                        var start = curve.Curve.GetEndPoint(0);
                        var end = curve.Curve.GetEndPoint(1);
                        return Tuple.Create(
                            new Coordinates3D { X = start.X, Y = start.Y, Z = start.Z },
                            new Coordinates3D { X = end.X, Y = end.Y, Z = end.Z },
                            true
                        );
                    }
                }
                
                // Handle ducts
                if (element is Duct duct)
                {
                    var curve = duct.Location as LocationCurve;
                    if (curve != null && curve.Curve != null)
                    {
                        var start = curve.Curve.GetEndPoint(0);
                        var end = curve.Curve.GetEndPoint(1);
                        return Tuple.Create(
                            new Coordinates3D { X = start.X, Y = start.Y, Z = start.Z },
                            new Coordinates3D { X = end.X, Y = end.Y, Z = end.Z },
                            true
                        );
                    }
                }
                
                // Handle cable trays
                if (element is CableTray cableTray)
                {
                    var curve = cableTray.Location as LocationCurve;
                    if (curve != null && curve.Curve != null)
                    {
                        var start = curve.Curve.GetEndPoint(0);
                        var end = curve.Curve.GetEndPoint(1);
                        return Tuple.Create(
                            new Coordinates3D { X = start.X, Y = start.Y, Z = start.Z },
                            new Coordinates3D { X = end.X, Y = end.Y, Z = end.Z },
                            true
                        );
                    }
                }
                
                // Handle conduits
                if (element is Conduit conduit)
                {
                    var curve = conduit.Location as LocationCurve;
                    if (curve != null && curve.Curve != null)
                    {
                        var start = curve.Curve.GetEndPoint(0);
                        var end = curve.Curve.GetEndPoint(1);
                        return Tuple.Create(
                            new Coordinates3D { X = start.X, Y = start.Y, Z = start.Z },
                            new Coordinates3D { X = end.X, Y = end.Y, Z = end.Z },
                            true
                        );
                    }
                }
                
                // Generic: check if element has LocationCurve
                var locCurve = element.Location as LocationCurve;
                if (locCurve != null && locCurve.Curve != null)
                {
                    var start = locCurve.Curve.GetEndPoint(0);
                    var end = locCurve.Curve.GetEndPoint(1);
                    return Tuple.Create(
                        new Coordinates3D { X = start.X, Y = start.Y, Z = start.Z },
                        new Coordinates3D { X = end.X, Y = end.Y, Z = end.Z },
                        true
                    );
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting linear coordinates: {ex.Message}");
            }
            
            return null;
        }

        /// <summary>
        /// Convert 3D model coordinates to 2D view coordinates
        /// </summary>
        private Coordinates2D ConvertToViewCoordinates(Coordinates3D modelCoords)
        {
            try
            {
                var point = new XYZ(modelCoords.X, modelCoords.Y, modelCoords.Z);
                var viewOrigin = _view.Origin;
                var viewRight = _view.RightDirection;
                var viewUp = _view.UpDirection;

                var relative = point - viewOrigin;
                var viewX = relative.DotProduct(viewRight);
                var viewY = relative.DotProduct(viewUp);

                return new Coordinates2D
                {
                    X = Math.Round(viewX, 4),
                    Y = Math.Round(viewY, 4)
                };
            }
            catch
            {
                return new Coordinates2D();
            }
        }

        /// <summary>
        /// Get system info for MEP elements
        /// </summary>
        private Tuple<long, string> GetElementSystemInfo(Element element)
        {
            try
            {
                // Try to get piping system
                if (element is Pipe pipe && pipe.MEPSystem != null)
                {
                    return Tuple.Create(
                        pipe.MEPSystem.Id.Value,
                        pipe.MEPSystem.Name ?? ""
                    );
                }

                // Try to get duct system
                if (element is Duct duct && duct.MEPSystem != null)
                {
                    return Tuple.Create(
                        duct.MEPSystem.Id.Value,
                        duct.MEPSystem.Name ?? ""
                    );
                }

                // Try to get system for family instances (fittings, terminals)
                if (element is FamilyInstance instance)
                {
                    // For MEP elements, try to get system from various properties
                    try
                    {
                        // Try different approaches for different element types
                        var mepModel = instance.MEPModel;
                        if (mepModel != null)
                        {
                            // Try to get system via ElementId parameter or other means
                            var systemParam = instance.LookupParameter("Система") ?? instance.LookupParameter("System");
                            if (systemParam != null)
                            {
                                var systemId = systemParam.AsElementId();
                                if (systemId != null && systemId != ElementId.InvalidElementId)
                                {
                                    var system = _document.GetElement(systemId) as MEPSystem;
                                    if (system != null)
                                    {
                                        return Tuple.Create(
                                            system.Id.Value,
                                            system.Name ?? ""
                                        );
                                    }
                                }
                            }
                        }
                    }
                    catch { /* Ignore */ }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting system info: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Get the element that a tag is tagging
        /// </summary>
        private Element GetTaggedElement(Element tag)
        {
            try
            {
                // For independent tags
                if (tag is IndependentTag indTag)
                {
                    var taggedIds = indTag.GetTaggedLocalElementIds();
                    if (taggedIds != null && taggedIds.Count > 0)
                    {
                        return _document.GetElement(taggedIds.First());
                    }
                }

                // For spatial element tags (rooms, spaces)
                if (tag is SpatialElementTag spatialTag)
                {
                    // SpatialElementTag in Revit 2024+ uses different API
                    try
                    {
                        var taggedParam = tag.LookupParameter("TaggedSpatialElement");
                        if (taggedParam != null)
                        {
                            var taggedId = taggedParam.AsElementId();
                            if (taggedId != null && taggedId != ElementId.InvalidElementId)
                            {
                                return _document.GetElement(taggedId);
                            }
                        }
                    }
                    catch { /* Ignore */ }
                }

                // Try to get via parameter
                var taggedElementParam = tag.LookupParameter("TaggedElement");
                if (taggedElementParam != null)
                {
                    var refValue = taggedElementParam.AsElementId();
                    if (refValue != null && refValue != ElementId.InvalidElementId)
                    {
                        return _document.GetElement(refValue);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting tagged element: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Get systems visible on the view with their elements
        /// </summary>
        private List<SystemData> GetSystemsOnView(List<Element> elements)
        {
            var systemsDict = new Dictionary<long, SystemData>();

            try
            {
                foreach (var element in elements)
                {
                    var systemInfo = GetElementSystemInfo(element);
                    if (systemInfo != null)
                    {
                        var systemId = systemInfo.Item1;

                        if (!systemsDict.ContainsKey(systemId))
                        {
                            var mepSystem = _document.GetElement(new ElementId(systemId)) as MEPSystem;
                            systemsDict[systemId] = new SystemData
                            {
                                SystemId = systemId,
                                SystemName = systemInfo.Item2,
                                SystemType = mepSystem?.GetType().Name ?? "Unknown"
                            };
                        }

                        systemsDict[systemId].ElementIds.Add(element.Id.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-COLLECTOR] Error getting systems: {ex.Message}");
            }

            return systemsDict.Values.ToList();
        }
    }
}
