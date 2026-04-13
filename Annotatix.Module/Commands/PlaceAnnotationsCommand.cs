using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Annotatix.Module.Core;
using Annotatix.Module.UI;
using PluginsManager.Core;
using Newtonsoft.Json;

namespace Annotatix.Module.Commands
{
    /// <summary>
    /// Command to place annotations from the last recording session
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class PlaceAnnotationsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uidoc = uiApp.ActiveUIDocument;
                var doc = uidoc.Document;

                // Get active view
                var activeView = doc.ActiveView;
                if (activeView == null)
                {
                    TaskDialog.Show("Annotatix", "Нет активного вида. Откройте вид для размещения аннотаций.");
                    return Result.Failed;
                }

                DebugLogger.Log("[ANNOTATIX-PLACE] Starting annotation placement...");

                // Get recordings directory
                string recordingsDir = RecordingState.RecordingsDirectory;
                if (string.IsNullOrEmpty(recordingsDir))
                {
                    recordingsDir = JsonExporter.GetDefaultRecordingsDirectory();
                }

                if (!Directory.Exists(recordingsDir))
                {
                    TaskDialog.Show("Annotatix", $"Папка записей не найдена:\n{recordingsDir}\n\nСначала выполните запись.");
                    return Result.Failed;
                }

                // Find the latest session folder
                var sessionFolders = Directory.GetDirectories(recordingsDir);
                if (sessionFolders.Length == 0)
                {
                    TaskDialog.Show("Annotatix", "Не найдено ни одной записи.\n\nСначала выполните запись.");
                    return Result.Failed;
                }

                // Sort by last write time (most recent first)
                var latestSession = sessionFolders
                    .OrderByDescending(f => Directory.GetLastWriteTime(f))
                    .First();

                DebugLogger.Log($"[ANNOTATIX-PLACE] Latest session folder: {latestSession}");

                // Find the end snapshot file
                var jsonFiles = Directory.GetFiles(latestSession, "*-end.json");
                if (jsonFiles.Length == 0)
                {
                    TaskDialog.Show("Annotatix", "В последней записи не найден файл конечного снимка.\n\nУбедитесь, что запись была завершена корректно.");
                    return Result.Failed;
                }

                var endSnapshotPath = jsonFiles[0];
                DebugLogger.Log($"[ANNOTATIX-PLACE] Loading end snapshot: {endSnapshotPath}");

                // Load the snapshot
                var snapshot = LoadSnapshot(endSnapshotPath);
                if (snapshot == null)
                {
                    TaskDialog.Show("Annotatix", "Не удалось загрузить файл записи.");
                    return Result.Failed;
                }

                DebugLogger.Log($"[ANNOTATIX-PLACE] Loaded snapshot with {snapshot.Annotations.Count} annotations");

                // Check if view matches
                if (snapshot.ViewId != activeView.Id.Value)
                {
                    var result = TaskDialog.Show("Annotatix", 
                        $"Внимание! Запись была сделана на другом виде.\n\n" +
                        $"Вид записи: {snapshot.ViewName} (ID: {snapshot.ViewId})\n" +
                        $"Текущий вид: {activeView.Name} (ID: {activeView.Id.Value})\n\n" +
                        "Продолжить размещение аннотаций?",
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                    if (result != TaskDialogResult.Yes)
                    {
                        return Result.Cancelled;
                    }
                }

                // Place annotations
                int placedCount = 0;
                int failedCount = 0;

                using (var transaction = new Transaction(doc, "Place Annotations from Recording"))
                {
                    transaction.Start();

                    foreach (var annotationData in snapshot.Annotations)
                    {
                        try
                        {
                            if (PlaceAnnotation(doc, activeView, annotationData, snapshot))
                            {
                                placedCount++;
                            }
                            else
                            {
                                failedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Error placing annotation {annotationData.ElementId}: {ex.Message}");
                            failedCount++;
                        }
                    }

                    transaction.Commit();
                }

                DebugLogger.Log($"[ANNOTATIX-PLACE] Placement complete: {placedCount} placed, {failedCount} failed");

                TaskDialog.Show("Annotatix", 
                    $"Размещение аннотаций завершено.\n\n" +
                    $"Успешно размещено: {placedCount}\n" +
                    $"Ошибок: {failedCount}\n\n" +
                    $"Файл записи: {Path.GetFileName(latestSession)}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] ERROR: {ex.Message}");
                TaskDialog.Show("Annotatix", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Load a ViewSnapshot from JSON file
        /// </summary>
        private ViewSnapshot LoadSnapshot(string filePath)
        {
            try
            {
                var json = File.ReadAllText(filePath);
                var settings = new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    DateFormatString = "yyyy-MM-ddTHH:mm:ss"
                };
                return JsonConvert.DeserializeObject<ViewSnapshot>(json, settings);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Error loading snapshot: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Place a single annotation on the view
        /// </summary>
        private bool PlaceAnnotation(Document doc, View view, AnnotationData annotationData, ViewSnapshot snapshot)
        {
            // Handle SpotDimension (elevation marks) separately
            if (annotationData.Category == "Высотные отметки" || annotationData.FamilyName == "SpotDimension")
            {
                return PlaceSpotDimension(doc, view, annotationData, snapshot);
            }
                    
            // Handle IndependentTag
            if (annotationData.TaggedElementId == null || annotationData.TaggedElementId == 0)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Annotation {annotationData.ElementId} has no tagged element, skipping");
                return false;
            }

            // Find the tagged element in the current document
            var taggedElement = doc.GetElement(new ElementId(annotationData.TaggedElementId.Value));
            if (taggedElement == null)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Tagged element {annotationData.TaggedElementId} not found in document");
                return false;
            }

            // Get the tag type (family symbol)
            FamilySymbol tagSymbol = FindTagSymbol(doc, annotationData.FamilyName, annotationData.TypeName);
            if (tagSymbol == null)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Tag symbol not found: {annotationData.FamilyName} - {annotationData.TypeName}");
                return false;
            }

            // Activate the symbol if needed
            if (!tagSymbol.IsActive)
            {
                tagSymbol.Activate();
            }

            // Create the tag
            try
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Placing tag for element {taggedElement.Id}, LeaderType={annotationData.LeaderType}, HasElbow={annotationData.HasElbow}");
                
                // Determine tag position - use model coordinates if available
                XYZ tagHeadPosition;
                
                // Check if we have new model coordinates (preferred)
                if (annotationData.HeadModelPosition != null && annotationData.HeadModelPosition.X != 0)
                {
                    // Use model coordinates directly
                    tagHeadPosition = new XYZ(
                        annotationData.HeadModelPosition.X,
                        annotationData.HeadModelPosition.Y,
                        annotationData.HeadModelPosition.Z
                    );
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using HeadModelPosition: ({tagHeadPosition.X:F2}, {tagHeadPosition.Y:F2}, {tagHeadPosition.Z:F2})");
                }
                else
                {
                    // Fallback to legacy coordinates (without Z)
                    tagHeadPosition = new XYZ(
                        annotationData.HeadPosition.X,
                        annotationData.HeadPosition.Y,
                        0 // Z coordinate unknown
                    );

                    // For 3D views, try to get Z from element
                    if (view.ViewType == ViewType.ThreeD)
                    {
                        Location elementLocation = taggedElement.Location;
                        if (elementLocation is LocationPoint locationPoint)
                        {
                            tagHeadPosition = new XYZ(annotationData.HeadPosition.X, annotationData.HeadPosition.Y, locationPoint.Point.Z);
                        }
                        else if (elementLocation is LocationCurve locationCurve)
                        {
                            var midPoint = locationCurve.Curve.Evaluate(0.5, true);
                            tagHeadPosition = new XYZ(annotationData.HeadPosition.X, annotationData.HeadPosition.Y, midPoint.Z);
                        }
                    }
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using legacy HeadPosition: ({tagHeadPosition.X:F2}, {tagHeadPosition.Y:F2}, {tagHeadPosition.Z:F2})");
                }

                // Create the IndependentTag
                Reference elementRef = GetReferenceForElement(taggedElement, view);
                if (elementRef == null)
                {
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Could not get reference for element {taggedElement.Id}");
                    return false;
                }

                // Determine orientation
                TagOrientation orientation = TagOrientation.Horizontal;
                if (!string.IsNullOrEmpty(annotationData.Orientation))
                {
                    if (annotationData.Orientation.Equals("Vertical", StringComparison.OrdinalIgnoreCase))
                    {
                        orientation = TagOrientation.Vertical;
                    }
                }

                // Create tag with or without leader based on LeaderType
                bool addLeader = annotationData.HasLeader || 
                                (annotationData.LeaderType?.Equals("Free", StringComparison.OrdinalIgnoreCase) == true);
                
                IndependentTag newTag = IndependentTag.Create(
                    doc,
                    tagSymbol.Id,
                    view.Id,
                    elementRef,
                    addLeader,
                    orientation,
                    tagHeadPosition
                );

                if (newTag != null)
                {
                    // Get the ACTUAL reference that the tag uses
                    // This is critical for SetLeaderEnd and SetLeaderElbow
                    var taggedRefs = newTag.GetTaggedReferences();
                    Reference tagRef = (taggedRefs != null && taggedRefs.Count > 0) ? taggedRefs[0] : elementRef;
                    
                    // Set leader end condition FIRST (before setting positions)
                    bool isFreeLeader = annotationData.LeaderType?.Equals("Free", StringComparison.OrdinalIgnoreCase) == true;
                    
                    if (isFreeLeader)
                    {
                        newTag.LeaderEndCondition = LeaderEndCondition.Free;
                    }
                    else if (annotationData.LeaderType?.Equals("Attached", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        newTag.LeaderEndCondition = LeaderEndCondition.Attached;
                    }
                                    
                    // Set tag head position after leader condition
                    newTag.TagHeadPosition = tagHeadPosition;
                    
                    // For Free leaders, set the leader end position (where the arrow points)
                    // This is the KEY to placing annotation at correct location, not on connector
                    if (isFreeLeader && tagRef != null)
                    {
                        try
                        {
                            XYZ leaderEndPosition;
                            
                            // Use LeaderEndModel if recorded
                            if (annotationData.LeaderEndModel != null && annotationData.LeaderEndModel.X != 0)
                            {
                                leaderEndPosition = new XYZ(
                                    annotationData.LeaderEndModel.X,
                                    annotationData.LeaderEndModel.Y,
                                    annotationData.LeaderEndModel.Z
                                );
                                
                                newTag.SetLeaderEnd(tagRef, leaderEndPosition);
                                DebugLogger.Log($"[ANNOTATIX-PLACE] Set leader end position: ({leaderEndPosition.X:F2}, {leaderEndPosition.Y:F2}, {leaderEndPosition.Z:F2})");
                            }
                        }
                        catch (Exception leaderEndEx)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Could not set leader end: {leaderEndEx.Message}");
                        }
                    }
                                    
                    // Set elbow position if recorded
                    if (annotationData.HasElbow && tagRef != null)
                    {
                        try
                        {
                            XYZ elbowPosition;
                            
                            // Use model coordinates if available
                            if (annotationData.ElbowModelPosition != null && annotationData.ElbowModelPosition.X != 0)
                            {
                                elbowPosition = new XYZ(
                                    annotationData.ElbowModelPosition.X,
                                    annotationData.ElbowModelPosition.Y,
                                    annotationData.ElbowModelPosition.Z
                                );
                            }
                            else
                            {
                                // Fallback to legacy coordinates
                                elbowPosition = new XYZ(
                                    annotationData.ElbowPosition.X,
                                    annotationData.ElbowPosition.Y,
                                    tagHeadPosition.Z
                                );
                            }
                            
                            newTag.SetLeaderElbow(tagRef, elbowPosition);
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Set elbow position: ({elbowPosition.X:F2}, {elbowPosition.Y:F2}, {elbowPosition.Z:F2})");
                        }
                        catch (Exception elbowEx)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Could not set elbow: {elbowEx.Message}");
                        }
                    }
                                    
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Created tag {newTag.Id} for element {taggedElement.Id}, LeaderType={newTag.LeaderEndCondition}, HasLeader={newTag.HasLeader}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Error creating tag: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Find a tag symbol by family name and type name
        /// </summary>
        private FamilySymbol FindTagSymbol(Document doc, string familyName, string typeName)
        {
            // Collect all FamilySymbol elements (tag types)
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>();

            // Log all available tag symbols for debugging
            var availableTags = new System.Collections.Generic.List<string>();
            foreach (var s in collector)
            {
                if (s.FamilyName != null)
                {
                    availableTags.Add($"{s.FamilyName} - {s.Name}");
                }
            }
            DebugLogger.Log($"[ANNOTATIX-PLACE] Available tag symbols: {string.Join(", ", availableTags.Take(20))}...");

            // Reset collector after enumeration
            collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>();

            // Try exact match first (family + type)
            if (!string.IsNullOrEmpty(familyName) && !string.IsNullOrEmpty(typeName))
            {
                foreach (var symbol in collector)
                {
                    if (symbol.FamilyName != null && 
                        symbol.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                        symbol.Name != null &&
                        symbol.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase))
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Found exact match: {symbol.FamilyName} - {symbol.Name}");
                        return symbol;
                    }
                }
            }

            // Try to find by type name only (most reliable for tags)
            if (!string.IsNullOrEmpty(typeName))
            {
                collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>();
                    
                foreach (var symbol in collector)
                {
                    if (symbol.Name != null && 
                        symbol.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase))
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Found by type name: {symbol.FamilyName} - {symbol.Name}");
                        return symbol;
                    }
                }
            }

            // Try to find by family name only
            if (!string.IsNullOrEmpty(familyName))
            {
                collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>();
                    
                foreach (var symbol in collector)
                {
                    if (symbol.FamilyName != null && 
                        symbol.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Found by family name: {symbol.FamilyName} - {symbol.Name}");
                        return symbol;
                    }
                }
            }

            DebugLogger.Log($"[ANNOTATIX-PLACE] No tag symbol found for family='{familyName}' type='{typeName}'");
            return null;
        }

        /// <summary>
        /// Get a reference for an element for tagging
        /// </summary>
        private Reference GetReferenceForElement(Element element, View view)
        {
            return GetReferenceForElement(element, view, false);
        }
        
        /// <summary>
        /// Get a reference for an element for tagging or SpotDimension
        /// </summary>
        /// <param name="element">The element to reference</param>
        /// <param name="view">The view context</param>
        /// <param name="forSpotDimension">If true, prefer face/surface references over curve references</param>
        private Reference GetReferenceForElement(Element element, View view, bool forSpotDimension)
        {
            try
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Getting reference for element {element.Id}, category: {element.Category?.Name}, forSpotDimension: {forSpotDimension}");

                // For MEPCurves (pipes, ducts, cable trays), get reference from geometry
                // IMPORTANT: LocationCurve.Curve.Reference returns null!
                // We need to get the Curve/Face from element geometry with ComputeReferences = true
                if (element is MEPCurve mepCurve)
                {
                    // Get geometry with compute references enabled
                    Options opt = new Options();
                    opt.ComputeReferences = true;
                    opt.IncludeNonVisibleObjects = false;
                    opt.View = view;
                    
                    GeometryElement geomElem = element.get_Geometry(opt);
                    if (geomElem != null)
                    {
                        Reference faceRef = null;
                        Reference curveRef = null;
                        
                        foreach (GeometryObject geoObj in geomElem)
                        {
                            // Check for solid geometry (surface references) - PREFERRED for SpotDimension
                            Solid solid = geoObj as Solid;
                            if (solid != null && solid.Faces.Size > 0)
                            {
                                // Find the best face - use the one closest to the pipe surface
                                // For SpotDimension, we need a face reference so LeaderEndPosition can be set
                                foreach (Face face in solid.Faces)
                                {
                                    if (face != null && face.Reference != null)
                                    {
                                        faceRef = face.Reference;
                                        DebugLogger.Log($"[ANNOTATIX-PLACE] Got face reference from solid for MEPCurve {element.Id}");
                                        break;
                                    }
                                }
                                if (faceRef != null) break;
                            }
                            
                            // Also look for Curve in geometry (centerline for pipes)
                            Curve cv = geoObj as Curve;
                            if (cv != null && cv.Reference != null)
                            {
                                curveRef = cv.Reference;
                                DebugLogger.Log($"[ANNOTATIX-PLACE] Got geometry curve reference for MEPCurve {element.Id}");
                            }
                        }
                        
                        // For SpotDimension, prefer FACE reference (allows LeaderEndPosition to work)
                        // For regular tags, curve reference is fine
                        if (forSpotDimension && faceRef != null)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Using FACE reference for SpotDimension on MEPCurve {element.Id}");
                            return faceRef;
                        }
                        else if (curveRef != null)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Using curve reference for MEPCurve {element.Id}");
                            return curveRef;
                        }
                        else if (faceRef != null)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Using face reference fallback for MEPCurve {element.Id}");
                            return faceRef;
                        }
                    }
                    
                    // Fallback: Create reference directly from element
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using element reference fallback for MEPCurve {element.Id}");
                    return new Reference(element);
                }

                // For FamilyInstance elements
                if (element is FamilyInstance familyInstance)
                {
                    // For most family instances, direct element reference works
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using element reference for FamilyInstance {element.Id}");
                    return new Reference(element);
                }

                // For other elements, try direct reference
                DebugLogger.Log($"[ANNOTATIX-PLACE] Using direct element reference for {element.Id}");
                return new Reference(element);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Error getting reference for {element.Id}: {ex.Message}");
                DebugLogger.Log($"[ANNOTATIX-PLACE] Stack trace: {ex.StackTrace}");
                return null;
            }
        }
        
        /// <summary>
        /// Place a SpotDimension (elevation mark)
        /// </summary>
        private bool PlaceSpotDimension(Document doc, View view, AnnotationData annotationData, ViewSnapshot snapshot)
        {
            try
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Attempting to place SpotDimension: {annotationData.TypeName}");
                
                // Find the SpotDimensionType
                var spotDimTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(SpotDimensionType))
                    .Cast<SpotDimensionType>();
                
                SpotDimensionType spotDimType = null;
                
                // Try to find by exact name first
                foreach (var t in spotDimTypes)
                {
                    if (t.Name == annotationData.TypeName)
                    {
                        spotDimType = t;
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Found exact SpotDimensionType: {t.Name}");
                        break;
                    }
                }
                
                // Fallback: try partial match
                if (spotDimType == null)
                {
                    spotDimTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(SpotDimensionType))
                        .Cast<SpotDimensionType>();
                        
                    foreach (var t in spotDimTypes)
                    {
                        if (t.Name.Contains("Отметка") || t.Name.Contains("Elevation") ||
                            t.Name.Contains("Схема"))
                        {
                            spotDimType = t;
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Found partial match SpotDimensionType: {t.Name}");
                            break;
                        }
                    }
                }
                
                if (spotDimType == null)
                {
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimensionType not found: {annotationData.TypeName}");
                    return false;
                }
                
                // Check if we have a tagged element
                if (annotationData.TaggedElementId == null || annotationData.TaggedElementId == 0)
                {
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension has no tagged element");
                    return false;
                }
                
                var taggedElement = doc.GetElement(new ElementId(annotationData.TaggedElementId.Value));
                if (taggedElement == null)
                {
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Tagged element {annotationData.TaggedElementId} not found");
                    return false;
                }
                
                // Get reference for the element
                // Use Curve Reference for SpotDimension - Face Reference requires origin ON the face surface
                Reference elemRef = GetReferenceForElement(taggedElement, view, false);
                if (elemRef == null)
                {
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Could not get reference for SpotDimension element");
                    return false;
                }
                
                // For SpotDimension (elevation marks), we need to track 3 control points:
                // 1. Origin (arrow position) - where the arrow points on the element
                // 2. End (text position) - where the text/annotation appears
                // 3. Bend (elbow) - the bend point on the leader line
                // 
                // IMPORTANT: In Revit API, NewSpotElevation parameters are:
                // - origin: the ARROW position (where dimension line starts)
                // - bend: elbow point on leader
                // - end: text position
                // - refPt: reference point for finding point on element
                //
                // SpotOriginModel = arrow position (origin parameter)
                // LeaderEndModel = attachment point to element (use as refPt or for finding origin)
                // HeadModelPosition = text position (end parameter)
                                                
                XYZ origin;  // Arrow position (where dimension line starts)
                XYZ end;     // Text position (where annotation appears)
                XYZ bend;    // Elbow point on leader
                XYZ refPt;   // Reference point for finding point on element
                                
                // CRITICAL: For SpotDimension on MEPCurve (pipes):
                // - 'origin' is the MEASUREMENT POINT (arrow position) = SpotOriginModel
                // - 'LeaderEndPosition' is where leader attaches = usually on pipe surface
                // 
                // The Reference should point to the pipe SURFACE (not centerline) for LeaderEndPosition to work!
                // When Reference is to centerline curve, LeaderEndPosition cannot be freely set.
                
                // CRITICAL: For SpotDimension with Curve Reference:
                // - 'origin' parameter is projected onto the element curve
                // - LeaderEndPosition is determined by this projection
                // - LeaderShoulderPosition controls the leader geometry (bend point)
                // - TextPosition controls where text appears
                
                // STRATEGY: 
                // 1. Pass SpotOriginModel as origin (arrow/measurement point)
                // 2. After creation, set LeaderShoulderPosition = Origin
                // 3. Then set TextPosition to desired location
                
                // Step 1: Determine origin (arrow point) - use SpotOriginModel for correct measurement
                if (annotationData.SpotOriginModel != null && annotationData.SpotOriginModel.X != 0)
                {
                    origin = new XYZ(
                        annotationData.SpotOriginModel.X,
                        annotationData.SpotOriginModel.Y,
                        annotationData.SpotOriginModel.Z
                    );
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using SpotOriginModel as origin (arrow): ({origin.X:F2}, {origin.Y:F2}, {origin.Z:F2})");
                }
                else if (annotationData.LeaderEndModel != null && annotationData.LeaderEndModel.X != 0)
                {
                    origin = new XYZ(
                        annotationData.LeaderEndModel.X,
                        annotationData.LeaderEndModel.Y,
                        annotationData.LeaderEndModel.Z
                    );
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using LeaderEndModel as origin fallback: ({origin.X:F2}, {origin.Y:F2}, {origin.Z:F2})");
                }
                else
                {
                    // Fallback: project onto curve
                    if (taggedElement is MEPCurve mepCurveElement && mepCurveElement.Location is LocationCurve lc)
                    {
                        origin = lc.Curve.Evaluate(0.5, true);
                    }
                    else if (taggedElement.Location is LocationPoint lp)
                    {
                        origin = lp.Point;
                    }
                    else
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Cannot determine SpotDimension origin");
                        return false;
                    }
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using fallback origin: ({origin.X:F2}, {origin.Y:F2}, {origin.Z:F2})");
                }
                                
                // Step 2: Determine reference point (refPt)
                // Use LeaderEndModel as refPt - this helps Revit find the correct geometry
                if (annotationData.LeaderEndModel != null && annotationData.LeaderEndModel.X != 0)
                {
                    refPt = new XYZ(
                        annotationData.LeaderEndModel.X,
                        annotationData.LeaderEndModel.Y,
                        annotationData.LeaderEndModel.Z
                    );
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using LeaderEndModel as refPt: ({refPt.X:F2}, {refPt.Y:F2}, {refPt.Z:F2})");
                }
                else
                {
                    refPt = origin;
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using origin as refPt fallback: ({refPt.X:F2}, {refPt.Y:F2}, {refPt.Z:F2})");
                }
                                
                // Store for later use
                XYZ pointOnElement = origin;
                
                // NOTE: Do NOT overwrite refPt here - it should remain as LeaderEndModel
                // refPt tells Revit WHERE on the element to attach the leader
                                
                // Determine text position (end) - where the annotation text appears
                // Priority: HeadModelPosition > offset from origin
                if (annotationData.HeadModelPosition != null && annotationData.HeadModelPosition.X != 0)
                {
                    end = new XYZ(
                        annotationData.HeadModelPosition.X,
                        annotationData.HeadModelPosition.Y,
                        annotationData.HeadModelPosition.Z
                    );
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using HeadModelPosition: ({end.X:F2}, {end.Y:F2}, {end.Z:F2})");
                }
                else
                {
                    // Fallback - offset from origin
                    end = origin + new XYZ(0.5, 0.5, 0);
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using offset fallback: ({end.X:F2}, {end.Y:F2}, {end.Z:F2})");
                }
                                
                // Determine bend point (elbow) - the bend on the leader line
                // For SpotDimension without elbow (HasElbow=false), use end point as bend for straight leader
                if (annotationData.HasElbow && annotationData.ElbowModelPosition != null 
                    && annotationData.ElbowModelPosition.X != 0)
                {
                    // Has explicit elbow position
                    bend = new XYZ(
                        annotationData.ElbowModelPosition.X,
                        annotationData.ElbowModelPosition.Y,
                        annotationData.ElbowModelPosition.Z
                    );
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using ElbowModelPosition: ({bend.X:F2}, {bend.Y:F2}, {bend.Z:F2})");
                }
                else
                {
                    // No elbow - use end point as bend for straight leader line
                    // This creates a straight line from origin to end through bend
                    bend = end;
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension using straight leader (bend=end): ({bend.X:F2}, {bend.Y:F2}, {bend.Z:F2})");
                }
                
                bool hasLeader = annotationData.HasLeader;
                
                // Store the projected point for LeaderEndPosition
                // CRITICAL: LeaderEndPosition must be a point ON the element geometry!
                // Revit will reject any point that is not on the referenced geometry
                XYZ projectedLeaderEnd = pointOnElement;
                
                // Create the SpotDimension
                SpotDimension spotDim = doc.Create.NewSpotElevation(
                    view,
                    elemRef,
                    origin,
                    bend,
                    end,
                    refPt,
                    hasLeader
                );
                
                if (spotDim != null)
                {
                    // Set the SpotDimensionType
                    spotDim.SpotDimensionType = spotDimType;
                    
                    // CRITICAL: Revit API constraints for SpotDimension with Curve Reference:
                    // - Setting LeaderShoulderPosition affects LeaderEndPosition and TextPosition
                    // - Setting LeaderEndPosition affects TextPosition
                    // - Setting TextPosition may be limited by geometric constraints
                    //
                    // CORRECT ORDER (based on testing):
                    // 1. Set LeaderShoulderPosition = Origin FIRST (establishes leader geometry)
                    // 2. Then set LeaderEndPosition (attachment point - this will be preserved)
                    // 3. Finally set TextPosition (text location)
                    
                    // Step 1: Set LeaderShoulderPosition to Origin FIRST
                    // This establishes the proper leader geometry where shoulder = arrow point
                    try
                    {
                        if (spotDim.LeaderHasShoulder)
                        {
                            spotDim.LeaderShoulderPosition = origin;
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Set LeaderShoulderPosition to origin: ({origin.X:F2}, {origin.Y:F2}, {origin.Z:F2})");
                        }
                    }
                    catch (Exception shoulderEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Could not set LeaderShoulderPosition: {shoulderEx.Message}");
                    }
                    
                    // Step 2: Set LeaderEndPosition AFTER LeaderShoulderPosition
                    // This is where the arrow attaches - should be offset from Origin
                    try
                    {
                        if (annotationData.LeaderEndModel != null && annotationData.LeaderEndModel.X != 0)
                        {
                            XYZ leaderEndPos = new XYZ(
                                annotationData.LeaderEndModel.X,
                                annotationData.LeaderEndModel.Y,
                                annotationData.LeaderEndModel.Z
                            );
                            spotDim.LeaderEndPosition = leaderEndPos;
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Set LeaderEndPosition to: ({leaderEndPos.X:F2}, {leaderEndPos.Y:F2}, {leaderEndPos.Z:F2})");
                        }
                    }
                    catch (Exception leaderEndEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Could not set LeaderEndPosition: {leaderEndEx.Message}");
                    }
                    
                    // Step 3: Set TextPosition AFTER LeaderEndPosition and LeaderShoulderPosition
                    // TextPosition controls where the text/annotation appears
                    try
                    {
                        // Check if text position is adjustable
                        if (spotDim.IsTextPositionAdjustable())
                        {
                            XYZ textPos = new XYZ(
                                annotationData.HeadModelPosition.X,
                                annotationData.HeadModelPosition.Y,
                                annotationData.HeadModelPosition.Z
                            );
                            spotDim.TextPosition = textPos;
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Set TextPosition to: ({textPos.X:F2}, {textPos.Y:F2}, {textPos.Z:F2})");
                        }
                        else
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] TextPosition is not adjustable for this SpotDimension");
                        }
                    }
                    catch (Exception textEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Could not set TextPosition: {textEx.Message}");
                    }
                    
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Created SpotDimension {spotDim.Id}:");
                    DebugLogger.Log($"[ANNOTATIX-PLACE]   Origin (measurement): ({origin.X:F2}, {origin.Y:F2}, {origin.Z:F2})");
                    DebugLogger.Log($"[ANNOTATIX-PLACE]   End (text): ({end.X:F2}, {end.Y:F2}, {end.Z:F2})");
                    DebugLogger.Log($"[ANNOTATIX-PLACE]   Bend (elbow): ({bend.X:F2}, {bend.Y:F2}, {bend.Z:F2})");
                    DebugLogger.Log($"[ANNOTATIX-PLACE]   RefPt: ({refPt.X:F2}, {refPt.Y:F2}, {refPt.Z:F2})");
                    
                    // Log actual SpotDimension properties after all modifications
                    try
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE]   Actual TextPosition: ({spotDim.TextPosition.X:F2}, {spotDim.TextPosition.Y:F2}, {spotDim.TextPosition.Z:F2})");
                        DebugLogger.Log($"[ANNOTATIX-PLACE]   Actual LeaderEndPosition: ({spotDim.LeaderEndPosition.X:F2}, {spotDim.LeaderEndPosition.Y:F2}, {spotDim.LeaderEndPosition.Z:F2})");
                        if (spotDim.LeaderHasShoulder)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE]   Actual LeaderShoulderPosition: ({spotDim.LeaderShoulderPosition.X:F2}, {spotDim.LeaderShoulderPosition.Y:F2}, {spotDim.LeaderShoulderPosition.Z:F2})");
                        }
                        DebugLogger.Log($"[ANNOTATIX-PLACE]   LeaderHasShoulder: {spotDim.LeaderHasShoulder}");
                    }
                    catch (Exception logEx)
                    {
                        DebugLogger.Log($"[ANNOTATIX-PLACE] Could not log actual positions: {logEx.Message}");
                    }
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Error placing SpotDimension: {ex.Message}");
                return false;
            }
        }
    }
}
