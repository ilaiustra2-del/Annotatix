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
                    // Set leader end condition FIRST (before setting head position)
                    // This ensures the leader type is correct
                    if (!string.IsNullOrEmpty(annotationData.LeaderType))
                    {
                        if (annotationData.LeaderType.Equals("Free", StringComparison.OrdinalIgnoreCase))
                        {
                            newTag.LeaderEndCondition = LeaderEndCondition.Free;
                        }
                        else if (annotationData.LeaderType.Equals("Attached", StringComparison.OrdinalIgnoreCase))
                        {
                            newTag.LeaderEndCondition = LeaderEndCondition.Attached;
                        }
                    }
                                    
                    // Set tag head position after leader condition
                    newTag.TagHeadPosition = tagHeadPosition;
                                    
                    // Set elbow position if recorded
                    if (annotationData.HasElbow)
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
                            
                            newTag.SetLeaderElbow(elementRef, elbowPosition);
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
            try
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Getting reference for element {element.Id}, category: {element.Category?.Name}");

                // For MEPCurves (pipes, ducts, cable trays), we need special handling
                if (element is MEPCurve mepCurve)
                {
                    LocationCurve locationCurve = mepCurve.Location as LocationCurve;
                    if (locationCurve != null && locationCurve.Curve != null)
                    {
                        // Try to get curve reference first
                        var curveRef = locationCurve.Curve.Reference;
                        if (curveRef != null)
                        {
                            DebugLogger.Log($"[ANNOTATIX-PLACE] Got curve reference for MEPCurve {element.Id}");
                            return curveRef;
                        }
                    }
                    
                    // Fallback: Create reference directly from element
                    // For 3D views, we can use the element itself as reference
                    DebugLogger.Log($"[ANNOTATIX-PLACE] Using element reference for MEPCurve {element.Id}");
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
                
                // SpotDimension creation requires specific geometry and references
                // This is a placeholder - SpotDimension creation is complex
                // For now, log the attempt and skip
                
                // Find the SpotDimensionType
                var spotDimTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(SpotDimensionType))
                    .Cast<SpotDimensionType>()
                    .Where(t => t.Name == annotationData.TypeName || t.Name.Contains("Отметка"))
                    .ToList();
                
                if (spotDimTypes.Count == 0)
                {
                    DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimensionType not found: {annotationData.TypeName}");
                    return false;
                }
                
                var spotDimType = spotDimTypes.First();
                DebugLogger.Log($"[ANNOTATIX-PLACE] Found SpotDimensionType: {spotDimType.Name}");
                
                // If we have a tagged element, try to create the spot dimension
                if (annotationData.TaggedElementId != null && annotationData.TaggedElementId > 0)
                {
                    var taggedElement = doc.GetElement(new ElementId(annotationData.TaggedElementId.Value));
                    if (taggedElement != null)
                    {
                        // Get the reference point
                        XYZ refPoint = null;
                        if (taggedElement.Location is LocationPoint locPoint)
                        {
                            refPoint = locPoint.Point;
                        }
                        else if (taggedElement.Location is LocationCurve locCurve)
                        {
                            refPoint = locCurve.Curve.Evaluate(0.5, true);
                        }
                        
                        if (refPoint != null)
                        {
                            // Spot dimensions need special handling based on view type
                            // For now, just log
                            DebugLogger.Log($"[ANNOTATIX-PLACE] SpotDimension placement not fully implemented yet. Would place at: {refPoint}");
                            // TODO: Implement actual SpotDimension creation
                            // doc.Create.NewSpotDimension(view, spotDimType, ref, point);
                        }
                    }
                }
                
                return false; // Not implemented yet
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATIX-PLACE] Error placing SpotDimension: {ex.Message}");
                return false;
            }
        }
    }
}
