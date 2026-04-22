using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Annotatix.Module.Core;
using PluginsManager.Core;

namespace Annotatix.Module.Commands
{
    /// <summary>
    /// Main command for deterministic ductwork annotation placement
    /// Replaces ML-based approach with algorithm-based placement
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class AnnotateDuctworkCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiApp = commandData.Application;
            var app = uiApp.Application;
            var uiDoc = uiApp.ActiveUIDocument;
            var doc = uiDoc.Document;
            var view = doc.ActiveView;
            
            try
            {
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Starting deterministic annotation placement...");
                
                // Validate view
                if (!IsValidView(view))
                {
                    TaskDialog.Show("Annotatix", "Команда работает только на видах с воздуховодами или трубами.");
                    return Result.Cancelled;
                }
                
                // Get UIView for the active view
                UIView uiView = null;
                try
                {
                    var uiViews = uiDoc.GetOpenUIViews();
                    uiView = uiViews?.FirstOrDefault(uv => uv.ViewId == view.Id);
                }
                catch { }
                
                // Stage 1: Collect data
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Stage 1: Collecting view data...");
                string sessionId = Guid.NewGuid().ToString();
                var collector = new ViewDataCollector(doc, view, uiView);
                var snapshot = collector.CollectSnapshot(sessionId, "start");
                
                if (snapshot.Elements.Count == 0)
                {
                    TaskDialog.Show("Annotatix", "На виде не найдено элементов для аннотирования.");
                    return Result.Succeeded;
                }
                
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Collected {snapshot.Elements.Count} elements, {snapshot.Systems?.Count ?? 0} systems");
                
                // Stage 2: Build system graphs
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Stage 2: Building system graphs...");
                var graphBuilder = new SystemGraphBuilder(doc);  // Pass document for Connector API
                var graphs = graphBuilder.BuildFromSnapshot(snapshot);
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Built {graphs.Count} system graphs");
                
                // Log graph details
                foreach (var graph in graphs)
                {
                    DebugLogger.Log($"[ANNOTATE-DUCTWORK] System '{graph.SystemName}': {graph.Nodes.Count} nodes, {graph.LinearElements.Count} ducts, {graph.Junctions.Count} junctions");
                }
                
                // Stage 3: Apply annotation rules
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Stage 3: Applying annotation rules...");
                var rulesEngine = new AnnotationRulesEngine();
                var allPlans = new List<AnnotationPlan>();
                
                foreach (var graph in graphs)
                {
                    var plans = rulesEngine.AnalyzeSystem(graph);
                    allPlans.AddRange(plans);
                    DebugLogger.Log($"[ANNOTATE-DUCTWORK] System '{graph.SystemName}': {plans.Count} annotations planned");
                }
                
                if (allPlans.Count == 0)
                {
                    TaskDialog.Show("Annotatix", "Не найдено элементов, требующих аннотирования по правилам.");
                    return Result.Succeeded;
                }
                
                // Deduplicate plans (same element may have multiple reasons)
                var uniquePlans = allPlans
                    .GroupBy(p => p.ElementId)
                    .Select(g => g.OrderBy(p => p.IsMandatory ? 0 : 1).First())
                    .ToList();
                
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Total unique annotations to place: {uniquePlans.Count}");
                
                // Stage 4: Measure annotation sizes (invisible to user)
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Stage 4: Measuring annotation sizes...");
                var sizer = new AnnotationSizer(doc, view);
                var sizes = sizer.MeasureSizes(uniquePlans);
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Measured {sizes.Count} annotation sizes");
                
                // Stage 5: Setup collision detection
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Stage 5: Setting up collision detection...");
                var collisionDetector = new CollisionDetector();
                collisionDetector.AddElementsFromSnapshot(snapshot);
                var intersections = collisionDetector.GetIntersections();
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Found {intersections.Count} visual intersections");
                
                // Stage 6: Place annotations (greedy algorithm)
                DebugLogger.Log("[ANNOTATE-DUCTWORK] Stage 6: Placing annotations...");
                double viewScale = view.Scale;
                if (viewScale < 1) viewScale = 1;
                
                var placementConfig = new PlacementConfig
                {
                    // All offset/height values are now calculated in paper space (mm) and converted
                    // to model feet using view scale inside GreedyPlacementService.
                    // These config values control the elbow height iteration loop.
                    BaseElbowHeight = 3.0 / 304.8 * viewScale,       // 3mm on paper base elbow height
                    BaseHorizontalOffset = 3.0 / 304.8 * viewScale,  // 3mm on paper base horizontal offset
                    MinLeaderLength = 1.0 / 304.8 * viewScale,       // 1mm on paper minimum leader
                    ElementSizeMultiplier = 0.75, // Scale offset by element size
                    MaxElbowHeight = 10.0 / 304.8 * viewScale,   // 10mm on paper max elbow height
                    ElbowHeightStep = 1.0 / 304.8 * viewScale    // 1mm on paper step size
                };
                
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] View scale {viewScale}, config MaxElbow={placementConfig.MaxElbowHeight:F2}ft, Step={placementConfig.ElbowHeightStep:F2}ft");
                
                var placementService = new GreedyPlacementService(
                    doc, view, collisionDetector, sizes, placementConfig);
                
                List<PlacementResult> results;
                
                using (var tx = new Transaction(doc, "Place Ductwork Annotations"))
                {
                    tx.Start();
                    
                    results = placementService.PlaceAll(uniquePlans);
                    
                    // Only commit if at least some succeeded
                    var successCount = results.Count(r => r.Success);
                    if (successCount > 0)
                    {
                        tx.Commit();
                    }
                    else
                    {
                        tx.RollBack();
                    }
                }
                
                // Stage 7: Report results
                var placed = results.Count(r => r.Success);
                var failed = results.Count(r => !r.Success);
                
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Placement complete: {placed} placed, {failed} failed");
                
                // Log failures for debugging
                foreach (var failure in results.Where(r => !r.Success))
                {
                    DebugLogger.Log($"[ANNOTATE-DUCTWORK] Failed to annotate element {failure.ElementId}: {failure.FailureReason}");
                }
                
                // Export structured CSV data for analysis
                try
                {
                    string logsDir = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "Autodesk", "Revit", "Addins", "2025",
                        "annotatix_dependencies", "logs");
                    
                    AnnotationCsvExporter.Export(
                        snapshot,
                        placementService.PlacementRecords,
                        placementService.IterationRecords,
                        logsDir);
                }
                catch (Exception csvEx)
                {
                    DebugLogger.Log($"[ANNOTATE-DUCTWORK] CSV export error: {csvEx.Message}");
                }
                
                // Show result to user
                string message2 = placed > 0 
                    ? $"Успешно размещено аннотаций: {placed}\nОшибок: {failed}"
                    : "Не удалось разместить ни одной аннотации.";
                
                TaskDialog.Show("Annotatix - Размещение аннотаций", message2);
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[ANNOTATE-DUCTWORK] Error: {ex.Message}\n{ex.StackTrace}");
                TaskDialog.Show("Annotatix - Ошибка", $"Произошла ошибка:\n{ex.Message}");
                return Result.Failed;
            }
        }
        
        private bool IsValidView(View view)
        {
            if (view == null) return false;
            
            // Accept 3D views and plan views
            var viewType = view.ViewType;
            return viewType == ViewType.ThreeD ||
                   viewType == ViewType.FloorPlan ||
                   viewType == ViewType.CeilingPlan ||
                   viewType == ViewType.EngineeringPlan ||
                   viewType == ViewType.Section ||
                   viewType == ViewType.Elevation;
        }
    }
}
