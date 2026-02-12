using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PluginsManager.Core;
using HVACSuperScheme;
using HVACSuperScheme.Commands.Settings;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using System.Windows;

namespace HVAC.Module.Commands
{
    /// <summary>
    /// External Event Handler for Complete Schema command
    /// Integrated with original HVACSuperScheme functionality
    /// </summary>
    public class CompleteSchemaHandler : IExternalEventHandler
    {
        private UIApplication _uiApp;
        private Document _doc;
        private ViewDrafting _draftView;
        private ElementId _annotationSymbolSymbolId;
        public ElementId _textNoteTypeId;

        public CompleteSchemaHandler(UIApplication uiApp)
        {
            _uiApp = uiApp;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Log("[HVAC-CMD] Executing CompleteSchema command...");
                
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                {
                    TaskDialog.Show("Ошибка", "Нет активного документа Revit");
                    return;
                }
                
                _doc = uidoc.Document;
                
                // Initialize Updater if not already initialized
                EnsureUpdaterInitialized(app);
                
                using (TransactionGroup trg = new TransactionGroup(_doc, "Созданием схемы"))
                {
                    try
                    {
                        SettingStorage.ReadSettings();
                        
                        // Only disable if Updater is initialized
                        if (Updater._m_updaterId != null)
                        {
                            UpdaterRegistry.DisableUpdater(Updater._m_updaterId);
                        }
                        trg.Start();

                        // Initialize cache dictionaries for this document
                        CashUtils.InitializeCash(_doc);

                        _draftView = CollectorUtils.GetExistHVACSchemaViews(_doc).FirstOrDefault();
                        if (_draftView == null)
                            throw new CustomException(ExceptionUtils.NotFoundValidDraftViews()); 

                        _annotationSymbolSymbolId = CollectorUtils.GetAnnotationSymbolSymbolId(_doc);
                        if (!ActivateUtils.IsSymbolActive(_doc, _annotationSymbolSymbolId))
                            ActivateUtils.ActivateFamilySymbol(_doc, _annotationSymbolSymbolId);
                        _textNoteTypeId = CollectorUtils.GetTextNoteTypeId(_doc, Constants.TEXT_NOTE_TYPE_NAME);
                        CheckUtils.CheckGeneral(_doc, _draftView, _annotationSymbolSymbolId);

                        List<ElementId> spacesIds = CollectorUtils.GetSpacesForAnnotations(_doc, withoutAnnotationInstance: true);

                        if (!spacesIds.Any())
                            throw new CustomException(ExceptionUtils.NotFoundValidSpacesWithoutAnnotations());

                        double topPointOfAnnotations = GetTopPointOfAnnotations(_doc, _draftView);
                        XYZ insertPoint = new XYZ(0, topPointOfAnnotations + Constants.HEIGHT_ANNOTATION * 3, 0);
                        SchemaUtils.CreateAnnotations(_doc, spacesIds, _annotationSymbolSymbolId, _textNoteTypeId, _draftView, insertPoint);
                        uidoc.ActiveView = _draftView;
                        trg.Assimilate();
                        
                        DebugLogger.Log("[HVAC-CMD] CompleteSchema completed successfully");
                    }
                    catch (CustomException ex)
                    {
                        trg.RollBack();
                        DebugLogger.Log($"[HVAC-CMD] CustomException in CompleteSchema: {ex.Message}");
                        MessageBox.Show(ExceptionUtils.Error(ex));
                    }
                    catch (Exception ex)
                    {
                        trg.RollBack();
                        DebugLogger.Log($"[HVAC-CMD] ERROR in CompleteSchema: {ex.Message}");
                        DebugLogger.Log($"[HVAC-CMD] Stack trace: {ex.StackTrace}");
                        MessageBox.Show(ExceptionUtils.SystemError(ex));
                    }
                    finally
                    {
                        // Only enable if Updater is initialized
                        if (Updater._m_updaterId != null)
                        {
                            UpdaterRegistry.EnableUpdater(Updater._m_updaterId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-CMD] FATAL ERROR in CompleteSchema: {ex.Message}");
                DebugLogger.Log($"[HVAC-CMD] Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Ошибка", $"Критическая ошибка при достроении схемы:\n{ex.Message}");
            }
        }
        
        public double GetTopPointOfAnnotations(Document doc, Autodesk.Revit.DB.View view)
        {
            double TopYOfAnnotation = double.MinValue;

            foreach (var annotation in CollectorUtils.GetAnnotationInstances(doc))
            {
                BoundingBoxXYZ bb = annotation.get_BoundingBox(view);
                if (bb == null)
                    continue;
                if (bb.Max.Y > TopYOfAnnotation)
                    TopYOfAnnotation = bb.Max.Y;
            }
            return TopYOfAnnotation;
        }
        
        private void EnsureUpdaterInitialized(UIApplication uiApp)
        {
            if (HVACSuperScheme.App._updater == null)
            {
                DebugLogger.Log("[HVAC-CMD] Initializing Updater...");
                
                // Store UIApplication for Idling events
                HVACSuperScheme.App._uiapp = uiApp;
                
                // Initialize element filters BEFORE creating Updater
                FilterUtils.InitElementFilters();
                DebugLogger.Log("[HVAC-CMD] Element filters initialized");
                
                // Get AddInId from the current application
                var addInId = uiApp.ActiveAddInId;
                
                // Create Updater instance
                HVACSuperScheme.App._updater = new Updater(addInId);
                
                // Register the updater
                UpdaterRegistry.RegisterUpdater(HVACSuperScheme.App._updater);
                
                DebugLogger.Log("[HVAC-CMD] Updater initialized and registered successfully");
            }
            else if (HVACSuperScheme.App._uiapp == null)
            {
                // Updater exists but UIApp not stored yet
                HVACSuperScheme.App._uiapp = uiApp;
                DebugLogger.Log("[HVAC-CMD] UIApplication stored for Idling events");
            }
        }

        public string GetName()
        {
            return "HVAC_CompleteSchemaHandler";
        }
    }
}
