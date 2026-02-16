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
    /// External Event Handler for Create Schema command
    /// Integrated with original HVACSuperScheme functionality
    /// </summary>
    public class CreateSchemaHandler : IExternalEventHandler
    {
        private UIApplication _uiApp;
        private Document _doc;
        private ViewDrafting _draftView;
        private ElementId _annotationSymbolSymbolId;
        public ElementId _textNoteTypeId;

        public CreateSchemaHandler()
        {
            // UIApplication will be provided in Execute()
        }
        
        public CreateSchemaHandler(UIApplication uiApp)
        {
            _uiApp = uiApp;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Log("[HVAC-CMD] Executing CreateSchema command...");
                
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                {
                    TaskDialog.Show("Ошибка", "Нет активного документа Revit");
                    return;
                }
                
                _doc = uidoc.Document;
                
                using (TransactionGroup trg = new TransactionGroup(_doc, "Создание схемы"))
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

                        List<ViewDrafting> HVACSchemaViews = CollectorUtils.GetExistHVACSchemaViews(_doc);
                        if (HVACSchemaViews.Any())
                            DeleteViewsDrafting(HVACSchemaViews);

                        CheckUtils.CheckDuplicateViewName(_doc);

                        _draftView = CreateViewDrafting();
                        _annotationSymbolSymbolId = CollectorUtils.GetAnnotationSymbolSymbolId(_doc);
                        if (!ActivateUtils.IsSymbolActive(_doc, _annotationSymbolSymbolId))
                            ActivateUtils.ActivateFamilySymbol(_doc, _annotationSymbolSymbolId);
                        _textNoteTypeId = CollectorUtils.GetTextNoteTypeId(_doc, Constants.TEXT_NOTE_TYPE_NAME);
                        CheckUtils.CheckGeneral(_doc, _draftView, _annotationSymbolSymbolId);

                        List<ElementId> spacesIds = CollectorUtils.GetSpacesForAnnotations(_doc, withoutAnnotationInstance: false);

                        XYZ insertPoint = new XYZ(0, 0, 0);
                        SchemaUtils.CreateAnnotations(_doc, spacesIds, _annotationSymbolSymbolId, _textNoteTypeId, _draftView, insertPoint);
                        uidoc.ActiveView = _draftView;
                        trg.Assimilate();
                        
                        DebugLogger.Log("[HVAC-CMD] CreateSchema completed successfully");
                    }
                    catch (CustomException ex)
                    {
                        trg.RollBack();
                        DebugLogger.Log($"[HVAC-CMD] CustomException in CreateSchema: {ex.Message}");
                        MessageBox.Show(ExceptionUtils.Error(ex));
                    }
                    catch (Exception ex)
                    {
                        trg.RollBack();
                        DebugLogger.Log($"[HVAC-CMD] ERROR in CreateSchema: {ex.Message}");
                        DebugLogger.Log($"[HVAC-CMD] Stack trace: {ex.StackTrace}");
                        MessageBox.Show(ExceptionUtils.SystemError(ex));
                    }
                    finally
                    {
                        // Only enable Updater if checkbox is enabled AND Updater is initialized
                        if (Updater._m_updaterId != null)
                        {
                            if (UI.HVACSyncState.IsSyncEnabled)
                            {
                                UpdaterRegistry.EnableUpdater(Updater._m_updaterId);
                                DebugLogger.Log("[HVAC-CMD] Updater enabled (sync checkbox ON)");
                            }
                            else
                            {
                                DebugLogger.Log("[HVAC-CMD] Updater NOT enabled (sync checkbox OFF)");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[HVAC-CMD] FATAL ERROR in CreateSchema: {ex.Message}");
                DebugLogger.Log($"[HVAC-CMD] Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Ошибка", $"Критическая ошибка при создании схемы:\n{ex.Message}");
            }
        }
        
        private void DeleteViewsDrafting(List<ViewDrafting> views)
        {
            using var t = new Transaction(_doc);

            t.Start("Удаление неактуальных чертежных видов");
            
            // Close views if they are open, then delete
            var uiDoc = new UIDocument(_doc);
            var successfullyDeleted = 0;
            var failedToDelete = 0;
            
            foreach (var view in views)
            {
                try
                {
                    // Save view name BEFORE any operations (view becomes invalid after close)
                    string viewName = "";
                    try
                    {
                        viewName = view.Name;
                    }
                    catch
                    {
                        viewName = "<unknown>";
                    }
                    
                    // Try to close the view if it's currently open
                    try
                    {
                        var openViews = uiDoc.GetOpenUIViews();
                        foreach (var openView in openViews)
                        {
                            if (openView.ViewId == view.Id)
                            {
                                openView.Close();
                                DebugLogger.Log($"[HVAC-CMD] Closed open view: {viewName}");
                                break;
                            }
                        }
                    }
                    catch
                    {
                        // Ignore errors closing view
                    }
                    
                    // Now try to delete
                    _doc.Delete(view.Id);
                    successfullyDeleted++;
                    DebugLogger.Log($"[HVAC-CMD] Deleted view: {viewName}");
                }
                catch (Exception ex)
                {
                    // Log but continue - some views might be in use
                    DebugLogger.Log($"[HVAC-CMD] WARNING: Could not delete view: {ex.Message}");
                    failedToDelete++;
                }
            }
            
            DebugLogger.Log($"[HVAC-CMD] Deleted {successfullyDeleted} views, failed to delete {failedToDelete} views");
            
            t.Commit();
        }
        
        private ViewDrafting CreateViewDrafting()
        {
            ViewFamilyType draftFamilyType = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .First(viewFamilyType => viewFamilyType.ViewFamily == ViewFamily.Drafting);

            ViewDrafting draftView;
            using var t = new Transaction(_doc);

            t.Start("Создание чертёжного вида");
            draftView = ViewDrafting.Create(_doc, draftFamilyType.Id);
            draftView.Name = Constants.DRAFT_VIEW_NAME;
            draftView.DetailLevel = ViewDetailLevel.Fine;
            draftView.Scale = 10;
            t.Commit();

            return draftView;
        }

        public string GetName()
        {
            return "HVAC_CreateSchemaHandler";
        }
    }
}
