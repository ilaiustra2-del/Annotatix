using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.DependencyInjection;
using HVACSuperScheme.Utils;
using HVACSuperScheme.Updaters;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection.Emit;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Annotations;
using System.Windows.Controls;
using HVACSuperScheme.Commands.Settings;
using View = Autodesk.Revit.DB.View;
using System.Text;
using System.Reflection;

namespace HVACSuperScheme.Commands.CreateSchema
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private Document _doc;
        private ViewDrafting _draftView;
        private ElementId _annotationSymbolSymbolId;
        public ElementId _textNoteTypeId;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            _doc = uidoc.Document;
            
            using (TransactionGroup trg = new TransactionGroup(_doc, "Создание схемы"))
            {
                try
                {
                    SettingStorage.ReadSettings();
                    UpdaterRegistry.DisableUpdater(Updater._m_updaterId);

                    trg.Start();

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
                }
                catch (CustomException ex)
                {
                    trg.RollBack();
                    MessageBox.Show(ExceptionUtils.Error(ex));
                }
                catch (Exception ex)
                {
                    trg.RollBack();
                    MessageBox.Show(ExceptionUtils.SystemError(ex));
                }
                finally
                {
                    UpdaterRegistry.EnableUpdater(Updater._m_updaterId);
                }
            }
            return Result.Succeeded;
        }
        private void DeleteViewsDrafting(List<ViewDrafting> views)
        {
            using var t = new Transaction(_doc);

            t.Start("Удаление неактуальных чертежных видов");
            _doc.Delete(views.Select(v => v.Id).ToList());
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
    }
}