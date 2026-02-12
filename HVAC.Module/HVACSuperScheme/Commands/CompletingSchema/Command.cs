using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HVACSuperScheme.Commands.Settings;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Xaml;

namespace HVACSuperScheme.Commands.CompletingSchema
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    internal class Command : IExternalCommand
    {
        private Document _doc;

        private ViewDrafting _draftView;

        private ElementId _annotationSymbolSymbolId;

        public ElementId _textNoteTypeId;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            _doc = uidoc.Document;
            using (TransactionGroup trg = new TransactionGroup(_doc, "Созданием схемы"))
            {
                try
                {
                    SettingStorage.ReadSettings();
                    UpdaterRegistry.DisableUpdater(Updater._m_updaterId);
                    trg.Start();

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
    }
}
