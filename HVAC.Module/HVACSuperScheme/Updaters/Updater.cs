using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using HVACSuperScheme.Commands.Settings;
using HVACSuperScheme.Utils;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Annotations;
using System.Xml.Linq;
using ParameterUtils = HVACSuperScheme.Utils.ParameterUtils;

namespace HVACSuperScheme.Updaters
{
    public class Updater : IUpdater
    {
        public static AddInId _m_appId;
        public static UpdaterId _m_updaterId;
        public Updater(AddInId id)
        {
            _m_appId = id;
            _m_updaterId = new UpdaterId(_m_appId, new Guid("96730022-636B-40BF-9F9A-CFB7458EE953"));
        }
        public void Execute(UpdaterData data)
        {
            Document doc = data.GetDocument();
            string filePath = doc.PathName;
            try
            {
                LoggingUtils.Logging(Warnings.UpdaterExecuteStart(), doc.PathName);

                List<ElementId> deletedElementIds = data.GetDeletedElementIds().ToList();
                if (deletedElementIds.Any())
                {
                    ProcessDeletedElements(doc, deletedElementIds, filePath);
                }
                List<ElementId> addedElementIds = data.GetAddedElementIds().ToList();
                if (addedElementIds.Any())
                {
                    ProcessAddedElements(doc, addedElementIds, filePath);
                }
                List<ElementId> modifiedElementIds = data.GetModifiedElementIds().ToList();
                if (modifiedElementIds.Any())
                {
                    ProcessModifiedElements(doc, modifiedElementIds, filePath);
                }
            }
            catch (Exception ex)
            {
                LoggingUtils.LoggingWithMessage(ExceptionUtils.UpdaterError(ex), doc.PathName);
            }
            doc.Regenerate();
            LoggingUtils.Logging(Warnings.UpdaterExecuteFinish(), doc.PathName);
        }

        private void ProcessModifiedElements(Document doc, List<ElementId> modifiedElementIds, string filePath)
        {
            LoggingUtils.Logging(Warnings.ModifiedElements(modifiedElementIds), doc.PathName);
            if (SettingStorage.Instance.IsUpdaterSync)
            {
                foreach (ElementId elementId in modifiedElementIds)
                {
                    Element element = doc.GetElement(elementId);
                    if (IsParamUpdaterOn(element))
                        ModifiedElementsParameterUpdaterOn(doc, element);
                    else
                        ModifiedElementsParameterUpdaterOff(doc, element);
                }
            }
            else
            {
                
            }
        }

        private void ProcessDeletedElements(Document doc, List<ElementId> deletedElementIds, string filePath)
        {
            LoggingUtils.Logging(Warnings.DeletedElements(deletedElementIds), doc.PathName);
            if (SettingStorage.Instance.IsUpdaterSync)
                DeletedElements(doc, deletedElementIds);
            else
                StorageHelper.ClearExtensibleStorageForMatchDeletionElement(doc, deletedElementIds);

            CashUtils.UpdateCashMatchedSpaceAndAnnotationAfterDeletion(doc, deletedElementIds);
            CashUtils.UpdateCashDuctTerminalsAfterDeletion(deletedElementIds, filePath);
        }

        private void ProcessAddedElements(Document doc, List<ElementId> addedElementIds, string filePath)
        {
            LoggingUtils.Logging(Warnings.AddedElements(addedElementIds), doc.PathName);
            CashUtils.UpdateCashDuctTerminalsAfterAdded(doc, addedElementIds, filePath);
            if (SettingStorage.Instance.IsUpdaterSync)
                ProcessAddedDuctTerminals(addedElementIds, doc);
        }

        private void ModifiedElementsParameterUpdaterOff(Document doc, Element element)
        {
            if (element is Space space)
                ProcessModifiedSpaceUpdaterOff(doc, space);

            else if (element is AnnotationSymbol annotation)
                ProcessModifiedAnnotationUpdaterOff(doc, annotation);

            else if (element is FamilyInstance ductTerminal && TerminalUtils.IsDuctTerminal(element))
                ProcessDuctTerminal(doc, ductTerminal);
        }
        private void ProcessModifiedAnnotationUpdaterOff(Document doc, AnnotationSymbol annotation)
        {
            ElementId spaceId = StorageHelper.GetMatchedSpaceId(annotation);

            Space space = doc.GetElement(spaceId) as Space;
            if (space == null)
            {
                LoggingUtils.LoggingWithMessage(Warnings.NotFoundSpaceMatchedWithAnnotation(annotation.Id), doc.PathName);
                return;
            }
            Parameter parameterUpdaterSpace = space.LookupParameter(Constants.PN_ADSK_UPDATER);
            ParameterUtils.SetParameterValueWithEqualCheck(parameterUpdaterSpace, 0, spaceId);
            LoggingUtils.LoggingWithMessage(Warnings.NotCompletedSynchronizeParametersWithLinkedSpace(), doc.PathName);
        }
        private void ProcessModifiedSpaceUpdaterOff(Document doc, Space space)
        {
            ElementId annotationId = StorageHelper.GetMatchedAnnotationId(space);

            Element annotation = doc.GetElement(annotationId);

            RecalculateDuctTerminals(doc, space);

            if (annotation == null)
            {
                LoggingUtils.LoggingWithMessage(Warnings.NotFoundAnnotationMatchedWithSpace(space.Id), doc.PathName);
                return;
            }
            Parameter parameterUpdaterAnnotation = annotation.LookupParameter(Constants.PN_ADSK_UPDATER);
            ParameterUtils.SetParameterValueWithEqualCheck(parameterUpdaterAnnotation, 0, annotationId);
            LoggingUtils.LoggingWithMessage(Warnings.NotCompletedSynchronizeParametersWithLinkedAnnotation(), doc.PathName);
        }


        private void ProcessAddedDuctTerminals(List<ElementId> addedElementIds, Document doc)
        {
            foreach (var addedElementId in addedElementIds)
            {
                Element element = doc.GetElement(addedElementId);

                if (!TerminalUtils.IsDuctTerminal(element))
                    continue;

                Space space = TerminalUtils.GetSpaceByDuctTerminal(doc, element as FamilyInstance);
                if (space == null)
                {
                    LoggingUtils.LoggingWithMessage(Warnings.CreatedDuctTerminalNotLocateInSpaceRecalculateAirFlowFailed(), doc.PathName);
                    continue;
                }
                RecalculateDuctTerminals(doc, space);
            }
        }
        private void DeletedElements(Document doc, List<ElementId> deletedElementIds)
        {
            string filePath = doc.PathName;
            foreach (ElementId deletedId in deletedElementIds)
            {
                bool deletedAnnotation = CashUtils.DictionaryMatchedAnnotationWithSpace[filePath].TryGetValue(deletedId, out ElementId matchSpaceId);
                bool deletedSpace = CashUtils.DictionaryMatchedSpaceWithAnnotation[filePath].TryGetValue(deletedId, out ElementId matchAnnotationId);
                bool deleteTerminal = CashUtils.DictionaryDuctTerminalBySpace[filePath].ContainsKey(deletedId);
                if (deletedAnnotation)
                {
                    doc.Delete(matchSpaceId);
                }
                else if (deletedSpace)
                {
                    doc.Delete(matchAnnotationId);
                }
                else if (deleteTerminal)
                {
                    ProcessAfterDeletedTerminal(deletedId, filePath, doc);
                }
                else
                {
                    LoggingUtils.Logging(Warnings.UnknownTypeRemovedElement(), doc.PathName);
                }
            }
        }
        private void ProcessAfterDeletedTerminal(ElementId deletedId, string filePath, Document doc)
        {
            ElementId spaceId = CashUtils.DictionaryDuctTerminalBySpace[filePath][deletedId];
            Space space = doc.GetElement(spaceId) as Space;
            if (space == null)
                LoggingUtils.LoggingWithMessage(Warnings.RecalculateDuctTerminalFailBecauseSpaceRemoved(deletedId), doc.PathName);
            else
                RecalculateDuctTerminals(doc, space);
        }
        private void RecalculateDuctTerminalsAirFlow(Document doc, Space space, List<FamilyInstance> ductTerminals, bool isSupply)
        {
            if (ductTerminals.Count == 0)
            {
                LoggingUtils.Logging(Warnings.InSpaceNotFoundDuctTerminals(space, isSupply), doc.PathName);
                return;
            }

            (List<FamilyInstance> ductTerminalsWithParamUpdaterOn, List<FamilyInstance> ductTerminalsWithParamUpdaterOff) =
                GroupTerminalByUpdaterOn(ductTerminals);

            if (ductTerminalsWithParamUpdaterOn.Count == 0)
            {
                LoggingUtils.Logging(Warnings.InSpaceNotFoundDuctTerminalsWithParamUpdaterOn(space, isSupply), doc.PathName);
                return;
            }

            string spaceAirflowParamName = GetSpaceAirflowParameterName(isSupply);

            double spaceTotalAirFlow = space.LookupParameter(spaceAirflowParamName).AsDouble();
            if (spaceTotalAirFlow == 0)
            {
                LoggingUtils.LoggingWithMessage(Warnings.SpaceTotalAirFlowEqualZero(spaceAirflowParamName, space), doc.PathName);
                return;
            }

            double airFlowForTerminalsWithUpdaterOff = GetAirFlowForTerminalsWithUpdaterOff(ductTerminalsWithParamUpdaterOff);

            double totalAirFlowWithUpdaterOn = spaceTotalAirFlow - airFlowForTerminalsWithUpdaterOff;
            if (totalAirFlowWithUpdaterOn < 0)
            {
                LoggingUtils.LoggingWithMessage(Warnings.ImpossibleRecalculateAirFlow(spaceAirflowParamName, ductTerminalsWithParamUpdaterOff.Select(d => d.Id), space.Id), doc.PathName);
                return;
            }

            double newAirFlowForOneDuctTerminalWithUpdaterOn = totalAirFlowWithUpdaterOn / ductTerminalsWithParamUpdaterOn.Count;
            foreach (var ductTerminal in ductTerminalsWithParamUpdaterOn)
            {
                ParameterUtils.SetParameterValueWithEqualCheck(ductTerminal.LookupParameter(Constants.PN_AIR_FLOW), newAirFlowForOneDuctTerminalWithUpdaterOn, ductTerminal.Id);
            }
        }
        private string GetSpaceAirflowParameterName(bool isSupply)
        {
            return isSupply ? Constants.PN_ADSK_CALCULATED_SUPPLY : Constants.PN_ADSK_CALCULATED_EXHAUST;
        }
        private double GetAirFlowForTerminalsWithUpdaterOff(List<FamilyInstance> ductTerminalsWithParamUpdaterOff)
        {
            return ductTerminalsWithParamUpdaterOff.Select(t => t.LookupParameter(Constants.PN_AIR_FLOW).AsDouble()).Sum();
        }
        private (List<FamilyInstance> ductTerminalsWithParamUpdaterOn, List<FamilyInstance> ductTerminalsWithParamUpdaterOff) GroupTerminalByUpdaterOn(List<FamilyInstance> ductTerminals)
        {
            List<FamilyInstance> ductTerminalsWithParamUpdaterOn = new();
            List<FamilyInstance> ductTerminalsWithParamUpdaterOff = new();
            foreach (var ductTerminal in ductTerminals)
            {
                if (IsParamUpdaterOn(ductTerminal))
                    ductTerminalsWithParamUpdaterOn.Add(ductTerminal);
                else
                    ductTerminalsWithParamUpdaterOff.Add(ductTerminal);
            }
            return (ductTerminalsWithParamUpdaterOn, ductTerminalsWithParamUpdaterOff);
        }
        private void ModifiedElementsParameterUpdaterOn(Document doc, Element element)
        {
            if (element is Space space)
                ProcessModifiedSpaceUpdaterOn(doc, space);

            else if (element is AnnotationSymbol annotation)
                ProcessModifiedAnnotationUpdaterOn(doc, annotation);

            else if (element is FamilyInstance ductTerminal && TerminalUtils.IsDuctTerminal(element))
                ProcessDuctTerminal(doc, ductTerminal);
        }
        private void ProcessDuctTerminal(Document doc, FamilyInstance ductTerminal)
        {
            Space space = TerminalUtils.GetSpaceByDuctTerminal(doc, ductTerminal);
            if (space == null)
            {
                LoggingUtils.LoggingWithMessage(Warnings.DuctTerminalChangeParameterValueAndNotLocateInSpace(), doc.PathName);
                return;
            }
            RecalculateDuctTerminals(doc, space);
        }
        private void RecalculateDuctTerminals(Document doc, Space space)
        {
            List<FamilyInstance> ductTerminals = GetDuctTerminalsBySpace(space);
           
            (List<FamilyInstance> supplyDuctTerminals, List<FamilyInstance> exhaustDuctTerminals) = TerminalUtils.GroupTerminalByType(ductTerminals);

            RecalculateDuctTerminalsAirFlow(doc, space, supplyDuctTerminals, isSupply: true);
            RecalculateDuctTerminalsAirFlow(doc, space, exhaustDuctTerminals, isSupply: false);

            LoggingUtils.Logging(Warnings.CompletedRecalculateAirFlowDuctTerminals(), doc.PathName);
        }


        public static bool IsParamUpdaterOn(Element element)
        {
            return element.LookupParameter(Constants.PN_ADSK_UPDATER).AsInteger() == 1;
        }

       
        private void ProcessModifiedAnnotationUpdaterOn(Document doc, AnnotationSymbol annotation)
        {
            ElementId spaceId = StorageHelper.GetMatchedSpaceId(annotation);

            Space space = doc.GetElement(spaceId) as Space;
            if (space == null)
            {
                LoggingUtils.LoggingWithMessage(Warnings.NotFoundSpaceMatchedWithAnnotation(annotation.Id), doc.PathName);
                return;
            }

            foreach (var kvp in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM)
            {
                string spaceParamName = kvp.Key;
                string annotationParamName = kvp.Value;

                Parameter annotationParameter = annotation.LookupParameter(annotationParamName);
                Parameter spaceParameter = space.LookupParameter(spaceParamName);

                object annotationParameterValue = ParameterUtils.NormalizeValue(ParameterUtils.GetParameterValue(annotationParameter));
                ParameterUtils.SetParameterValueWithEqualCheck(spaceParameter, annotationParameterValue, spaceId);
            }
            RecalculateDuctTerminals(doc, space); 
            LoggingUtils.Logging(Warnings.CompletedTransferParameterValueAnnotationInMatchedSpace(), doc.PathName);
        }
        private void ProcessModifiedSpaceUpdaterOn(Document doc, Space space)
        {
            ElementId annotationId = StorageHelper.GetMatchedAnnotationId(space);

            Element annotation = doc.GetElement(annotationId);

            RecalculateDuctTerminals(doc, space);

            if (annotation == null)
            {
                LoggingUtils.LoggingWithMessage(Warnings.NotFoundAnnotationMatchedWithSpace(space.Id), doc.PathName);
                return;
            }

            foreach (var keyValuePair in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM)
            {
                string spaceParamName = keyValuePair.Key;
                string annotationParamName = keyValuePair.Value;

                Parameter spaceParameter = space.LookupParameter(spaceParamName);
                Parameter annotationParameter = annotation.LookupParameter(annotationParamName);

                object spaceValue = ParameterUtils.NormalizeValue(ParameterUtils.GetParameterValue(spaceParameter));
                ParameterUtils.SetParameterValueWithEqualCheck(annotationParameter, spaceValue, annotationId);
            }
            LoggingUtils.Logging(Warnings.CompletedTransferParameterValueSpaceInMatchedAnnotation(), doc.PathName);
        }
        public static List<FamilyInstance> GetDuctTerminalsBySpace(Space space)
        {
            return CollectorUtils.GetDuctTerminals(space.Document)
                .Cast<FamilyInstance>()
                .Where(ductTerminal => ductTerminal.Space != null && ductTerminal.Space.Id == space.Id)
                .ToList();
        }

        public string GetUpdaterName()
        {
            return "Annotation Updater";
        }
        public string GetAdditionalInformation()
        {
            return "Этот апдейтер обновляет параметры аннотаций";
        }
        public ChangePriority GetChangePriority()
        {
            return ChangePriority.Annotations;
        }
        public UpdaterId GetUpdaterId()
        {
            return _m_updaterId;
        }
        public static void AddChangeParameterForDuctTerminalsTriggers(Element ductTerminal)
        {
            AddChangeParameterAirFlowForDuctTerminalsTriggers(ductTerminal);
            AddChangeParameterUpdaterForDuctTerminalsTriggers(ductTerminal);
        }

        public static void AddChangeParameterUpdaterForDuctTerminalsTriggers(Element ductTerminal)
        {
            ElementId parameterId = ductTerminal.LookupParameter(Constants.PN_ADSK_UPDATER).Id;
            UpdaterRegistry.AddTrigger(App._updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeParameter(parameterId));
        }
        public static void CreateDuctTerminalChangeSystemNameParameterTrigger(Element ductTerminal)
        {
            ElementId parameterId = ductTerminal.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).Id;
            UpdaterRegistry.AddTrigger(App._updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeParameter(parameterId));
        }
        public static void AddChangeParameterAirFlowForDuctTerminalsTriggers(Element ductTerminal)
        {
            ElementId parameterId = ductTerminal.LookupParameter(Constants.PN_AIR_FLOW).Id;
            UpdaterRegistry.AddTrigger(App._updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeParameter(parameterId));
        }
        public static void AddChangeParameterValueForSpacesTriggers(Element space)
        {
            foreach (string parameterName in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM.Keys)
            {
                ElementId parameterId = space.LookupParameter(parameterName).Id;
                UpdaterRegistry.AddTrigger(App._updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeParameter(parameterId));
            }
        }
        public static void AddChangeParameterValueForAnnotationsTriggers(Element annotationInstance)
        {
            foreach (string parameterName in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM.Values)
            {
                ElementId parameterId = annotationInstance.LookupParameter(parameterName).Id;
                UpdaterRegistry.AddTrigger(App._updater.GetUpdaterId(), FilterUtils._annotationFilter, Element.GetChangeTypeParameter(parameterId));
            }
        }
        public static void RemoveAllTriggers()
        {
            UpdaterRegistry.RemoveAllTriggers(App._updater.GetUpdaterId());
            App._triggerForSpaceParameterChangedCreated = false;
            App._triggerForAnnotationParameterChangedCreated = false;
            App._triggerForDuctTerminalParameterChangedCreated = false;
        }
    }
}