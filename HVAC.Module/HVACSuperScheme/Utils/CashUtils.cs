using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using HVACSuperScheme.Updaters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace HVACSuperScheme.Utils
{
    internal class CashUtils
    {
        public static Dictionary<string, Dictionary<ElementId, ElementId>> DictionaryMatchedAnnotationWithSpace { get; set; } = new();
        public static Dictionary<string, Dictionary<ElementId, ElementId>> DictionaryMatchedSpaceWithAnnotation { get; set; } = new();
        public static Dictionary<string, Dictionary<ElementId, List<ElementId>>> DictionarySpaceByDuctTerminals { get; set; } = new();
        public static Dictionary<string, Dictionary<ElementId, ElementId>> DictionaryDuctTerminalBySpace { get; set; } = new();

        public static void InitializeCash(Document doc)
        {
            string filePath = doc.PathName;

            if (!DictionaryMatchedAnnotationWithSpace.ContainsKey(filePath))
                DictionaryMatchedAnnotationWithSpace[filePath] = new Dictionary<ElementId, ElementId>();

            if (!DictionaryMatchedSpaceWithAnnotation.ContainsKey(filePath))
                DictionaryMatchedSpaceWithAnnotation[filePath] = new Dictionary<ElementId, ElementId>();

            if (!DictionarySpaceByDuctTerminals.ContainsKey(filePath))
                DictionarySpaceByDuctTerminals[filePath] = new Dictionary<ElementId, List<ElementId>>();

            if (!DictionaryDuctTerminalBySpace.ContainsKey(filePath))
                DictionaryDuctTerminalBySpace[filePath] = new Dictionary<ElementId, ElementId>();


            InitializeAnnotationsCash(doc, filePath);
            InitializeSpacesCash(doc, filePath);
            InitializeDuctTerminalsCash(doc, filePath);
        }
        private static void InitializeDuctTerminalsCash(Document doc, string filePath)
        {
            foreach (FamilyInstance ductTerminal in CollectorUtils.GetDuctTerminals(doc))
            {
                Space space = TerminalUtils.GetSpaceByDuctTerminal(doc, ductTerminal);
                ElementId spaceId;
                if (space == null)
                {
                    spaceId = ElementId.InvalidElementId;
                }
                else
                {
                    spaceId = space.Id;

                    if (!DictionarySpaceByDuctTerminals[filePath].ContainsKey(spaceId))
                        DictionarySpaceByDuctTerminals[filePath][spaceId] = new List<ElementId>();
                    DictionarySpaceByDuctTerminals[filePath][spaceId].Add(ductTerminal.Id);

                }
                DictionaryDuctTerminalBySpace[filePath][ductTerminal.Id] = spaceId;
            }
        }
        public static void UpdateCashMatchedSpaceAndAnnotationAfterDeletion(Document doc, List<ElementId> deletedElementIds)
        {
            string filePath = doc.PathName;
            foreach (var deletedId in deletedElementIds)
            {
                bool deleteAnnotation = DictionaryMatchedAnnotationWithSpace[filePath].TryGetValue(deletedId, out ElementId spaceId);
                bool deleteSpace = DictionaryMatchedSpaceWithAnnotation[filePath].TryGetValue(deletedId, out ElementId annotationId);
                if (deleteAnnotation)
                {
                    DictionaryMatchedAnnotationWithSpace[filePath].Remove(deletedId);
                    DictionaryMatchedSpaceWithAnnotation[filePath].Remove(spaceId);
                }
                else if (deleteSpace)
                {
                    ElementId spaceDeletedId = deletedId;
                    DictionaryMatchedSpaceWithAnnotation[filePath].Remove(deletedId);
                    DictionaryMatchedAnnotationWithSpace[filePath].Remove(annotationId);
                }
                else
                {
                    
                }
            }
        }
        public static void UpdateCashDuctTerminalsAfterDeletion(List<ElementId> deletedElementIds, string filePath)
        {
            foreach (var deletedId in deletedElementIds)
            {
                bool deleteTerminal = DictionaryDuctTerminalBySpace[filePath].ContainsKey(deletedId);
                bool deleteSpace = DictionarySpaceByDuctTerminals[filePath].ContainsKey(deletedId);
                if (deleteTerminal)
                {
                    ElementId terminalDeletedId = deletedId;
                    ElementId spaceId = DictionaryDuctTerminalBySpace[filePath][terminalDeletedId];

                    DictionaryDuctTerminalBySpace[filePath].Remove(terminalDeletedId);
                    if (spaceId != ElementId.InvalidElementId)
                    {
                        DictionarySpaceByDuctTerminals[filePath][spaceId].Remove(terminalDeletedId);
                    }
                }
                else if (deleteSpace)
                {
                    ElementId spaceDeletedId = deletedId;
                    List<ElementId> ductTerminalIdsInSpace = DictionarySpaceByDuctTerminals[filePath][spaceDeletedId];
                    foreach (var terminalId in ductTerminalIdsInSpace)
                    {
                        DictionaryDuctTerminalBySpace[filePath][terminalId] = ElementId.InvalidElementId;
                    }
                    DictionarySpaceByDuctTerminals[filePath].Remove(spaceDeletedId);
                }
                else
                {
                    
                }
            }
        }
        public static void UpdateCashDuctTerminalsAfterAdded(Document doc, List<ElementId> addedElementIds, string filePath)
        {
            foreach (var addedId in addedElementIds)
            {
                var element = doc.GetElement(addedId);

                if (element is Space newSpace)
                {
                    var spaceId = newSpace.Id;
                    List<FamilyInstance> ductTerminalsInSpace = Updater.GetDuctTerminalsBySpace(newSpace);

                    DictionarySpaceByDuctTerminals[filePath][spaceId] = ductTerminalsInSpace.Select(t => t.Id).ToList();

                    foreach (var ductTerminal in ductTerminalsInSpace)
                    {
                        ElementId ductTerminalId = ductTerminal.Id;
                        DictionaryDuctTerminalBySpace[filePath][ductTerminalId] = spaceId;
                    }
                }
                else if (TerminalUtils.IsDuctTerminal(element))
                {
                    ElementId ductTerminalId = addedId;
                    Space space = TerminalUtils.GetSpaceByDuctTerminal(doc, element as FamilyInstance);

                    ElementId spaceId = space == null ? ElementId.InvalidElementId : space.Id;

                    DictionaryDuctTerminalBySpace[filePath][ductTerminalId] = spaceId;

                    if (!DictionarySpaceByDuctTerminals[filePath].ContainsKey(spaceId))
                        DictionarySpaceByDuctTerminals[filePath][spaceId] = new List<ElementId>();
                    DictionarySpaceByDuctTerminals[filePath][spaceId].Add(ductTerminalId);
                }
                else
                {
                    LoggingUtils.LoggingWithMessage(Warnings.UnknownTypeAddedElement(), filePath);
                }
            }
        }
        private static void InitializeSpacesCash(Document doc, string filePath)
        {
            foreach (Element space in CollectorUtils.GetSpaces(doc))
            {
                ElementId annotationId = StorageHelper.GetMatchedAnnotationId(space);
                if (annotationId == ElementId.InvalidElementId)
                {
                    continue;
                }

                Element annotation = doc.GetElement(annotationId);
                if (annotation == null)
                {
                    StorageHelper.TryClearExtensibleStorageWithTransaction(doc, space);
                    continue;
                }

                ElementId candidateSpaceId = StorageHelper.GetMatchedSpaceId(annotation);
                if (candidateSpaceId.IntegerValue == space.Id.IntegerValue)
                {
                    DictionaryMatchedSpaceWithAnnotation[filePath][space.Id] = annotationId;
                }
                else
                {
                    StorageHelper.TryClearExtensibleStorageWithTransaction(doc, space);
                }
            }
        }
        private static void InitializeAnnotationsCash(Document doc, string filePath)
        {
            foreach (Element annotation in CollectorUtils.GetAnnotationInstances(doc))
            {
                ElementId spaceId = StorageHelper.GetMatchedSpaceId(annotation);
                if (spaceId == ElementId.InvalidElementId) 
                    continue;

                Element space = doc.GetElement(spaceId); 
                if (space == null)
                {
                    StorageHelper.TryClearExtensibleStorageWithTransaction(doc, annotation);
                    continue; 
                }

                ElementId annotationId = StorageHelper.GetMatchedAnnotationId(space); 
                if (annotationId.IntegerValue == annotation.Id.IntegerValue)
                {
                    DictionaryMatchedAnnotationWithSpace[filePath][annotation.Id] = spaceId; 
                }
                else
                {
                    StorageHelper.TryClearExtensibleStorageWithTransaction(doc, annotation);
                }
            }
        }
        public static void RemoveCashSpacesAndAnnotations(Document doc)
        {
            string filePath = doc.PathName;
            DictionaryMatchedAnnotationWithSpace.Remove(filePath);
            DictionaryMatchedSpaceWithAnnotation.Remove(filePath);
            DictionaryDuctTerminalBySpace.Remove(filePath);
            DictionarySpaceByDuctTerminals.Remove(filePath);
        }
    }
}
