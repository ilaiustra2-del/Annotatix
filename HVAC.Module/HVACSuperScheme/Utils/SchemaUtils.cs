using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public static class SchemaUtils
    {
        public static void CreateAnnotations(Document doc, IEnumerable<ElementId> spaceIds, ElementId annotationSymbolSymbolId,
            ElementId textNoteTypeId, ViewDrafting draftView, XYZ insertPoint)
        {
            string filePath = doc.PathName;
            AnnotationSymbolType annotationSymbol = doc.GetElement(annotationSymbolSymbolId) as AnnotationSymbolType;

            List<Space> spaces = spaceIds.Select(doc.GetElement).Cast<Space>().ToList();
            int curLevelNumber = 0;
            string groupSpaceParameterName = Constants.PN_LEVEL;
            using var t = new Transaction(doc);
            foreach (var groupByLevel in spaces
                .GroupBy(s => s.LookupParameter(groupSpaceParameterName).AsElementId())
                .OrderByDescending(group => (doc.GetElement(group.Key) as Level).Elevation))
            {
                string level = (doc.GetElement(groupByLevel.Key) as Level).Name;
                List<Space> spacesByLevel = groupByLevel.ToList();

                t.Start("Создание текста уровня");
                TextNote.Create(doc, draftView.Id, new XYZ(-3.5, insertPoint.Y, 0), (doc.GetElement(groupByLevel.Key) as Level).Name, textNoteTypeId);
                t.Commit();

                List<List<string>> systemNamesBySpaces = GetSystemNamesForSpaces(spacesByLevel);

                List<List<string>> uniqueSystemNamesBySpaces = new SystemUtils().GetKeysForSpaces(systemNamesBySpaces);

                foreach (var uniqueSystemNamesBySpace in uniqueSystemNamesBySpaces)
                {
                    List<Space> spacesBySystem = GetSpacesBySystemNames(spacesByLevel, uniqueSystemNamesBySpace)
                        .OrderBy(space => ExtractAndCombineNumbers(space.Number)).ToList();
                    foreach (var space in spacesBySystem)
                    {
                        FamilyInstance annotationInstance;

                        t.Start("Создание аннотаций");
                        annotationInstance = doc.Create.NewFamilyInstance(insertPoint, annotationSymbol, draftView);
                        StorageHelper.FillSchemaData(doc, space, annotationInstance);
                        CashUtils.DictionaryMatchedSpaceWithAnnotation[filePath][space.Id] = annotationInstance.Id;
                        CashUtils.DictionaryMatchedAnnotationWithSpace[filePath][annotationInstance.Id] = space.Id;

                        t.Commit();

                        t.Start("Предварительная очистка параметров");
                        ParameterUtils.CleanParameters(annotationInstance);
                        t.Commit();

                        t.Start("Заполнение параметров");
                        SetParameterForAnnotation(annotationInstance, space);
                        t.Commit();
                        insertPoint += new XYZ(Constants.WIDTH_ANNOTATION, 0, 0);
                    }
                    TextNote systemNameText;

                    t.Start("Создание текста имен системы");
                    systemNameText = TextNote.Create(doc, draftView.Id,
                        new XYZ(insertPoint.X - (spacesBySystem.Count * Constants.WIDTH_ANNOTATION / 2) - Constants.WIDTH_ANNOTATION / 2,
                            insertPoint.Y - Constants.VERTICAL_DISTANCE_TO_SYSTEM_NAME_TEXT,
                            0),
                        string.Join("/", uniqueSystemNamesBySpace), textNoteTypeId);
                    t.Commit();

                    t.Start("Центрирование текста имен систем");
                    BoundingBoxXYZ bb = systemNameText.get_BoundingBox(draftView);
                    double moveByX = (bb.Max.X - bb.Min.X) / 2;
                    ElementTransformUtils.MoveElement(doc, systemNameText.Id, new XYZ(-moveByX, 0, 0));
                    t.Commit();

                    insertPoint += new XYZ(Constants.HORIZONTAL_DISTANCE_BETWEEN_GROUP_IN_ROW, 0, 0);
                }
                insertPoint = new XYZ(0, insertPoint.Y - Constants.HEIGHT_ANNOTATION - Constants.VERTICAL_DISTANCE_BETWEEN_ROW, 0);
                curLevelNumber++;
            }
        }
        public static void SetParameterForAnnotation(FamilyInstance annotationInstance, Space space)
        {
            double height = 10.0 / 304.8;
            foreach (var parametersPair in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM)
            {
                string spacePN = parametersPair.Key;
                string annotationPN = parametersPair.Value;

                object value = ParameterUtils.GetParameterValue(space.LookupParameter(spacePN));
                bool isValueEmpty = ParameterUtils.IsValueEmpty(value);

                if (!isValueEmpty)
                    ParameterUtils.SetValue(annotationInstance.LookupParameter(annotationPN), value);

                if (spacePN == Constants.PN_ADSK_NAME_EXHAUST_SYSTEM_FROM_LOCAL_SUCTION)
                    if (isValueEmpty)
                        annotationInstance.LookupParameter(Constants.PN_VISIBLE_LOCAL_SUCTION).Set(0);
                    else
                        annotationInstance.LookupParameter(Constants.PN_VISIBLE_LOCAL_SUCTION).Set(1);

                if (spacePN == Constants.PN_ADSK_NAME_EXHAUST_SYSTEM)
                    if (isValueEmpty)
                        annotationInstance.LookupParameter(Constants.PN_VISIBLE_EXHAUST).Set(0); 
                    else
                        annotationInstance.LookupParameter(Constants.PN_VISIBLE_EXHAUST).Set(1);

                if (spacePN == Constants.PN_ADSK_NAME_SUPPLY_SYSTEM)
                    if (isValueEmpty)
                        annotationInstance.LookupParameter(Constants.PN_VISIBLE_SUPPLY).Set(0); 
                    else
                        annotationInstance.LookupParameter(Constants.PN_VISIBLE_SUPPLY).Set(1);
            }
            annotationInstance.LookupParameter(Constants.PN_LEADER_LENGTH_SUPPLY).Set(height);
            annotationInstance.LookupParameter(Constants.PN_LEADER_LENGTH_LOCAL_SUCTION).Set(height);
            annotationInstance.LookupParameter(Constants.PN_LEADER_LENGTH_EXHAUST).Set(height);
            annotationInstance.LookupParameter(Constants.PN_COUNT_LOCAL_SUCTION)?.Set(GetCountLocalSuctions(space));
        }
        public static int GetCountLocalSuctions(Space space)
        {
            int countLocalSuction = 0;
            List<string> localSuctionParameterNames = new List<string>()
            {
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_1_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_2_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_3_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_4_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_5_WITH_UNDERSCORE,
            };
            foreach (var pn in localSuctionParameterNames)
            {
                double value = space.LookupParameter(pn).AsDouble();
                if (value != 0.0)
                    countLocalSuction++;
            }
            return countLocalSuction;
        }
        private static long ExtractAndCombineNumbers(string input)
        {
            string combinedNumberStr = "";

            Regex regex = new Regex(@"\d+");
            MatchCollection matches = regex.Matches(input);

            foreach (Match match in matches)
            {
                combinedNumberStr += match.Value;
            }

            if (long.TryParse(combinedNumberStr, out long combinedNumber))
            {
                return combinedNumber;
            }
            else
            {
                throw new CustomException(ExceptionUtils.ConversionFailedStringToNumber());
            }
        }
        private static List<Space> GetSpacesBySystemNames(List<Space> spacesByLevel, List<string> uniqueSystemNamesBySpace)
        {
            List<Space> filterSpaces = new List<Space>();
            foreach (var space in spacesByLevel)
            {
                string exhaustSystemName = space.LookupParameter(Constants.PN_ADSK_NAME_EXHAUST_SYSTEM).AsString();
                string supplySystemName = space.LookupParameter(Constants.PN_ADSK_NAME_SUPPLY_SYSTEM).AsString();
                if (!string.IsNullOrWhiteSpace(exhaustSystemName)
                    && uniqueSystemNamesBySpace.Contains(exhaustSystemName)
                    && !filterSpaces.Select(s => s.Id).Contains(space.Id))
                {
                    filterSpaces.Add(space);
                }
                if (!string.IsNullOrWhiteSpace(supplySystemName)
                    && uniqueSystemNamesBySpace.Contains(supplySystemName)
                    && !filterSpaces.Select(s => s.Id).Contains(space.Id))
                {
                    filterSpaces.Add(space);
                }
            }
            return filterSpaces;
        }
        private static List<List<string>> GetSystemNamesForSpaces(List<Space> spacesByLevel)
        {
            List<List<string>> systemNamesForSpaces = new List<List<string>>();
            foreach (var space in spacesByLevel)
            {
                List<string> systemNamesSpace = new List<string>();
                string exhaustSystemName = space.LookupParameter(Constants.PN_ADSK_NAME_EXHAUST_SYSTEM).AsString();
                string supplySystemName = space.LookupParameter(Constants.PN_ADSK_NAME_SUPPLY_SYSTEM).AsString();
                if (!string.IsNullOrWhiteSpace(exhaustSystemName))
                {
                    systemNamesSpace.Add(exhaustSystemName);
                }
                if (!string.IsNullOrWhiteSpace(supplySystemName))
                {
                    systemNamesSpace.Add(supplySystemName);
                }
                systemNamesForSpaces.Add(systemNamesSpace);
            }
            return systemNamesForSpaces;
        }
    }
}
