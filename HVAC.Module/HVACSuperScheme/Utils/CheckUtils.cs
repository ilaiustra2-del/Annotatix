using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using HVACSuperScheme.Updaters;

namespace HVACSuperScheme.Utils
{
    public static class CheckUtils
    {
        public static bool DocumentHasRequiredSharedParameters(Document doc)
        {
            List<string> sharedParameter = new List<string>()
            {
                Constants.PN_ADSK_NAME_EXHAUST_SYSTEM_FROM_LOCAL_SUCTION,
                Constants.PN_ADSK_CALCULATED_EXHAUST,
                Constants.PN_ADSK_NAME_EXHAUST_SYSTEM,
                Constants.PN_ADSK_CALCULATED_SUPPLY,
                Constants.PN_ADSK_NAME_SUPPLY_SYSTEM,
                Constants.PN_AIR_FLOW,
                Constants.PN_ADSK_UPDATER,
                Constants.PN_ADSK_ROOM_CATEGORY,
                Constants.PN_ADSK_ROOM_TEMPERATURE,
                Constants.PN_CLEAN_CLASS,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_1_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_2_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_3_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_4_WITH_UNDERSCORE,
                Constants.PN_AIR_FLOW_LOCAL_SUCTION_5_WITH_UNDERSCORE
            };

            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator it = bindingMap.ForwardIterator();

            HashSet<string> existingParameter = new();

            while (it.MoveNext())
            {
                Definition def = it.Key;
                if (def != null)
                    existingParameter.Add(def.Name);
            }
            foreach (string parameter in sharedParameter)
            {
                if (!existingParameter.Contains(parameter))
                {
                    LoggingUtils.LoggingWithMessage(Warnings.NotFoundSharedParametersInProject(parameter), doc.PathName);
                    return false;
                }
            }
            return true;
        }
        public static bool DocumentHasRequiredSharedParametersByCategory(Document doc, BuiltInCategory builtInCategory, List<string> requiredParameters)
        {
            Category category = doc.Settings.Categories.get_Item(builtInCategory);
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();

            HashSet<string> addedParameterNames = new HashSet<string>();
            
            while (iterator.MoveNext())
            {
                Definition definition = iterator.Key;
                ElementBinding binding = iterator.Current as ElementBinding;

                if (binding != null
                    && binding.Categories.Contains(category))
                {
                    addedParameterNames.Add(definition.Name);
                }
            }
            foreach (string requiredParameter in requiredParameters)
            {
                if (!addedParameterNames.Contains(requiredParameter))
                {
                    LoggingUtils.LoggingWithMessage(Warnings.NotFoundSharedParametersForCategory(requiredParameter, category.Name), doc.PathName);
                    return false;
                }
            }
            return true;
        }
        public static void CheckGeneral(Document doc, ViewDrafting draftView, ElementId annotationSymbolSymbolId)
        {
            CheckParameterAnnotation(doc, annotationSymbolSymbolId, draftView);
            CheckExistSpaceParameters(doc);
        }
        private static void CheckExistSpaceParameters(Document doc)
        {
            Space space = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .FirstOrDefault();

            if (space == null)
                throw new CustomException(ExceptionUtils.SpacesNotFound());
            foreach (var nameParameter in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM.Keys)
                if (space.LookupParameter(nameParameter) == null)
                    throw new CustomException(ExceptionUtils.ParameterNotFoundInSpace(nameParameter));
        }
        private static void CheckParameterAnnotation(Document doc, ElementId annotationSymbolSymbolId, ViewDrafting draftView)
        {
            AnnotationSymbolType annotationSymbol = doc.GetElement(annotationSymbolSymbolId) as AnnotationSymbolType;
            using var t = new Transaction(doc);
            
            t.Start("Создание временной аннотации для проверки параметров");
            FamilyInstance annotationInstance = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), annotationSymbol, draftView);
            var checkedPN = Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM.Values.ToList();
            checkedPN.Add(Constants.PN_VISIBLE_SUPPLY);
            checkedPN.Add(Constants.PN_VISIBLE_EXHAUST);
            checkedPN.Add(Constants.PN_VISIBLE_LOCAL_SUCTION);

            checkedPN.Add(Constants.PN_LEADER_LENGTH_EXHAUST);
            checkedPN.Add(Constants.PN_LEADER_LENGTH_LOCAL_SUCTION);
            checkedPN.Add(Constants.PN_LEADER_LENGTH_SUPPLY);

            
            checkedPN.Add(Constants.PN_COUNT_LOCAL_SUCTION);

            foreach (var nameParameter in checkedPN)
                if (annotationInstance.LookupParameter(nameParameter) == null)
                    throw new CustomException(ExceptionUtils.ParameterNotFoundInFamilySymbol(nameParameter));

            t.RollBack();  
        }
        public static void CheckDuplicateViewName(Document doc)
        {
            var multiclassFilter = new ElementMulticlassFilter(new List<Type>
                {
                    typeof(ViewPlan),
                    typeof(ViewSection),
                    typeof(View3D),
                    typeof(ViewSchedule)
                });

            List<View> existingViews = new FilteredElementCollector(doc)
                .WherePasses(multiclassFilter)
                .Cast<View>()
                .Where(v =>
                    !v.IsTemplate &&
                    v.Name.Equals(Constants.DRAFT_VIEW_NAME, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (existingViews.Any())
            {
                throw new CustomException(ExceptionUtils.DuplicateViewName());
            }
        }
    }
}
