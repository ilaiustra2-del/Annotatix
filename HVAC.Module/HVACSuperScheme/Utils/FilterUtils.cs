using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public class FilterUtils
    {
        public static ElementCategoryFilter _spaceFilter;
        public static LogicalAndFilter _ductTerminalFilter;
        public static LogicalAndFilter _annotationFilter;
        public static void InitElementFilters()
        {
            _spaceFilter = CreateSpaceFilter();
            _ductTerminalFilter = CreateDuctTerminalFilter();
            _annotationFilter = CreateAnnotationFilter();
        }
        public static ElementCategoryFilter CreateSpaceFilter()
        {
            return new ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces);
        }
        public static LogicalAndFilter CreateDuctTerminalFilter()
        {
            return new LogicalAndFilter(
                new ElementCategoryFilter(BuiltInCategory.OST_DuctTerminal),
                new ElementClassFilter(typeof(FamilyInstance)));
        }
        public static LogicalAndFilter CreateAnnotationFilter()
        {
            var annotationSymbolCategoryAndClassFilter = new LogicalAndFilter(
                new ElementCategoryFilter(BuiltInCategory.OST_GenericAnnotation),
                new ElementClassFilter(typeof(FamilyInstance)));

            var annotationSymbolTypeNameRule = ParameterFilterRuleFactory.CreateEqualsRule(new ElementId(
                BuiltInParameter.SYMBOL_NAME_PARAM), Constants.ANNOTATION_SYMBOL_TYPE_NAME, true);
            var annotationSymbolTypeNameFilter = new ElementParameterFilter(annotationSymbolTypeNameRule);

            var annotationSymbolFamilyNameRule = ParameterFilterRuleFactory.CreateEqualsRule(new ElementId(
                BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM), Constants.ANNOTATION_SYMBOL_FAMILY_NAME, true);
            var annotationSymbolFamilyNameFilter = new ElementParameterFilter(annotationSymbolFamilyNameRule);

            var annotationSymbolTypeAndFamilyNameFilter = new LogicalAndFilter(annotationSymbolTypeNameFilter, annotationSymbolFamilyNameFilter);

            return new LogicalAndFilter(annotationSymbolCategoryAndClassFilter, annotationSymbolTypeAndFamilyNameFilter);
        }
    }
}
