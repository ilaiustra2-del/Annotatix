using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public static class CollectorUtils
    {
        public static List<Element> GetAnnotationInstances(Document doc)
        {
            return new FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_GenericAnnotation)
               .OfClass(typeof(FamilyInstance))
               .WhereElementIsNotElementType()
               .Where(elem =>
               {
                   if (elem is AnnotationSymbol annotation)
                   {
                       return annotation.Symbol.FamilyName == Constants.ANNOTATION_SYMBOL_FAMILY_NAME &&
                              annotation.Symbol.Name == Constants.ANNOTATION_SYMBOL_TYPE_NAME;
                   }
                   return false;
               })
               .ToList();
        }
        public static List<ViewDrafting> GetExistHVACSchemaViews(Document doc)
        {
            return new FilteredElementCollector(doc)
               .OfClass(typeof(ViewDrafting))
               .Cast<ViewDrafting>()
               .Where(v => v.Name.IndexOf(Constants.DRAFT_VIEW_NAME, StringComparison.OrdinalIgnoreCase) >= 0)
               .ToList();
        }
        public static ElementId GetTextNoteTypeId(Document doc, string textNoteTypeName)
        {
            IEnumerable<ElementId> textNoteTypeIds = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .Where(type => type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == textNoteTypeName)
                .Select(type => type.Id);

            if (!textNoteTypeIds.Any())
                throw new CustomException(ExceptionUtils.NotFoundFamilySymbol(textNoteTypeName)); 
            else if (textNoteTypeIds.Count() == 1)
                return textNoteTypeIds.First();
            else
                throw new CustomException(ExceptionUtils.FindDuplicateFamilySymbol(textNoteTypeName));
        }
        public static ElementId GetAnnotationSymbolSymbolId(Document doc)
        {
            Element annotationSymbolSymbol = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(type => type.Name == Constants.ANNOTATION_SYMBOL_TYPE_NAME
                    && type.FamilyName == Constants.ANNOTATION_SYMBOL_FAMILY_NAME);

            if (annotationSymbolSymbol == null)
                throw new CustomException(ExceptionUtils.NotFoundFamilySymbol(Constants.ANNOTATION_SYMBOL_TYPE_NAME));
            else
                return annotationSymbolSymbol.Id;
        }
        public static List<ElementId> GetSpacesForAnnotations(Document doc, bool withoutAnnotationInstance)
        {
            List<ElementId> spacesIds = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .Where(space => !string.IsNullOrWhiteSpace(space.LookupParameter(Constants.PN_NAME).AsString())
                    && !string.IsNullOrWhiteSpace(space.LookupParameter(Constants.PN_NUMBER).AsString()) 
                    && space.Area > 0
                    && (!string.IsNullOrWhiteSpace(space.LookupParameter(Constants.PN_ADSK_NAME_EXHAUST_SYSTEM).AsString())
                        || !string.IsNullOrWhiteSpace(space.LookupParameter(Constants.PN_ADSK_NAME_SUPPLY_SYSTEM).AsString()))
                    && (!withoutAnnotationInstance || StorageHelper.GetMatchedAnnotationId(space) == ElementId.InvalidElementId))
                .Select(space => space.Id)
                .ToList();
            if (!spacesIds.Any())
            {
                if (withoutAnnotationInstance)
                    throw new CustomException(ExceptionUtils.NotFoundedSpacesWithParamtersWithoutAnnotations());
                else
                    throw new CustomException(ExceptionUtils.NotFoundedSpacesWithParamters());
            } 
            return spacesIds;
        }
        public static List<Space> GetSpaces(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .Where(space => space.Area > 0)
                .ToList();
        }
        public static List<Element> GetDuctTerminals(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctTerminal)
                .WhereElementIsNotElementType()
                .ToList();
        }
    }
}
