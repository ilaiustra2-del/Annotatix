using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Mechanical;
using System.Windows;
using System.Reflection;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using HVACSuperScheme.Updaters;

namespace HVACSuperScheme.Utils
{
    public static class StorageHelper
    {
        private static readonly Guid SchemaGuid = new Guid("63BBF72E-1A12-4D75-951B-B0CB7B4A870E");

        private static Schema GetOrCreateSchema()
        {
            var existing = Schema.Lookup(SchemaGuid);
            if (existing != null)
                return existing;

            SchemaBuilder builder = new SchemaBuilder(SchemaGuid);
            builder.SetSchemaName(Constants.KW_ANNOTATION_SPACE_LINK);

            builder.AddSimpleField(Constants.KW_ANNOTATION_ID, typeof(ElementId));
            builder.AddSimpleField(Constants.KW_SPACE_ID, typeof(ElementId));

            return builder.Finish();
        }
        public static void ClearExtensibleStorageForMatchDeletionElement(Document doc, List<ElementId> deletedElementIds)
        {
            string filePath = doc.PathName;

            foreach (ElementId deletedId in deletedElementIds)
            {
                bool deleteAnnotation = CashUtils.DictionaryMatchedAnnotationWithSpace[filePath].TryGetValue(deletedId, out ElementId spaceId);
                bool deleteSpace = CashUtils.DictionaryMatchedSpaceWithAnnotation[filePath].TryGetValue(deletedId, out ElementId annotationId);
                if (deleteAnnotation)
                {
                    Element space = doc.GetElement(spaceId);
                    TryClearExtensibleStorage(doc, space);
                }
                else if (deleteSpace)
                {
                    Element annotation = doc.GetElement(annotationId);
                    TryClearExtensibleStorage(doc, annotation);
                }
                else
                {
                    LoggingUtils.Logging(Warnings.DeletedSpaceOrAnnotationWithoutMatchOrDeletedDuctTerminal(deletedId), doc.PathName);
                }
            }
        }
        public static ElementId GetMatchedSpaceId(Element annotationInstance)
        {
            Schema schema = GetOrCreateSchema();
            Entity entity = annotationInstance.GetEntity(schema);

            if (entity.IsValid())
            {
                Field spaceField = schema.GetField(Constants.KW_SPACE_ID);
                return entity.Get<ElementId>(spaceField);
            }

            return ElementId.InvalidElementId;
        }
        public static ElementId GetMatchedAnnotationId(Element space)
        {
            Schema schema = GetOrCreateSchema();
            Entity entity = space.GetEntity(schema);

            if (entity.IsValid())
            {
                Field annField = schema.GetField(Constants.KW_ANNOTATION_ID);
                return entity.Get<ElementId>(annField);
            }

            return ElementId.InvalidElementId;
        }
        public static void FillSchemaData(Document doc, Space space, FamilyInstance annotationInstance)
        {
            Schema schema = GetOrCreateSchema();

            Field annotationField = schema.GetField(Constants.KW_ANNOTATION_ID);
            Field spaceField = schema.GetField(Constants.KW_SPACE_ID);

            Entity entityAnnotation = new Entity(schema);
            entityAnnotation.Set(spaceField, space.Id);
            annotationInstance.SetEntity(entityAnnotation);

            Entity entitySpace = new Entity(schema);
            entitySpace.Set(annotationField, annotationInstance.Id);
            space.SetEntity(entitySpace);
        }
        public static void TryClearExtensibleStorageWithTransaction(Document doc, Element element)
        {
            Schema schema = GetOrCreateSchema();
            Entity entity = element.GetEntity(schema);

            if (!entity.IsValid())
                return;

            using var t = new Transaction(doc);
            t.Start("Очищаем Extensible Storage");
            element.DeleteEntity(schema);
            t.Commit();
        }
        public static void TryClearExtensibleStorage(Document doc, Element element)
        {
            Schema schema = GetOrCreateSchema();
            Entity entity = element.GetEntity(schema);

            if (!entity.IsValid())
                return;
            element.DeleteEntity(schema);
        }
    }
}
