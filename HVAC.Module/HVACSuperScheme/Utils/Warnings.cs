using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public static class Warnings
    {
        public static string NotFoundAnnotationMatchedWithSpace(ElementId spaceId)
        {
            return $"Изменен параметр пространства {spaceId.IntegerValue}, у которого не найдено связанная аннотация" +
                $"(пространство создано вручную или была удалена аннотация с выключенным апдейтером).";
        }
        public static string NotFoundSpaceMatchedWithAnnotation(ElementId annotationId)
        {
            return $"Не найдено пространство связанное с аннотацией {annotationId.IntegerValue}";
        }
        public static string SpaceTotalAirFlowEqualZero(string spaceAirflowParamName, Space space)
        {
            return $"Значение параметра {spaceAirflowParamName} у пространства {space.Id} равно 0. Расчёт расходов у ВРУ невозможен.";
        }
        public static string ImpossibleRecalculateAirFlow(string spaceAirflowParamName, IEnumerable<ElementId> ductTerminalsWithParamUpdaterOffIds, ElementId spaceId)
        {
            return $"Невозможно пересчитать расходы, так как сумма расходов ВРУ: {string.Join("; ", ductTerminalsWithParamUpdaterOffIds)} " +
                $"с выключеным параметром {Constants.PN_ADSK_UPDATER} больше, чем значение параметра {spaceAirflowParamName} пространства {spaceId}.";
        }
        public static string NotFoundSharedParametersInProject(string parameter)
        {
            return $"Общий параметр {parameter} не найден в проекте. Синхронизация принудительно отключена. Требуется добавить параметр, сохранить модель и перезапустить ревит.";
        }
        public static string InSpaceNotFoundDuctTerminals(Space space, bool isSupply)
        {
            string supplyOrExhaust = isSupply ? "приточных" : "вытяжных";
            return $"В пространстве {space.Id} не обнаружено {supplyOrExhaust} ВРУ";
        }
        public static string InSpaceNotFoundDuctTerminalsWithParamUpdaterOn(Space space, bool isSupply)
        {
            string supplyOrExhaust = isSupply ? "приточной" : "вытяжной";
            return $"В пространстве {space.Id} отсутсвуют ВРУ с включенным параметром {Constants.PN_ADSK_UPDATER} у {supplyOrExhaust} системы. " +
                $"Требуется проверить расход пространства и суммарный расход ВРУ вручную. Возможны расхождения в сумме расходов по ВРУ и в расходе пространства.";
        }
        public static string RecalculateDuctTerminalFailBecauseSpaceRemoved(ElementId deletedId)
        {
            return $"Удалено ВРУ {deletedId}, но выполнить перерасчёт расходов невозможно, так как одновременно было удалено пространство, в котором находилось это ВРУ.";
        }
        public static string UnknownTypeRemovedElement()
        {
            return "Неизвестный тип удаляемого элемента, или пришедшая пара удалённого элемента, пространство или аннотация из словаря";
        }
        public static string UnknownTypeAddedElement()
        {
            return "Неизвестный тип добавляемого элемента";
        }
        public static string DuctTerminalChangeParameterValueAndNotLocateInSpace()
        {
            return "ВРУ изменило значение какого-то параметра, но оно находится не в пространстве. Изменение не оказало влияние на другие элементы.";
        }
        public static string CompletedRecalculateAirFlowDuctTerminals()
        {
            return "Произведен перерасчёт расходов ВРУ.";
        }
        public static string CompletedTransferParameterValueAnnotationInMatchedSpace()
        {
            return "Произведен перенос значений параметров из аннотации в связанное пространство.";
        }
        public static string CompletedTransferParameterValueSpaceInMatchedAnnotation()
        {
            return "Произведен перенос значений параметров из пространства в связанную аннотацию.";
        }
        public static string DeletedSpaceOrAnnotationWithoutMatchOrDeletedDuctTerminal(ElementId deletedId)
        {
            return $"Удалили {deletedId} пространство или аннотацию без метча или удалили ВРУ.";
        }
        public static string CreatedDuctTerminalNotLocateInSpaceRecalculateAirFlowFailed()
        {
            return "Созданный ВРУ не находится в пространстве, перерасчёт расходов не был произведен.";
        }
        public static string NotCompletedSynchronizeParametersWithLinkedSpace()
        {
            return $"Был изменен параметр аннотации, но {Constants.PN_ADSK_UPDATER} выключен у аннотации, поэтому синхронизация параметров не была произведена со связанным пространством.";
        }
        public static string NotCompletedSynchronizeParametersWithLinkedAnnotation()
        {
            return $"Был изменен параметр пространства, но {Constants.PN_ADSK_UPDATER} выключен у пространства, поэтому синхронизация параметров не была произведена со связанной аннотацией.";
        }
        public static string ModifiedElements(List<ElementId> modifiedElementIds)
        {
            return $"Изменился/-лись следующий/-ие элемент/-ы {string.Join(", ", modifiedElementIds.Select(id => id.IntegerValue))}";
        }
        public static string AddedElements(List<ElementId> addedElementIds)
        {
            return $"Добавился/-лись следующий/-ие элемент/-ы {string.Join(", ", addedElementIds.Select(id => id.IntegerValue))}";
        }
        public static string DeletedElements(List<ElementId> deletedElementIds)
        {
            return $"Удалился/-лись следующий/-ие элемент/-ы {string.Join(", ", deletedElementIds.Select(id => id.IntegerValue))}";
        }
        public static string SubscribeToIdlingComplete()
        {
            return "Подписка на idling выполнена";
        }
        public static string TriggersOnDeleteElementsAdded()
        {
            return "Триггеры на удаление элемнтов добавлены.";
        }
        public static string TriggersOnAddElementsAdded()
        {
            return "Тригеры на добавления элемнтов добавлены.";
        }
        public static string TriggersOnChangeDuctTermilalParametersCreated()
        {
            return "Триггеры на изменение параметров ВРУ созданы.";
        }
        public static string TriggersOnChangeSpaceParametersCreated()
        {
            return "Триггеры на изменение параметров пространств созданы.";
        }
        public static string TriggersOnChangeAnnotationParametersCreated()
        {
            return "Триггеры на изменение параметров аннотаций созданы.";
        }
        public static string TriggersOnChangeSuccessfulCreated()
        {
            return "Триггеры на изменение параметров успешно созданы.";
        }
        public static string TriggersDeleted()
        {
            return "Триггеры удалены";
        }
        public static string NotFoundSharedParametersForCategory(string parameter, string categoryName)
        {
            return $"Общий параметр {parameter} не найден в категории {categoryName}. Синхронизация принудительно отключена. Требуется добавить параметр для корректной работы апдейтера.";
        }
        public static string FileNotFoundSetDefaultSettings()
        {
            return $"Файл настроек не существует, приняты настройки по умолчанию.";
        }
        public static string FileSettingsChangedSetNewSettings()
        {
            return $"Формат файла настроек был изменен, принят новый формат. Требуется пересохранить настройки.";
        }
        public static string InProjectNotFoundDuctTerminal()
        {
            return "В проекте не найден воздухораспределитель, добавьте его в проект, для корректной работы Апдейтера";
        }
        public static string DuctTerminalHasNotParameter()
        {
            return $"У воздухораспределителя отсутствует параметр {Constants.PN_AIR_FLOW}.";
        }
        public static string ElementParameterValueNotChanged(Parameter parameter, ElementId id, string oldValue)
        {
            return $"Значение параметра {parameter.Definition.Name} элемента {id.IntegerValue} не было изменено, значение {oldValue}";
        }
        public static string ElementParameterValueChanged(Parameter parameter, ElementId id, string oldValue, string newValue)
        {
            return $"Значение параметра {parameter.Definition.Name} элемента {id.IntegerValue} изменено, значение cтарое {oldValue}, значение новое {newValue}";
        }
        public static string UnsubscribeToIdlingComplete()
        {
            return "Принудительное завершение работы IdlingHandler так как он был активен, привязка к определённым триггерам не была завершена, апдейтер может работать нестабильно";
        }
        public static string UpdaterExecuteStart()
        {
            return "\nUpdater Execute начал работу\n";
        }
        public static string UpdaterExecuteFinish()
        {
            return "\nUpdater Execute закончил работу\n";
        }
    }
}
