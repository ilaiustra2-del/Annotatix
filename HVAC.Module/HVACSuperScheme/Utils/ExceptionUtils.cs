using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using HVACSuperScheme.Updaters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public static class ExceptionUtils
    {
        public static string NotFoundFamilySymbol(string familySymbolName)
        {
            return $"Не найден типоразмер с именем {familySymbolName}";
        }
        public static string FindDuplicateFamilySymbol(string familySymbolSymbolName)
        {
            return $"Найдено несколько типоразмеров с одинаковыми именами - {familySymbolSymbolName}";
        }
        public static string ParameterNotFoundInFamilySymbol(string nameParameter)
        {
            return $"Параметр экземпляра с именем {nameParameter} не существует в семействе с типоразмером {Constants.ANNOTATION_SYMBOL_TYPE_NAME}. Добавьте необходимый параметр и повторите запуск плагина.";
        }
        public static string ParameterNotFoundInSpace(string nameParameter)
        {
            return $"Параметр экземпляра с именем {nameParameter} не существует у пространства. Добавьте необходимый параметр и повторите запуск плагина.";
        }
        public static string SpacesNotFound()
        {
            return "Пространства отсутствуют в проекте";
        }
        public static string NotFoundedSpacesWithParamtersWithoutAnnotations()
        {
            return $"В файле не найдены размещенные пространства, у которых нет аннотаций, и у которых заполнены параметры {Constants.PN_NAME}, {Constants.PN_NUMBER}, {Constants.PN_ADSK_NAME_EXHAUST_SYSTEM}, {Constants.PN_ADSK_NAME_SUPPLY_SYSTEM}.";
        }
        public static string NotFoundedSpacesWithParamters()
        {
            return $"В файле не найдены размещенные пространства, у которых заполнены параметры {Constants.PN_NAME}, {Constants.PN_NUMBER}, {Constants.PN_ADSK_NAME_EXHAUST_SYSTEM}, {Constants.PN_ADSK_NAME_SUPPLY_SYSTEM}.";
        }
        public static string ConversionFailedStringToNumber()
        {
            return "Не удалось преобразовать строку в число.";
        }
        public static string DuplicateViewName()
        {
            return $"В проекте уже существует нечертежный вид с именем {Constants.DRAFT_VIEW_NAME}, для корректной работы не должно быть повторяющихся видов.";
        }
        public static string NotFoundValidDraftViews()
        {
            return "В файле не найдены подходящие чертёжные виды.";
        }
        public static string NotFoundValidSpacesWithoutAnnotations()
        {
            return "В файле не найдены подходящие пространства, у которых не создана аннотация.";
        }
        public static string IdlingHandlerError(Exception ex)
        {
            return $"Произошла ошибка при выполнении IdlingHandler. {ex.Message + ex.StackTrace}";
        }
        public static string UpdaterError(Exception ex)
        {
            return $"Ошибка апдейтер. {ex.Message + ex.StackTrace}";
        }
        public static string Error(Exception ex)
        {
            return $"Ошибка. {ex.Message}";
        }
        public static string SystemError(Exception ex)
        {
            return $"Системная ошибка. {ex.Message + ex.StackTrace}";
        }
    }
}
