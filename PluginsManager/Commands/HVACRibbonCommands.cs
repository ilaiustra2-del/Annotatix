using System;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Базовый класс для ribbon команд HVAC
    /// </summary>
    public abstract class HvacRibbonCommandBase : IExternalCommand
    {
        protected abstract string CommandName { get; }
        protected abstract string HandlerTypeName { get; }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc?.Document;

                if (doc == null)
                {
                    TaskDialog.Show("HVAC", "Нет активного документа Revit");
                    return Result.Failed;
                }

                Core.DebugLogger.Log($"[HVAC-RIBBON] {CommandName} started");

                // Получаем модуль HVAC
                var module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                if (module == null)
                {
                    TaskDialog.Show("HVAC", "Модуль HVAC не загружен. Откройте Plugins Hub для авторизации.");
                    return Result.Failed;
                }

                // Получаем сборку модуля
                var moduleAssembly = module.GetType().Assembly;
                Core.DebugLogger.Log($"[HVAC-RIBBON] Module assembly: {moduleAssembly.FullName}");

                // Получаем тип handler'а
                var handlerType = moduleAssembly.GetType(HandlerTypeName);
                if (handlerType == null)
                {
                    Core.DebugLogger.Log($"[HVAC-RIBBON] ERROR: Handler type not found: {HandlerTypeName}");
                    TaskDialog.Show("HVAC", $"Не найден тип обработчика: {HandlerTypeName}");
                    return Result.Failed;
                }

                // Создаем экземпляр handler'а
                var handler = Activator.CreateInstance(handlerType) as IExternalEventHandler;
                if (handler == null)
                {
                    Core.DebugLogger.Log($"[HVAC-RIBBON] ERROR: Could not create handler instance");
                    TaskDialog.Show("HVAC", "Не удалось создать обработчик");
                    return Result.Failed;
                }

                // Создаём и запускаем ExternalEvent
                var extEvent = ExternalEvent.Create(handler);
                extEvent.Raise();

                Core.DebugLogger.Log($"[HVAC-RIBBON] {CommandName} event raised");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[HVAC-RIBBON] ERROR: {ex.Message}\n{ex.StackTrace}");
                TaskDialog.Show("HVAC", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Ribbon команда "Построить схему"
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class HvacCreateSchemaRibbonCommand : HvacRibbonCommandBase
    {
        protected override string CommandName => "HvacCreateSchema";
        protected override string HandlerTypeName => "HVAC.Module.Commands.CreateSchemaHandler";
    }

    /// <summary>
    /// Ribbon команда "Достроить схему"
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class HvacCompleteSchemaRibbonCommand : HvacRibbonCommandBase
    {
        protected override string CommandName => "HvacCompleteSchema";
        protected override string HandlerTypeName => "HVAC.Module.Commands.CompleteSchemaHandler";
    }

    /// <summary>
    /// Ribbon команда для переключения синхронизации чертежа с моделью
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class HvacSyncToggleRibbonCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Получаем модуль HVAC
                var module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                if (module == null)
                {
                    TaskDialog.Show("HVAC", "Модуль HVAC не загружен. Откройте Plugins Hub для авторизации.");
                    return Result.Failed;
                }

                // Получаем сборку модуля
                var moduleAssembly = module.GetType().Assembly;

                // Получаем тип SettingStorage
                var settingStorageType = moduleAssembly.GetType("HVACSuperScheme.Commands.Settings.SettingStorage");
                if (settingStorageType == null)
                {
                    Core.DebugLogger.Log($"[HVAC-RIBBON] SettingStorage type not found");
                    TaskDialog.Show("HVAC", "Не найден тип SettingStorage");
                    return Result.Failed;
                }

                // Вызываем ReadSettings
                var readSettingsMethod = settingStorageType.GetMethod("ReadSettings");
                readSettingsMethod?.Invoke(null, null);

                // Получаем Instance и IsUpdaterSync
                var instanceProperty = settingStorageType.GetProperty("Instance");
                var instance = instanceProperty?.GetValue(null);
                var isUpdaterSyncProperty = instance?.GetType().GetProperty("IsUpdaterSync");
                bool currentValue = (bool)(isUpdaterSyncProperty?.GetValue(instance) ?? false);
                
                // Переключаем состояние
                bool newState = !currentValue;
                isUpdaterSyncProperty?.SetValue(instance, newState);

                // Сохраняем настройки
                var saveSettingsMethod = settingStorageType.GetMethod("SaveSettings");
                saveSettingsMethod?.Invoke(null, null);

                Core.DebugLogger.Log($"[HVAC-RIBBON] Sync toggled to: {newState}");

                // Показываем сообщение о новом состоянии
                string status = newState ? "включена" : "выключена";
                TaskDialog.Show("HVAC", $"Синхронизация чертежа с моделью {status}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[HVAC-RIBBON] ERROR in sync toggle: {ex.Message}");
                TaskDialog.Show("HVAC", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
