using System;
using System.IO;
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
                
                // Если модуль не загружен - пытаемся загрузить автоматически
                if (module == null)
                {
                    Core.DebugLogger.Log("[HVAC-RIBBON] Module not loaded, attempting auto-load...");
                    
                    // Путь к модулю HVAC (annotatix_dependencies/main/ -> annotatix_dependencies/)
                    string assemblyPath = typeof(HvacRibbonCommandBase).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);  // Переход на уровень выше
                    string hvacModulePath = Path.Combine(modulesPath, "hvac", "HVAC.Module.dll");
                    
                    Core.DebugLogger.Log($"[HVAC-RIBBON] Assembly dir: {assemblyDir}");
                    Core.DebugLogger.Log($"[HVAC-RIBBON] Modules path: {modulesPath}");
                    Core.DebugLogger.Log($"[HVAC-RIBBON] HVAC module path: {hvacModulePath}");
                    
                    if (!File.Exists(hvacModulePath))
                    {
                        Core.DebugLogger.Log($"[HVAC-RIBBON] Module not found at: {hvacModulePath}");
                        TaskDialog.Show("HVAC", $"Модуль HVAC не найден по пути:\n{hvacModulePath}");
                        return Result.Failed;
                    }
                    
                    Core.DebugLogger.Log($"[HVAC-RIBBON] Loading module from: {hvacModulePath}");
                    if (!Core.DynamicModuleLoader.LoadModule("hvac", hvacModulePath))
                    {
                        Core.DebugLogger.Log("[HVAC-RIBBON] Failed to load module");
                        TaskDialog.Show("HVAC", "Не удалось загрузить модуль HVAC. Откройте Plugins Hub.");
                        return Result.Failed;
                    }
                    
                    module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                    Core.DebugLogger.Log("[HVAC-RIBBON] Module auto-loaded successfully");
                }

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
                    Core.DebugLogger.Log("[HVAC-RIBBON] ERROR: Could not create handler instance");
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
        /// <summary>
        /// Ссылка на кнопку синхронизации для динамического изменения текста
        /// </summary>
        public static PushButton SyncButton { get; set; }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Получаем модуль HVAC
                var module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                
                // Если модуль не загружен - пытаемся загрузить автоматически
                if (module == null)
                {
                    Core.DebugLogger.Log("[HVAC-RIBBON-SYNC] Module not loaded, attempting auto-load...");
                    
                    string assemblyPath = typeof(HvacSyncToggleRibbonCommand).Assembly.Location;
                    string assemblyDir = Path.GetDirectoryName(assemblyPath);
                    string modulesPath = Path.GetDirectoryName(assemblyDir);
                    string hvacModulePath = Path.Combine(modulesPath, "hvac", "HVAC.Module.dll");
                    
                    if (File.Exists(hvacModulePath))
                    {
                        Core.DynamicModuleLoader.LoadModule("hvac", hvacModulePath);
                        module = Core.DynamicModuleLoader.GetModuleInstance("hvac");
                    }
                }
                
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
                    Core.DebugLogger.Log("[HVAC-RIBBON-SYNC] SettingStorage type not found");
                    TaskDialog.Show("HVAC", "Не найден тип SettingStorage");
                    return Result.Failed;
                }

                // Получаем Instance - если null, инициализируем через ReadSettings
                var instanceProperty = settingStorageType.GetProperty("Instance");
                var instance = instanceProperty?.GetValue(null);
                
                if (instance == null)
                {
                    Core.DebugLogger.Log("[HVAC-RIBBON-SYNC] Instance is null, calling ReadSettings...");
                    var readSettingsMethod = settingStorageType.GetMethod("ReadSettings");
                    readSettingsMethod?.Invoke(null, null);
                    instance = instanceProperty?.GetValue(null);
                }
                
                if (instance == null)
                {
                    Core.DebugLogger.Log("[HVAC-RIBBON-SYNC] Failed to initialize Instance");
                    TaskDialog.Show("HVAC", "Не удалось инициализировать настройки");
                    return Result.Failed;
                }

                // Получаем текущее состояние IsUpdaterSync
                var isUpdaterSyncProperty = instance.GetType().GetProperty("IsUpdaterSync");
                bool currentValue = (bool)(isUpdaterSyncProperty?.GetValue(instance) ?? false);
                Core.DebugLogger.Log($"[HVAC-RIBBON-SYNC] Current IsUpdaterSync: {currentValue}");
                
                // Переключаем состояние
                bool newState = !currentValue;
                isUpdaterSyncProperty?.SetValue(instance, newState);
                Core.DebugLogger.Log($"[HVAC-RIBBON-SYNC] Set IsUpdaterSync to: {newState}");

                // Сохраняем настройки
                var saveSettingsMethod = settingStorageType.GetMethod("SaveSettings");
                saveSettingsMethod?.Invoke(null, null);
                Core.DebugLogger.Log("[HVAC-RIBBON-SYNC] Settings saved");

                // Обновляем HVACSyncState.IsSyncEnabled (используется в handlers)
                var hvacSyncStateType = moduleAssembly.GetType("HVAC.Module.UI.HVACSyncState");
                if (hvacSyncStateType != null)
                {
                    var isSyncEnabledProperty = hvacSyncStateType.GetProperty("IsSyncEnabled");
                    isSyncEnabledProperty?.SetValue(null, newState);
                    Core.DebugLogger.Log($"[HVAC-RIBBON-SYNC] HVACSyncState.IsSyncEnabled set to: {newState}");
                }

                // Обновляем текст кнопки
                if (SyncButton != null)
                {
                    SyncButton.ItemText = newState 
                        ? "Синхронизация\nвключена" 
                        : "Синхронизация\nвыключена";
                    Core.DebugLogger.Log($"[HVAC-RIBBON-SYNC] Button text updated: {SyncButton.ItemText}");
                }

                // Запускаем/останавливаем IdlingHandler через ToggleIdlingHandler
                var toggleHandlerType = moduleAssembly.GetType("HVAC.Module.Commands.ToggleIdlingHandler");
                if (toggleHandlerType != null)
                {
                    var toggleHandler = Activator.CreateInstance(toggleHandlerType) as IExternalEventHandler;
                    if (toggleHandler != null)
                    {
                        // Устанавливаем состояние (start/stop)
                        var setStartMethod = toggleHandlerType.GetMethod("SetStart");
                        setStartMethod?.Invoke(toggleHandler, new object[] { newState });
                        
                        // Запускаем через ExternalEvent
                        var toggleEvent = ExternalEvent.Create(toggleHandler);
                        toggleEvent.Raise();
                        Core.DebugLogger.Log($"[HVAC-RIBBON-SYNC] ToggleIdlingHandler raised with start={newState}");
                    }
                }
                else
                {
                    Core.DebugLogger.Log("[HVAC-RIBBON-SYNC] ToggleIdlingHandler type not found");
                }

                // Показываем сообщение о новом состоянии
                string status = newState ? "включена" : "выключена";
                TaskDialog.Show("HVAC", $"Синхронизация чертежа с моделью {status}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Core.DebugLogger.Log($"[HVAC-RIBBON-SYNC] ERROR: {ex.Message}");
                TaskDialog.Show("HVAC", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
