using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using HVACSuperScheme.Commands.Settings;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using System;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace HVACSuperScheme
{
    public class App : IExternalApplication
    { 
        public static UIControlledApplication _uiapp;
        public static UIApplication _uiApplication; // For plugin mode (non-standalone)
        public static bool _idlingHandlerIsActive;
        public static bool _triggerForSpaceParameterChangedCreated;
        public static bool _triggerForAnnotationParameterChangedCreated;
        public static bool _triggerForDuctTerminalParameterChangedCreated;

        public static Updater _updater;
        public Result OnStartup(UIControlledApplication application)
        {
            _uiapp = application;
            _updater = new Updater(application.ActiveAddInId);
            UpdaterRegistry.RegisterUpdater(_updater);

            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            application.ControlledApplication.DocumentClosing += OnDocumentClosing;

            RibbonPanel ribbonPanel;
            PushButtonData pushButton;
            PushButton ribbonItem;

            #region HVAC

            application.CreateRibbonTab("HVAC");
            ribbonPanel = application.CreateRibbonPanel("HVAC", "HVAC");
            pushButton = new PushButtonData(
                "PIDScheme", "Построение схемы",
                typeof(Commands.CreateSchema.Command).Assembly.Location,
                typeof(Commands.CreateSchema.Command).FullName);
            ribbonItem = ribbonPanel.AddItem(pushButton) as PushButton;
            ribbonItem.ToolTip = "Нажимаешь на кнопку и происходит магия";

            pushButton = new PushButtonData(
                "CompletionScheme", "Достроение схемы",
                typeof(Commands.CompletingSchema.Command).Assembly.Location,
                typeof(Commands.CompletingSchema.Command).FullName);
            ribbonItem = ribbonPanel.AddItem(pushButton) as PushButton;
            ribbonItem.ToolTip = "";

            pushButton = new PushButtonData(
                "Settings", "Настройки",
                typeof(Commands.Settings.Command).Assembly.Location,
                typeof(Commands.Settings.Command).FullName);
            ribbonItem = ribbonPanel.AddItem(pushButton) as PushButton;
            ribbonItem.ToolTip = "";

            #endregion

            FilterUtils.InitElementFilters();

            SettingStorage.ReadSettings();

            CreateDeletionTriggers();
            CreateAdditionTriggers();

            if (SettingStorage.Instance.IsUpdaterSync)
            {
                LoggingUtils.Logging(Warnings.SubscribeToIdlingComplete(), "-");
                CreateIdlingHandler();
            }

            return Result.Succeeded;
        }
        private void OnDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            Document doc = e.Document;
            CashUtils.RemoveCashSpacesAndAnnotations(doc);
        }
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            Document doc = e.Document;
            CashUtils.InitializeCash(doc);
        }
        public static void CreateDeletionTriggers()
        {
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeElementDeletion());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeElementDeletion());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._annotationFilter, Element.GetChangeTypeElementDeletion());
            LoggingUtils.Logging(Warnings.TriggersOnDeleteElementsAdded(), "-");
        }

        public static void CreateAdditionTriggers()
        {
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeElementAddition());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeElementAddition());
            LoggingUtils.Logging(Warnings.TriggersOnAddElementsAdded(), "-");
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            UnregisterUpdater();
            Updater.RemoveAllTriggers();
            application.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            application.ControlledApplication.DocumentClosing -= OnDocumentClosing;
            return Result.Succeeded;
        }
        private void UnregisterUpdater()
        {
            if (UpdaterRegistry.IsUpdaterRegistered(_updater.GetUpdaterId()))
                UpdaterRegistry.UnregisterUpdater(_updater.GetUpdaterId());
        }

        public static void IdlingHandler(object sender, IdlingEventArgs e)
        {
            Document doc = (sender as UIApplication)?.ActiveUIDocument?.Document;
            try
            {
                if (doc == null)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Document is null, skipping");
                    return;
                }
                
                PluginsManager.Core.DebugLogger.Log($"[HVAC-IDLING] IdlingHandler executing for document: {doc.Title}");

                // Check required parameters
                bool hasGeneral = CheckUtils.DocumentHasRequiredSharedParameters(doc);
                bool hasSpace = CheckUtils.DocumentHasRequiredSharedParametersByCategory(doc, BuiltInCategory.OST_MEPSpaces, Constants.REQUIRED_PARAMETERS_FOR_SPACE);
                bool hasDuct = CheckUtils.DocumentHasRequiredSharedParametersByCategory(doc, BuiltInCategory.OST_DuctTerminal, Constants.REQUIRED_PARAMETERS_FOR_DUCT_TERMINAL);
                
                PluginsManager.Core.DebugLogger.Log($"[HVAC-IDLING] Parameter checks: General={hasGeneral}, Space={hasSpace}, Duct={hasDuct}");
                
                if (!hasGeneral || !hasSpace || !hasDuct)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Required parameters missing, removing triggers and disabling sync");
                    Updater.RemoveAllTriggers();
                    RemoveIdlingHandler();
                    SyncOff();
                    return;
                }

                if (!_triggerForDuctTerminalParameterChangedCreated)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Creating DuctTerminal parameter triggers...");
                    Element ductTerminal = CollectorUtils.GetDuctTerminals(doc).FirstOrDefault();
                    if (ductTerminal != null)
                    {
                        Updater.AddChangeParameterForDuctTerminalsTriggers(ductTerminal);
                        _triggerForDuctTerminalParameterChangedCreated = true;
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] DuctTerminal triggers created");
                        LoggingUtils.Logging(Warnings.TriggersOnChangeDuctTermilalParametersCreated(), doc.PathName);
                    }
                    else
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] No DuctTerminals found in document");
                    }
                }

                if (!_triggerForSpaceParameterChangedCreated)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Creating Space parameter triggers...");
                    Space space = CollectorUtils.GetSpaces(doc).FirstOrDefault();
                    if (space != null)
                    {
                        Updater.AddChangeParameterValueForSpacesTriggers(space);
                        _triggerForSpaceParameterChangedCreated = true;
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Space triggers created");
                        LoggingUtils.Logging(Warnings.TriggersOnChangeSpaceParametersCreated(), doc.PathName);
                    }
                    else
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] No Spaces found in document");
                    }
                }

                if (!_triggerForAnnotationParameterChangedCreated)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Creating Annotation parameter triggers...");
                    Element annotation = CollectorUtils.GetAnnotationInstances(doc).FirstOrDefault();
                    if (annotation != null)
                    {
                        Updater.AddChangeParameterValueForAnnotationsTriggers(annotation);
                        _triggerForAnnotationParameterChangedCreated = true;
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Annotation triggers created");
                        LoggingUtils.Logging(Warnings.TriggersOnChangeAnnotationParametersCreated(), doc.PathName);
                    }
                    else
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] No Annotations found in document");
                    }
                }

                if (_triggerForSpaceParameterChangedCreated
                    && _triggerForAnnotationParameterChangedCreated
                    && _triggerForDuctTerminalParameterChangedCreated)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] All triggers created successfully, removing IdlingHandler");
                    LoggingUtils.Logging(Warnings.TriggersOnChangeSuccessfulCreated(), doc.PathName);
                    RemoveIdlingHandler();
                }
            }
            catch (Exception ex)
            {
                LoggingUtils.LoggingWithMessage(ExceptionUtils.IdlingHandlerError(ex), doc?.PathName);
            }
        }
        public static void CreateIdlingHandler()
        {
            // Use _uiApplication for Idling event (only UIApplication has Idling event)
            if (_uiApplication != null)
            {
                _uiApplication.Idling += IdlingHandler;
                _idlingHandlerIsActive = true;
                PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] IdlingHandler subscribed to UIApplication.Idling");
            }
            else
            {
                PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] ERROR: Cannot create IdlingHandler - UIApplication is null");
            }
        }
        public static void RemoveIdlingHandler()
        {
            // Remove from UIApplication only (UIControlledApplication doesn't have Idling event)
            if (_uiApplication != null)
            {
                _uiApplication.Idling -= IdlingHandler;
                PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] IdlingHandler unsubscribed from UIApplication.Idling");
            }
            _idlingHandlerIsActive = false;
        }

        private static void SyncOff()
        {
            SettingStorage.Instance.IsUpdaterSync = false;
            SettingStorage.SaveSettings();
        }
    }
}
