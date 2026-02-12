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
                    return;

                if (!CheckUtils.DocumentHasRequiredSharedParameters(doc)
                    || !CheckUtils.DocumentHasRequiredSharedParametersByCategory(doc, BuiltInCategory.OST_MEPSpaces, Constants.REQUIRED_PARAMETERS_FOR_SPACE)
                    || !CheckUtils.DocumentHasRequiredSharedParametersByCategory(doc, BuiltInCategory.OST_DuctTerminal, Constants.REQUIRED_PARAMETERS_FOR_DUCT_TERMINAL))
                {
                    Updater.RemoveAllTriggers();
                    RemoveIdlingHandler();
                    SyncOff();
                    return;
                }

                if (!_triggerForDuctTerminalParameterChangedCreated)
                {
                    Element ductTerminal = CollectorUtils.GetDuctTerminals(doc).FirstOrDefault();
                    Updater.AddChangeParameterForDuctTerminalsTriggers(ductTerminal);
                    _triggerForDuctTerminalParameterChangedCreated = true;
                    LoggingUtils.Logging(Warnings.TriggersOnChangeDuctTermilalParametersCreated(), doc.PathName);
                }

                if (!_triggerForSpaceParameterChangedCreated)
                {
                    Space space = CollectorUtils.GetSpaces(doc).FirstOrDefault();
                    if (space != null)
                    {
                        Updater.AddChangeParameterValueForSpacesTriggers(space);
                        _triggerForSpaceParameterChangedCreated = true;
                        LoggingUtils.Logging(Warnings.TriggersOnChangeSpaceParametersCreated(), doc.PathName);
                    }
                }

                if (!_triggerForAnnotationParameterChangedCreated)
                {
                    Element annotation = CollectorUtils.GetAnnotationInstances(doc).FirstOrDefault();
                    if (annotation != null)
                    {
                        Updater.AddChangeParameterValueForAnnotationsTriggers(annotation);
                        _triggerForAnnotationParameterChangedCreated = true;
                        LoggingUtils.Logging(Warnings.TriggersOnChangeAnnotationParametersCreated(), doc.PathName);
                    }
                }

                if (_triggerForSpaceParameterChangedCreated
                    && _triggerForAnnotationParameterChangedCreated
                    && _triggerForDuctTerminalParameterChangedCreated)
                {
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
            _uiapp.Idling += IdlingHandler;
            _idlingHandlerIsActive = true;
        }
        public static void RemoveIdlingHandler()
        {
            _uiapp.Idling -= IdlingHandler;
            _idlingHandlerIsActive = false;
        }

        private static void SyncOff()
        {
            SettingStorage.Instance.IsUpdaterSync = false;
            SettingStorage.SaveSettings();
        }
    }
}
