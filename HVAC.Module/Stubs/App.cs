using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Linq;
using PluginsManager.Core;

namespace HVACSuperScheme
{
    /// <summary>
    /// Stub App class to hold static Updater instance and trigger flags
    /// Replaces the original App.cs IExternalApplication
    /// </summary>
    public static class App
    {
        public static Updater _updater;
        public static bool _triggerForSpaceParameterChangedCreated = false;
        public static bool _triggerForAnnotationParameterChangedCreated = false;
        public static bool _triggerForDuctTerminalParameterChangedCreated = false;
        public static bool _idlingHandlerIsActive = false;
        public static UIApplication _uiapp;

        public static void CreateDeletionTriggers()
        {
            if (_updater == null) return;
            
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeElementDeletion());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeElementDeletion());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._annotationFilter, Element.GetChangeTypeElementDeletion());
            LoggingUtils.Logging(Warnings.TriggersOnDeleteElementsAdded(), "-");
            DebugLogger.Log("[HVAC-APP] Deletion triggers created");
        }

        public static void CreateAdditionTriggers()
        {
            if (_updater == null) return;
            
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeElementAddition());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeElementAddition());
            LoggingUtils.Logging(Warnings.TriggersOnAddElementsAdded(), "-");
            DebugLogger.Log("[HVAC-APP] Addition triggers created");
        }

        public static void CreateIdlingHandler()
        {
            if (_uiapp != null)
            {
                _uiapp.Idling += IdlingHandler;
                _idlingHandlerIsActive = true;
                DebugLogger.Log("[HVAC-APP] Idling handler subscribed");
            }
        }

        public static void RemoveIdlingHandler()
        {
            if (_uiapp != null)
            {
                _uiapp.Idling -= IdlingHandler;
                _idlingHandlerIsActive = false;
                DebugLogger.Log("[HVAC-APP] Idling handler unsubscribed");
            }
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
                    DebugLogger.Log("[HVAC-APP] Missing required parameters, sync disabled");
                    return;
                }

                if (!_triggerForDuctTerminalParameterChangedCreated)
                {
                    Element ductTerminal = CollectorUtils.GetDuctTerminals(doc).FirstOrDefault();
                    if (ductTerminal != null)
                    {
                        Updater.AddChangeParameterForDuctTerminalsTriggers(ductTerminal);
                        _triggerForDuctTerminalParameterChangedCreated = true;
                        LoggingUtils.Logging(Warnings.TriggersOnChangeDuctTermilalParametersCreated(), doc.PathName);
                    }
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
                    DebugLogger.Log("[HVAC-APP] All triggers created successfully, idling handler removed");
                }
            }
            catch (Exception ex)
            {
                LoggingUtils.LoggingWithMessage(ExceptionUtils.IdlingHandlerError(ex), doc?.PathName);
                DebugLogger.Log($"[HVAC-APP] Idling handler error: {ex.Message}");
            }
        }

        private static void SyncOff()
        {
            HVACSuperScheme.Commands.Settings.SettingStorage.Instance.IsUpdaterSync = false;
            HVACSuperScheme.Commands.Settings.SettingStorage.SaveSettings();
            DebugLogger.Log("[HVAC-APP] Sync disabled due to missing parameters");
        }
    }
}
