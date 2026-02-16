using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using System;
using System.Linq;

namespace HVACSuperScheme
{
    /// <summary>
    /// Stub App class for HVAC.Module (non-standalone mode)
    /// Real App functionality handled by reflection from PluginsManager
    /// </summary>
    public class App : IExternalApplication
    {
        public static UIControlledApplication _uiapp;
        public static UIApplication _uiApplication; // CRITICAL: Required for IdlingHandler
        public static bool _idlingHandlerIsActive;
        public static bool _triggerForSpaceParameterChangedCreated;
        public static bool _triggerForAnnotationParameterChangedCreated;
        public static bool _triggerForDuctTerminalParameterChangedCreated;

        public static Updater _updater;

        public Result OnStartup(UIControlledApplication application)
        {
            // This method is never called in plugin mode
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // This method is never called in plugin mode
            return Result.Succeeded;
        }

        public static void CreateDeletionTriggers()
        {
            if (_updater == null || FilterUtils._spaceFilter == null) return;
            
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeElementDeletion());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeElementDeletion());
            PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Deletion triggers added");
        }

        public static void CreateAdditionTriggers()
        {
            if (_updater == null || FilterUtils._spaceFilter == null) return;
            
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._spaceFilter, Element.GetChangeTypeElementAddition());
            UpdaterRegistry.AddTrigger(_updater.GetUpdaterId(), FilterUtils._ductTerminalFilter, Element.GetChangeTypeElementAddition());
            PluginsManager.Core.DebugLogger.Log("[HVAC-INIT] Addition triggers added");
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
                    
                    // Ensure filter is initialized
                    if (FilterUtils._ductTerminalFilter == null)
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] DuctTerminal filter is NULL, reinitializing filters...");
                        FilterUtils.InitElementFilters();
                    }
                    
                    Element ductTerminal = CollectorUtils.GetDuctTerminals(doc).FirstOrDefault();
                    if (ductTerminal != null)
                    {
                        try
                        {
                            // Check if required parameters exist
                            if (ductTerminal.LookupParameter(Constants.PN_ADSK_UPDATER) == null ||
                                ductTerminal.LookupParameter(Constants.PN_AIR_FLOW) == null)
                            {
                                PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] DuctTerminal missing required parameters (schema not built yet), skipping trigger creation");
                                return; // Will try again on next Idling
                            }
                            
                            Updater.AddChangeParameterForDuctTerminalsTriggers(ductTerminal);
                            _triggerForDuctTerminalParameterChangedCreated = true;
                            PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] DuctTerminal triggers created");
                        }
                        catch (Exception ex)
                        {
                            PluginsManager.Core.DebugLogger.Log($"[HVAC-IDLING] ERROR creating DuctTerminal triggers: {ex.Message}");
                            return; // Will try again on next Idling
                        }
                    }
                    else
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] No DuctTerminals found in document");
                    }
                }

                if (!_triggerForSpaceParameterChangedCreated)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Creating Space parameter triggers...");
                    
                    // Ensure filter is initialized
                    if (FilterUtils._spaceFilter == null)
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Space filter is NULL, reinitializing filters...");
                        FilterUtils.InitElementFilters();
                    }
                    
                    Space space = CollectorUtils.GetSpaces(doc).FirstOrDefault();
                    if (space != null)
                    {
                        try
                        {
                            // Check if required parameter exists
                            if (space.LookupParameter(Constants.PN_ADSK_UPDATER) == null)
                            {
                                PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Space missing required parameter (schema not built yet), skipping trigger creation");
                                return; // Will try again on next Idling
                            }
                            
                            Updater.AddChangeParameterValueForSpacesTriggers(space);
                            _triggerForSpaceParameterChangedCreated = true;
                            PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Space triggers created");
                        }
                        catch (Exception ex)
                        {
                            PluginsManager.Core.DebugLogger.Log($"[HVAC-IDLING] ERROR creating Space triggers: {ex.Message}");
                            return; // Will try again on next Idling
                        }
                    }
                    else
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] No Spaces found in document");
                    }
                }

                if (!_triggerForAnnotationParameterChangedCreated)
                {
                    PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Creating Annotation parameter triggers...");
                    
                    // Ensure filter is initialized
                    if (FilterUtils._annotationFilter == null)
                    {
                        PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Annotation filter is NULL, reinitializing filters...");
                        FilterUtils.InitElementFilters();
                    }
                    
                    Element annotation = CollectorUtils.GetAnnotationInstances(doc).FirstOrDefault();
                    if (annotation != null)
                    {
                        try
                        {
                            // Check if required parameter exists
                            if (annotation.LookupParameter(Constants.PN_ADSK_UPDATER) == null)
                            {
                                PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Annotation missing required parameter (schema not built yet), skipping trigger creation");
                                return; // Will try again on next Idling
                            }
                            
                            Updater.AddChangeParameterValueForAnnotationsTriggers(annotation);
                            _triggerForAnnotationParameterChangedCreated = true;
                            PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Annotation triggers created");
                        }
                        catch (Exception ex)
                        {
                            PluginsManager.Core.DebugLogger.Log($"[HVAC-IDLING] ERROR creating Annotation triggers: {ex.Message}");
                            return; // Will try again on next Idling
                        }
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
                    RemoveIdlingHandler();
                }
            }
            catch (Exception ex)
            {
                PluginsManager.Core.DebugLogger.Log($"[HVAC-IDLING] ERROR: {ex.Message}");
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
            HVACSuperScheme.Commands.Settings.SettingStorage.Instance.IsUpdaterSync = false;
            HVACSuperScheme.Commands.Settings.SettingStorage.SaveSettings();
            PluginsManager.Core.DebugLogger.Log("[HVAC-IDLING] Sync disabled");
        }
    }
}
