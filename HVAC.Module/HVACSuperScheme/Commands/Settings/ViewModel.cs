using Autodesk.Revit.DB;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HVACSuperScheme.Updaters;
using HVACSuperScheme.Utils;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows;

namespace HVACSuperScheme.Commands.Settings
{
    public partial class ViewModel : ObservableValidator
    {
        public SettingStorage Settings => SettingStorage.Instance;
        public View View { get; set; }
        public Document Document {  get; set; }

        private bool _oldValueUpdaterSync;

        public ViewModel(Document doc)
        {
            Document = doc;
            SettingStorage.ReadSettings();
            _oldValueUpdaterSync = Settings.IsUpdaterSync;

            View = new View() { DataContext = this };
        }

        [RelayCommand]
        private void Save()
        {
            SettingStorage.SaveSettings();
            bool updaterSyncNow = Settings.IsUpdaterSync;
            bool changeSyncStatus = _oldValueUpdaterSync != updaterSyncNow; 
            
            if (changeSyncStatus)
            {
                LoggingUtils.Logging(Warnings.TriggersDeleted(), Document.PathName);
                Updater.RemoveAllTriggers();
                App.CreateAdditionTriggers();
                App.CreateDeletionTriggers();

                if (updaterSyncNow)
                {
                    if (!App._idlingHandlerIsActive)
                    {
                        LoggingUtils.Logging(Warnings.SubscribeToIdlingComplete(), Document.PathName);
                        App.CreateIdlingHandler();
                    }
                }
                else
                {
                    if (App._idlingHandlerIsActive)
                    {
                        LoggingUtils.Logging(Warnings.UnsubscribeToIdlingComplete(), Document.PathName);
                        App.RemoveIdlingHandler();
                    }
                }
            }
            View.Close();
        }

        [RelayCommand]
        private void Cancel()
        {
            View.Close();
        }
    }
}
