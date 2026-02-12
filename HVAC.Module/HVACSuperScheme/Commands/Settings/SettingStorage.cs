using CommunityToolkit.Mvvm.ComponentModel;
using HVACSuperScheme.Utils;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows;

namespace HVACSuperScheme.Commands.Settings
{
    public partial class SettingStorage : ObservableValidator
    {
        private static string _defaultSettingPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Autodesk",
            "Revit",
            "Addins",
            "HVACSuperSchemeSettings.cfg");

        [ObservableProperty]
        private bool isUpdaterSync = true;
        public SettingStorage() { }
        public static SettingStorage Instance { get; set; }
        public static void ReadSettings()
        {
            try
            {
                if (File.Exists(_defaultSettingPath))
                {
                    string jsonString = File.ReadAllText(_defaultSettingPath);
                    Instance = JsonConvert.DeserializeObject<SettingStorage>(jsonString);
                }
                else
                {
                    LoggingUtils.LoggingWithMessage(Warnings.FileNotFoundSetDefaultSettings(), "-");
                    Instance = new SettingStorage();
                }
            }
            catch (Exception ex)
            {
                LoggingUtils.LoggingWithMessage(Warnings.FileSettingsChangedSetNewSettings(), "-");
                Instance = new SettingStorage();
            }
        }
        public static void SaveSettings()
        {
            var json = JsonConvert.SerializeObject(SettingStorage.Instance, Formatting.Indented);
            File.WriteAllText(_defaultSettingPath, json);
        }
    }
}
