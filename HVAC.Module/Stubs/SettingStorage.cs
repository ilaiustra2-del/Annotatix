using System;
using System.IO;
using Newtonsoft.Json;

namespace HVACSuperScheme.Commands.Settings
{
    /// <summary>
    /// Simplified Settings Storage without MVVM Toolkit dependencies
    /// </summary>
    public class SettingStorage
    {
        private static SettingStorage _instance;
        
        public static SettingStorage Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SettingStorage();
                }
                return _instance;
            }
        }
        
        public bool IsUpdaterSync { get; set; }
        
        private SettingStorage()
        {
            IsUpdaterSync = false;
        }
        
        public static void ReadSettings()
        {
            try
            {
                string settingsPath = GetSettingsPath();
                if (File.Exists(settingsPath))
                {
                    string json = File.ReadAllText(settingsPath);
                    var settings = JsonConvert.DeserializeObject<SettingStorage>(json);
                    if (settings != null)
                    {
                        Instance.IsUpdaterSync = settings.IsUpdaterSync;
                    }
                }
            }
            catch
            {
                // If reading fails, use default settings
            }
        }
        
        public static void SaveSettings()
        {
            try
            {
                string settingsPath = GetSettingsPath();
                string json = JsonConvert.SerializeObject(Instance, Formatting.Indented);
                
                string directory = Path.GetDirectoryName(settingsPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                
                File.WriteAllText(settingsPath, json);
            }
            catch
            {
                // Ignore save errors
            }
        }
        
        private static string GetSettingsPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "Annotatix", "HVAC", "settings.json");
        }
    }
}
