using System.IO;
using Newtonsoft.Json;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Shared settings for Annotatix module.
    /// Persisted to a JSON file so both Plugin Hub panel and ribbon commands
    /// use the same values. Updated when user changes values in the panel.
    /// </summary>
    public static class AnnotatixSettings
    {
        private const string SETTINGS_FILE = "AnnotatixSettings.json";

        /// <summary>Grid step in mm (default 3.0). Controls raster export cell size.</summary>
        public static double GridStepMm { get; set; } = 3.0;

        /// <summary>Occupancy threshold (0.0-1.0). Default 0.10 = 10%.</summary>
        public static double OccupancyThreshold { get; set; } = 0.10;

        /// <summary>
        /// Edge margin in mm (default 0.0). Expands the view boundary outward
        /// beyond the model element extents before grid alignment.
        /// Positive values increase the crop box on all four sides.
        /// </summary>
        public static double EdgeMarginMm { get; set; } = 0.0;

        static AnnotatixSettings()
        {
            Load();
        }

        /// <summary>
        /// Path to the settings file stored alongside the module DLL.
        /// </summary>
        private static string GetSettingsPath()
        {
            string assemblyDir = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            return Path.Combine(assemblyDir, SETTINGS_FILE);
        }

        /// <summary>
        /// Load settings from JSON file. Falls back to defaults silently.
        /// </summary>
        public static void Load()
        {
            try
            {
                string path = GetSettingsPath();
                if (File.Exists(path))
                {
                    string json = File.ReadAllText(path);
                    var data = JsonConvert.DeserializeObject<SettingsData>(json);
                    if (data != null)
                    {
                        if (data.GridStepMm > 0) GridStepMm = data.GridStepMm;
                        if (data.OccupancyThreshold >= 0) OccupancyThreshold = data.OccupancyThreshold;
                        if (data.EdgeMarginMm >= 0) EdgeMarginMm = data.EdgeMarginMm;
                    }
                }
            }
            catch { /* use defaults */ }
        }

        /// <summary>
        /// Save current settings to JSON file.
        /// </summary>
        public static void Save()
        {
            try
            {
                string path = GetSettingsPath();
                var data = new SettingsData
                {
                    GridStepMm = GridStepMm,
                    OccupancyThreshold = OccupancyThreshold,
                    EdgeMarginMm = EdgeMarginMm
                };
                string json = JsonConvert.SerializeObject(data, Formatting.Indented);
                File.WriteAllText(path, json);
            }
            catch { }
        }

        private class SettingsData
        {
            public double GridStepMm { get; set; } = 3.0;
            public double OccupancyThreshold { get; set; } = 0.10;
            public double EdgeMarginMm { get; set; } = 0.0;
        }
    }
}
