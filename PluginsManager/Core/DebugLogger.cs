using System;
using System.IO;

namespace PluginsManager.Core
{
    /// <summary>
    /// Centralized debug logger for tracking module loading lifecycle
    /// </summary>
    public static class DebugLogger
    {
        private static string _logFilePath;
        private static readonly object _lockObject = new object();

        static DebugLogger()
        {
            // Log to annotatix_dependencies/logs folder
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logFolder = Path.Combine(appDataPath, "Autodesk", "Revit", "Addins", "2024", "annotatix_dependencies", "logs");
            
            // Create logs folder if it doesn't exist
            try
            {
                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }
            }
            catch
            {
                // Fallback to temp if can't create directory
                logFolder = Path.GetTempPath();
            }
            
            _logFilePath = Path.Combine(logFolder, "annotatix_debug.log");
            
            // Clear log on application start
            try
            {
                File.WriteAllText(_logFilePath, $"=== Annotatix Debug Log Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n\n");
            }
            catch { }
        }

        /// <summary>
        /// Log message to both Debug output and file
        /// </summary>
        public static void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var fullMessage = $"[{timestamp}] {message}";
            
            // Log to Debug output (visible in Visual Studio)
            System.Diagnostics.Debug.WriteLine(fullMessage);
            
            // Log to file (visible in any text editor)
            try
            {
                lock (_lockObject)
                {
                    File.AppendAllText(_logFilePath, fullMessage + "\n");
                }
            }
            catch { }
        }

        /// <summary>
        /// Log separator line
        /// </summary>
        public static void LogSeparator(char character = '=', int length = 80)
        {
            var separator = new string(character, length);
            System.Diagnostics.Debug.WriteLine(separator);
            
            try
            {
                lock (_lockObject)
                {
                    File.AppendAllText(_logFilePath, separator + "\n");
                }
            }
            catch { }
        }

        /// <summary>
        /// Get log file path
        /// </summary>
        public static string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// Open log file in default text editor
        /// </summary>
        public static void OpenLogFile()
        {
            try
            {
                if (File.Exists(_logFilePath))
                {
                    System.Diagnostics.Process.Start("notepad.exe", _logFilePath);
                }
            }
            catch { }
        }
    }
}
