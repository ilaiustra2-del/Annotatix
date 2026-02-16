using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace HVACSuperScheme.Utils
{
    public class LoggingUtils
    { 
        // Use same log file as DebugLogger for unified logging
        private static string _logPath;
        private static readonly object _lockObject = new object();
            
        static LoggingUtils()
        {
            // Use same path as DebugLogger: annotatix_dependencies/logs
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logFolder = Path.Combine(appDataPath, "Autodesk", "Revit", "Addins", "2024", "annotatix_dependencies", "logs");
                
                // Create logs folder if it doesn't exist
                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }
                
                _logPath = Path.Combine(logFolder, "annotatix_debug.log");
            }
            catch
            {
                // Fallback to temp folder
                _logPath = Path.Combine(Path.GetTempPath(), "annotatix_debug.log");
            }
        }
        private static string BuildLogString(string docPath, string filePath, int lineNumber, string memberName, string message)
        {
            return $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
                   $"{docPath ?? "-"} " +
                   $"{Path.GetFileName(filePath)}:{lineNumber} " +
                   $"{memberName} " +
                   $"{message} ";
        }
        public static void LoggingWithMessage(string message,
            string docPath,
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0,
            [CallerMemberName] string memberName = "")
        {
            try
            {
                MessageBox.Show(message);
                lock (_lockObject)
                {
                    File.AppendAllLines(_logPath, [BuildLogString(docPath, filePath, lineNumber, memberName, message)]);
                }
            }
            catch
            {
                // Ignore logging errors to prevent breaking Updater
            }
        }
        public static void Logging(string message,
            string docPath,
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0,
            [CallerMemberName] string memberName = "")
        {
            try
            {
                lock (_lockObject)
                {
                    File.AppendAllLines(_logPath, [BuildLogString(docPath, filePath, lineNumber, memberName, message)]);
                }
            }
            catch
            {
                // Ignore logging errors to prevent breaking Updater
            }
        }
        
        /// <summary>
        /// Get log file path
        /// </summary>
        public static string GetLogFilePath()
        {
            return _logPath;
        }
        
        /// <summary>
        /// Open log file in default text editor
        /// </summary>
        public static void OpenLogFile()
        {
            try
            {
                if (File.Exists(_logPath))
                {
                    System.Diagnostics.Process.Start("notepad.exe", _logPath);
                }
            }
            catch
            {
                // Ignore open errors
            }
        }
    }
}
