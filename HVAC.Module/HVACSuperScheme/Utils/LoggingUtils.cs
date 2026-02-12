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
        private static string _logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Autodesk",
            "Revit",
            "Addins",
            "HVACSuperSchemeLog.txt");
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
            MessageBox.Show(message);
            File.AppendAllLines(_logPath, [BuildLogString(docPath, filePath, lineNumber, memberName, message)]);
        }
        public static void Logging(string message,
            string docPath,
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0,
            [CallerMemberName] string memberName = "")
        {
            File.AppendAllLines(_logPath, [BuildLogString(docPath, filePath, lineNumber, memberName, message)]);
        }
    }
}
