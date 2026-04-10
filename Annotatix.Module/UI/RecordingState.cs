using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace Annotatix.Module.UI
{
    /// <summary>
    /// Static class to track recording state across the module
    /// </summary>
    public static class RecordingState
    {
        /// <summary>
        /// Whether recording is currently active
        /// </summary>
        public static bool IsRecording { get; set; }

        /// <summary>
        /// Unique identifier for the current recording session
        /// </summary>
        public static string SessionId { get; set; }

        /// <summary>
        /// Path to the start snapshot file
        /// </summary>
        public static string StartSnapshotPath { get; set; }

        /// <summary>
        /// Reference to the record button for dynamic text update
        /// </summary>
        public static PushButton RecordButton { get; set; }

        /// <summary>
        /// Path to the recordings directory
        /// </summary>
        public static string RecordingsDirectory { get; set; }

        /// <summary>
        /// Reset the recording state
        /// </summary>
        public static void Reset()
        {
            IsRecording = false;
            SessionId = null;
            StartSnapshotPath = null;
            DebugLogger.Log("[ANNOTATIX-STATE] Recording state reset");
        }

        /// <summary>
        /// Start a new recording session
        /// </summary>
        public static void StartNewSession()
        {
            SessionId = System.Guid.NewGuid().ToString();
            IsRecording = true;
            DebugLogger.Log($"[ANNOTATIX-STATE] New recording session started: {SessionId}");
        }
    }
}
