using System;
using System.Collections.Generic;

namespace dwg2rvt.Core
{
    public class AnalysisResult
    {
        public bool Success { get; set; }
        public int BlockCount { get; set; }
        public string LogFilePath { get; set; }  // Kept for backward compatibility, but optional now
        public string ErrorMessage { get; set; }
        public string Summary { get; set; }
        
        // New: Store complete block data in memory
        public List<BlockData> BlockData { get; set; } = new List<BlockData>();
        public Dictionary<string, List<BlockData>> BlocksByType { get; set; } = new Dictionary<string, List<BlockData>>();
        public DateTime AnalysisTimestamp { get; set; } = DateTime.Now;
        public string ImportInstanceName { get; set; }
        public int ImportInstanceId { get; set; }
    }
}
