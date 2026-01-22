using System;

namespace dwg2rvt.Core
{
    public class AnalysisResult
    {
        public bool Success { get; set; }
        public int BlockCount { get; set; }
        public string LogFilePath { get; set; }
        public string ErrorMessage { get; set; }
        public string Summary { get; set; }
    }
}
