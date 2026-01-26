using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using dwg2rvt.Core;

namespace dwg2rvt.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AnnotateBlocksCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Get data from cache instead of log files
                var analysisResult = Core.AnalysisDataCache.GetLastAnalysisResult();
                
                if (analysisResult == null || !analysisResult.Success)
                {
                    TaskDialog.Show("Error", "No analysis data found in cache. Please run analysis first.");
                    return Result.Failed;
                }
                
                if (analysisResult.BlockData == null || analysisResult.BlockData.Count == 0)
                {
                    TaskDialog.Show("Error", "No blocks found in analysis results.");
                    return Result.Failed;
                }
                
                // Create text notes for each block
                using (Transaction trans = new Transaction(doc, "Annotate DWG Blocks"))
                {
                    trans.Start();
                    
                    View activeView = doc.ActiveView;
                    
                    // Get default text note type
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    TextNoteType textNoteType = collector
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .FirstOrDefault();
                    
                    if (textNoteType == null)
                    {
                        TaskDialog.Show("Error", "No text note type found in the document.");
                        trans.RollBack();
                        return Result.Failed;
                    }
                    
                    int annotatedCount = 0;
                    
                    foreach (var block in analysisResult.BlockData)
                    {
                        try
                        {
                            // Create XYZ point for text note location (center of block)
                            XYZ location = new XYZ(block.CenterX, block.CenterY, 0);
                            
                            // Create text note
                            TextNote textNote = TextNote.Create(doc, activeView.Id, location, block.Name, textNoteType.Id);
                            
                            annotatedCount++;
                        }
                        catch (Exception ex)
                        {
                            // Skip this block if annotation fails
                            continue;
                        }
                    }
                    
                    trans.Commit();
                    
                    TaskDialog.Show("Success", $"Annotated {annotatedCount} blocks successfully!");
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }
        
        private System.Collections.Generic.List<BlockAnnotationInfo> ParseLogFile(string logFilePath)
        {
            return ParseLogFileStatic(logFilePath);
        }
        
        public static System.Collections.Generic.List<BlockAnnotationInfo> ParseLogFileStatic(string logFilePath)
        {
            var blocks = new System.Collections.Generic.List<BlockAnnotationInfo>();
            
            try
            {
                string[] lines = File.ReadAllLines(logFilePath, System.Text.Encoding.UTF8);
                bool inBlocksSection = false;
                
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    string trimmed = line.Trim();
                    
                    if (trimmed == "Блоки:")
                    {
                        inBlocksSection = true;
                        continue;
                    }

                    if (!inBlocksSection) continue;
                    
                    // New format: "BlockName №Number Координаты центра блока: (X, Y)"
                    if (line.StartsWith("    ") && !line.StartsWith("        ") && line.Contains("№") && line.Contains("Координаты центра блока:"))
                    {
                        try
                        {
                            // Extract block name (before №)
                            int hashIndex = trimmed.IndexOf("№");
                            if (hashIndex > 0)
                            {
                                string blockName = trimmed.Substring(0, hashIndex).Trim();
                                
                                // Extract coordinates from parentheses
                                int coordStart = trimmed.IndexOf("(");
                                int coordEnd = trimmed.IndexOf(")");
                                if (coordStart > 0 && coordEnd > coordStart)
                                {
                                    string coordSection = trimmed.Substring(coordStart + 1, coordEnd - coordStart - 1);
                                    
                                    // Split by comma separator (not decimal separator)
                                    // Format: (X, Y) where comma separates coordinates
                                    string[] coordParts = coordSection.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    if (coordParts.Length >= 2)
                                    {
                                        // Get first two parts (X and Y coordinates)
                                        string xStr = coordParts[0].Trim();
                                        string yStr = coordParts[1].Trim();
                                        
                                        if (double.TryParse(xStr, System.Globalization.NumberStyles.Any, 
                                            System.Globalization.CultureInfo.InvariantCulture, out double x) &&
                                            double.TryParse(yStr, System.Globalization.NumberStyles.Any, 
                                            System.Globalization.CultureInfo.InvariantCulture, out double y))
                                        {
                                            BlockAnnotationInfo block = new BlockAnnotationInfo
                                            {
                                                Name = blockName,
                                                CenterX = x,
                                                CenterY = y
                                            };
                                            blocks.Add(block);
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // Skip malformed lines
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Return empty list if parsing fails
            }
            
            return blocks;
        }
    }
    
    public class BlockAnnotationInfo
    {
        public string Name { get; set; }
        public double CenterX { get; set; }
        public double CenterY { get; set; }
    }
}
