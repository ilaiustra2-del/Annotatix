using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using PluginsManager.Core;

namespace PluginsManager.Commands
{
    /// <summary>
    /// Universal External Event Handler for numbering risers
    /// Works with AutoNumbering.Module through reflection
    /// Created in PluginsManager to avoid loading order issues
    /// </summary>
    public class NumberRisersHandler : IExternalEventHandler
    {
        public object RiserGroups { get; set; }
        public string ParameterName { get; set; }
        public string Prefix { get; set; }
        public string Postfix { get; set; }
        public int SuccessCount { get; private set; }
        public int TotalCount { get; private set; }
        public string ErrorMessage { get; private set; }

        public void Execute(UIApplication app)
        {
            SuccessCount = 0;
            TotalCount = 0;
            ErrorMessage = null;

            try
            {
                var doc = app.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    ErrorMessage = "Документ не открыт";
                    return;
                }

                if (RiserGroups == null)
                {
                    ErrorMessage = "Нет данных для нумерации";
                    return;
                }

                DebugLogger.Log($"[RISER-NUMBERING] Starting numbering with prefix='{Prefix}', postfix='{Postfix}', param='{ParameterName}'");

                using (Transaction trans = new Transaction(doc, "Нумерация стояков"))
                {
                    trans.Start();

                    try
                    {
                        // RiserGroups is List<RiserGroup>, iterate using reflection
                        var groupsList = RiserGroups as System.Collections.IList;
                        if (groupsList == null)
                        {
                            ErrorMessage = "Неверный формат данных групп";
                            return;
                        }

                        foreach (var group in groupsList)
                        {
                            // Get GroupNumber property
                            var groupType = group.GetType();
                            int groupNumber = (int)groupType.GetProperty("GroupNumber").GetValue(group);
                            
                            // Get Risers property (List<RiserInfo>)
                            var risers = groupType.GetProperty("Risers").GetValue(group) as System.Collections.IList;
                            
                            string numberValue = $"{Prefix}{groupNumber}{Postfix}";

                            foreach (var riser in risers)
                            {
                                TotalCount++;
                                
                                // Get Pipe property from RiserInfo
                                var riserType = riser.GetType();
                                var pipe = riserType.GetProperty("Pipe").GetValue(riser) as Pipe;
                                
                                if (pipe != null)
                                {
                                    var param = pipe.LookupParameter(ParameterName);
                                    if (param != null && !param.IsReadOnly)
                                    {
                                        param.Set(numberValue);
                                        SuccessCount++;
                                        DebugLogger.Log($"[RISER-NUMBERING] Set pipe {pipe.Id.IntegerValue} = {numberValue}");
                                    }
                                    else
                                    {
                                        DebugLogger.Log($"[RISER-NUMBERING] Parameter '{ParameterName}' not found or readonly for pipe {pipe.Id.IntegerValue}");
                                    }
                                }
                            }
                        }

                        trans.Commit();
                        DebugLogger.Log($"[RISER-NUMBERING] Numbering complete: {SuccessCount}/{TotalCount} successful");
                        
                        // Show success message from handler (executed in API context)
                        System.Windows.MessageBox.Show(
                            $"Нумерация завершена!\n\nПронумеровано: {SuccessCount} из {TotalCount} стояков",
                            "Успех",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = ex.Message;
                        DebugLogger.Log($"[RISER-NUMBERING] ERROR: {ex.Message}");
                        
                        // Show error message from handler
                        System.Windows.MessageBox.Show(
                            $"Ошибка нумерации:\n{ex.Message}",
                            "Ошибка",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                DebugLogger.Log($"[RISER-NUMBERING] EXCEPTION: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "NumberRisersHandler";
        }
    }
}
