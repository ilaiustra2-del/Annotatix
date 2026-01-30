using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace PluginsManager.Core
{
    /// <summary>
    /// Information about module file for download
    /// </summary>
    public class ModuleFileInfo
    {
        public string ModuleTag { get; set; }
        public string DllDownloadUrl { get; set; }
        public string PdbDownloadUrl { get; set; }
        public string Version { get; set; }
        public long FileSize { get; set; }
    }
    
    /// <summary>
    /// Download and install module files from server
    /// </summary>
    public class ModuleDownloader
    {
        private readonly string _annotatixDependenciesPath;
        private static readonly HttpClient _httpClient = new HttpClient();
        
        public ModuleDownloader()
        {
            // Get path to annotatix_dependencies folder
            _annotatixDependenciesPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk", "Revit", "Addins", "2024", "annotatix_dependencies"
            );
            
            DebugLogger.Log($"[MODULE-DOWNLOADER] Dependencies path: {_annotatixDependenciesPath}");
        }
        
        /// <summary>
        /// Download and install a single module
        /// </summary>
        public async Task<bool> DownloadAndInstallModule(ModuleFileInfo moduleInfo, IProgress<string> progress = null)
        {
            try
            {
                var moduleTag = moduleInfo.ModuleTag.ToLower();
                progress?.Report($"Загрузка модуля {moduleTag}...");
                
                DebugLogger.Log($"[MODULE-DOWNLOADER] Starting download for module: {moduleTag}");
                
                // 1. Create module folder if not exists
                var moduleFolderPath = Path.Combine(_annotatixDependenciesPath, moduleTag);
                if (!Directory.Exists(moduleFolderPath))
                {
                    Directory.CreateDirectory(moduleFolderPath);
                    DebugLogger.Log($"[MODULE-DOWNLOADER] Created folder: {moduleFolderPath}");
                }
                
                // 2. Download DLL file
                var dllFileName = $"{moduleTag}.Module.dll";
                var dllPath = Path.Combine(moduleFolderPath, dllFileName);
                
                if (!await DownloadFile(moduleInfo.DllDownloadUrl, dllPath, progress))
                {
                    DebugLogger.Log($"[MODULE-DOWNLOADER] Failed to download DLL for {moduleTag}");
                    return false;
                }
                
                DebugLogger.Log($"[MODULE-DOWNLOADER] Downloaded DLL: {dllPath}");
                
                // 3. Download PDB file (optional, for debugging)
                if (!string.IsNullOrEmpty(moduleInfo.PdbDownloadUrl))
                {
                    var pdbFileName = $"{moduleTag}.Module.pdb";
                    var pdbPath = Path.Combine(moduleFolderPath, pdbFileName);
                    
                    if (await DownloadFile(moduleInfo.PdbDownloadUrl, pdbPath, progress))
                    {
                        DebugLogger.Log($"[MODULE-DOWNLOADER] Downloaded PDB: {pdbPath}");
                    }
                }
                
                progress?.Report($"Модуль {moduleTag} установлен");
                DebugLogger.Log($"[MODULE-DOWNLOADER] ✓ Module {moduleTag} installed successfully");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MODULE-DOWNLOADER] ERROR installing module {moduleInfo.ModuleTag}: {ex.Message}");
                DebugLogger.Log($"[MODULE-DOWNLOADER] Stack trace: {ex.StackTrace}");
                return false;
            }
        }
        
        /// <summary>
        /// Download all modules for user
        /// </summary>
        public async Task<(int success, int total)> DownloadAllModules(
            List<ModuleFileInfo> moduleFiles, 
            IProgress<string> progress = null)
        {
            if (moduleFiles == null || moduleFiles.Count == 0)
            {
                DebugLogger.Log("[MODULE-DOWNLOADER] No module files to download");
                return (0, 0);
            }
            
            DebugLogger.Log($"[MODULE-DOWNLOADER] Starting download of {moduleFiles.Count} modules");
            
            int successCount = 0;
            int totalCount = moduleFiles.Count;
            
            foreach (var moduleFile in moduleFiles)
            {
                progress?.Report($"Загрузка {successCount + 1}/{totalCount}: {moduleFile.ModuleTag}");
                
                if (await DownloadAndInstallModule(moduleFile, progress))
                {
                    successCount++;
                }
            }
            
            DebugLogger.Log($"[MODULE-DOWNLOADER] Download complete: {successCount}/{totalCount} successful");
            return (successCount, totalCount);
        }
        
        /// <summary>
        /// Download file from URL
        /// </summary>
        private async Task<bool> DownloadFile(string url, string destinationPath, IProgress<string> progress = null)
        {
            try
            {
                if (string.IsNullOrEmpty(url))
                {
                    DebugLogger.Log("[MODULE-DOWNLOADER] Download URL is empty");
                    return false;
                }
                
                DebugLogger.Log($"[MODULE-DOWNLOADER] Downloading from: {url}");
                DebugLogger.Log($"[MODULE-DOWNLOADER] Destination: {destinationPath}");
                
                var response = await _httpClient.GetAsync(url);
                
                if (!response.IsSuccessStatusCode)
                {
                    DebugLogger.Log($"[MODULE-DOWNLOADER] HTTP error: {response.StatusCode} - {response.ReasonPhrase}");
                    
                    // If 404, file doesn't exist on server (module not ready yet)
                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        DebugLogger.Log($"[MODULE-DOWNLOADER] File not found on server (404), skipping");
                    }
                    
                    return false;
                }
                
                var fileBytes = await response.Content.ReadAsByteArrayAsync();
                File.WriteAllBytes(destinationPath, fileBytes);
                
                DebugLogger.Log($"[MODULE-DOWNLOADER] File saved: {fileBytes.Length} bytes");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MODULE-DOWNLOADER] Download error: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Check if module folder exists
        /// </summary>
        public bool IsModuleInstalled(string moduleTag)
        {
            var moduleFolderPath = Path.Combine(_annotatixDependenciesPath, moduleTag.ToLower());
            var dllPath = Path.Combine(moduleFolderPath, $"{moduleTag.ToLower()}.Module.dll");
            
            return File.Exists(dllPath);
        }
        
        /// <summary>
        /// Get list of modules that need to be downloaded
        /// </summary>
        public List<ModuleFileInfo> GetModulesToDownload(List<ModuleFileInfo> allModules)
        {
            var toDownload = new List<ModuleFileInfo>();
            
            foreach (var module in allModules)
            {
                if (!IsModuleInstalled(module.ModuleTag))
                {
                    toDownload.Add(module);
                    DebugLogger.Log($"[MODULE-DOWNLOADER] Module {module.ModuleTag} needs download");
                }
                else
                {
                    DebugLogger.Log($"[MODULE-DOWNLOADER] Module {module.ModuleTag} already installed");
                }
            }
            
            return toDownload;
        }
    }
}
