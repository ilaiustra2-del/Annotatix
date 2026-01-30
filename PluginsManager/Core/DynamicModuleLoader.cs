using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PluginsManager.Core
{
    /// <summary>
    /// Dynamic loader for compiled module DLLs
    /// </summary>
    public class DynamicModuleLoader
    {
        private static Dictionary<string, Assembly> _loadedModules = new Dictionary<string, Assembly>();
        private static Dictionary<string, IModule> _moduleInstances = new Dictionary<string, IModule>();
        private static bool _assemblyResolverInitialized = false;
        private static string _mainFolderPath = null;

        /// <summary>
        /// Initialize AssemblyResolve handler for module dependencies
        /// </summary>
        private static void InitializeAssemblyResolver()
        {
            if (_assemblyResolverInitialized)
                return;

            // Get main folder path (where PluginsManager.dll is located)
            var currentAssembly = Assembly.GetExecutingAssembly();
            _mainFolderPath = Path.GetDirectoryName(currentAssembly.Location);
            
            DebugLogger.Log($"[MODULE-LOADER] Initializing AssemblyResolver for main folder: {_mainFolderPath}");

            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                // Parse the assembly name
                var assemblyName = new AssemblyName(args.Name);
                var fileName = assemblyName.Name + ".dll";
                
                // Look for assembly in main folder
                var mainPath = Path.Combine(_mainFolderPath, fileName);
                
                if (File.Exists(mainPath))
                {
                    DebugLogger.Log($"[ASSEMBLY-RESOLVER] Resolving {assemblyName.Name} from main folder");
                    return Assembly.LoadFrom(mainPath);
                }
                
                return null;
            };

            _assemblyResolverInitialized = true;
            DebugLogger.Log("[MODULE-LOADER] AssemblyResolver initialized");
        }

        /// <summary>
        /// Load a compiled module DLL
        /// </summary>
        /// <param name="moduleTag">Module tag (e.g., "dwg2rvt", "hvac")</param>
        /// <param name="moduleDllPath">Path to the DLL file</param>
        /// <returns>True if module loaded successfully</returns>
        public static bool LoadModule(string moduleTag, string moduleDllPath)
        {
            try
            {
                // Initialize AssemblyResolver on first module load
                InitializeAssemblyResolver();
                
                DebugLogger.Log($"[MODULE-LOADER] Loading module: {moduleTag} from {moduleDllPath}");

                // Check if module already loaded
                if (_loadedModules.ContainsKey(moduleTag))
                {
                    DebugLogger.Log($"[MODULE-LOADER] Module {moduleTag} already loaded");
                    return true;
                }

                // Check if DLL exists
                if (!File.Exists(moduleDllPath))
                {
                    DebugLogger.Log($"[MODULE-LOADER] ERROR: Module DLL not found: {moduleDllPath}");
                    return false;
                }

                // Load the assembly
                Assembly moduleAssembly = Assembly.LoadFrom(moduleDllPath);
                DebugLogger.Log($"[MODULE-LOADER] Assembly loaded: {moduleAssembly.GetName().Name} v{moduleAssembly.GetName().Version}");

                // Find IModule implementation
                var moduleTypes = moduleAssembly.GetTypes()
                    .Where(t => typeof(IModule).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract)
                    .ToList();

                if (moduleTypes.Count == 0)
                {
                    DebugLogger.Log($"[MODULE-LOADER] ERROR: No IModule implementation found in {moduleTag}");
                    return false;
                }

                // Create module instance
                var moduleType = moduleTypes.First();
                var moduleInstance = (IModule)Activator.CreateInstance(moduleType);
                
                DebugLogger.Log($"[MODULE-LOADER] Module instance created: {moduleInstance.ModuleName} v{moduleInstance.ModuleVersion}");
                
                // Initialize module
                moduleInstance.Initialize();
                DebugLogger.Log($"[MODULE-LOADER] Module initialized: {moduleTag}");

                // Store the loaded module
                _loadedModules[moduleTag] = moduleAssembly;
                _moduleInstances[moduleTag] = moduleInstance;
                
                DebugLogger.Log($"[MODULE-LOADER] ✓ Module {moduleTag} loaded successfully");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MODULE-LOADER] EXCEPTION loading module {moduleTag}: {ex.Message}");
                DebugLogger.Log($"[MODULE-LOADER] Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Check if module is loaded
        /// </summary>
        public static bool IsModuleLoaded(string moduleTag)
        {
            return _loadedModules.ContainsKey(moduleTag);
        }

        /// <summary>
        /// Get module instance
        /// </summary>
        public static IModule GetModuleInstance(string moduleTag)
        {
            return _moduleInstances.ContainsKey(moduleTag) ? _moduleInstances[moduleTag] : null;
        }

        /// <summary>
        /// Get loaded module assembly
        /// </summary>
        public static Assembly GetModuleAssembly(string moduleTag)
        {
            return _loadedModules.ContainsKey(moduleTag) ? _loadedModules[moduleTag] : null;
        }

        /// <summary>
        /// Unload module (clear from cache)
        /// </summary>
        public static void UnloadModule(string moduleTag)
        {
            if (_loadedModules.ContainsKey(moduleTag))
            {
                _loadedModules.Remove(moduleTag);
                DebugLogger.Log($"[MODULE-LOADER] Module {moduleTag} unloaded");
            }

            if (_moduleInstances.ContainsKey(moduleTag))
            {
                _moduleInstances.Remove(moduleTag);
            }
        }

        /// <summary>
        /// Get list of all loaded modules
        /// </summary>
        public static List<string> GetLoadedModules()
        {
            return _loadedModules.Keys.ToList();
        }

        /// <summary>
        /// Get loaded modules info for debugging
        /// </summary>
        public static string GetLoadedModulesInfo()
        {
            if (_moduleInstances.Count == 0)
                return "No modules loaded";

            var info = "Loaded modules:\n";
            foreach (var kvp in _moduleInstances)
            {
                var module = kvp.Value;
                var assembly = _loadedModules[kvp.Key];
                info += $"  - {module.ModuleName} ({kvp.Key}) v{module.ModuleVersion}\n";
                info += $"    Assembly: {assembly.Location}\n";
            }
            return info;
        }
    }
}
