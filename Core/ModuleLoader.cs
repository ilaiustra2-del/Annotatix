using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.CSharp;

namespace dwg2rvt.Core
{
    /// <summary>
    /// Dynamic module loader for runtime compilation and loading of .cs module files
    /// </summary>
    public class ModuleLoader
    {
        private static Dictionary<string, Assembly> _loadedModules = new Dictionary<string, Assembly>();
        private static Dictionary<string, object> _moduleInstances = new Dictionary<string, object>();

        /// <summary>
        /// Load and compile a module from .cs file
        /// </summary>
        /// <param name="moduleTag">Module tag (e.g., "dwg2rvt", "hvac")</param>
        /// <param name="moduleFilePath">Path to the .cs file</param>
        /// <returns>True if module loaded successfully</returns>
        public static bool LoadModule(string moduleTag, string moduleFilePath)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Loading module: {moduleTag} from {moduleFilePath}");

                // Check if module already loaded
                if (_loadedModules.ContainsKey(moduleTag))
                {
                    System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Module {moduleTag} already loaded");
                    return true;
                }

                // Check if file exists
                if (!File.Exists(moduleFilePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] ERROR: Module file not found: {moduleFilePath}");
                    return false;
                }

                // Read source code
                string sourceCode = File.ReadAllText(moduleFilePath);
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Source code read: {sourceCode.Length} characters");

                // Compile the code
                var assembly = CompileCode(sourceCode, moduleTag);
                if (assembly == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] ERROR: Failed to compile module {moduleTag}");
                    return false;
                }

                // Store the loaded module
                _loadedModules[moduleTag] = assembly;
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Module {moduleTag} loaded and compiled successfully");

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] EXCEPTION loading module {moduleTag}: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Compile C# source code into assembly
        /// </summary>
        private static Assembly CompileCode(string sourceCode, string moduleName)
        {
            try
            {
                // Create C# code provider
                var provider = new CSharpCodeProvider();

                // Set compiler parameters
                var parameters = new CompilerParameters
                {
                    GenerateInMemory = true,
                    GenerateExecutable = false,
                    TreatWarningsAsErrors = false
                };

                // Add required references
                parameters.ReferencedAssemblies.Add("System.dll");
                parameters.ReferencedAssemblies.Add("System.Core.dll");
                parameters.ReferencedAssemblies.Add("System.Xaml.dll");
                parameters.ReferencedAssemblies.Add("WindowsBase.dll");
                parameters.ReferencedAssemblies.Add("PresentationCore.dll");
                parameters.ReferencedAssemblies.Add("PresentationFramework.dll");
                
                // Add Revit API references
                parameters.ReferencedAssemblies.Add("C:\\Program Files\\Autodesk\\Revit 2024\\RevitAPI.dll");
                parameters.ReferencedAssemblies.Add("C:\\Program Files\\Autodesk\\Revit 2024\\RevitAPIUI.dll");

                // Add reference to current assembly (for accessing main plugin classes)
                var currentAssembly = Assembly.GetExecutingAssembly();
                parameters.ReferencedAssemblies.Add(currentAssembly.Location);

                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Compiling module {moduleName}...");

                // Compile the source code
                CompilerResults results = provider.CompileAssemblyFromSource(parameters, sourceCode);

                // Check for compilation errors
                if (results.Errors.HasErrors)
                {
                    StringBuilder errorMsg = new StringBuilder();
                    errorMsg.AppendLine($"Compilation errors for module {moduleName}:");
                    foreach (CompilerError error in results.Errors)
                    {
                        errorMsg.AppendLine($"  Line {error.Line}: {error.ErrorText}");
                        System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Compilation error: Line {error.Line}: {error.ErrorText}");
                    }
                    System.Diagnostics.Debug.WriteLine(errorMsg.ToString());
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Module {moduleName} compiled successfully");
                return results.CompiledAssembly;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] EXCEPTION during compilation: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Create instance of a class from loaded module
        /// </summary>
        public static object CreateModuleInstance(string moduleTag, string className, params object[] constructorArgs)
        {
            try
            {
                if (!_loadedModules.ContainsKey(moduleTag))
                {
                    System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] ERROR: Module {moduleTag} not loaded");
                    return null;
                }

                var assembly = _loadedModules[moduleTag];
                var type = assembly.GetTypes().FirstOrDefault(t => t.Name == className || t.FullName == className);

                if (type == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] ERROR: Class {className} not found in module {moduleTag}");
                    return null;
                }

                var instance = Activator.CreateInstance(type, constructorArgs);
                _moduleInstances[moduleTag] = instance;

                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Instance of {className} created successfully");
                return instance;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] EXCEPTION creating instance: {ex.Message}");
                return null;
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
        /// Get loaded module assembly
        /// </summary>
        public static Assembly GetModuleAssembly(string moduleTag)
        {
            return _loadedModules.ContainsKey(moduleTag) ? _loadedModules[moduleTag] : null;
        }

        /// <summary>
        /// Get module instance
        /// </summary>
        public static object GetModuleInstance(string moduleTag)
        {
            return _moduleInstances.ContainsKey(moduleTag) ? _moduleInstances[moduleTag] : null;
        }

        /// <summary>
        /// Unload module (clear from cache)
        /// </summary>
        public static void UnloadModule(string moduleTag)
        {
            if (_loadedModules.ContainsKey(moduleTag))
            {
                _loadedModules.Remove(moduleTag);
                System.Diagnostics.Debug.WriteLine($"[MODULE-LOADER] Module {moduleTag} unloaded");
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
    }
}
