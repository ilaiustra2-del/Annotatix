using System;
using System.Windows;

namespace PluginsManager.Core
{
    /// <summary>
    /// Interface for dynamically loaded plugin modules
    /// </summary>
    public interface IModule
    {
        /// <summary>
        /// Module unique identifier (e.g., "dwg2rvt", "hvac")
        /// </summary>
        string ModuleId { get; }
        
        /// <summary>
        /// Module display name
        /// </summary>
        string ModuleName { get; }
        
        /// <summary>
        /// Module version
        /// </summary>
        string ModuleVersion { get; }
        
        /// <summary>
        /// Initialize the module
        /// </summary>
        void Initialize();
        
        /// <summary>
        /// Create and return the main panel for this module
        /// </summary>
        Window CreatePanel(object[] parameters);
    }
}
