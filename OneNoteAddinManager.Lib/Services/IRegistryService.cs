using System;
using System.Collections.Generic;
using OneNoteAddinManager.Lib.Models;

namespace OneNoteAddinManager.Lib.Services
{
    /// <summary>
    /// Abstraction for registry operations to enable testability and dependency injection.
    /// This interface wraps the core registry functionality needed by the application.
    /// </summary>
    public interface IRegistryService
    {
        /// <summary>
        /// Gets all installed OneNote add-ins from the registry.
        /// </summary>
        /// <returns>List of add-in information</returns>
        /// <exception cref="InvalidOperationException">Thrown when registry access fails</exception>
        List<AddinInfo> GetInstalledAddins();

        /// <summary>
        /// Sets the enabled state of an add-in by modifying its LoadBehavior registry value.
        /// </summary>
        /// <param name="addin">The add-in to enable or disable</param>
        /// <param name="enabled">True to enable, false to disable</param>
        /// <exception cref="InvalidOperationException">Thrown when registry access fails</exception>
        void SetAddinEnabled(AddinInfo addin, bool enabled);

        /// <summary>
        /// Registers a new add-in in the registry with the specified information.
        /// </summary>
        /// <param name="name">The add-in name</param>
        /// <param name="friendlyName">The friendly display name</param>
        /// <param name="description">The add-in description</param>
        /// <param name="dllPath">Path to the add-in DLL</param>
        /// <param name="guid">The add-in GUID</param>
        /// <exception cref="InvalidOperationException">Thrown when registry access fails</exception>
        void RegisterAddin(string name, string friendlyName, string description, string dllPath, string guid);

        /// <summary>
        /// Unregisters an add-in from the registry, removing all associated entries.
        /// </summary>
        /// <param name="addin">The add-in to unregister</param>
        /// <exception cref="InvalidOperationException">Thrown when registry access fails</exception>
        void UnregisterAddin(AddinInfo addin);

        /// <summary>
        /// Checks if the current user has administrator privileges for registry operations.
        /// </summary>
        /// <returns>True if running as administrator, false otherwise</returns>
        bool IsRunningAsAdministrator();
    }
}