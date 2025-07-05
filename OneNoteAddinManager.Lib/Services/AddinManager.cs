using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OneNoteAddinManager.Lib.Models;
using System.Runtime.Versioning;

namespace OneNoteAddinManager.Lib.Services
{
    [SupportedOSPlatform("windows")]
    public class AddinManager
    {
        private readonly IRegistryService _registryService;

        /// <summary>
        /// Initializes a new instance of the AddinManager with dependency injection.
        /// </summary>
        /// <param name="registryService">The registry service to use for registry operations</param>
        public AddinManager(IRegistryService registryService)
        {
            _registryService = registryService ?? throw new ArgumentNullException(nameof(registryService));
        }

        /// <summary>
        /// Default constructor that uses the Windows registry service.
        /// </summary>
        public AddinManager() : this(new RegistryService(new DotNetWindowsRegistry.WindowsRegistry()))
        {
        }

        public List<AddinInfo> GetAllAddins()
        {
            return _registryService.GetInstalledAddins();
        }

        public void EnableAddin(AddinInfo addin)
        {
            _registryService.SetAddinEnabled(addin, true);
        }

        public void DisableAddin(AddinInfo addin)
        {
            _registryService.SetAddinEnabled(addin, false);
        }

        public void RegisterNewAddin(string dllPath)
        {
            if (!File.Exists(dllPath))
            {
                throw new FileNotFoundException($"DLL file not found: {dllPath}");
            }

            // Extract information from the DLL path
            var fileName = Path.GetFileNameWithoutExtension(dllPath);
            var friendlyName = fileName.Replace(".", " ");
            var description = $"Add-in loaded from {dllPath}";

            // Generate a new GUID for the add-in
            var guid = Guid.NewGuid().ToString("B").ToUpper();

            _registryService.RegisterAddin(fileName, friendlyName, description, dllPath, guid);
        }

        public void UnregisterAddin(AddinInfo addin)
        {
            _registryService.UnregisterAddin(addin);
        }

        public bool IsRunningAsAdministrator()
        {
            return _registryService.IsRunningAsAdministrator();
        }

        public List<AddinInfo> FindOrphanedEntries()
        {
            var addins = GetAllAddins();
            return addins.Where(a => !string.IsNullOrEmpty(a.DllPath) && !File.Exists(a.DllPath)).ToList();
        }

        public void CleanupOrphanedEntries()
        {
            var orphaned = FindOrphanedEntries();
            foreach (var addin in orphaned)
            {
                try
                {
                    UnregisterAddin(addin);
                }
                catch (Exception ex)
                {
                    // Log the error but continue with other entries
                    System.Diagnostics.Debug.WriteLine($"Failed to cleanup {addin.Name}: {ex.Message}");
                }
            }
        }

        public void RefreshAddinStatus(AddinInfo addin)
        {
            var allAddins = GetAllAddins();
            var updated = allAddins.FirstOrDefault(a => a.Name == addin.Name);
            if (updated != null)
            {
                addin.IsEnabled = updated.IsEnabled;
                addin.LoadBehavior = updated.LoadBehavior;
                addin.DllPath = updated.DllPath;
            }
        }
    }
}
