using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OneNoteAddinManager.Models;

namespace OneNoteAddinManager.Services
{
    public class AddinManager
    {
        private readonly RegistryManager _registryManager;

        public AddinManager()
        {
            _registryManager = new RegistryManager();
        }

        public List<AddinInfo> GetAllAddins()
        {
            return _registryManager.GetInstalledAddins();
        }

        public void EnableAddin(AddinInfo addin)
        {
            _registryManager.SetAddinEnabled(addin, true);
        }

        public void DisableAddin(AddinInfo addin)
        {
            _registryManager.SetAddinEnabled(addin, false);
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

            _registryManager.RegisterAddin(fileName, friendlyName, description, dllPath, guid);
        }

        public void UnregisterAddin(AddinInfo addin)
        {
            _registryManager.UnregisterAddin(addin);
        }


        public bool IsRunningAsAdministrator()
        {
            return _registryManager.IsRunningAsAdministrator();
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
