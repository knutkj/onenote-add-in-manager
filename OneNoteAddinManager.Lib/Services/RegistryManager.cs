using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Runtime.Versioning;
using Microsoft.Win32;
using OneNoteAddinManager.Lib.Models;

namespace OneNoteAddinManager.Lib.Services
{
    [SupportedOSPlatform("windows")]
    public class RegistryManager
    {
        private const string OFFICE_ADDINS_PATH = @"SOFTWARE\Microsoft\Office\OneNote\AddIns";
        private const string CLASSES_ROOT_APPID = @"AppID";
        private const string CLASSES_ROOT_CLSID = @"CLSID";
        private const string WOW64_CLSID = @"WOW6432Node\CLSID";

        public List<AddinInfo> GetInstalledAddins()
        {
            var addins = new List<AddinInfo>();

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(OFFICE_ADDINS_PATH))
                {
                    if (key != null)
                    {
                        foreach (string subKeyName in key.GetSubKeyNames())
                        {
                            using (var subKey = key.OpenSubKey(subKeyName))
                            {
                                if (subKey != null)
                                {
                                    var addin = CreateAddinFromRegistry(subKeyName, subKey);
                                    if (addin != null)
                                    {
                                        addins.Add(addin);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (SecurityException ex)
            {
                throw new InvalidOperationException("Access denied to registry. Please run as administrator.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error reading registry: {ex.Message}", ex);
            }

            return addins;
        }

        private AddinInfo CreateAddinFromRegistry(string keyName, RegistryKey subKey)
        {
            try
            {
                var addin = new AddinInfo
                {
                    Name = keyName,
                    FriendlyName = subKey.GetValue("FriendlyName")?.ToString() ?? keyName,
                    Description = subKey.GetValue("Description")?.ToString() ?? "",
                    LoadBehavior = Convert.ToInt32(subKey.GetValue("LoadBehavior") ?? 0)
                };

                addin.IsEnabled = addin.LoadBehavior == 3;

                // Try to find the CLSID and DLL path
                var guid = FindAddinGuid(keyName);
                if (!string.IsNullOrEmpty(guid))
                {
                    addin.Guid = guid;
                    addin.DllPath = FindDllPath(guid);
                }

                return addin;
            }
            catch (Exception)
            {
                return null!;
            }
        }

        private string FindAddinGuid(string addinName)
        {
            try
            {
                // Look for CLSID in HKEY_CLASSES_ROOT\AddinName\CLSID
                using (var key = Registry.ClassesRoot.OpenSubKey($@"{addinName}\CLSID"))
                {
                    return key?.GetValue("")?.ToString()!;
                }
            }
            catch
            {
                return null!;
            }
        }

        private string FindDllPath(string guid)
        {
            try
            {
                // Try 64-bit first
                using (var key = Registry.ClassesRoot.OpenSubKey($@"{CLASSES_ROOT_CLSID}\{guid}\InprocServer32"))
                {
                    var codeBase = key?.GetValue("CodeBase")?.ToString();
                    if (!string.IsNullOrEmpty(codeBase))
                    {
                        return codeBase.Replace("file:///", "").Replace("/", "\\");
                    }
                }

                // Try WOW64 (32-bit)
                using (var key = Registry.ClassesRoot.OpenSubKey($@"{WOW64_CLSID}\{guid}\InprocServer32"))
                {
                    var codeBase = key?.GetValue("CodeBase")?.ToString();
                    if (!string.IsNullOrEmpty(codeBase))
                    {
                        return codeBase.Replace("file:///", "").Replace("/", "\\");
                    }
                }
            }
            catch
            {
                // Ignore errors
            }

            return null!;
        }


        public void SetAddinEnabled(AddinInfo addin, bool enabled)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey($@"{OFFICE_ADDINS_PATH}\{addin.Name}", true))
                {
                    if (key != null)
                    {
                        int loadBehavior = enabled ? 3 : 0;
                        key.SetValue("LoadBehavior", loadBehavior, RegistryValueKind.DWord);
                        addin.LoadBehavior = loadBehavior;
                        addin.IsEnabled = enabled;
                    }
                    else
                    {
                        throw new InvalidOperationException($"Add-in registry key not found: {addin.Name}");
                    }
                }
            }
            catch (SecurityException ex)
            {
                throw new InvalidOperationException("Access denied to registry. Please run as administrator.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error updating registry: {ex.Message}", ex);
            }
        }

        public void RegisterAddin(string name, string friendlyName, string description, string dllPath, string guid)
        {
            try
            {
                // Register in CurrentUser AddIns
                using (var addinsKey = Registry.CurrentUser.CreateSubKey($@"{OFFICE_ADDINS_PATH}\{name}"))
                {
                    addinsKey.SetValue("FriendlyName", friendlyName);
                    addinsKey.SetValue("Description", description);
                    addinsKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }

                // Register CLSID entries
                RegisterCLSID(guid, name, dllPath);
            }
            catch (SecurityException ex)
            {
                throw new InvalidOperationException("Access denied to registry. Please run as administrator.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error registering add-in: {ex.Message}", ex);
            }
        }

        private void RegisterCLSID(string guid, string progId, string dllPath)
        {
            // Register AppID
            using (var appIdKey = Registry.ClassesRoot.CreateSubKey($@"{CLASSES_ROOT_APPID}\{guid}"))
            {
                appIdKey.SetValue("DllSurrogate", "");
            }

            // Register CLSID
            using (var clsidKey = Registry.ClassesRoot.CreateSubKey($@"{CLASSES_ROOT_CLSID}\{guid}"))
            {
                clsidKey.SetValue("", $"{progId}.AddIn");
                clsidKey.SetValue("AppID", guid);
            }

            // Register InprocServer32
            using (var inprocKey = Registry.ClassesRoot.CreateSubKey($@"{CLASSES_ROOT_CLSID}\{guid}\InprocServer32"))
            {
                inprocKey.SetValue("", "mscoree.dll");
                inprocKey.SetValue("ThreadingModel", "Both");
                inprocKey.SetValue("CodeBase", $"file:///{dllPath.Replace("\\", "/")}");
                inprocKey.SetValue("Class", $"{progId}.AddIn");
                inprocKey.SetValue("RuntimeVersion", "v4.0.30319");
            }

            // Register ProgID
            using (var progIdKey = Registry.ClassesRoot.CreateSubKey($@"{CLASSES_ROOT_CLSID}\{guid}\ProgID"))
            {
                progIdKey.SetValue("", progId);
            }
        }

        public void UnregisterAddin(AddinInfo addin)
        {
            try
            {
                // Remove from CurrentUser AddIns
                Registry.CurrentUser.DeleteSubKeyTree($@"{OFFICE_ADDINS_PATH}\{addin.Name}", false);

                // Remove CLSID entries if we have the GUID
                if (!string.IsNullOrEmpty(addin.Guid))
                {
                    Registry.ClassesRoot.DeleteSubKeyTree($@"{CLASSES_ROOT_APPID}\{addin.Guid}", false);
                    Registry.ClassesRoot.DeleteSubKeyTree($@"{CLASSES_ROOT_CLSID}\{addin.Guid}", false);
                }
            }
            catch (SecurityException ex)
            {
                throw new InvalidOperationException("Access denied to registry. Please run as administrator.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error unregistering add-in: {ex.Message}", ex);
            }
        }

        public bool IsRunningAsAdministrator()
        {
            try
            {
                // Try to write to HKEY_CLASSES_ROOT to test permissions
                using (var testKey = Registry.ClassesRoot.CreateSubKey("OneNoteAddinManager_Test"))
                {
                    if (testKey != null)
                    {
                        Registry.ClassesRoot.DeleteSubKey("OneNoteAddinManager_Test", false);
                        return true;
                    }
                }
            }
            catch
            {
                // If we can't write, we don't have admin privileges
            }

            return false;
        }
    }
}
