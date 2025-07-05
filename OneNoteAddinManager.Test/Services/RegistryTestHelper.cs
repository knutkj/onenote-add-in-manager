using System;
using DotNetWindowsRegistry;
using Microsoft.Win32;
using OneNoteAddinManager.Lib.Models;

namespace OneNoteAddinManager.Test.Services;

/// <summary>
/// Helper class to set up test data in InMemoryRegistry for testing purposes.
/// </summary>
public static class RegistryTestHelper
{
    private const string OFFICE_ADDINS_PATH = @"SOFTWARE\Microsoft\Office\OneNote\AddIns";
    private const string CLASSES_ROOT_APPID = @"AppID";
    private const string CLASSES_ROOT_CLSID = @"CLSID";

    public static void SetupTestData(IRegistry registry, params AddinInfo[] testAddins)
    {
        foreach (var addin in testAddins)
        {
            // Create the add-in registry entry
            using (var addinsKey = registry.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).CreateSubKey($@"{OFFICE_ADDINS_PATH}\{addin.Name}"))
            {
                addinsKey.SetValue("FriendlyName", addin.FriendlyName);
                addinsKey.SetValue("Description", addin.Description);
                addinsKey.SetValue("LoadBehavior", addin.LoadBehavior, RegistryValueKind.DWord);
            }

            // If we have a GUID, set up the CLSID entries
            if (!string.IsNullOrEmpty(addin.Guid))
            {
                // Set up CLSID lookup path
                using (var clsidLookupKey = registry.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default).CreateSubKey($@"{addin.Name}\CLSID"))
                {
                    clsidLookupKey.SetValue("", addin.Guid);
                }

                // Set up the full CLSID registration
                RegisterCLSID(registry, addin.Guid, addin.Name, addin.DllPath ?? "");
            }
        }
    }

    private static void RegisterCLSID(IRegistry registry, string guid, string progId, string dllPath)
    {
        // Register AppID
        using (var appIdKey = registry.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default).CreateSubKey($@"{CLASSES_ROOT_APPID}\{guid}"))
        {
            appIdKey.SetValue("DllSurrogate", "");
        }

        // Register CLSID
        using (var clsidKey = registry.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default).CreateSubKey($@"{CLASSES_ROOT_CLSID}\{guid}"))
        {
            clsidKey.SetValue("", $"{progId}.AddIn");
            clsidKey.SetValue("AppID", guid);
        }

        // Register InprocServer32
        using (var inprocKey = registry.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default).CreateSubKey($@"{CLASSES_ROOT_CLSID}\{guid}\InprocServer32"))
        {
            inprocKey.SetValue("", "mscoree.dll");
            inprocKey.SetValue("ThreadingModel", "Both");
            inprocKey.SetValue("CodeBase", $"file:///{dllPath.Replace("\\", "/")}");
            inprocKey.SetValue("Class", $"{progId}.AddIn");
            inprocKey.SetValue("RuntimeVersion", "v4.0.30319");
        }

        // Register ProgID
        using (var progIdKey = registry.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default).CreateSubKey($@"{CLASSES_ROOT_CLSID}\{guid}\ProgID"))
        {
            progIdKey.SetValue("", progId);
        }
    }
}