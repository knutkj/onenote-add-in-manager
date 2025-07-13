// Windows API core definitions and types.
#include <windows.h>

// COM interface definitions (IDispatch, IUnknown, etc.).
#include <oaidl.h>

// Safe string manipulation functions (StringCchPrintfW, etc.).
#include <strsafe.h>

// OLE/COM helper functions and types.
#include <olectl.h>

// String stream for building log messages.
#include <sstream>

// Utility functions for logging and module path
#include "utils.h"

// COM classes
#include "MinimalComObject.h"
#include "MinimalFactory.h"

// Globals
HMODULE g_hModule; // Handle to the DLL module, used for path resolution
ULONG g_lockCount = 0; // COM server lock count
const CLSID CLSID_GUID = // COM class ID
    {0xea8d3bc3, 0x3109, 0x4f38, {0xa0, 0x85, 0xbd, 0x21, 0xcb, 0x4f, 0x1d, 0xd5}};

const wchar_t* const PROGID = L"MinimalComServer.Object";
const wchar_t* const PROGID_DESCRIPTION = L"Minimal COM Server Object";

//
// COM Stuff.
//



extern "C" STDAPI DllCanUnloadNow() {
    return (g_lockCount == 0) ? S_OK : S_FALSE;
}

extern "C" STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv) {
    if (clsid != CLSID_GUID) return CLASS_E_CLASSNOTAVAILABLE;
    MinimalFactory* factory = new MinimalFactory();
    HRESULT hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}

/**
 * @brief Called by regsvr32.exe to register the COM server.
 * 
 * This function creates the necessary registry entries for the COM server,
 * including the CLSID, the path to the DLL, and the threading model.
 * These entries allow COM clients to find and use the server.
 * @return HRESULT 
 */
extern "C" STDAPI DllRegisterServer() {
    HKEY hKey;
    std::wstring modulePath = GetDllModulePath();

    // Convert the CLSID to a string.
    wchar_t clsidStr[64];
    StringFromGUID2(CLSID_GUID, clsidStr, 64);
    
    // Create the registry key path: Software\\Classes\\CLSID\\{...}\\InprocServer32 (per-user)
    wchar_t keyPath[256];
    StringCchPrintfW(keyPath, 256, L"Software\\Classes\\CLSID\\%s\\InprocServer32", clsidStr);

    // Create the registry key under HKEY_CURRENT_USER (per-user, no admin rights needed)
    if (RegCreateKeyW(HKEY_CURRENT_USER, keyPath, &hKey) != ERROR_SUCCESS)
        return SELFREG_E_CLASS;

    // Set the default value of the key to the path of the DLL.
    RegSetValueExW(hKey, NULL, 0, REG_SZ,
        (BYTE*)modulePath.c_str(), (DWORD)((modulePath.length()+1)*sizeof(wchar_t)));

    // Set the threading model to "Both".
    RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, (BYTE*)L"Both", sizeof(L"Both"));

    // Close the registry key.
    RegCloseKey(hKey);

    // Register ProgID: MinimalComServer.Object
    wchar_t progIdKey[256];
    StringCchPrintfW(progIdKey, 256, L"Software\\Classes\\%s", PROGID);
    if (RegCreateKeyW(HKEY_CURRENT_USER, progIdKey, &hKey) == ERROR_SUCCESS) {
        // Set the default value of the ProgID key to a human-readable description.
        // This is what appears in registry tools and helps identify the COM object.
        RegSetValueExW(hKey, NULL, 0, REG_SZ,
            (BYTE*)PROGID_DESCRIPTION,
            (DWORD)((wcslen(PROGID_DESCRIPTION)+1)*sizeof(wchar_t)));

        // Create CLSID subkey under ProgID
        HKEY hClsidKey;
        if (RegCreateKeyW(hKey, L"CLSID", &hClsidKey) == ERROR_SUCCESS) {
            RegSetValueExW(hClsidKey, NULL, 0, REG_SZ, (BYTE*)clsidStr, (DWORD)((wcslen(clsidStr)+1)*sizeof(wchar_t)));
            RegCloseKey(hClsidKey);
        }
        RegCloseKey(hKey);
    }

    // Log registration
    std::wstringstream msg;
    msg << L"DllRegisterServer called. Registered at " <<
        keyPath << L" and ProgID " << PROGID;

    Log(msg.str().c_str());

    return S_OK;
}

/**
 * @brief Called by regsvr32.exe /u to unregister the COM server.
 * 
 * This function removes the registry entries created by DllRegisterServer.
 * This ensures that the component is cleanly uninstalled from the system.
 * @return HRESULT 
 */
extern "C" STDAPI DllUnregisterServer() {
    // Convert the CLSID to a string.
    wchar_t clsidStr[64];
    StringFromGUID2(CLSID_GUID, clsidStr, 64);

    // Create the registry key path: Software\\Classes\\CLSID\\{...} (per-user)
    wchar_t keyPath[256];
    StringCchPrintfW(keyPath, 256, L"Software\\Classes\\CLSID\\%s", clsidStr);

    // Delete the registry key under HKEY_CURRENT_USER (per-user)
    BOOL deleted = (RegDeleteTreeW(HKEY_CURRENT_USER, keyPath) == ERROR_SUCCESS);

    // Delete ProgID key
    wchar_t progIdKey[256];
    StringCchPrintfW(progIdKey, 256, L"Software\\Classes\\%s", PROGID);
    BOOL progidDeleted =
        (RegDeleteTreeW(HKEY_CURRENT_USER, progIdKey) == ERROR_SUCCESS);

    // Log unregistration
    std::wstringstream msg;
    msg << L"DllUnregisterServer called. Unregistered at " <<
        keyPath << L" and ProgID " << PROGID;

    Log(msg.str().c_str());

    return (deleted && progidDeleted) ? S_OK : S_FALSE;
}

/**
 * @brief Entry point for the DLL.
 *
 * DllMain is called by the system when processes and threads are initialized
 * and terminated, or upon calls to LoadLibrary and FreeLibrary. On
 * DLL_PROCESS_ATTACH, it stores the module handle for use in path resolution
 * and logs the loading process's PID and name. Returns TRUE to indicate
 * successful initialization.
 *
 * @param hModule Handle to the DLL module.
 * @param reason Reason code for the call (e.g., DLL_PROCESS_ATTACH,
 * DLL_PROCESS_DETACH).
 * @param lpReserved Reserved, not used.
 * @return BOOL TRUE if successful, FALSE otherwise.
 */
BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        g_hModule = hModule;

        // Get current process ID
        DWORD pid = GetCurrentProcessId();

        // Get the full path of the executable that loaded this DLL
        wchar_t processPath[MAX_PATH];
        GetModuleFileNameW(NULL, processPath, MAX_PATH);

        // Extract the process name from the full path
        std::wstring wsProcessPath(processPath);
        size_t lastSlash = wsProcessPath.find_last_of(L"\\/");
        std::wstring processName = (lastSlash == std::wstring::npos) ?
            wsProcessPath : wsProcessPath.substr(lastSlash + 1);

        // Format the log message
        std::wstringstream ss;
        ss << L"MinimalComServer.dll loaded into process: " <<
            processName << L" (PID: " << pid << L") at path: " << processPath;

        // Output to debugger (e.g., DebugView)
        OutputDebugStringW(ss.str().c_str());

        // Also log to file if file logging is desired/configured
        Log(ss.str().c_str());

    } else if (reason == DLL_PROCESS_DETACH) {
        Log(L"DllMain: DLL_PROCESS_DETACH - DLL unloading.");
    }

    // Optionally log other reasons if needed
    // (DLL_THREAD_ATTACH, DLL_THREAD_DETACH).
    return TRUE;
}
