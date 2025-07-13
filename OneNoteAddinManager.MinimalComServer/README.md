# Minimal COM Server

This project is a minimal implementation of a COM server in C++. It is intended
to demonstrate the core concepts of how a COM component is registered and used
by applications at runtime.

## How COM Components Work at Runtime

The primary way an application uses a COM component is through a process called
"activation," which is managed by the COM runtime system and the Windows
Registry. This process does **not** require the application to link against a
`.lib` file.

1.  **Registration is Key:** First, the COM server (`.dll`) must be registered
    on the system. A tool like `regsvr32.exe` is used, which calls the
    `DllRegisterServer` function inside the DLL. This function writes critical
    information to the Windows Registry, including a unique ID for the component
    (`CLSID`) and the full path to the `.dll` file.

2.  **The Activation Process:**
    - An application needs to use the component, so it calls a COM function like
      `CoCreateInstance()` with the component's `CLSID`.
    - The COM runtime looks up this `CLSID` in the Windows Registry
      (`HKEY_CLASSES_ROOT\CLSID\{...}`).
    - The registry entry tells the COM runtime where to find the `.dll` file on
      disk.
    - The COM runtime loads the `.dll` into the application's process.
    - Finally, it calls another standard function in the DLL,
      `DllGetClassObject`, to get the component's "class factory," which is then
      used to create an instance of the actual object.

This entire runtime mechanism is orchestrated by the registry. The client
application only needs to know the `CLSID` to get a fully functional object.

## Building the COM Server DLL

The Dynamic-Link Library (`.dll`) is the core of our COM server. A DLL is a type
of executable file that contains functions and resources that can be used by
other programs. In the context of COM, this `.dll` file has several key
responsibilities:

- **Hosting COM Objects:** It contains the actual implementation of the COM
  objects (like `MinimalComObject` in our case) that other applications will
  use.
- **Providing Standard Entry Points:** For a DLL to function as a COM server, it
  must export specific functions that the COM runtime can call. These include:
  - `DllRegisterServer`: Used to write the component's registration information
    to the Windows Registry.
  - `DllUnregisterServer`: Used to remove the component's registration
    information from the Windows Registry.
  - `DllGetClassObject`: Called by the COM runtime to obtain a pointer to the
    component's class factory, which is responsible for creating instances of
    the COM object.
  - `DllCanUnloadNow`: Called by the COM runtime to determine if the DLL can be
    safely unloaded from memory (i.e., if no objects or locks are outstanding).
- **Managing Object Lifecycles:** Internally, the DLL is responsible for
  creating and destroying instances of its COM objects, and managing their
  reference counts to ensure proper memory management.

This `.dll` file is produced using the Microsoft Visual C++ compiler (`cl.exe`)
and linker. The build process involves:

1.  **Compilation:** The `cl.exe` command compiles the `MinimalComServer.cpp`
    source file into an object file (`.obj`).
2.  **Linking:** The linker then takes this object file, along with the
    `MinimalComServer.def` file (which specifies the functions to be exported
    from the DLL), and links them with necessary system libraries (like
    `advapi32.lib` for registry functions and `ole32.lib` for COM functions).
    The `/LD` flag passed to `cl.exe` instructs it to produce a DLL.

To build the COM server, navigate to the project directory in a Visual Studio
Developer Command Prompt and run the following command:

```cmd
cl /LD MinimalComServer.cpp /link /DEF:MinimalComServer.def advapi32.lib ole32.lib
```

## Diagnostics: MinimalComServer.log

When the COM server DLL is loaded or used, it writes diagnostic messages to a
log file named `MinimalComServer.log` in the same directory as the DLL. This log
records key events such as DLL load/unload, COM registration/unregistration, and
object/class factory creation requests, including the requested interface IDs
(IIDs). Each entry is timestamped for troubleshooting and auditing purposes.

- **Location:** The log file is created in the DLL's directory, ensuring it is
  easy to find alongside the server binary.
- **Contents:** Events such as `DllMain` (process attach/detach),
  `DllRegisterServer`, `DllUnregisterServer`, and class factory/object creation
  are logged with relevant details and timestamps.
- **Purpose:** The log helps developers and administrators diagnose registration
  issues, activation failures, and interface negotiation problems without
  needing to attach a debugger.
- **Source:** Logging is implemented in the C++ code using a utility function
  that appends messages to the file. The log file is ignored by version control
  (see `.gitignore`).

---

## Appendix: Build Artifacts

When you compile a C++ project like this one, the compiler and linker generate
several files. Here is a breakdown of the files created during the build process
and their roles:

- **`.obj` (Object File):** An intermediate file generated by the compiler from
  a single `.cpp` source file. It contains machine code that the linker will
  combine with other `.obj` files to create the final DLL.

- **`.exp` (Export File):** An intermediate file generated by the linker that
  lists all the functions the DLL "exports" for external use. The linker uses
  this file to create the `.lib` import library.

- **`.lib` (Import Library):** This file is an **import library**, and it is
  used for a more traditional, compile-time method of using a DLL. In that
  model, a developer would directly link their application against the `.lib`
  file. This would allow their code to call functions exported from the DLL by
  name, as if they were part of the application itself.

  However, this is **not** how COM activation works. As described above, COM
  uses the registry and a class factory system. A client application that wants
  to use our `MinimalComObject` does not need to link against our `.lib` file to
  do so.

  So, why is a `.lib` file created in this project at all? It's because our
  `.def` file explicitly lists functions for export (`DllRegisterServer`, etc.).
  The C++ linker sees these exports and automatically generates a corresponding
  `.lib` file as part of the standard DLL build process. It is an artifact of
  the build tools, but it is not essential for the primary task of COM
  activation.

## Appendix: About regsvr32.exe

`regsvr32.exe` is a command-line utility provided by Windows for registering and
unregistering COM DLLs. When you run `regsvr32 MinimalComServer.dll`, the tool
loads the DLL and calls its `DllRegisterServer` function. This function writes
the necessary entries to the Windows Registry so that the COM runtime can locate
and activate the component using its CLSID. Unregistering is done with the `/u`
flag, which calls `DllUnregisterServer` to remove those registry entries. This
process is essential for making a COM server available to client applications
via COM activation.

### Bitness and Registry Views

Windows provides both 32-bit and 64-bit versions of `regsvr32.exe`:

- `C:\Windows\System32\regsvr32.exe` is the 64-bit version.
- `C:\Windows\SysWow64\regsvr32.exe` is the 32-bit version.

The version you run determines which registry view is used for COM registration:

- The 64-bit version registers the DLL in the 64-bit registry view.
- The 32-bit version registers the DLL in the 32-bit registry view.

You must use the version of `regsvr32.exe` that matches the bitness of your DLL.
This ensures the DLL is registered in the correct registry view and can be
activated by applications of the same bitness. If the 32-bit version is not in
your PATH, specify its full path when running it. This distinction allows both
32-bit and 64-bit COM servers to coexist on the same system without conflict.

## Appendix: The Windows Registry

The Windows Registry is a hierarchical database used by Windows to store
configuration settings and options for the operating system, applications, and
hardware devices. It consists of several top-level sections called "hives," such
as `HKEY_LOCAL_MACHINE` and `HKEY_CURRENT_USER`, each containing keys, subkeys,
and values. The registry enables Windows and applications to retrieve and update
settings efficiently. Changes to the registry can affect system behavior, user
preferences, and how software components like COM objects are registered and
activated. Understanding the registry structure is important for managing
system-wide and user-specific configurations.

## Appendix: HKEY_LOCAL_MACHINE and HKEY_CURRENT_USER

`HKEY_LOCAL_MACHINE\Software\Classes` is the system-wide registry location for
COM registration and file associations. Entries here apply to all users on the
computer and typically require administrator privileges to modify. Most COM
servers register themselves here by default, which is why admin rights are often
needed for COM registration.

`HKEY_CURRENT_USER\Software\Classes` is the user-specific registry location.
Entries here apply only to the current user and do not require administrator
rights. If a value exists in both `HKEY_CURRENT_USER\Software\Classes` and
`HKEY_LOCAL_MACHINE\Software\Classes`, the user's value takes precedence. This
allows individual users to override system-wide settings for COM components and
file associations.

`HKEY_CLASSES_ROOT` presents a merged view of both locations, showing
user-specific values first and falling back to system-wide values if needed.
This structure enables both per-user and system-wide COM registrations to
coexist and be resolved correctly by the COM runtime.

## Appendix: HKEY_CLASSES_ROOT

`HKEY_CLASSES_ROOT` (HKCR) is a top-level registry hive in Windows, but
technically it is a "merged view" or "virtual view" rather than a standalone
hive. HKCR combines data from both `HKEY_LOCAL_MACHINE\Software\Classes`
(system-wide) and `HKEY_CURRENT_USER\Software\Classes` (user-specific). When a
value exists in both locations, the user-specific value takes precedence. This
merged view allows Windows and applications to resolve file associations and COM
registrations efficiently, presenting a unified interface for lookups and
configuration.

For COM, HKCR contains the `CLSID` subkeys, which map unique class identifiers
(CLSIDs) to the details needed for activation, such as the path to the server
DLL and configuration settings. When a COM client requests an object by CLSID,
the COM runtime looks under `HKEY_CLASSES_ROOT\CLSID\{...}` to find the
necessary information, including the `InprocServer32` subkey. This structure
enables Windows to locate, load, and configure COM components for use by
applications, regardless of whether the registration is per-user or system-wide.

## Appendix: The InprocServer32 Subkey

The `InprocServer32` subkey is a critical part of COM registration in the
Windows Registry. It appears under the CLSID of a COM component (e.g.,
`HKEY_CLASSES_ROOT\CLSID\{...}\InprocServer32`) and specifies the path to the
DLL that implements the COM server. When a COM client requests an object by
CLSID, the COM runtime looks up this subkey to locate and load the DLL into the
client's process ("in-process server").

The name "InprocServer32" reflects its historical and architectural context:

- **"Inproc"** stands for "in-process," indicating that the server DLL is loaded
  directly into the client's process for fast, direct communication.
- **"Server"** refers to the component providing services (the COM object).
- **"32"** was introduced with 32-bit Windows to distinguish 32-bit in-process
  servers from earlier 16-bit ones (which used `InprocServer`).

Even for 64-bit COM servers, the subkey remains `InprocServer32` for
compatibility. The Windows registry and COM runtime use separate registry views
for 32-bit and 64-bit components, not different key names.

This means that Windows maintains two distinct sections ("views") of the
registry for COM registration: one for 32-bit processes and one for 64-bit
processes. When a 32-bit application accesses the registry, it sees the 32-bit
view; when a 64-bit application accesses the registry, it sees the 64-bit view.
This allows both 32-bit and 64-bit COM servers to be registered on the same
system without conflict, even though they use the same subkey name
(`InprocServer32`). The COM runtime automatically selects the correct view based
on the application's bitness.
