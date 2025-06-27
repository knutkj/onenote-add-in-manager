# Add-in Field Reference

## Basic Information Fields

### Name
The internal name of the add-in, typically the COM class name. This is used internally by OneNote to identify the add-in component.

### Friendly Name  
The display name shown to users in OneNote's COM Add-ins dialog. This should be descriptive and user-friendly.

### Status
Current operational state of the add-in:
- **Enabled**: Add-in is active and will load with OneNote
- **Disabled**: Add-in is installed but inactive
- **Not Loaded**: Add-in failed to load (check DLL path and dependencies)

### Environment
Indicates whether the add-in is configured for:
- **Development**: Points to debug/development DLL
- **Production**: Points to release/production DLL

### GUID
Unique identifier for the add-in. This GUID must match the one declared in the add-in's source code and is used for COM registration.

### DLL Path
File system path to the add-in's compiled library. OneNote uses this path to load the add-in at startup.

## LoadBehavior Values

The LoadBehavior registry value controls how OneNote loads the add-in:

- **0**: Do not load
- **1**: Load on startup  
- **2**: Load on demand
- **3**: Load on startup (default for most add-ins)
- **8**: Connect on next startup only
- **9**: Load on startup and stay connected
- **16**: Connect first time, then load on demand

## Registry Key Purposes

### HKEY_CLASSES_ROOT\CLSID\{GUID}
COM class registration containing the add-in's implementation details and server information.

### HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\Addins\{ProgID}
User-specific add-in settings including LoadBehavior and friendly name.

### HKEY_LOCAL_MACHINE\Software\Microsoft\Office\OneNote\Addins\{ProgID}  
Machine-wide add-in settings (requires administrator privileges to modify).