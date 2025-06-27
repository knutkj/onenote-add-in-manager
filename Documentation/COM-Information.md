# COM Registration and CLSID Information

## Overview

COM (Component Object Model) is the underlying technology that enables OneNote add-ins to integrate with the Office application. Understanding COM registration is essential for troubleshooting add-in issues and managing development environments.

## What is COM?

COM is Microsoft's binary-interface standard for software components. OneNote add-ins use COM to:
- Register themselves with Windows
- Communicate with the OneNote application
- Provide a standardized interface for Office integration

## Key COM Components

### CLSID (Class Identifier)
A unique GUID that identifies your add-in component in the Windows registry.

**Example:** `{12345678-1234-1234-1234-123456789ABC}`

### ProgID (Programmatic Identifier)
A human-readable identifier that maps to a CLSID.

**Example:** `MyAddin.Connect`

### InprocServer32
Registry key that tells Windows where to find the DLL file for your add-in.

## Registry Structure

COM registration involves several registry locations:

### 1. HKEY_CLASSES_ROOT\AppID\{GUID}
- **Purpose:** Application identifier registration
- **Values:**
  - `DllSurrogate` = "" (empty string for in-process server)

### 2. HKEY_CLASSES_ROOT\CLSID\{GUID}
- **Purpose:** Main COM class registration
- **Values:**
  - `(Default)` = "YourAddin.Connect"
  - `AppID` = "{GUID}"

### 3. HKEY_CLASSES_ROOT\CLSID\{GUID}\InprocServer32
- **Purpose:** Specifies the DLL location and runtime information
- **Values:**
  - `(Default)` = "mscoree.dll" (for .NET assemblies)
  - `ThreadingModel` = "Both"
  - `CodeBase` = "file:///C:/Path/To/Your/Assembly.dll"
  - `Class` = "YourAddin.Connect"
  - `RuntimeVersion` = "v4.0.30319"

### 4. HKEY_CLASSES_ROOT\CLSID\{GUID}\ProgID
- **Purpose:** Maps CLSID to ProgID
- **Values:**
  - `(Default)` = "YourAddin.Connect"

## Common COM Issues

### Missing Registry Entries
- **Symptom:** Add-in doesn't appear in OneNote
- **Solution:** Verify all required registry keys exist

### Incorrect CodeBase Path
- **Symptom:** Add-in fails to load with "file not found" errors
- **Solution:** Ensure CodeBase points to the correct DLL location

### Wrong ThreadingModel
- **Symptom:** Crashes or instability
- **Solution:** Use "Both" for Office add-ins

### Permission Issues
- **Symptom:** Cannot register or modify COM entries
- **Solution:** Run registration tools as Administrator

## Development vs Production

### Development Environment
- DLL typically located in `bin\Debug` or `bin\Release` folders
- CodeBase points to local development path
- Frequent re-registration during development

### Production Environment
- DLL installed to Program Files or similar location
- CodeBase points to final installation path
- Registered once during installation

## Troubleshooting Commands

### View COM Registration
```cmd
reg query "HKCR\CLSID\{YOUR-GUID}"
```

### Register Assembly
```cmd
regasm YourAddin.dll /codebase
```

### Unregister Assembly
```cmd
regasm YourAddin.dll /unregister
```

## Best Practices

1. **Always use strong names** for production assemblies
2. **Register with /codebase** during development
3. **Use installer for production** deployment
4. **Test in clean environment** before release
5. **Handle registration errors** gracefully in installers

## Security Considerations

- COM registration requires administrative privileges
- Be cautious with CodeBase paths containing spaces
- Validate DLL signatures in production environments
- Use least-privilege principles for add-in execution