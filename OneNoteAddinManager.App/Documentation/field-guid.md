# GUID Field

## What is the GUID field?

The **GUID** (Globally Unique Identifier) is a unique 128-bit identifier that serves as the primary key for COM registration of your OneNote add-in. This identifier is crucial for Windows to properly locate and instantiate your add-in.

## GUID Format and Structure

### Standard Format
GUIDs are displayed in the standard format:
```
{12345678-1234-1234-1234-123456789ABC}
```

### Components
- **32 hexadecimal digits** (0-9, A-F)
- **Four hyphens** separating into five groups
- **Curly braces** enclosing the entire identifier
- **Total length**: 38 characters including braces and hyphens

### Case Sensitivity
- Typically displayed in uppercase
- Registry searches are case-insensitive
- Best practice: maintain consistent case

## Role in COM Registration

### CLSID (Class Identifier)
The GUID serves as the CLSID for COM registration:
```
HKEY_CLASSES_ROOT\CLSID\{YOUR-GUID-HERE}
```

### AppID Registration
May also be used as AppID for application-level COM settings:
```
HKEY_CLASSES_ROOT\AppID\{YOUR-GUID-HERE}
```

### Unique Identification
- **No Two Alike**: Mathematically guaranteed to be unique
- **Global Scope**: Unique across all computers and time
- **Persistent**: Should never change for a given add-in

## GUID Generation

### Visual Studio
When creating COM-visible classes in Visual Studio:
```csharp
[Guid("12345678-1234-1234-1234-123456789ABC")]
[ComVisible(true)]
public class MyAddinConnect
{
    // Add-in implementation
}
```

### Manual Generation
Tools for generating GUIDs:
- **Visual Studio**: Tools → Create GUID
- **PowerShell**: `[System.Guid]::NewGuid()`
- **Online Tools**: Various GUID generators
- **uuidgen.exe**: Windows SDK utility

### Registry Format
In the registry, GUIDs appear in various formats:
- **With Braces**: `{12345678-1234-1234-1234-123456789ABC}`
- **Without Braces**: `12345678-1234-1234-1234-123456789ABC`
- **Different Cases**: Upper or lowercase hexadecimal digits

## Discovery Methods

### Assembly Reflection
GUIDs can be found by examining the add-in assembly:
- **GuidAttribute**: Applied to COM-visible classes
- **Type.GUID Property**: Runtime GUID discovery
- **Assembly Metadata**: Embedded GUID information

### Registry Search
This tool discovers GUIDs by searching:
1. **HKEY_CLASSES_ROOT\AddinName\CLSID**: Direct CLSID reference
2. **HKEY_CLASSES_ROOT\CLSID**: Scanning for matching assemblies
3. **Progressive Search**: Multiple registry locations

### File Analysis
Alternative discovery methods:
- **TLB Files**: Type library GUID information
- **Manifest Files**: Registration-free COM manifests
- **Assembly Attributes**: Examining .NET assembly metadata

## GUID-Related Registry Entries

### Primary CLSID Entry
```
HKEY_CLASSES_ROOT\CLSID\{GUID}
├── (Default) = "AddinName.Connect"
├── AppID = "{GUID}"
├── InprocServer32\
│   ├── (Default) = "mscoree.dll"
│   ├── ThreadingModel = "Both"
│   ├── Class = "AddinName.Connect"
│   ├── Assembly = "AddinName, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
│   └── RuntimeVersion = "v4.0.30319"
└── ProgID\
    └── (Default) = "AddinName.Connect"
```

### AppID Entry
```
HKEY_CLASSES_ROOT\AppID\{GUID}
└── DllSurrogate = ""
```

## Troubleshooting GUID Issues

### GUID Not Found
If no GUID is displayed:
- **Incomplete Registration**: Add-in may not be fully registered
- **Missing COM Attributes**: Assembly lacks proper COM visibility
- **Registry Corruption**: COM registration entries damaged
- **Permission Issues**: Cannot access required registry areas

### Wrong GUID Displayed
If incorrect GUID appears:
- **Multiple Registrations**: Conflicting registry entries
- **Cached Information**: Old registry data not cleaned up
- **Assembly Conflicts**: Multiple versions with different GUIDs

### Registration Failures
Common GUID-related registration problems:
- **GUID Conflicts**: Same GUID used by multiple applications
- **Invalid Format**: Malformed GUID in registry
- **Permission Denied**: Cannot write to CLSID registry keys

## GUID Best Practices

### Generation
1. **Always Generate New**: Never reuse GUIDs from examples or other projects
2. **Use Proper Tools**: Rely on standard GUID generation utilities
3. **Document GUIDs**: Keep records of GUIDs used in projects
4. **Version Consistency**: Same GUID across versions of same add-in

### Management
1. **Source Control**: Include GUID in source code attributes
2. **Build Process**: Ensure consistent GUID across builds
3. **Documentation**: Record GUID in project documentation
4. **Backup**: Maintain GUID information for disaster recovery

### Security
1. **GUID Privacy**: GUIDs may reveal information about development environment
2. **Predictability**: Never use sequential or predictable GUIDs
3. **Validation**: Verify GUID format before registration
4. **Cleanup**: Remove unused GUID registrations

## Advanced GUID Topics

### GUID Versions
Different GUID generation algorithms:
- **Version 1**: Time-based with MAC address
- **Version 4**: Random or pseudo-random
- **Version 5**: Name-based with SHA-1 hashing

### Registry-Free COM
Modern alternatives to GUID registration:
- **Manifest-based**: Registration-free COM with manifests
- **Side-by-side**: Isolated COM component deployment
- **Application Context**: App-specific COM registration

### Cross-Platform Considerations
- **GUID Portability**: Same concepts apply across platforms
- **.NET Core**: GUID usage in cross-platform applications
- **Interoperability**: GUID translation between systems

## Debugging GUID Issues

### Registry Analysis Tools
- **RegEdit**: Manual registry inspection
- **Process Monitor**: Track registry access
- **Registry Compare**: Identify registration differences

### Development Tools
- **OLEView**: COM object browser and registry viewer
- **Visual Studio**: COM registration debugging
- **PowerShell**: Automated GUID and registry analysis

### Logging and Diagnostics
- **Event Logs**: Windows Event Log COM errors
- **Debug Output**: Add-in initialization logging
- **Registry Auditing**: Track GUID registration changes