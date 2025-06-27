# OneNote Add-in Manager

## Purpose and Audience

This tool is designed for **OneNote add-in developers** who want visibility into the registration process that makes add-ins work with OneNote. It provides a user interface for managing registry operations related to add-in registration.

## Why This Tool Exists

OneNote add-in development involves COM registration and registry configuration. This tool helps with:

- Seeing why an add-in doesn't appear in OneNote after building
- Managing registry and COM registration
- Handling DLL paths and LoadBehavior settings
- Cleaning up registration entries

This tool provides visibility into OneNote's add-in registration system.

## Key Capabilities

### 1. Registration Management
- **Register new add-ins** with COM and registry setup
- **Enable/disable existing add-ins** for testing
- **Unregister add-ins** for removal
- **Clean up orphaned entries** to maintain registry health

### 2. Technical Visibility
- **View registry structure** showing OneNote add-in entries
- **See COM registration details** including CLSID and InprocServer32 configuration
- **Check DLL path validity** to verify assembly accessibility
- **Understand LoadBehavior** and add-in status

### 3. Development Support
- **Diagnose status** identifying why add-ins may not load
- **Validate registration** completeness
- **Access field information** with help for each component

## OneNote Add-in Registration Requirements

For OneNote to load an add-in, these elements should be configured:

### 1. Add-in Registry Entry
Location: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\[AddinName]`
- **FriendlyName**: Display name shown to users
- **LoadBehavior**: Controls how and when OneNote loads the add-in
- **Description**: Optional descriptive text

### 2. COM Class Registration  
Location: `HKEY_CLASSES_ROOT\CLSID\{GUID}`
- **Default value**: ProgID reference
- **AppID**: Application identifier
- **InprocServer32**: Assembly location and runtime details

### 3. Assembly Accessibility
- **DLL path**: Must be accessible to OneNote process
- **Dependencies**: Required assemblies available
- **Runtime version**: Compatible .NET Framework version

### 4. Unique Identification
- **GUID**: Globally unique identifier for COM registration
- **Consistency**: Same GUID across builds and versions

## Common Registration Scenarios

### Missing COM Registration
- Add-in class may lack COM visibility attributes
- Assembly not registered with COM subsystem
- GUID attribute missing or incorrectly formatted

### Path Issues
- DLL moved after registration
- Access permission issues

### LoadBehavior Settings
- LoadBehavior value may prevent loading
- OneNote may disable add-in after failures
- Administrative policies may affect add-in activation

### Registry Entries
- Incomplete registration entries
- Multiple or conflicting registrations
- Leftover entries from previous installations

## Development Workflow

### Initial Registration
1. **Build add-in** in your environment
2. **Browse to assembly** using the registration interface
3. **Register add-in** with COM and registry setup
4. **Check registration** in the details view
5. **Test** in OneNote

### Ongoing Development
1. **Rebuild** after code changes
2. **Refresh** tool to check current status
3. **Review** status and path validation
4. **Re-register** if paths or configuration changed
5. **Test** updated functionality

### Troubleshooting
1. **Check status** for specific information
2. **Validate paths** ensuring DLL accessibility
3. **Review COM registration** for completeness
4. **Clean and re-register** if needed
5. **Test** with known working configuration

## Interface Organization

### Left Panel: Add-in Management
- **Add-in list** showing registered add-ins
- **Refresh and cleanup** operations
- **Enable/disable and unregister** actions
- **New add-in registration** interface

### Center Panel: Technical Details
- **Add-in Information tab**: Basic properties and LoadBehavior
- **COM Registration tab**: Registry structure
- **Field-specific help** via information buttons

### Right Panel: Documentation
- **Help content** based on current context
- **Field documentation** 
- **Technical reference** for registration components

## Getting Started

**For new add-ins**: Use the registration section in the left panel to browse to your built assembly and register it with OneNote.

**For existing add-ins**: Select an add-in from the list to view its technical configuration.

This tool provides visibility into OneNote add-in registration to help with development and troubleshooting.