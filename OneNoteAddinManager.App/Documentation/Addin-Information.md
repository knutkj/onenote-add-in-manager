# OneNote Add-in Management Guide

## Overview

OneNote add-ins extend the functionality of Microsoft OneNote by providing custom features, automation, and integration capabilities. This guide covers add-in management, configuration, and troubleshooting.

## What are OneNote Add-ins?

OneNote add-ins are COM-based extensions that integrate directly with the OneNote application. They can:
- Add custom ribbon buttons and menus
- Automate OneNote operations
- Integrate with external systems
- Provide custom functionality and workflows

## Add-in Types

### Office Add-ins (VSTO)
- **Technology:** Visual Studio Tools for Office
- **Language:** Typically C# or VB.NET
- **Registration:** Registry-based COM registration
- **Deployment:** Windows Installer (MSI) or ClickOnce

### Web Add-ins (Office.js)
- **Technology:** JavaScript API
- **Language:** HTML, CSS, JavaScript
- **Registration:** Manifest-based
- **Deployment:** Web-based or AppSource

*Note: This tool primarily manages VSTO/COM-based add-ins.*

## Add-in Fields Explained

### Basic Information

**Name**
- The internal identifier for the add-in
- Used in registry keys
- Usually matches the assembly name

**Friendly Name**
- Display name shown to users
- Appears in OneNote's add-in lists
- Can be localized for different languages

**Description**
- Brief explanation of add-in functionality
- Helps users understand the add-in's purpose
- Optional but recommended

**Status**
- **Enabled:** Add-in loads when OneNote starts
- **Disabled:** Add-in is installed but not loaded
- **Error:** Add-in failed to load (check DLL path)

### Technical Information

**GUID (Globally Unique Identifier)**
- Unique identifier for COM registration
- Format: `{12345678-1234-1234-1234-123456789ABC}`
- Required for proper COM registration

**DLL Path**
- Location of the add-in assembly file
- Must be accessible by OneNote process
- Location varies based on installation method

## LoadBehavior Values

The LoadBehavior registry value controls how OneNote handles the add-in:

### LoadBehavior = 0 (Disconnected)
- Add-in is not loaded
- Manual intervention required to enable
- Used for disabled add-ins

### LoadBehavior = 1 (Connected)
- Add-in is loaded and connected
- Rare in modern Office versions
- Legacy behavior

### LoadBehavior = 2 (Bootload)
- Load at application startup
- Not commonly used for OneNote
- Historical significance

### LoadBehavior = 3 (Demand Load - Recommended)
- **Default and recommended setting**
- Add-in loads when OneNote starts
- Automatic error recovery
- Best user experience

### LoadBehavior = 8 (Demand Load)
- Similar to 3 but with different error handling
- Less commonly used
- May cause issues with some add-ins

### LoadBehavior = 16 (Connect First Time)
- Load add-in on first use
- Not recommended for OneNote
- Can cause user confusion

## Registry Locations

### Add-in Registration
```
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\[AddinName]
```

**Values:**
- `FriendlyName` (String): Display name
- `Description` (String): Add-in description
- `LoadBehavior` (DWORD): Load behavior setting

### COM Registration
See the COM Information tab for detailed COM registry structure.

## Common Management Tasks

**Enabling/Disabling Add-ins**
- Change LoadBehavior between 0 (disabled) and 3 (enabled)
- Restart OneNote to apply changes
- Use for testing different configurations

**Updating Builds**
- Stop OneNote completely
- Build new version
- Re-register if assembly location changed
- Restart OneNote

## Troubleshooting Common Issues

### Add-in Doesn't Appear in OneNote

**Possible Causes:**
- Incorrect registry entries
- Wrong LoadBehavior value
- Missing COM registration
- Insufficient permissions

**Solutions:**
1. Verify registry entries exist
2. Set LoadBehavior to 3
3. Check COM registration
4. Run as Administrator

### Add-in Fails to Load

**Possible Causes:**
- DLL file not found
- Dependency missing
- .NET Framework version mismatch
- Assembly not signed properly

**Solutions:**
1. Verify DLL path is correct
2. Check all dependencies are available
3. Ensure .NET Framework compatibility
4. Re-register the assembly

### Permission Errors

**Symptoms:**
- Cannot register add-in
- Registry access denied
- Installation failures

**Solutions:**
1. Run tools as Administrator
2. Check user account permissions
3. Verify registry key access rights
4. Use proper deployment method

## Best Practices

### Development
- Use strong naming for assemblies
- Test in clean virtual machines
- Handle OneNote API errors gracefully
- Implement proper logging

### Installation
- Use Windows Installer (MSI) packages
- Include all dependencies
- Test installation on target machines
- Provide clear uninstall procedures

### Maintenance
- Monitor add-in performance
- Update for new OneNote versions
- Maintain backwards compatibility
- Document configuration changes

## Security Considerations

- Add-ins run with user privileges
- Validate all external inputs
- Use secure communication protocols
- Follow Microsoft security guidelines
- Test with security software enabled