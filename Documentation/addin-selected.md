# Add-in Information

You have selected an add-in! The main panel shows detailed registry information.

## What You're Seeing

### Basic Information
- **Name**: Internal COM class name
- **Friendly Name**: Display name shown to users
- **Status**: Current operational state (Enabled/Disabled)
- **Environment**: Development or Production mode
- **GUID**: Unique identifier for COM registration
- **DLL Path**: Location of the add-in's compiled library

### LoadBehavior Configuration
This controls when and how OneNote loads your add-in:

- **0**: Disabled - Don't load
- **1**: Load at startup only  
- **2**: Load on demand
- **3**: Load at startup (most common)
- **8**: Connect on next startup only
- **9**: Load at startup and stay connected
- **16**: Connect first time, then load on demand

### Registry Keys
Shows the actual Windows Registry entries that OneNote uses to find and load your add-in. These keys contain the configuration data that tells OneNote:

- Where to find the DLL file
- What COM class to instantiate
- How to load the add-in (LoadBehavior)
- User-specific vs machine-wide settings

## Available Actions

- **Enable/Disable**: Toggle add-in without uninstalling
- **Switch to Dev/Prod**: Change between development and production versions
- **Unregister**: Completely remove all registry entries

## Troubleshooting Tips

If your add-in isn't working:

1. **Check the DLL path** - Ensure the file exists at the specified location
2. **Verify LoadBehavior** - Should be 3 for most add-ins
3. **Run as Administrator** - Required for machine-wide registry changes
4. **Restart OneNote** - Required after making changes