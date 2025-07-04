# Status Field

## What is the Status field?

The **Status** field indicates the current operational state of your OneNote add-in. It reflects whether the add-in is enabled, disabled, or experiencing issues that prevent it from loading properly.

## Possible Status Values

### Enabled
- **Meaning**: Add-in is active and loaded in OneNote
- **LoadBehavior**: Typically 3 (Demand Load)
- **Indicators**: 
  - Add-in UI elements are visible in OneNote
  - Add-in functionality is available to users
  - No error conditions detected

### Disabled
- **Meaning**: Add-in is installed but not loaded
- **LoadBehavior**: Typically 0 (Disconnected)
- **Indicators**:
  - Add-in UI elements are hidden
  - Add-in functionality is not available
  - Manually disabled by user or administrator

### Error
- **Meaning**: Add-in failed to load due to an issue
- **LoadBehavior**: May be automatically changed by OneNote
- **Common Causes**:
  - DLL file not found or inaccessible
  - Missing dependencies
  - .NET Framework version mismatch
  - Security restrictions
  - Corrupted installation

### Not Available
- **Meaning**: Status cannot be determined
- **Possible Causes**:
  - Incomplete registry entries
  - Permission issues accessing registry
  - Corrupted add-in registration

## How Status is Determined

### Registry Analysis
The status is calculated by examining:
1. **LoadBehavior Value**: Primary indicator of intended state
2. **DLL Path Accessibility**: Whether the add-in file exists
3. **Registry Completeness**: All required entries are present
4. **Dependency Check**: Required components are available

### OneNote Integration
OneNote may automatically modify LoadBehavior based on:
- Load success/failure history
- User actions (disable/enable)
- Administrative policies
- Security settings

## Status Change Scenarios

### User Actions
- **Enable/Disable**: User toggles add-in in OneNote COM Add-ins dialog
- **Trust Settings**: User modifies security settings affecting add-ins
- **Manual Registry**: Advanced users modify registry directly

### System Events
- **Software Updates**: OneNote or Windows updates affecting compatibility
- **File Moves**: DLL relocated without updating registry
- **Permission Changes**: Security policies affecting add-in access
- **Dependency Updates**: .NET Framework or other component changes

### Error Conditions
- **Load Failures**: Add-in throws exceptions during startup
- **Security Blocks**: Antivirus or security software intervention
- **Compatibility Issues**: OneNote version changes breaking add-in

## Troubleshooting by Status

### Enabled but Not Working
If status shows "Enabled" but add-in doesn't function:
1. Check OneNote version compatibility
2. Verify add-in UI hasn't moved to different ribbon location
3. Look for error messages in Windows Event Log
4. Test with other Office applications if applicable

### Disabled Status
To re-enable a disabled add-in:
1. Use OneNote's COM Add-ins dialog (File → Options → Add-ins)
2. Change LoadBehavior to 3 in registry
3. Use this tool's Enable/Disable button
4. Check for administrative policies preventing enabling

### Error Status
For add-ins showing error status:
1. **Verify DLL Path**: Ensure file exists and is accessible
2. **Check Dependencies**: Confirm .NET Framework version
3. **Review Permissions**: Ensure proper file and registry access
4. **Reinstall Add-in**: Clean uninstall and fresh installation
5. **Event Log**: Check Windows Event Log for specific error details

### Not Available Status
When status cannot be determined:
1. Check registry permissions
2. Verify complete add-in registration
3. Run as Administrator to access all registry areas
4. Re-register the add-in completely

## Status Monitoring

### Automatic Detection
This tool automatically detects status by:
- Reading current LoadBehavior value
- Checking DLL file existence
- Validating registry entry completeness
- Analyzing file accessibility

### Real-Time Updates
Status may change due to:
- OneNote startup/shutdown cycles
- User interactions with add-in settings
- System configuration changes
- File system modifications

## Best Practices

### For Developers
1. **Robust Error Handling**: Prevent add-ins from causing load failures
2. **Graceful Degradation**: Handle missing dependencies elegantly
3. **Status Reporting**: Provide clear feedback about add-in state
4. **Installation Validation**: Verify proper registration during install

### For Administrators
1. **Monitor Status Changes**: Track add-in health across organization
2. **Policy Management**: Use Group Policy for consistent add-in states
3. **Troubleshooting Documentation**: Maintain procedures for common issues
4. **Testing Protocols**: Validate add-ins in controlled environments

### For Users
1. **Check Status Regularly**: Monitor add-in health in this tool
2. **Report Issues**: Provide detailed status information when seeking help
3. **Avoid Manual Changes**: Use proper tools rather than direct registry editing
4. **Backup Settings**: Document working configurations for recovery