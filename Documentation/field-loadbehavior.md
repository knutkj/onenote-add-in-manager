# LoadBehavior Configuration

## What is LoadBehavior?

**LoadBehavior** is a registry value that controls how and when OneNote loads your add-in. This numeric value determines the add-in's startup behavior, error handling, and user interaction patterns.

## LoadBehavior Values Explained

### LoadBehavior = 0 (Disconnected)
**Status:** Disabled
- **Behavior**: Add-in is not loaded
- **When Used**: Manually disabled by user or administrator
- **User Impact**: Add-in functionality completely unavailable
- **Recovery**: Requires manual re-enabling
- **Registry**: `LoadBehavior = 0 (DWORD)`

### LoadBehavior = 1 (Connected)
**Status:** Enabled (Legacy)
- **Behavior**: Add-in is loaded and connected
- **When Used**: Rarely used in modern Office versions
- **User Impact**: Add-in loads but may not handle errors well
- **Recovery**: Limited automatic error recovery
- **Recommendation**: Avoid for new add-ins

### LoadBehavior = 2 (Bootload)
**Status:** Load at Startup
- **Behavior**: Load when host application starts
- **When Used**: Historical, not recommended for OneNote
- **User Impact**: May slow OneNote startup
- **Recovery**: Limited error handling
- **Modern Alternative**: Use LoadBehavior = 3

### LoadBehavior = 3 (Demand Load - Recommended)
**Status:** Enabled with Error Recovery
- **Behavior**: Load on demand with automatic error handling
- **When Used**: **Recommended for all OneNote add-ins**
- **User Impact**: Best user experience and reliability
- **Recovery**: Automatic retry after failures
- **Benefits**:
  - Graceful error handling
  - Automatic reconnection attempts
  - User notification of issues
  - No permanent disabling

### LoadBehavior = 8 (Demand Load Alternative)
**Status:** Enabled (Alternative Mode)
- **Behavior**: Similar to 3 but different error handling
- **When Used**: Less common, specific scenarios
- **User Impact**: Good reliability but different from standard
- **Recovery**: Different error recovery patterns
- **Recommendation**: Use LoadBehavior = 3 unless specific need

### LoadBehavior = 16 (Connect First Time)
**Status:** Load on First Use
- **Behavior**: Load only when first accessed by user
- **When Used**: Rarely appropriate for OneNote add-ins
- **User Impact**: May confuse users expecting immediate availability
- **Recovery**: Standard error handling
- **Recommendation**: Not suitable for most OneNote scenarios

## How OneNote Uses LoadBehavior

### Startup Process
1. **Registry Scan**: OneNote reads all add-in LoadBehavior values
2. **Load Decision**: Determines which add-ins to load based on LoadBehavior
3. **Error Handling**: Applies appropriate error recovery based on value
4. **User Notification**: May inform user of add-in status changes

### Error Recovery Mechanism
**LoadBehavior = 3 Automatic Recovery:**
- **First Failure**: OneNote temporarily disables add-in
- **Retry Logic**: Attempts to reload on next OneNote start
- **Success**: Restores normal operation
- **Persistent Failure**: May prompt user for action
- **No Permanent Disable**: LoadBehavior remains 3

**Other LoadBehavior Values:**
- **Less Resilient**: May permanently disable on errors
- **Manual Recovery**: User intervention often required
- **Limited Retry**: Fewer automatic recovery attempts

## LoadBehavior State Changes

### User Actions
- **COM Add-ins Dialog**: User can enable/disable add-ins
- **Trust Center**: Security settings may affect loading
- **Group Policy**: Administrative policies override user settings

### Automatic Changes
OneNote may automatically modify LoadBehavior:
- **Load Failures**: Temporary or permanent disabling
- **Security Issues**: Blocking untrusted add-ins
- **Performance Problems**: Disabling problematic add-ins
- **User Requests**: Responding to user disable actions

### Administrative Control
- **Group Policy**: Centralized LoadBehavior management
- **Registry Scripts**: Automated LoadBehavior configuration
- **Deployment Tools**: Installation-time LoadBehavior setup

## Best Practices for LoadBehavior

### Recommended Configuration
**Always Use LoadBehavior = 3 for OneNote Add-ins**

Reasons:
1. **Automatic Error Recovery**: Best resilience to failures
2. **User Experience**: Smooth operation with minimal user intervention
3. **Microsoft Recommendation**: Officially recommended by Microsoft
4. **Future Compatibility**: Best support in future Office versions

### Development Considerations
- **Testing**: Test add-in behavior with LoadBehavior = 3
- **Error Handling**: Implement robust error handling in add-in code
- **Performance**: Ensure add-in loads quickly and efficiently
- **Debugging**: Test recovery scenarios during development

### Deployment Strategy
- **Initial Installation**: Set LoadBehavior = 3 during installation
- **Updates**: Maintain LoadBehavior = 3 across version updates
- **Migration**: Change from other values to 3 when updating
- **Validation**: Verify LoadBehavior = 3 after installation

## Troubleshooting LoadBehavior Issues

### Add-in Not Loading
If LoadBehavior = 3 but add-in doesn't load:
1. **Check DLL Path**: Verify assembly file exists and is accessible
2. **Review Dependencies**: Ensure all required components available
3. **Event Logs**: Check Windows Event Log for load errors
4. **Permissions**: Verify proper file and registry access
5. **OneNote Version**: Confirm compatibility with OneNote version

### LoadBehavior Reset to 0
If OneNote automatically sets LoadBehavior to 0:
1. **Identify Root Cause**: Check Event Log for specific errors
2. **Fix Underlying Issue**: Address DLL, dependency, or code problems
3. **Re-enable**: Set LoadBehavior back to 3
4. **Test Thoroughly**: Verify issue is resolved before deploying

### Performance Impact
If LoadBehavior causes startup delays:
1. **Optimize Initialization**: Minimize add-in startup time
2. **Lazy Loading**: Defer heavy operations until needed
3. **Background Processing**: Move intensive tasks off UI thread
4. **Profile Performance**: Measure and optimize bottlenecks

## LoadBehavior and Security

### Security Policies
- **Administrative Control**: Group Policy can override LoadBehavior
- **Trust Settings**: OneNote trust center affects add-in loading
- **Code Signing**: Signed add-ins may have different LoadBehavior treatment
- **Macro Security**: High security settings may impact add-in loading

### Network Environments
- **Domain Policies**: Centralized LoadBehavior management
- **Roaming Profiles**: LoadBehavior settings follow user across machines
- **Terminal Server**: Special considerations for multi-user environments
- **Virtualization**: Container/VDI specific LoadBehavior behavior

## Advanced LoadBehavior Topics

### Registry Location
LoadBehavior is stored at:
```
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\[AddinName]
Value: LoadBehavior (DWORD)
```

### Programmatic Control
Add-ins can check their own LoadBehavior:
```csharp
// Example: Check current LoadBehavior value
var keyPath = @"SOFTWARE\Microsoft\Office\OneNote\AddIns\MyAddin";
using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
{
    var loadBehavior = key?.GetValue("LoadBehavior");
    // Handle LoadBehavior value
}
```

### Monitoring and Logging
- **Registry Monitoring**: Track LoadBehavior changes
- **Event Logging**: Log LoadBehavior modifications
- **Performance Counters**: Monitor add-in load success rates
- **User Analytics**: Track LoadBehavior impact on user experience

## Migration and Updates

### Changing LoadBehavior
When updating add-ins:
1. **Backup Current Settings**: Save existing LoadBehavior values
2. **Update Strategy**: Plan migration to LoadBehavior = 3
3. **Testing**: Validate new LoadBehavior configuration
4. **Rollback Plan**: Ability to restore previous settings
5. **User Communication**: Inform users of behavior changes

### Legacy Add-in Updates
For older add-ins not using LoadBehavior = 3:
1. **Assessment**: Evaluate current LoadBehavior effectiveness
2. **Code Review**: Ensure add-in compatible with LoadBehavior = 3
3. **Gradual Migration**: Phased rollout of new LoadBehavior
4. **Support Planning**: Prepare for migration-related issues