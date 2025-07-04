# Troubleshooting Add-ins

## Common Issues

### Add-in Not Loading
**Symptoms**: Add-in appears in list but OneNote doesn't show its features

**Possible Causes**:
- Invalid DLL path - file moved or deleted
- Missing dependencies - required libraries not installed
- Architecture mismatch - 32-bit vs 64-bit compatibility
- .NET Framework version incompatibility

**Solutions**:
1. Verify DLL exists at specified path
2. Check if add-in targets correct .NET Framework version
3. Ensure all dependencies are available in GAC or same directory
4. Test with Process Monitor to identify file access issues

### Registry Permission Errors  
**Symptoms**: "Access denied" errors when enabling/disabling add-ins

**Solutions**:
- Run application as Administrator
- Check user permissions on registry keys
- Use machine-wide registration if user-specific fails

### Orphaned Registry Entries
**Symptoms**: Add-ins listed that don't actually exist

**Causes**:
- Incomplete uninstallation
- Manual deletion of DLL files
- Failed installation/registration

**Solutions**:
- Use "Cleanup" button to remove orphaned entries
- Manually verify registry entries match actual files
- Re-install add-in properly if needed

## Development Mode Issues

### Switching Between Versions
When switching between development and production versions:

1. Ensure OneNote is closed
2. Verify target DLL exists and is accessible  
3. Check that new DLL has same GUID as registered version
4. Restart OneNote to load new version

### Loading Development Builds
- Development DLLs may have different dependencies
- Debug assemblies require Visual Studio runtime
- Ensure development environment matches target system