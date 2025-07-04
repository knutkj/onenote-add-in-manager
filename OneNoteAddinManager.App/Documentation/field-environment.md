# Environment Field

## What is the Environment field?

The **Environment** field indicates whether the add-in is currently configured for development or production use. This classification helps developers and administrators understand the add-in's intended deployment context and manage different versions appropriately.

## Environment Classifications

### Development
- **Meaning**: Add-in is configured for development and testing
- **Indicators**:
  - DLL located in development folders (bin, debug, release)
  - Path contains common development directory patterns
  - Typically registered with `/codebase` flag
  - May point to local build output directories

### Production
- **Meaning**: Add-in is configured for production deployment
- **Indicators**:
  - DLL installed in Program Files or similar system locations
  - Proper installation package deployment
  - May be in Global Assembly Cache (GAC)
  - Follows standard software installation patterns

### Unknown
- **Meaning**: Environment cannot be determined
- **Possible Causes**:
  - Custom installation location
  - DLL path not available
  - Non-standard directory structure
  - Incomplete registration information

## How Environment is Determined

### Path Analysis
The environment classification is based on the DLL path patterns:

**Development Indicators:**
- Contains `\bin\` directory
- Contains `\debug\` directory  
- Contains `\release\` directory
- Located in user profile directories
- Path includes common IDE output folders

**Production Indicators:**
- Located in `Program Files` directories
- System-wide installation paths
- Standard application installation locations
- Lacks development-specific folder patterns

### Registry Analysis
Additional factors considered:
- **CodeBase Registration**: Development add-ins often use file:// URLs
- **GAC Registration**: Production add-ins may use strong names
- **Installation Method**: MSI vs. development registration

## Development Environment Characteristics

### Typical Paths
```
C:\Users\Developer\Source\MyProject\bin\Debug\MyAddin.dll
C:\Dev\Projects\OneNoteTools\bin\Release\OneNoteTools.dll
D:\Source\Addins\MyAddin\bin\x64\Debug\MyAddin.dll
```

### Registration Method
- Often registered using `regasm.exe /codebase`
- Registry CodeBase points to local file system
- May be re-registered frequently during development
- LoadBehavior typically set to 3 for immediate testing

### Development Benefits
- **Easy Updates**: Simply rebuild and restart OneNote
- **Debugging Support**: Can attach debugger to OneNote process
- **Rapid Iteration**: Quick code-test-debug cycles
- **Version Control**: Source code and builds in version control

### Development Considerations
- **Performance**: Debug builds may be slower
- **Dependencies**: Development tools must be installed
- **Stability**: May include experimental or untested features
- **Security**: Local file access required

## Production Environment Characteristics

### Typical Paths
```
C:\Program Files\MyCompany\OneNote Add-ins\MyAddin.dll
C:\Program Files (x86)\Productivity Tools\OneNoteExtensions.dll
```

### Registration Method
- Installed via MSI or similar installer
- May use strong-named assemblies in GAC
- Registry entries created by installer
- Proper uninstall support

### Production Benefits
- **Stability**: Tested and validated releases
- **Performance**: Optimized release builds
- **Security**: Signed assemblies and trusted locations
- **Maintenance**: Proper versioning and update mechanisms

### Production Considerations
- **Updates**: Require installer packages or deployment tools
- **Debugging**: Limited debugging capabilities
- **Distribution**: Requires package management
- **Rollback**: Should support clean uninstall/rollback

## Environment Management

### Development to Production Migration
When moving from development to production:

1. **Build Release Version**
   - Use Release configuration
   - Enable optimizations
   - Sign assemblies if required

2. **Create Installer Package**
   - MSI or similar installation package
   - Include all dependencies
   - Proper registry entries

3. **Unregister Development Version**
   - Clean up development registry entries
   - Remove development DLL references
   - Verify complete cleanup

4. **Install Production Version**
   - Use installer package
   - Verify proper registration
   - Test functionality

### Dual Environment Setup
Some developers maintain both environments:

- **Development**: For active development and testing
- **Production**: For stable feature validation
- **Switching**: Tools can help switch between versions
- **Isolation**: Different registry entries and paths

## Troubleshooting by Environment

### Development Environment Issues
- **Build Failures**: Check Visual Studio configuration
- **Path Changes**: Update registry if project moved
- **Dependency Missing**: Ensure all references available
- **Permission Issues**: Development paths may need special access

### Production Environment Issues
- **Installation Problems**: Use proper installer packages
- **Update Failures**: Ensure clean uninstall before update
- **Path Issues**: Verify installation completed successfully
- **Permission Problems**: Check system-wide access rights

### Mixed Environment Problems
- **Registry Conflicts**: Development and production entries interfering
- **DLL Conflicts**: Multiple versions in different locations
- **Version Confusion**: Unclear which version is active

## Best Practices

### For Developers
1. **Clear Separation**: Keep development and production clearly separated
2. **Consistent Paths**: Use standard development directory structures  
3. **Documentation**: Document environment setup and requirements
4. **Version Control**: Include environment configuration in source control

### For Deployment
1. **Clean Installation**: Remove development versions before production deploy
2. **Verification**: Test environment classification after deployment
3. **Documentation**: Provide clear installation and configuration guides
4. **Rollback Plans**: Maintain ability to revert to previous versions

### For Administration
1. **Environment Auditing**: Regular checks of add-in environments
2. **Policy Enforcement**: Ensure production systems use production versions
3. **Change Management**: Control environment changes through proper processes
4. **Monitoring**: Track environment changes and issues