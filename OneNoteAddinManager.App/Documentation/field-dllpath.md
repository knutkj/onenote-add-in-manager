# DLL Path Field

## What is the DLL Path field?

The **DLL Path** field shows the location of the actual assembly file (.dll) that contains your OneNote add-in code. This path is crucial for OneNote to locate and load your add-in when needed.

## Path Sources and Storage

### Registry Location
The DLL path is typically stored in the COM registration under:
```
HKEY_CLASSES_ROOT\CLSID\{GUID}\InprocServer32
Value: CodeBase (REG_SZ)
```

### Path Formats
The path may appear in different formats:

**File URI Format:**
```
file:///C:/Program Files/MyCompany/MyAddin/MyAddin.dll
```

**Local File Path:**
```
C:\Program Files\MyCompany\MyAddin\MyAddin.dll
```

**UNC Path:**
```
\\ServerName\Share\Applications\MyAddin.dll
```

## Path Types and Contexts

### Common DLL Locations
DLL files may be located in various directories:
```
C:\Users\Developer\Source\MyProject\bin\Debug\MyAddin.dll
C:\Program Files\MyCompany\OneNote Add-ins\MyAddin.dll
C:\Program Files (x86)\Productivity Tools\OneNoteExtensions.dll
D:\Source\Addins\MyAddin\bin\x64\Debug\MyAddin.dll
```

**Characteristics:**
- May be in project output directories during development
- Can be in system-wide installation directories
- Location depends on how the add-in was registered
- Access permissions vary by location

### Alternative Locations
Less common but valid paths:
- **User Profile**: `%USERPROFILE%\Documents\MyAddin\MyAddin.dll`
- **AppData**: `%APPDATA%\MyCompany\MyAddin\MyAddin.dll`
- **Custom Directories**: Organization-specific install locations
- **Network Shares**: Centrally deployed add-ins

## Path Validation and Issues

### File Existence Check
This tool verifies:
- **File exists** at the specified path
- **File is accessible** with current permissions
- **Path format is valid**
- **No invalid characters** in path

### Common Path Problems

**File Not Found:**
- DLL moved or deleted after registration
- Incorrect path in registry
- Network path temporarily unavailable
- Permission denied to access location

**Invalid Path Format:**
- Malformed file:// URI
- Invalid characters in path
- Path length exceeding Windows limits
- Incorrect escape sequences

**Permission Issues:**
- Insufficient rights to access file
- Security software blocking access
- Network authentication problems
- File locked by another process

## Path Discovery Methods

### COM Registry Analysis
Primary method for finding DLL paths:
1. **Locate GUID**: Find add-in's CLSID in registry
2. **InprocServer32 Key**: Navigate to COM server registration
3. **CodeBase Value**: Extract DLL path from CodeBase entry
4. **Assembly Value**: Alternative path in Assembly registration

### Progressive Search
When CodeBase is not available:
1. **GAC Search**: Check Global Assembly Cache
2. **System Paths**: Search standard system directories
3. **Application Paths**: Check OneNote installation directory
4. **Registry Scan**: Look for alternative path references

### File System Validation
After path discovery:
1. **Existence Check**: Verify file exists
2. **Access Test**: Confirm read permissions
3. **Format Validation**: Ensure proper DLL format
4. **Dependency Check**: Verify required dependencies

## Path Considerations

### Build Locations
- **Frequent Changes**: Path may update with each build
- **Build Information**: May include debug symbols and additional files
- **Local Access**: Direct file system access required

### Installation Locations
- **Stable Paths**: Consistent location across installations
- **System Integration**: Proper installer and uninstaller support

## Path Management Best Practices

### For Developers
1. **Consistent Paths**: Use standard build output directories
2. **Source Control**: Don't commit absolute paths
3. **Relative References**: Use project-relative paths where possible
4. **Documentation**: Document expected deployment paths

### For Installation
1. **Standard Locations**: Use conventional installation directories
2. **Path Validation**: Verify paths during installation
3. **Uninstall Support**: Clean up paths during removal
4. **Permission Setup**: Ensure proper access rights

### For Administration
1. **Path Auditing**: Regular verification of DLL paths
2. **Security Policies**: Control allowed installation locations
3. **Change Tracking**: Monitor path modifications
4. **Backup Procedures**: Include DLL files in backup strategies

## Troubleshooting Path Issues

### Path Not Found
When DLL path shows "Not Available":
1. **Check COM Registration**: Verify CLSID entries exist
2. **Registry Permissions**: Ensure read access to registry
3. **Alternative Paths**: Look for Assembly value instead of CodeBase
4. **Reinstall Add-in**: Complete re-registration may be needed

### Access Denied
For permission-related path issues:
1. **Run as Administrator**: Elevated permissions may be required
2. **Security Software**: Check antivirus exclusions
3. **File Permissions**: Verify read access to DLL
4. **Network Issues**: Confirm network path accessibility

### Invalid Path Format
For malformed paths:
1. **Registry Repair**: Fix incorrect registry entries
2. **Re-registration**: Use proper registration tools
3. **Path Encoding**: Check for encoding issues in paths
4. **Length Limits**: Verify path doesn't exceed Windows limits

## Advanced Path Scenarios

### Network Deployment
Considerations for network-based DLL paths:
- **UNC Path Support**: Full network path specification
- **Authentication**: User credentials for network access
- **Availability**: Network path must be consistently available
- **Performance**: Network latency affects add-in loading

### Side-by-Side Deployment
Multiple version support:
- **Version-Specific Paths**: Different paths for different versions
- **Registration Isolation**: Separate registry entries per version
- **Conflict Prevention**: Avoid version conflicts

### Registry-Free COM
Alternative deployment without registry:
- **Manifest Files**: COM registration in manifest
- **Application Directory**: DLL in application-specific location
- **Isolation Benefits**: No system-wide registration required

## Security Considerations

### Path Security
- **Trusted Locations**: Deploy to trusted directories
- **Code Signing**: Sign DLL files for authenticity
- **Access Control**: Limit write access to DLL locations
- **Audit Trail**: Monitor DLL file changes

### Network Security
- **Encrypted Paths**: Use secure network protocols
- **Authentication**: Require proper credentials
- **Integrity Checks**: Verify DLL hasn't been tampered with
- **Fallback Plans**: Handle network unavailability gracefully