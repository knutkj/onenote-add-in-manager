# Add-in Name Field

## What is the Name field?

The **Name** field represents the internal identifier used to register the add-in with OneNote and Windows. This is the technical name that the system uses to identify your add-in in registry entries and internal operations.

## Key Characteristics

### Registry Key Name
- The Name value becomes the registry key name under:
  ```
  HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\[Name]
  ```

### Technical Identifier
- Used as the primary identifier in all registry operations
- Must be unique across all installed add-ins
- Cannot contain special characters or spaces

## Common Patterns

### Assembly-Based Names
Most add-ins use names that match their .NET assembly:
- `MyCompany.OneNoteAddin`
- `ProductivityTools.OneNoteExtensions`
- `DataSync.OneNoteConnector`

### ProgID Format
Some add-ins follow the ProgID naming convention:
- `MyAddin.Connect`
- `CompanyName.ProductName`

## Important Notes

### Cannot Be Changed Easily
Once an add-in is installed, changing the Name requires:
1. Unregistering the old add-in completely
2. Cleaning up all registry entries
3. Re-registering with the new name

### Case Sensitivity
- Names are case-sensitive in the registry
- Best practice: Use consistent capitalization
- Avoid mixing cases unnecessarily

### Relationship to Other Fields
- **Different from Friendly Name**: The Name is technical, while Friendly Name is user-facing
- **Used in COM Registration**: Forms part of CLSID and ProgID entries
- **Referenced in Error Messages**: Appears in system logs and error reports

## Troubleshooting

### Name Conflicts
If you see registration errors, check for:
- Duplicate names with existing add-ins
- Case variations of the same name
- Orphaned registry entries from previous installations

### Invalid Characters
Avoid these characters in add-in names:
- Spaces (use dots or underscores instead)
- Special symbols (/, \, ?, <, >, |, etc.)
- Unicode characters (stick to ASCII)

## Best Practices

1. **Use Reverse Domain Notation**: `com.company.product.addin`
2. **Include Company/Developer Name**: Helps avoid conflicts
3. **Keep It Descriptive**: Should hint at the add-in's purpose
4. **Use Consistent Naming**: Match your assembly and namespace names
5. **Document the Name**: Keep a record for future reference and updates