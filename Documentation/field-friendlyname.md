# Friendly Name Field

## What is the Friendly Name field?

The **Friendly Name** is the user-facing display name for your OneNote add-in. This is what users see in OneNote's add-in management interfaces, error messages, and any user-facing lists or dialogs.

## Key Characteristics

### User-Visible Name
- Displayed in OneNote's COM Add-ins dialog
- Shown in Office Trust Center add-in lists
- Appears in error messages and notifications
- Used in administrative tools and reports

### Localization Support
- Can be localized for different languages
- Supports Unicode characters and special symbols
- No strict character restrictions like the technical Name field

## Registry Storage

The Friendly Name is stored in the registry at:
```
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\[AddinName]
Value: FriendlyName (REG_SZ)
```

## Common Examples

### Product-Focused Names
- "OneNote Web Clipper"
- "Meeting Notes Organizer"
- "Task Manager Integration"
- "PDF Export Tools"

### Company-Branded Names
- "Contoso OneNote Extensions"
- "Acme Productivity Suite"
- "Enterprise Data Connector"

### Descriptive Names
- "Quick Text Formatter"
- "Advanced Search Tools"
- "Auto-Backup Manager"
- "Template Library"

## Best Practices

### Clear and Descriptive
- Immediately convey the add-in's purpose
- Use terms that users will understand
- Avoid technical jargon or acronyms
- Keep it concise but informative

### Professional Naming
- Use proper capitalization
- Avoid ALL CAPS or excessive punctuation
- Consider including your company name
- Make it memorable and searchable

### Length Considerations
- Aim for 20-40 characters
- Consider how it displays in narrow UI spaces
- Test in different OneNote interface contexts
- Ensure readability at various font sizes

## Relationship to Other Fields

### vs. Technical Name
- **Technical Name**: `MyCompany.OneNoteAddin.Connect`
- **Friendly Name**: "MyCompany OneNote Productivity Tools"

### vs. Assembly Name
- **Assembly**: `OneNoteExtensions.dll`
- **Friendly Name**: "OneNote Extensions Suite"

### vs. Product Name
- May match your product's marketing name
- Should be consistent with other product branding
- Can include version information if appropriate

## Localization Considerations

### Multi-Language Support
- Can store different friendly names for different locales
- Registry supports Unicode strings
- Consider cultural naming conventions
- Test with longer translations (German, etc.)

### Character Support
- Full Unicode support available
- Can include accented characters
- Supports symbols and emojis (use sparingly)
- Right-to-left language compatibility

## Troubleshooting

### Missing Friendly Name
If no Friendly Name is set:
- OneNote may display the technical Name instead
- Some interfaces might show "(Unknown Add-in)"
- Users may have difficulty identifying the add-in

### Display Issues
- Check for extremely long names causing truncation
- Verify special characters display correctly
- Test in different Windows display scaling settings
- Ensure readability in high-contrast themes

## Updating Friendly Names

### Registry Update
Can be changed by modifying the registry value:
1. Navigate to the add-in's registry key
2. Edit the "FriendlyName" string value
3. Restart OneNote to see changes

### Programmatic Update
Add-ins can update their own friendly name:
```csharp
// Example: Update friendly name programmatically
Registry.SetValue(registryPath, "FriendlyName", "New Display Name");
```

### Installation Updates
- Installer packages should handle friendly name updates
- Consider user preferences for customized names
- Maintain consistency across version updates