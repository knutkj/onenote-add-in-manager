# OneNote Add-in Manager

A WPF application for managing OneNote add-ins on Windows. This tool helps
developers and users to easily register, enable/disable, and manage OneNote
add-ins through a user-friendly interface.

## Features

- **View Installed Add-ins**: See all currently registered OneNote add-ins with
  their status and location
- **Enable/Disable Add-ins**: Toggle add-ins on or off without manual registry
  editing
- **Register New Add-ins**: Register new add-ins from DLL files for development
  or testing
- **Cleanup Orphaned Entries**: Remove registry entries for add-ins whose files
  no longer exist

## Usage

### Managing Add-ins

1. **View Add-ins**: The main grid shows all registered OneNote add-ins
2. **Enable/Disable**: Click the Enable/Disable button next to any add-in
3. **Register New Add-in**:
   - Click "..." to browse for a DLL file
   - Click "Register Add-in" to register it
4. **Cleanup**: Click "Cleanup Orphaned" to remove entries for missing DLL files

## Registry Locations

The application manages these registry keys:

- `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\` - Add-in
  registration
- `HKEY_CLASSES_ROOT\AppID\{GUID}` - AppID entries
- `HKEY_CLASSES_ROOT\CLSID\{GUID}` - CLSID entries
- `HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{GUID}` - 32-bit CLSID entries (if
  applicable)

## Building from Source

```bash
# Restore dependencies
dotnet restore

# Build the application
dotnet build

# Run the application
dotnet run
```

## License

This project is provided as-is for educational and development purposes.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve the
application.
