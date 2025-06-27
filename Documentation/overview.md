# OneNote Add-in Manager

This application helps you manage OneNote COM add-ins by providing a user-friendly interface to view, enable, disable, and register add-ins.

## What are COM Add-ins?

COM add-ins are Component Object Model extensions that integrate with Microsoft Office applications like OneNote. They extend functionality by:

- Adding custom ribbon buttons and menus
- Automating tasks and workflows  
- Integrating with external systems
- Providing custom features not available in the base application

## Registry Integration

OneNote add-ins are registered in the Windows Registry under specific keys that tell OneNote how to load and interact with them. This application provides visibility into these registry entries and allows safe manipulation of add-in settings.

## Safety Features

- **Administrator Detection**: Warns when not running with administrator privileges
- **Orphaned Entry Cleanup**: Identifies and removes broken registry entries
- **Development/Production Switching**: Safely switch between different versions of add-ins
- **Validation**: Prevents invalid operations that could break OneNote integration