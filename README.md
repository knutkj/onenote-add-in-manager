# OneNote Add-In Manager

[![Code Formatting Check](https://github.com/knutkj/onenote-add-in-manager/actions/workflows/format-check.yml/badge.svg)](https://github.com/knutkj/onenote-add-in-manager/actions/workflows/format-check.yml)

OneNote Add-In Manager is a Windows Presentation Foundation (WPF) based
application for Windows that explains and manages the Windows Registry entries
required for Microsoft OneNote add-ins to load and work correctly.

This app is designed both as a **learning tool** and a **practical utility** for
developers and IT administrators who need to understand or configure OneNote
add-in registration.

## What It Does

- **Teaches** the basics of the Windows Registry, with a focus on OneNote
  add-ins.
- **Explains** the specific registry keys and values used by OneNote to discover
  and load add-ins.
- **Shows** COM registration concepts and how OneNote uses them to instantiate
  add-in classes.
- **Lets you** browse, view, and edit add-in registry entries on your local
  machine.
- **Includes** a built-in example add-in registration to help you get started
  even if you have no existing add-ins.

## How It Works

OneNote add-ins rely on COM components. For OneNote to find and load these
components, their CLSIDs and settings must be correctly registered in the
Windows Registry.

This app displays:

- Registry paths like
  `HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\Addins`.
- Individual registry keys and their data types (e.g., String, DWORD).
- Common values such as `LoadBehavior` and their meanings.
- COM Class registration under `HKEY_CLASSES_ROOT\CLSID`.

It provides explanations in clear, organized **Markdown** pages so users can
learn as they explore.

## Features

- 📖 Educational content about the Windows Registry and COM registration for
  OneNote add-ins.
- 🔍 View existing OneNote add-in registrations on your machine.
- ✏️ Edit registry entries safely within the app.
- 🧭 Built-in sample add-in to demonstrate typical registration.

## Who Should Use This

- Developers building OneNote add-ins.
- IT administrators deploying or troubleshooting add-ins.
- Anyone wanting to learn about how OneNote integrates with the Windows Registry
  and COM.

## Development

### Code Formatting

This project uses automated code formatting to maintain consistent style:

- **Automatic checks**: GitHub Actions runs `dotnet format --verify-no-changes`
  on all pushes and pull requests
- **Manual formatting**: Run `dotnet format` to format all files manually
- **Formatting verification**: Use `dotnet format --verify-no-changes` to check
  formatting without making changes

Pull requests will be blocked if formatting issues are found. Simply run
`dotnet format` to fix any issues and push the changes.
