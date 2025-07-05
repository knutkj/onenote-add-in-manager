# OneNote Add-In Manager

[![CI Pipeline](https://github.com/knutkj/onenote-add-in-manager/actions/workflows/ci.yml/badge.svg)](https://github.com/knutkj/onenote-add-in-manager/actions/workflows/ci.yml)

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

### CI Pipeline

This project uses automated CI/CD to ensure code quality:

- **Code formatting**: Enforces consistent style with
  `dotnet format --verify-no-changes`
- **Solution build**: Compiles the entire solution to catch build errors
- **Unit tests**: Runs all tests to ensure functionality works correctly
- **Manual commands**:
  - `dotnet format` - Format all files manually
  - `dotnet build` - Build the solution locally
  - `dotnet test` - Run unit tests locally

Pull requests will be blocked if any CI checks fail (formatting, build, or
tests). Fix issues locally and push the changes to pass all checks.
