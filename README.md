# OneNote Add-In Manager

OneNote Add-In Manager is a Windows Presentation Foundation (WPF) application
for Windows that explains and manages the Windows Registry entries required for
Microsoft OneNote add-ins to load and work correctly.

This app is designed as both a **learning tool** and a **practical utility** for
developers and IT administrators who need to understand, configure, or
troubleshoot OneNote add-in registration.

## What It Does

- **Teaches** the basics of the Windows Registry, focusing on OneNote add-ins.
- **Explains** the specific registry keys and values OneNote uses to discover
  and load add-ins.
- **Shows** COM registration concepts and how OneNote uses them to instantiate
  add-in classes.
- **Lets you** browse, view, and edit add-in registry entries on your local
  machine.
- **Includes** a built-in example to help you get started, even with no existing
  add-ins.

## How It Works

OneNote add-ins rely on COM components. For OneNote to find and load these
components, their CLSIDs and settings must be correctly registered in the
Windows Registry.

This app displays:

- Add-in registration under
  `HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\Addins`.
- COM Class registration under `HKEY_CLASSES_ROOT\CLSID` and
  `HKEY_CLASSES_ROOT\WOW6432Node\CLSID` for 32-bit add-ins on 64-bit Windows.
- Individual registry keys and their data types (e.g., String, DWORD).
- Common values such as `LoadBehavior` and their meanings.

It provides explanations in clear, organized **Markdown** pages so users can
learn as they explore.

## Features

- 📖 Educational content on the Windows Registry and COM registration for
  OneNote add-ins.
- 🔍 View existing OneNote add-in registrations on your machine.
- ✏️ Enable, disable, or unregister add-ins safely within the app.
- 🧹 Find and remove orphaned registry entries for uninstalled add-ins.
- 🚀 Start and stop the OneNote process directly from the app for quick testing.
- 🧭 Built-in sample add-in to demonstrate typical registration.

## How to Use

1.  **Run as Administrator** for full functionality.
2.  Select an add-in from the list to view its details.
3.  Use the buttons to **Enable**, **Disable**, or **Unregister** the selected
    add-in.
4.  Click **Refresh** to reload the add-in list.
5.  Use the **Cleanup** button to find and remove orphaned registry entries.
6.  The **OneNote Control** shows the running status of OneNote and lets you
    start or stop it.

## Who Should Use This

- Developers building OneNote add-ins.
- IT administrators deploying or troubleshooting add-ins.
- Anyone wanting to learn how OneNote integrates with the Windows Registry and
  COM.
