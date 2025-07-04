# Welcome to OneNote Add-In Manager

This application helps you understand how Microsoft OneNote add-ins are
registered and managed using the Windows Registry and COM (Component Object
Model) technology. It also provides tools to view and edit add-in registration
information.

## What is the Windows Registry?

The Windows Registry is a central database that stores configuration settings
and options for the Windows operating system and many applications. It organizes
information in a hierarchy of keys and values, allowing programs like OneNote to
find details about installed components, such as add-ins. For OneNote add-ins,
the registry holds information about how and when each add-in should be loaded.

## What is COM?

COM (Component Object Model) is a Microsoft technology that allows software
components to communicate and interact, regardless of the language they were
written in. OneNote add-ins are implemented as COM components, which means they
must be registered in Windows so that OneNote can find and load them. The
registry contains the information needed for OneNote to create and use these COM
objects.

## Purpose of This App

This app helps you:

- Identify which registry keys are relevant to OneNote add-ins
- Understand the values and data types these keys use
- See how COM registration enables OneNote to find and load add-ins
- Browse registered add-ins for OneNote
- View or edit their registry entries
- Learn the structure and purpose of each setting

## Getting Started

If you don't have any registered add-ins yet, you can start by exploring the
built-in example add-in included with this app. This example shows typical
registry keys and values used for OneNote add-ins.

To continue, select an add-in from the list on the left. The details pane will
show the registry information for that add-in along with explanations and
editing options.
