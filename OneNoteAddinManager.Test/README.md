# OneNote Add-in Manager - Unit Tests

This directory contains comprehensive unit tests for the OneNote Add-in Manager
application using MSTest and System.IO.Abstractions frameworks.

## Testing Framework

- **MSTest**: The de facto testing framework for .NET applications, especially
  XAML/WPF applications
- **System.IO.Abstractions**: Industry-standard file system abstraction for
  testable code
- **System.IO.Abstractions.TestingHelpers**: MockFileSystem for isolated testing

## Key Testing Patterns

### Clean Dependency Injection with System.IO.Abstractions

```csharp
var mockFileSystem = new MockFileSystem();
var viewModel = new DllInfoViewModel(testPath, mockFileSystem);
```

### File System Mocking

```csharp
mockFileSystem.AddFile(testPath, new MockFileData("test content"));
Assert.IsTrue(mockFileSystem.File.Exists(testPath));
```

## Running Tests

### Command Line

```bash
dotnet test OneNoteAddinManager.Test
```

### Visual Studio

- Use Test Explorer to run individual or all tests
- Right-click on test methods/classes for context menu options
