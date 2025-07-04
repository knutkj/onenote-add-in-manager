using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using OneNoteAddinManager.Lib.Models;
using OneNoteAddinManager.Lib.Services;

namespace OneNoteAddinManager.App.ViewModels
{
    /// <summary>
    /// ViewModel for Add-in information - contains ALL presentation logic
    /// </summary>
    public class AddInInfoViewModel : INotifyPropertyChanged, IDisposable
    {
        private readonly AddinInfo? _addinInfo;
        private readonly string? _registryPath;

        public AddInInfoViewModel(string? registryPath)
        {
            _registryPath = registryPath;
            
            // Load add-in information from registry path if provided
            if (!string.IsNullOrEmpty(registryPath))
            {
                try
                {
                    // Extract add-in name from registry path and load its info
                    var addinName = ExtractAddinNameFromPath(registryPath);
                    if (!string.IsNullOrEmpty(addinName))
                    {
                        var registryManager = new RegistryManager();
                        var addins = registryManager.GetInstalledAddins();
                        _addinInfo = addins.Find(a => a.Name.Equals(addinName, StringComparison.OrdinalIgnoreCase));
                    }
                }
                catch (Exception)
                {
                    // If we can't load the add-in info, _addinInfo remains null
                }
            }
            
            // Initialize commands
            OpenDllFolderCommand = new RelayCommand(OpenDllFolder, () => CanOpenDllFolder);
            OpenRegistryEditorCommand = new RelayCommand(OpenRegistryEditor, () => HasValidRegistryPath);
            CopyRegistryPathCommand = new RelayCommand(CopyRegistryPath, () => HasValidRegistryPath);
            ShowInfoCommand = new RelayCommand<string>(ShowInfo);
        }

        private string? ExtractAddinNameFromPath(string registryPath)
        {
            // Extract add-in name from path like "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\MyAddin"
            var parts = registryPath.Split('\\');
            return parts.Length > 0 ? parts[parts.Length - 1] : null;
        }

        // Properties for display
        public string Name => _addinInfo?.Name ?? "Not Available";
        public string FriendlyName => _addinInfo?.FriendlyName ?? "Not Available";
        public string Status => _addinInfo?.Status ?? "Not Available";
        public string Guid => _addinInfo?.Guid ?? "Not Available";
        public string DllPath => _addinInfo?.DllPath ?? "Not Available";
        public string RegistryPath => _registryPath ?? "Not Available";

        // Command enablement properties
        public bool CanOpenDllFolder
        {
            get
            {
                if (_addinInfo?.DllPath == null) return false;
                try
                {
                    var directory = Path.GetDirectoryName(_addinInfo.DllPath);
                    return !string.IsNullOrEmpty(directory) && Directory.Exists(directory);
                }
                catch
                {
                    return false;
                }
            }
        }

        public bool HasValidRegistryPath => !string.IsNullOrEmpty(_registryPath);

        // Commands
        public ICommand OpenDllFolderCommand { get; }
        public ICommand OpenRegistryEditorCommand { get; }
        public ICommand CopyRegistryPathCommand { get; }
        public ICommand ShowInfoCommand { get; }

        // Command handlers
        private void OpenDllFolder()
        {
            if (_addinInfo?.DllPath == null) return;

            try
            {
                var directory = Path.GetDirectoryName(_addinInfo.DllPath);
                if (!Directory.Exists(directory))
                {
                    MessageBox.Show($"The DLL folder does not exist:\n{directory}", "Folder Not Found",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{directory}\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening folder: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenRegistryEditor()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "regedit.exe",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening Registry Editor: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyRegistryPath()
        {
            if (string.IsNullOrEmpty(_registryPath)) return;

            try
            {
                Clipboard.SetText(_registryPath);
                MessageBox.Show($"Copied to clipboard: {_registryPath}", "Copied",
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying to clipboard: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ShowInfo(string fieldName)
        {
            // Trigger event for parent to handle documentation loading
            InfoRequested?.Invoke(this, new InfoRequestedEventArgs(fieldName));
        }

        // Event for requesting documentation
        public event EventHandler<InfoRequestedEventArgs>? InfoRequested;

        public void Dispose()
        {
            // No resources to dispose in this implementation
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Event args for info requests
    public class InfoRequestedEventArgs : EventArgs
    {
        public string FieldName { get; }

        public InfoRequestedEventArgs(string fieldName)
        {
            FieldName = fieldName;
        }
    }

    // Simple RelayCommand implementation
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) => _execute();
    }

    // RelayCommand with parameter
    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool>? _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter)
        {
            if (parameter is T typedParameter)
                return _canExecute?.Invoke(typedParameter) ?? true;
            return false;
        }

        public void Execute(object? parameter)
        {
            if (parameter is T typedParameter)
                _execute(typedParameter);
        }
    }
}