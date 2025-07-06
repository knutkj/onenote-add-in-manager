using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace OneNoteAddinManager.App.ViewModels
{
    /// <summary>
    /// ViewModel for the RegistryKeyGrid control
    /// </summary>
    public class RegistryKeyGridViewModel(string registryPath)
    {
        public RegistryKeyGridViewModel() : this(GetDesignTimeRegistryPath())
        {
            if (DesignerProperties.GetIsInDesignMode(new DependencyObject()))
            {
                LoadDesignTimeData();
            }
        }

        public string Description { get; set; } = string.Empty;
        
        public ObservableCollection<RegistryEntry> Entries { get; set; } = [];
        
        public string[] PathSegments => registryPath
            .Split('\\', StringSplitOptions.RemoveEmptyEntries)
            .Select(segment => segment.Trim())
            .ToArray();

        private static string GetDesignTimeRegistryPath()
        {
            if (DesignerProperties.GetIsInDesignMode(new DependencyObject()))
            {
                return @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\MyAddin";
            }
            return string.Empty;
        }

        private void LoadDesignTimeData()
        {
            Description = "This registry key contains the main OneNote add-in configuration. The LoadBehavior value determines how OneNote loads the add-in: 0 = disabled, 2 = load on demand, 3 = load at startup. The FriendlyName appears in OneNote's COM Add-ins dialog.";

            Entries = [
                new() { Name = "(Default)", Type = "REG_SZ", Data = "(value not set)", Icon = "📄", IsKey = false },
                new() { Name = "Description", Type = "REG_SZ", Data = "Extension for OneNote", Icon = "📄", IsKey = false },
                new() { Name = "FriendlyName", Type = "REG_SZ", Data = "MyAddin Demo", Icon = "📄", IsKey = false },
                new() { Name = "LoadBehavior", Type = "REG_DWORD", Data = "0x00000003 (3)", Icon = "📄", IsKey = false }
            ];
        }

    }

    /// <summary>
    /// Represents a single registry entry (key or value)
    /// </summary>
    public class RegistryEntry
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Data { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;
        public bool IsKey { get; set; }

        /// <summary>
        /// Display name with icon for binding
        /// </summary>
        public string DisplayName => $"{Icon} {Name}";
    }
}