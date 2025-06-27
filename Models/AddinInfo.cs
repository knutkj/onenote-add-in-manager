using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;

namespace OneNoteAddinManager.Models
{
    public class AddinInfo : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _friendlyName = string.Empty;
        private string _description = string.Empty;
        private string _dllPath = string.Empty;
        private string _guid = string.Empty;
        private bool _isEnabled;
        private int _loadBehavior;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string FriendlyName
        {
            get => _friendlyName;
            set
            {
                _friendlyName = value;
                OnPropertyChanged(nameof(FriendlyName));
            }
        }

        public string Description
        {
            get => _description;
            set
            {
                _description = value;
                OnPropertyChanged(nameof(Description));
            }
        }

        public string DllPath
        {
            get => _dllPath;
            set
            {
                _dllPath = value;
                OnPropertyChanged(nameof(DllPath));
            }
        }

        public string Guid
        {
            get => _guid;
            set
            {
                _guid = value;
                OnPropertyChanged(nameof(Guid));
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                _isEnabled = value;
                OnPropertyChanged(nameof(IsEnabled));
            }
        }


        public int LoadBehavior
        {
            get => _loadBehavior;
            set
            {
                _loadBehavior = value;
                OnPropertyChanged(nameof(LoadBehavior));
            }
        }

        public string Status => IsEnabled ? "Enabled" : "Disabled";

        // Registry information for educational purposes
        public string OfficeAddinRegistryPath => $@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\{Name}";
        public string AppIdRegistryPath => !string.IsNullOrEmpty(Guid) ? $@"HKEY_CLASSES_ROOT\AppID\{Guid}" : "Not Available";
        public string ClsidRegistryPath => !string.IsNullOrEmpty(Guid) ? $@"HKEY_CLASSES_ROOT\CLSID\{Guid}" : "Not Available";
        public string ProgIdRegistryPath => $@"HKEY_CLASSES_ROOT\{Name}";
        
        // COM information
        public string ComClassName => $"{Name}.AddIn";
        public string ThreadingModel => "Both";
        public string RuntimeVersion => "v4.0.30319";
        public string InprocServer => "mscoree.dll";
        
        // LoadBehavior explanations
        public string LoadBehaviorExplanation => LoadBehavior switch
        {
            0 => "Disabled - Add-in is not loaded",
            1 => "Loaded once - Add-in is loaded only on demand",
            2 => "Loaded at startup - Add-in is loaded when the application starts",
            3 => "Loaded at startup and on demand - Add-in is loaded at startup and remains loaded",
            8 => "Connected on demand - Add-in is loaded only when requested by the user",
            9 => "Connected at startup - Add-in is loaded at startup and connected",
            16 => "Connected with first document - Add-in is loaded when the first document is opened",
            _ => $"Unknown behavior ({LoadBehavior})"
        };

        // Registry keys that should exist for this add-in
        public List<RegistryKeyInfo> RegistryKeys
        {
            get
            {
                var keys = new List<RegistryKeyInfo>
                {
                    new RegistryKeyInfo
                    {
                        Path = OfficeAddinRegistryPath,
                        Purpose = "Office Add-in Registration",
                        Description = "Registers the add-in with OneNote. Contains LoadBehavior, FriendlyName, and Description.",
                        Values = new Dictionary<string, string>
                        {
                            { "LoadBehavior", $"{LoadBehavior} ({LoadBehaviorExplanation})" },
                            { "FriendlyName", FriendlyName ?? "Not set" },
                            { "Description", Description ?? "Not set" }
                        }
                    }
                };

                if (!string.IsNullOrEmpty(Guid))
                {
                    keys.Add(new RegistryKeyInfo
                    {
                        Path = AppIdRegistryPath,
                        Purpose = "AppID Registration",
                        Description = "Registers the application ID for COM activation. DllSurrogate enables out-of-process activation.",
                        Values = new Dictionary<string, string>
                        {
                            { "DllSurrogate", "(empty string) - Enables out-of-process activation" }
                        }
                    });

                    keys.Add(new RegistryKeyInfo
                    {
                        Path = ClsidRegistryPath,
                        Purpose = "CLSID Registration",
                        Description = "Registers the COM class ID. This is the main COM registration for the add-in.",
                        Values = new Dictionary<string, string>
                        {
                            { "(Default)", ComClassName },
                            { "AppID", Guid }
                        }
                    });

                    keys.Add(new RegistryKeyInfo
                    {
                        Path = $@"{ClsidRegistryPath}\InprocServer32",
                        Purpose = "In-Process Server Registration",
                        Description = "Specifies how the COM object is loaded. Points to .NET runtime and the actual DLL.",
                        Values = new Dictionary<string, string>
                        {
                            { "(Default)", InprocServer + " - .NET Runtime host" },
                            { "ThreadingModel", ThreadingModel + " - Supports both STA and MTA" },
                            { "CodeBase", DllPath ?? "Not set" },
                            { "Class", ComClassName },
                            { "RuntimeVersion", RuntimeVersion + " - .NET Framework version" }
                        }
                    });

                    keys.Add(new RegistryKeyInfo
                    {
                        Path = $@"{ClsidRegistryPath}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}}",
                        Purpose = ".NET Category Registration",
                        Description = "Identifies this as a .NET component. Required for .NET COM interop.",
                        Values = new Dictionary<string, string>
                        {
                            { "Category", ".NET Category" }
                        }
                    });

                    keys.Add(new RegistryKeyInfo
                    {
                        Path = $@"{ClsidRegistryPath}\ProgID",
                        Purpose = "Programmatic Identifier",
                        Description = "Human-readable name for the COM class. Used for CreateObject calls.",
                        Values = new Dictionary<string, string>
                        {
                            { "(Default)", Name }
                        }
                    });
                }

                keys.Add(new RegistryKeyInfo
                {
                    Path = ProgIdRegistryPath,
                    Purpose = "ProgID Class Registration",
                    Description = "Maps the friendly name to the CLSID. Allows creation by name instead of GUID.",
                    Values = new Dictionary<string, string>
                    {
                        { "(Default)", ComClassName },
                        { "CLSID", Guid ?? "Not set" }
                    }
                });

                return keys;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }    public class RegistryKeyInfo
    {
        public string Path { get; set; } = string.Empty;
        public string Purpose { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>();
        
        public string ValuesText => string.Join("; ", Values.Select(kv => $"{kv.Key}: {kv.Value}"));
    }
}
