using System.ComponentModel;
using System.IO;
using System.IO.Abstractions;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Reflection;
using System.Diagnostics;

namespace OneNoteAddinManager.App.ViewModels
{
    /// <summary>
    /// ViewModel for DLL information - contains ALL presentation logic
    /// </summary>
    public class DllInfoViewModel : INotifyPropertyChanged
    {
        private readonly IFileSystem _fileSystem;
        private readonly IFileInfo? _assemblyFile;
        private readonly string? _assemblyPath;
        private System.Windows.Threading.DispatcherTimer? _lockStatusTimer;

        // ViewModel-specific state (not part of the immutable model)
        private bool _isLocked = false;
        private string _lockDetails = string.Empty;

        // Assembly information cache
        private Assembly? _loadedAssembly;
        private FileVersionInfo? _versionInfo;
        private bool _assemblyLoadAttempted = false;

        public DllInfoViewModel(string? assemblyPath, IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
            _assemblyPath = assemblyPath;

            // Create FileInfo only if we have a valid path
            if (assemblyPath != null)
            {
                _assemblyFile = _fileSystem.FileInfo.New(assemblyPath);

                // Initialize lock status
                if (_assemblyFile.Exists)
                {
                    var (isLocked, details) = CheckFileLock(assemblyPath);
                    _isLocked = isLocked;
                    _lockDetails = details;

                    // Load assembly information on initialization
                    LoadAssemblyInformation();

                    // Notify UI that assembly properties may have changed
                    NotifyAssemblyPropertiesChanged();
                }
            }

            // Set up periodic lock status checking
            _lockStatusTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            _lockStatusTimer.Tick += (s, e) => UpdateLockStatusIfNeeded();
            _lockStatusTimer.Start();
        }

        // Read-only properties based on FileInfo

        // ViewModel-managed lock state (separate from immutable model)
        public bool IsLocked
        {
            get => _isLocked;
            private set
            {
                if (_isLocked != value)
                {
                    _isLocked = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(FileLockedText));
                    OnPropertyChanged(nameof(FileLockedBrush));
                }
            }
        }

        public string LockDetails
        {
            get => _lockDetails;
            private set
            {
                if (_lockDetails != value)
                {
                    _lockDetails = value;
                    OnPropertyChanged();
                }
            }
        }

        // Assembly information properties
        public string AssemblyVersionText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetAssemblyVersion();
            }
        }

        public string FileVersionText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetFileVersion();
            }
        }

        public string TargetFrameworkText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetTargetFramework();
            }
        }

        public string ImplementedInterfacesText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetImplementedInterfaces();
            }
        }

        public string AssemblyArchitectureText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetAssemblyArchitecture();
            }
        }

        public string CompanyText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetCompany();
            }
        }

        public string ComVisibleText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetComVisible();
            }
        }

        public string GuidText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetGuid();
            }
        }

        public string ProgIdText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return GetProgId();
            }
        }

        // Computed properties for presentation
        public string FileExistsText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "❌ No";
                return "✓ Yes";
            }
        }

        public Brush FileExistsBrush
        {
            get
            {
                if (_assemblyFile?.Exists != true) return Brushes.Red;
                return Brushes.Green;
            }
        }

        public string FileSizeText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return FormatFileSize(_assemblyFile.Length);
            }
        }

        public string FileLockedText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return IsLocked ? "🔒 Yes" : "🔓 No";
            }
        }

        public Brush FileLockedBrush
        {
            get
            {
                if (_assemblyFile?.Exists != true) return Brushes.Gray;
                return IsLocked ? Brushes.Red : Brushes.Green;
            }
        }

        public string LastModifiedText
        {
            get
            {
                if (_assemblyFile?.Exists != true) return "N/A";
                return _assemblyFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
        }

        private void NotifyAssemblyPropertiesChanged()
        {
            OnPropertyChanged(nameof(AssemblyVersionText));
            OnPropertyChanged(nameof(FileVersionText));
            OnPropertyChanged(nameof(TargetFrameworkText));
            OnPropertyChanged(nameof(ImplementedInterfacesText));
            OnPropertyChanged(nameof(AssemblyArchitectureText));
            OnPropertyChanged(nameof(CompanyText));
            OnPropertyChanged(nameof(ComVisibleText));
            OnPropertyChanged(nameof(GuidText));
            OnPropertyChanged(nameof(ProgIdText));
        }

        private void LoadAssemblyInformation()
        {
            if (_assemblyLoadAttempted || _assemblyPath == null || _assemblyFile?.Exists != true)
                return;

            _assemblyLoadAttempted = true;

            try
            {
                // Load version information from file
                _versionInfo = FileVersionInfo.GetVersionInfo(_assemblyPath);

                // Try to load the assembly for reflection
                try
                {
                    _loadedAssembly = Assembly.LoadFrom(_assemblyPath);
                }
                catch
                {
                    // If we can't load the assembly, we'll still have file version info
                    _loadedAssembly = null;
                }
            }
            catch
            {
                // If anything fails, leave both null
                _versionInfo = null;
                _loadedAssembly = null;
            }
        }

        private string GetAssemblyVersion()
        {
            try
            {
                if (_loadedAssembly != null)
                {
                    var version = _loadedAssembly.GetName().Version;
                    return version?.ToString() ?? "Unknown";
                }
                return "Unable to load assembly";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetFileVersion()
        {
            try
            {
                if (_versionInfo != null)
                {
                    return _versionInfo.FileVersion ?? "Unknown";
                }
                return "Unable to read version";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetTargetFramework()
        {
            try
            {
                if (_loadedAssembly != null)
                {
                    var targetFrameworkAttribute = _loadedAssembly
                        .GetCustomAttributes(typeof(System.Runtime.Versioning.TargetFrameworkAttribute), false)
                        .FirstOrDefault() as System.Runtime.Versioning.TargetFrameworkAttribute;

                    if (targetFrameworkAttribute != null)
                    {
                        return targetFrameworkAttribute.FrameworkName;
                    }

                    // Fallback to runtime version
                    return $".NET Framework {_loadedAssembly.ImageRuntimeVersion}";
                }
                return "Unable to load assembly";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetImplementedInterfaces()
        {
            try
            {
                if (_loadedAssembly != null)
                {
                    var types = _loadedAssembly.GetTypes();
                    var relevantInterfaces = new List<string>();
                    var commonInterfaces = new List<string>();

                    foreach (var type in types)
                    {
                        if (type.IsClass && type.IsPublic)
                        {
                            var interfaces = type.GetInterfaces();
                            foreach (var iface in interfaces)
                            {
                                var interfaceName = iface.Name;
                                var fullName = iface.FullName ?? interfaceName;

                                // Look for Office/OneNote specific interfaces
                                if (interfaceName.Contains("Extensibility") ||
                                    interfaceName.Contains("Office") ||
                                    interfaceName.Contains("OneNote") ||
                                    interfaceName.Contains("IDTExtensibility") ||
                                    interfaceName.Contains("IRibbonExtensibility") ||
                                    interfaceName.Contains("ICustomTaskPaneConsumer"))
                                {
                                    if (!relevantInterfaces.Contains(interfaceName))
                                    {
                                        relevantInterfaces.Add(interfaceName);
                                    }
                                }
                                // Track common COM interfaces
                                else if (interfaceName.Contains("IDisposable") ||
                                        interfaceName.Contains("IUnknown") ||
                                        interfaceName.Contains("IDispatch"))
                                {
                                    if (!commonInterfaces.Contains(interfaceName))
                                    {
                                        commonInterfaces.Add(interfaceName);
                                    }
                                }
                            }
                        }
                    }

                    if (relevantInterfaces.Any())
                    {
                        return $"Office/OneNote: {string.Join(", ", relevantInterfaces)}";
                    }
                    else if (commonInterfaces.Any())
                    {
                        return $"COM: {string.Join(", ", commonInterfaces)}";
                    }

                    return "No relevant interfaces found";
                }
                return "Unable to load assembly";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetAssemblyArchitecture()
        {
            try
            {
                if (_loadedAssembly != null)
                {
                    // Use Module.GetPEKind for more reliable architecture detection
                    var module = _loadedAssembly.GetModules()[0];
                    module.GetPEKind(out var peKind, out var machine);

                    return (peKind, machine) switch
                    {
                        (PortableExecutableKinds.ILOnly, ImageFileMachine.I386) => "AnyCPU (MSIL)",
                        (PortableExecutableKinds.Required32Bit, ImageFileMachine.I386) => "x86 (32-bit)",
                        (PortableExecutableKinds.ILOnly, ImageFileMachine.AMD64) => "x64 (64-bit)",
                        (PortableExecutableKinds.ILOnly, ImageFileMachine.IA64) => "IA64 (Itanium)",
                        (PortableExecutableKinds.ILOnly, ImageFileMachine.ARM) => "ARM",
                        _ => $"Unknown ({peKind}, {machine})"
                    };
                }
                return "Unable to load assembly";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetCompany()
        {
            try
            {
                if (_versionInfo != null)
                {
                    return !string.IsNullOrEmpty(_versionInfo.CompanyName) ? _versionInfo.CompanyName : "Not specified";
                }
                return "Unable to read version info";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetComVisible()
        {
            try
            {
                if (_loadedAssembly == null)
                    return "Unable to load assembly";

                var comVisibleValues = new List<string>();

                // Check assembly-level ComVisible attribute
                var assemblyComVisible = _loadedAssembly.GetCustomAttributes(typeof(System.Runtime.InteropServices.ComVisibleAttribute), false)
                    .FirstOrDefault() as System.Runtime.InteropServices.ComVisibleAttribute;

                if (assemblyComVisible != null)
                {
                    comVisibleValues.Add($"Assembly: {assemblyComVisible.Value}");
                }

                // Check types for ComVisible attributes
                var types = _loadedAssembly.GetTypes().Where(t => t.IsPublic);
                var typeComVisible = new List<string>();

                foreach (var type in types)
                {
                    var typeComVisibleAttr = type.GetCustomAttributes(typeof(System.Runtime.InteropServices.ComVisibleAttribute), false)
                        .FirstOrDefault() as System.Runtime.InteropServices.ComVisibleAttribute;

                    if (typeComVisibleAttr != null)
                    {
                        typeComVisible.Add($"{type.Name}: {typeComVisibleAttr.Value}");
                    }
                }

                if (typeComVisible.Any())
                {
                    comVisibleValues.Add($"Types: {string.Join(", ", typeComVisible)}");
                }

                return comVisibleValues.Any() ? string.Join("; ", comVisibleValues) : "Not specified";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetGuid()
        {
            try
            {
                if (_loadedAssembly == null)
                    return "Unable to load assembly";

                var guids = new List<string>();

                // Check assembly-level GUID attribute
                var assemblyGuid = _loadedAssembly.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false)
                    .FirstOrDefault() as System.Runtime.InteropServices.GuidAttribute;

                if (assemblyGuid != null)
                {
                    guids.Add($"Assembly: {assemblyGuid.Value}");
                }

                // Check types for GUID attributes (interfaces and classes)
                var types = _loadedAssembly.GetTypes().Where(t => t.IsPublic && (t.IsInterface || t.IsClass));
                var typeGuids = new List<string>();

                foreach (var type in types)
                {
                    var typeGuidAttr = type.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false)
                        .FirstOrDefault() as System.Runtime.InteropServices.GuidAttribute;

                    if (typeGuidAttr != null)
                    {
                        typeGuids.Add($"{type.Name}: {typeGuidAttr.Value}");
                    }
                }

                if (typeGuids.Any())
                {
                    guids.Add($"Types: {string.Join(", ", typeGuids)}");
                }

                return guids.Any() ? string.Join("; ", guids) : "Not specified";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string GetProgId()
        {
            try
            {
                if (_loadedAssembly == null)
                    return "Unable to load assembly";

                var progIds = new List<string>();

                // Check types for ProgId attributes (typically on classes)
                var types = _loadedAssembly.GetTypes().Where(t => t.IsPublic && t.IsClass);

                foreach (var type in types)
                {
                    var progIdAttr = type.GetCustomAttributes(typeof(System.Runtime.InteropServices.ProgIdAttribute), false)
                        .FirstOrDefault() as System.Runtime.InteropServices.ProgIdAttribute;

                    if (progIdAttr != null)
                    {
                        progIds.Add($"{type.Name}: {progIdAttr.Value}");
                    }
                }

                return progIds.Any() ? string.Join(", ", progIds) : "Not specified";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        // No UpdateFromPath method - path is immutable after construction!

        private void UpdateLockStatusIfNeeded()
        {
            if (_assemblyPath == null || _assemblyFile?.Exists != true)
                return;

            try
            {
                var (isLocked, details) = CheckFileLock(_assemblyPath);

                // Only update if lock status actually changed
                if (IsLocked != isLocked)
                {
                    IsLocked = isLocked;
                    LockDetails = details;
                }
            }
            catch (Exception)
            {
                // Ignore errors during lock status updates
            }
        }

        private (bool IsLocked, string Details) CheckFileLock(string filePath)
        {
            try
            {
                using (var stream = _fileSystem.File.Open(filePath, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    return (false, "File is not locked - available for writing");
                }
            }
            catch (UnauthorizedAccessException)
            {
                return (true, "File is locked - access denied (may be in use by OneNote or another process)");
            }
            catch (IOException ex)
            {
                if (ex.Message.Contains("being used by another process"))
                {
                    return (true, "File is locked - currently being used by another process (likely OneNote)");
                }
                return (true, $"File is locked - {ex.Message}");
            }
            catch (Exception ex)
            {
                return (false, $"Could not determine lock status: {ex.Message}");
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] suffixes = { "B", "KB", "MB", "GB", "TB" };
            int counter = 0;
            decimal number = bytes;
            while (Math.Round(number / 1024) >= 1)
            {
                number /= 1024;
                counter++;
            }
            return $"{number:n1} {suffixes[counter]}";
        }

        public void Dispose()
        {
            _lockStatusTimer?.Stop();
            _lockStatusTimer = null;

            // Note: In .NET Core/5+, Assembly instances don't need explicit disposal
            // The LoadFrom context will be cleaned up by GC
            _loadedAssembly = null;
            _versionInfo = null;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}