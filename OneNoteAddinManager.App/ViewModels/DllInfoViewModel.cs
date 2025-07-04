using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using OneNoteAddinManager.Lib.Models;

namespace OneNoteAddinManager.App.ViewModels
{
    /// <summary>
    /// ViewModel for DLL information - contains ALL presentation logic
    /// </summary>
    public class DllInfoViewModel : INotifyPropertyChanged
    {
        private readonly FileInfo? _assemblyFile;
        private readonly string? _assemblyPath;
        private System.Windows.Threading.DispatcherTimer? _lockStatusTimer;
        
        // ViewModel-specific state (not part of the immutable model)
        private bool _isLocked = false;
        private string _lockDetails = string.Empty;

        public DllInfoViewModel(string? assemblyPath)
        {
            _assemblyPath = assemblyPath;
            
            // Create FileInfo only if we have a valid path
            if (assemblyPath != null)
            {
                _assemblyFile = new FileInfo(assemblyPath);
                
                // Initialize lock status
                if (_assemblyFile.Exists)
                {
                    var (isLocked, details) = CheckFileLock(assemblyPath);
                    _isLocked = isLocked;
                    _lockDetails = details;
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
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Write, FileShare.None))
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
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}