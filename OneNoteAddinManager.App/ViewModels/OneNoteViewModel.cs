using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Media;

namespace OneNoteAddinManager.App.ViewModels
{
    /// <summary>
    /// ViewModel for OneNote control - contains ALL OneNote management logic
    /// </summary>
    public class OneNoteViewModel : INotifyPropertyChanged, IDisposable
    {
        private System.Windows.Threading.DispatcherTimer? _statusUpdateTimer;
        private string _statusText = "🔴 OneNote Not Running";
        private Brush _statusBrush = Brushes.LightCoral;
        private string _buttonText = "▶ Start OneNote";
        private string _pidText = "";
        private string _startTimeText = "";
        private bool _showDetails = false;

        public OneNoteViewModel()
        {
            StartStopCommand = new RelayCommand(ExecuteStartStop);
            
            // Initial status update
            UpdateOneNoteStatus();
            
            // Set up periodic status updates
            _statusUpdateTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            _statusUpdateTimer.Tick += (s, e) => UpdateOneNoteStatus();
            _statusUpdateTimer.Start();
        }

        public ICommand StartStopCommand { get; }

        public string StatusText
        {
            get => _statusText;
            private set
            {
                if (_statusText != value)
                {
                    _statusText = value;
                    OnPropertyChanged();
                }
            }
        }

        public Brush StatusBrush
        {
            get => _statusBrush;
            private set
            {
                if (_statusBrush != value)
                {
                    _statusBrush = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ButtonText
        {
            get => _buttonText;
            private set
            {
                if (_buttonText != value)
                {
                    _buttonText = value;
                    OnPropertyChanged();
                }
            }
        }

        public string PidText
        {
            get => _pidText;
            private set
            {
                if (_pidText != value)
                {
                    _pidText = value;
                    OnPropertyChanged();
                }
            }
        }

        public string StartTimeText
        {
            get => _startTimeText;
            private set
            {
                if (_startTimeText != value)
                {
                    _startTimeText = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool ShowDetails
        {
            get => _showDetails;
            private set
            {
                if (_showDetails != value)
                {
                    _showDetails = value;
                    OnPropertyChanged();
                }
            }
        }

        // Event for communicating status messages to the main window
        public event EventHandler<string>? StatusChanged;

        public void UpdateOneNoteStatus()
        {
            try
            {
                var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");
                if (oneNoteProcesses.Length > 0)
                {
                    StatusText = "🟢 OneNote Running";
                    StatusBrush = Brushes.LightGreen;
                    ButtonText = "⏹ Close OneNote";
                    ShowDetails = true;

                    try
                    {
                        // Always show all PIDs
                        var allPids = string.Join(", ", oneNoteProcesses.Select(p => p.Id.ToString()));
                        PidText = oneNoteProcesses.Length == 1 ? $"PID: {allPids}" : $"PIDs: {allPids}";

                        // Show the start time of the oldest OneNote process
                        var oldestProcess = oneNoteProcesses.OrderBy(p => p.StartTime).First();
                        StartTimeText = $"Started: {oldestProcess.StartTime:HH:mm:ss}";
                    }
                    catch (Exception ex)
                    {
                        PidText = "PID: Access denied";
                        StartTimeText = $"Error: {ex.Message}";
                    }
                }
                else
                {
                    StatusText = "🔴 OneNote Not Running";
                    StatusBrush = Brushes.LightCoral;
                    ButtonText = "▶ Start OneNote";
                    ShowDetails = false;
                    PidText = "";
                    StartTimeText = "";
                }
            }
            catch (Exception ex)
            {
                StatusText = $"❓ Unknown: {ex.Message}";
                StatusBrush = Brushes.Gray;
                ShowDetails = false;
                PidText = "";
                StartTimeText = "";
            }
        }

        private void ExecuteStartStop()
        {
            try
            {
                var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");

                if (oneNoteProcesses.Length > 0)
                {
                    // Close OneNote
                    StatusChanged?.Invoke(this, $"Closing {oneNoteProcesses.Length} OneNote process{(oneNoteProcesses.Length > 1 ? "es" : "")}...");

                    // Close all OneNote processes
                    int closedCount = 0;
                    int killedCount = 0;

                    foreach (var process in oneNoteProcesses)
                    {
                        try
                        {
                            Console.WriteLine($"Attempting to close OneNote process PID {process.Id}");

                            // First try graceful close
                            if (process.CloseMainWindow())
                            {
                                if (process.WaitForExit(5000))
                                {
                                    closedCount++;
                                    Console.WriteLine($"Process PID {process.Id} closed gracefully");
                                }
                                else
                                {
                                    // Force kill if graceful close didn't work
                                    process.Kill();
                                    process.WaitForExit(2000);
                                    killedCount++;
                                    Console.WriteLine($"Process PID {process.Id} was force killed");
                                }
                            }
                            else
                            {
                                // No main window, force kill
                                process.Kill();
                                process.WaitForExit(2000);
                                killedCount++;
                                Console.WriteLine($"Process PID {process.Id} had no main window, force killed");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error closing OneNote process PID {process.Id}: {ex.Message}");
                        }
                    }

                    // Update status message
                    var statusMessage = $"Closed {closedCount} process{(closedCount != 1 ? "es" : "")}";
                    if (killedCount > 0)
                        statusMessage += $", killed {killedCount}";

                    StatusChanged?.Invoke(this, statusMessage);

                    // Update UI after a delay to allow processes to fully terminate
                    var timer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(2)
                    };
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        UpdateOneNoteStatus();
                        StatusChanged?.Invoke(this, "Ready - OneNote closed");
                    };
                    timer.Start();
                }
                else
                {
                    // No OneNote running, start it
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "onenote",
                        UseShellExecute = true
                    };

                    Process.Start(startInfo);
                    StatusChanged?.Invoke(this, "Starting OneNote...");

                    // Update status after a delay
                    var timer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(3)
                    };
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        UpdateOneNoteStatus();
                        StatusChanged?.Invoke(this, "Ready - OneNote started");
                    };
                    timer.Start();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error managing OneNote: {ex.Message}", "Error",
                                               System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        public void Dispose()
        {
            _statusUpdateTimer?.Stop();
            _statusUpdateTimer = null;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}