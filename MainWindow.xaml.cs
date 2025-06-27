using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using OneNoteAddinManager.Models;
using OneNoteAddinManager.Services;

namespace OneNoteAddinManager;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private readonly AddinManager _addinManager;
    private readonly ObservableCollection<AddinInfo> _addins;
    private AddinInfo? _selectedAddin;

    // Live monitoring components
    private ManagementEventWatcher? _processStartWatcher;
    private ManagementEventWatcher? _processStopWatcher;
    private FileSystemWatcher? _fileWatcher;
    private System.Windows.Threading.DispatcherTimer? _statusUpdateTimer;

    public MainWindow()
    {
        Console.WriteLine("Starting MainWindow initialization...");

        // Initialize fields first
        _addinManager = new AddinManager();
        _addins = new ObservableCollection<AddinInfo>();

        try
        {
            InitializeComponent();
            Console.WriteLine("InitializeComponent completed");

            AddinsListBox.ItemsSource = _addins;
            Console.WriteLine("ListBox bound");

            CheckAdminPrivileges();
            Console.WriteLine("Admin privileges checked");

            LoadAddins();
            Console.WriteLine("Add-ins loaded");

            // Load initial documentation when window is loaded
            Loaded += MainWindow_Loaded;

            // Setup live monitoring
            SetupLiveMonitoring();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in MainWindow constructor: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            MessageBox.Show($"Error initializing application: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CheckAdminPrivileges()
    {
        if (!_addinManager.IsRunningAsAdministrator())
        {
            AdminStatusText.Text = "⚠ Not running as administrator - Some operations may fail";
            AdminStatusText.Foreground = System.Windows.Media.Brushes.Yellow;
        }
        else
        {
            AdminStatusText.Text = "✓ Running as administrator";
            AdminStatusText.Foreground = System.Windows.Media.Brushes.LightGreen;
        }
    }

    private void LoadAddins()
    {
        try
        {
            StatusText.Text = "Loading add-ins...";
            _addins.Clear();

            var addins = _addinManager.GetAllAddins();
            foreach (var addin in addins)
            {
                _addins.Add(addin);
            }

            StatusText.Text = $"Loaded {addins.Count} add-in(s)";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading add-ins: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "Error loading add-ins";
        }
    }

    private void AddinsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _selectedAddin = AddinsListBox.SelectedItem as AddinInfo;

        if (_selectedAddin != null)
        {
            ShowAddinDetails(_selectedAddin);
            EnableDisableButton.IsEnabled = true;
            UnregisterButton.IsEnabled = true;

            // Update button text
            EnableDisableButton.Content = _selectedAddin.IsEnabled ? "Disable" : "Enable";

            // Load context-specific documentation based on current details tab
            LoadDocumentationForCurrentTab();
        }
        else
        {
            HideAddinDetails();
            EnableDisableButton.IsEnabled = false;
            UnregisterButton.IsEnabled = false;

            // Load appropriate documentation when no add-in is selected
            LoadDocumentation("welcome");
        }
    }

    private void ShowAddinDetails(AddinInfo addin)
    {
        WelcomeText.Visibility = Visibility.Collapsed;
        DetailsTabControl.Visibility = Visibility.Visible;

        // Populate basic information
        NameText.Text = addin.Name;
        FriendlyNameText.Text = addin.FriendlyName;
        AddinStatusText.Text = addin.Status;
        GuidText.Text = addin.Guid ?? "Not Available";
        DllPathText.Text = addin.DllPath ?? "Not Available";
        RegistryPathText.Text = addin.OfficeAddinRegistryPath;

        // Populate LoadBehavior information
        LoadBehaviorText.Text = $"LoadBehavior: {addin.LoadBehavior}";
        LoadBehaviorExplanationText.Text = addin.LoadBehaviorExplanation;

        // Populate registry keys
        RegistryKeysItemsControl.ItemsSource = addin.RegistryKeys;

        // Enable/disable Open Folder button based on DLL path validity
        UpdateOpenFolderButton(addin.DllPath);

        // Update DLL information
        UpdateDllInformation(addin.DllPath);

        // Setup file monitoring for the selected add-in's DLL
        SetupFileMonitoring(addin.DllPath);

        // Enable registry buttons when add-in is selected
        UpdateRegistryButtons(addin);
    }

    private void HideAddinDetails()
    {
        WelcomeText.Visibility = Visibility.Visible;
        DetailsTabControl.Visibility = Visibility.Collapsed;
        OpenDllFolderButton.IsEnabled = false;
        OpenRegistryButton.IsEnabled = false;
        CopyRegistryPathButton.IsEnabled = false;
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        HeaderRefreshStatusText.Text = "Refreshing...";
        HeaderRefreshStatusText.Foreground = Brushes.Yellow;
        
        try
        {
            LoadAddins();
            HeaderRefreshStatusText.Text = $"Updated {DateTime.Now:HH:mm:ss}";
            HeaderRefreshStatusText.Foreground = Brushes.LightGreen;
        }
        catch (Exception ex)
        {
            HeaderRefreshStatusText.Text = $"Error: {ex.Message}";
            HeaderRefreshStatusText.Foreground = Brushes.LightCoral;
        }
    }

    private void CleanupButton_Click(object sender, RoutedEventArgs e)
    {
        HeaderCleanupStatusText.Text = "Scanning...";
        HeaderCleanupStatusText.Foreground = Brushes.Yellow;
        
        try
        {
            var orphaned = _addinManager.FindOrphanedEntries();
            if (orphaned.Count == 0)
            {
                HeaderCleanupStatusText.Text = "No orphans found";
                HeaderCleanupStatusText.Foreground = Brushes.LightGreen;
                MessageBox.Show("No orphaned entries found.", "Cleanup",
                               MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            HeaderCleanupStatusText.Text = $"Found {orphaned.Count} orphan(s)";
            HeaderCleanupStatusText.Foreground = Brushes.Orange;

            var result = MessageBox.Show(
                $"Found {orphaned.Count} orphaned entries. Remove them?\n\n" +
                string.Join("\n", orphaned.Select(a => $"- {a.FriendlyName}")),
                "Cleanup Orphaned Entries",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                HeaderCleanupStatusText.Text = "Cleaning up...";
                HeaderCleanupStatusText.Foreground = Brushes.Yellow;
                
                _addinManager.CleanupOrphanedEntries();
                LoadAddins();
                StatusText.Text = $"Cleaned up {orphaned.Count} orphaned entries";
                
                HeaderCleanupStatusText.Text = $"Cleaned {orphaned.Count} orphan(s)";
                HeaderCleanupStatusText.Foreground = Brushes.LightGreen;
            }
            else
            {
                HeaderCleanupStatusText.Text = "Cleanup cancelled";
                HeaderCleanupStatusText.Foreground = Brushes.Gray;
            }
        }
        catch (Exception ex)
        {
            HeaderCleanupStatusText.Text = $"Error: {ex.Message}";
            HeaderCleanupStatusText.Foreground = Brushes.LightCoral;
            MessageBox.Show($"Error during cleanup: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ToggleAddin_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedAddin == null)
        {
            MessageBox.Show("Please select an add-in from the list.", "Selection Required",
                           MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            if (_selectedAddin.IsEnabled)
            {
                _addinManager.DisableAddin(_selectedAddin);
                StatusText.Text = $"Disabled {_selectedAddin.FriendlyName}";
            }
            else
            {
                _addinManager.EnableAddin(_selectedAddin);
                StatusText.Text = $"Enabled {_selectedAddin.FriendlyName}";
            }

            // Refresh the display
            ShowAddinDetails(_selectedAddin);
            EnableDisableButton.Content = _selectedAddin.IsEnabled ? "Disable" : "Enable";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error toggling add-in: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BrowseDllButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "DLL Files (*.dll)|*.dll|All Files (*.*)|*.*",
            Title = "Select Add-in DLL"
        };

        if (dialog.ShowDialog() == true)
        {
            DllPathTextBox.Text = dialog.FileName;
        }
    }

    private void RegisterButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(DllPathTextBox.Text))
            {
                MessageBox.Show("Please select a DLL file to register.", "Validation",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _addinManager.RegisterNewAddin(DllPathTextBox.Text);
            DllPathTextBox.Clear();
            LoadAddins();
            StatusText.Text = "Add-in registered successfully";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error registering add-in: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }


    private void UnregisterButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedAddin == null)
        {
            MessageBox.Show("Please select an add-in from the list.", "Selection Required",
                           MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var result = MessageBox.Show(
            $"Are you sure you want to unregister '{_selectedAddin.FriendlyName}'?\n\n" +
            "This will remove all registry entries for this add-in.",
            "Confirm Unregister",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            try
            {
                _addinManager.UnregisterAddin(_selectedAddin);
                LoadAddins();
                HideAddinDetails();
                StatusText.Text = $"Unregistered {_selectedAddin.FriendlyName}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error unregistering add-in: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        // Load welcome documentation when application starts
        LoadDocumentation("welcome");
    }

    private void OverviewLinkButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentationForCurrentTab();
    }

    private void FieldsLinkButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("fields");
    }

    private void TroubleshootingLinkButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("troubleshooting");
    }

    private void DetailsTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Load appropriate documentation when switching between details tabs
        LoadDocumentationForCurrentTab();
    }

    private void LoadDocumentationForCurrentTab()
    {
        if (DetailsTabControl?.SelectedItem == AddinDetailsTab)
        {
            LoadDocumentation("Addin-Information");
        }
        else if (DetailsTabControl?.SelectedItem == COMDetailsTab)
        {
            LoadDocumentation("COM-Information");
        }
        else
        {
            // Fallback when no tab is selected or details are hidden
            LoadDocumentation("welcome");
        }
    }

    private void LoadDocumentation(string topic)
    {
        if (MarkdownContent == null)
            return; // UI not initialized yet

        try
        {
            var resourceName = $"OneNoteAddinManager.Documentation.{topic}.md";
            var assembly = Assembly.GetExecutingAssembly();

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                // Fallback to file system
                var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documentation", $"{topic}.md");
                if (File.Exists(filePath))
                {
                    var content = File.ReadAllText(filePath);
                    DisplaySimpleMarkdown(content);
                }
                else
                {
                    MarkdownContent.Children.Clear();
                    MarkdownContent.Children.Add(new TextBlock
                    {
                        Text = $"Documentation for '{topic}' not found.",
                        FontFamily = new FontFamily("Segoe UI"),
                        FontSize = 12
                    });
                }
                return;
            }

            using var reader = new StreamReader(stream);
            var markdownContent = reader.ReadToEnd();
            DisplaySimpleMarkdown(markdownContent);
        }
        catch (Exception ex)
        {
            MarkdownContent.Children.Clear();
            MarkdownContent.Children.Add(new TextBlock
            {
                Text = $"Error loading documentation: {ex.Message}",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                Foreground = Brushes.Red
            });
        }
    }

    private void DisplaySimpleMarkdown(string markdownContent)
    {
        MarkdownContent.Children.Clear();

        var lines = markdownContent.Split('\n');
        var currentParagraph = new System.Text.StringBuilder();

        foreach (var line in lines)
        {
            var trimmedLine = line.TrimEnd();

            if (string.IsNullOrWhiteSpace(trimmedLine))
            {
                // End current paragraph if it has content
                if (currentParagraph.Length > 0)
                {
                    AddParagraph(currentParagraph.ToString());
                    currentParagraph.Clear();
                }
                continue;
            }

            // Headers
            if (trimmedLine.StartsWith("# "))
            {
                // Finish current paragraph first
                if (currentParagraph.Length > 0)
                {
                    AddParagraph(currentParagraph.ToString());
                    currentParagraph.Clear();
                }

                AddHeader1(trimmedLine.Substring(2));
                continue;
            }

            if (trimmedLine.StartsWith("## "))
            {
                if (currentParagraph.Length > 0)
                {
                    AddParagraph(currentParagraph.ToString());
                    currentParagraph.Clear();
                }

                AddHeader2(trimmedLine.Substring(3));
                continue;
            }

            if (trimmedLine.StartsWith("### "))
            {
                if (currentParagraph.Length > 0)
                {
                    AddParagraph(currentParagraph.ToString());
                    currentParagraph.Clear();
                }

                AddHeader3(trimmedLine.Substring(4));
                continue;
            }

            // List items
            if (trimmedLine.StartsWith("- ") || trimmedLine.StartsWith("* "))
            {
                if (currentParagraph.Length > 0)
                {
                    AddParagraph(currentParagraph.ToString());
                    currentParagraph.Clear();
                }

                AddListItem(trimmedLine.Substring(2));
                continue;
            }

            // Code blocks (simple detection)
            if (trimmedLine.StartsWith("```"))
            {
                continue; // Skip markdown code fences
            }

            // Regular paragraph text
            if (currentParagraph.Length > 0)
            {
                currentParagraph.AppendLine();
            }
            currentParagraph.Append(ProcessBoldText(trimmedLine));
        }

        // Add final paragraph if exists
        if (currentParagraph.Length > 0)
        {
            AddParagraph(currentParagraph.ToString());
        }
    }

    private void AddHeader1(string text)
    {
        var textBlock = new TextBlock
        {
            Text = text,
            FontSize = 20,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x2B, 0x57, 0x9A)),
            Margin = new Thickness(0, 15, 0, 10),
            TextWrapping = TextWrapping.Wrap
        };
        MarkdownContent.Children.Add(textBlock);
    }

    private void AddHeader2(string text)
    {
        var textBlock = new TextBlock
        {
            Text = text,
            FontSize = 16,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x2B, 0x57, 0x9A)),
            Margin = new Thickness(0, 12, 0, 8),
            TextWrapping = TextWrapping.Wrap
        };
        MarkdownContent.Children.Add(textBlock);
    }

    private void AddHeader3(string text)
    {
        var textBlock = new TextBlock
        {
            Text = text,
            FontSize = 14,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
            Margin = new Thickness(0, 10, 0, 6),
            TextWrapping = TextWrapping.Wrap
        };
        MarkdownContent.Children.Add(textBlock);
    }

    private void AddListItem(string text)
    {
        var textBlock = new TextBlock
        {
            Text = "• " + ProcessBoldText(text),
            FontSize = 12,
            Margin = new Thickness(20, 2, 0, 2),
            TextWrapping = TextWrapping.Wrap,
            FontFamily = new FontFamily("Segoe UI")
        };
        MarkdownContent.Children.Add(textBlock);
    }

    private void AddParagraph(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return;

        var textBlock = new TextBlock
        {
            FontSize = 12,
            Margin = new Thickness(0, 0, 0, 8),
            TextWrapping = TextWrapping.Wrap,
            FontFamily = new FontFamily("Segoe UI"),
            LineHeight = 18
        };

        // Process bold formatting
        ProcessBoldInlines(textBlock, text);

        MarkdownContent.Children.Add(textBlock);
    }

    private string ProcessBoldText(string text)
    {
        // Simple bold removal for plain text scenarios
        return text.Replace("**", "");
    }

    private void ProcessBoldInlines(TextBlock textBlock, string text)
    {
        var parts = text.Split(new[] { "**" }, StringSplitOptions.None);
        bool isBold = false;

        foreach (var part in parts)
        {
            if (string.IsNullOrEmpty(part))
            {
                isBold = !isBold;
                continue;
            }

            var run = new Run(part);
            if (isBold)
            {
                run.FontWeight = FontWeights.Bold;
            }

            textBlock.Inlines.Add(run);
            isBold = !isBold;
        }
    }

    private void UpdateOpenFolderButton(string? dllPath)
    {
        if (string.IsNullOrWhiteSpace(dllPath))
        {
            OpenDllFolderButton.IsEnabled = false;
            OpenDllFolderButton.ToolTip = "DLL path not available";
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(dllPath);
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
            {
                OpenDllFolderButton.IsEnabled = true;
                OpenDllFolderButton.ToolTip = $"Open folder: {directory}";
            }
            else
            {
                OpenDllFolderButton.IsEnabled = false;
                OpenDllFolderButton.ToolTip = "DLL folder does not exist";
            }
        }
        catch
        {
            OpenDllFolderButton.IsEnabled = false;
            OpenDllFolderButton.ToolTip = "Invalid DLL path";
        }
    }

    private void OpenDllFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedAddin?.DllPath == null)
        {
            MessageBox.Show("No DLL path available for the selected add-in.", "Information",
                           MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(_selectedAddin.DllPath);
            if (string.IsNullOrEmpty(directory))
            {
                MessageBox.Show("Could not determine the DLL folder path.", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!Directory.Exists(directory))
            {
                MessageBox.Show($"The DLL folder does not exist:\n{directory}", "Folder Not Found",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Open the folder in Windows Explorer
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{directory}\"",
                UseShellExecute = true
            });

            StatusText.Text = $"Opened folder: {directory}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening folder: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // Field Information Button Click Handlers
    private void NameInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-name");
    }

    private void FriendlyNameInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-friendlyname");
    }

    private void StatusInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-status");
    }

    private void GuidInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-guid");
    }

    private void DllPathInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-dllpath");
    }

    private void LoadBehaviorInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-loadbehavior");
    }

    private void UpdateDllInformation(string? dllPath)
    {
        if (string.IsNullOrWhiteSpace(dllPath))
        {
            DllExistsText.Text = "N/A";
            DllLockedText.Text = "N/A";
            DllSizeText.Text = "N/A";
            DllModifiedText.Text = "N/A";
        }
        else
        {
            try
            {
                if (File.Exists(dllPath))
                {
                    DllExistsText.Text = "✓ Yes";
                    DllExistsText.Foreground = Brushes.Green;

                    var fileInfo = new FileInfo(dllPath);
                    DllSizeText.Text = FormatFileSize(fileInfo.Length);
                    DllModifiedText.Text = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");

                    // Check if file is locked with detailed information
                    var (isLocked, lockDetails) = CheckFileLock(dllPath);
                    if (isLocked)
                    {
                        DllLockedText.Text = "🔒 Yes";
                        DllLockedText.Foreground = Brushes.Orange;
                        DllLockedText.ToolTip = lockDetails;
                    }
                    else
                    {
                        DllLockedText.Text = "🔓 No";
                        DllLockedText.Foreground = Brushes.Green;
                        DllLockedText.ToolTip = lockDetails;
                    }
                }
                else
                {
                    DllExistsText.Text = "❌ No";
                    DllExistsText.Foreground = Brushes.Red;
                    DllLockedText.Text = "N/A";
                    DllSizeText.Text = "N/A";
                    DllModifiedText.Text = "N/A";
                }
            }
            catch (Exception ex)
            {
                DllExistsText.Text = $"Error: {ex.Message}";
                DllExistsText.Foreground = Brushes.Red;
                DllLockedText.Text = "Unknown";
                DllSizeText.Text = "Unknown";
                DllModifiedText.Text = "Unknown";
            }
        }

        // Update OneNote status
        UpdateOneNoteStatus();
    }

    private (bool IsLocked, string Details) CheckFileLock(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return (false, "File does not exist");

            // Simple file lock check - try to open file for writing
            try
            {
                using var stream = File.Open(filePath, FileMode.Open, FileAccess.Write, FileShare.None);
                return (false, "Not locked");
            }
            catch (IOException)
            {
                return (true, "File is locked");
            }
        }
        catch (Exception ex)
        {
            return (false, $"Error checking lock: {ex.Message}");
        }
    }





    private bool IsFileLocked(string filePath)
    {
        var (isLocked, _) = CheckFileLock(filePath);
        return isLocked;
    }

    private string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }

    private void UpdateOneNoteStatus()
    {
        try
        {
            var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");
            if (oneNoteProcesses.Length > 0)
            {
                HeaderOneNoteStatusText.Text = "🟢 OneNote Running";
                HeaderOneNoteStatusText.Foreground = Brushes.LightGreen;
                HeaderStartOneNoteButton.Content = "⏹ Close OneNote";
                HeaderStartOneNoteButton.ToolTip = "Close all OneNote processes";

                try
                {
                    // Always show all PIDs
                    var allPids = string.Join(", ", oneNoteProcesses.Select(p => p.Id.ToString()));
                    HeaderOneNotePidText.Text = oneNoteProcesses.Length == 1 ? $"PID: {allPids}" : $"PIDs: {allPids}";

                    // Show the start time of the OneNote process
                    var oldestProcess = oneNoteProcesses.OrderBy(p => p.StartTime).First();
                    HeaderOneNoteStartTimeText.Text = $"Started: {oldestProcess.StartTime:HH:mm:ss}";
                }
                catch (Exception ex)
                {
                    HeaderOneNotePidText.Text = "PID: Access denied";
                    HeaderOneNoteStartTimeText.Text = $"Error: {ex.Message}";
                }
            }
            else
            {
                HeaderOneNoteStatusText.Text = "🔴 OneNote Not Running";
                HeaderOneNoteStatusText.Foreground = Brushes.LightCoral;
                HeaderStartOneNoteButton.Content = "▶ Start OneNote";
                HeaderStartOneNoteButton.ToolTip = "Start Microsoft OneNote";
                HeaderOneNotePidText.Text = "";
                HeaderOneNoteStartTimeText.Text = "";
            }
        }
        catch (Exception ex)
        {
            HeaderOneNoteStatusText.Text = $"❓ Unknown: {ex.Message}";
            HeaderOneNoteStatusText.Foreground = Brushes.Gray;
            HeaderOneNotePidText.Text = "";
            HeaderOneNoteStartTimeText.Text = "";
        }
    }

    private void StartOneNoteButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var oneNoteProcesses = Process.GetProcessesByName("ONENOTE");

            if (oneNoteProcesses.Length > 0)
            {
                // Close OneNote without warning
                StatusText.Text = $"Closing {oneNoteProcesses.Length} OneNote process{(oneNoteProcesses.Length > 1 ? "es" : "")}...";

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

                StatusText.Text = statusMessage;

                // Update UI after a delay to allow processes to fully terminate
                var timer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2)
                };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    UpdateOneNoteStatus();
                    if (_selectedAddin?.DllPath != null)
                    {
                        UpdateDllInformation(_selectedAddin.DllPath);
                    }
                    StatusText.Text = "Ready - OneNote closed";
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
                StatusText.Text = "Starting OneNote...";

                // Update status after a delay
                var timer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(3)
                };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    UpdateOneNoteStatus();
                    if (_selectedAddin?.DllPath != null)
                    {
                        UpdateDllInformation(_selectedAddin.DllPath);
                    }
                    StatusText.Text = "Ready - OneNote started";
                };
                timer.Start();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error managing OneNote: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void RegistryPathInfoButton_Click(object sender, RoutedEventArgs e)
    {
        LoadDocumentation("field-registrypath");
    }

    private void UpdateRegistryButtons(AddinInfo addin)
    {
        if (addin == null || string.IsNullOrWhiteSpace(addin.Name))
        {
            OpenRegistryButton.IsEnabled = false;
            CopyRegistryPathButton.IsEnabled = false;
            OpenRegistryButton.ToolTip = "Add-in name not available";
            CopyRegistryPathButton.ToolTip = "Add-in name not available";
            return;
        }

        OpenRegistryButton.IsEnabled = true;
        CopyRegistryPathButton.IsEnabled = true;
        OpenRegistryButton.ToolTip = "Open Registry Editor";
        CopyRegistryPathButton.ToolTip = $"Copy registry path to clipboard: {addin.OfficeAddinRegistryPath}";
    }

    private void OpenRegistryButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Simply open Registry Editor
            Process.Start(new ProcessStartInfo
            {
                FileName = "regedit.exe",
                UseShellExecute = true
            });

            StatusText.Text = "Opened Registry Editor";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening Registry Editor: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SetupLiveMonitoring()
    {
        try
        {
            // Setup process monitoring for OneNote
            SetupProcessMonitoring();

            // Setup periodic status updates
            _statusUpdateTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2) // Update every 2 seconds
            };
            _statusUpdateTimer.Tick += StatusUpdateTimer_Tick;
            _statusUpdateTimer.Start();

            Console.WriteLine("Live monitoring setup completed");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error setting up live monitoring: {ex.Message}");
            // Continue without live monitoring
        }
    }

    private void SetupProcessMonitoring()
    {
        try
        {
            // Monitor process start events for OneNote
            var startQuery = new WqlEventQuery(
                "SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName = 'ONENOTE.EXE'");
            _processStartWatcher = new ManagementEventWatcher(startQuery);
            _processStartWatcher.EventArrived += OnOneNoteProcessStarted;
            _processStartWatcher.Start();

            // Monitor process stop events for OneNote
            var stopQuery = new WqlEventQuery(
                "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName = 'ONENOTE.EXE'");
            _processStopWatcher = new ManagementEventWatcher(stopQuery);
            _processStopWatcher.EventArrived += OnOneNoteProcessStopped;
            _processStopWatcher.Start();

            Console.WriteLine("Process monitoring setup completed");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Process monitoring setup failed: {ex.Message}");
            // Fall back to timer-based monitoring only
        }
    }

    private void SetupFileMonitoring(string? dllPath)
    {
        // Clean up existing file watcher
        _fileWatcher?.Dispose();
        _fileWatcher = null;

        if (string.IsNullOrWhiteSpace(dllPath) || !File.Exists(dllPath))
            return;

        try
        {
            var directory = Path.GetDirectoryName(dllPath);
            var fileName = Path.GetFileName(dllPath);

            if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(fileName))
                return;

            _fileWatcher = new FileSystemWatcher(directory, fileName)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.Attributes,
                EnableRaisingEvents = true
            };

            _fileWatcher.Changed += OnDllFileChanged;
            _fileWatcher.Error += OnFileWatcherError;

            Console.WriteLine($"File monitoring setup for: {dllPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"File monitoring setup failed: {ex.Message}");
        }
    }

    private void OnOneNoteProcessStarted(object sender, EventArrivedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            UpdateOneNoteStatus();
            StatusText.Text = "OneNote process started";

            // Update DLL lock status after a delay (give OneNote time to load)
            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(3)
            };
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                if (_selectedAddin?.DllPath != null)
                {
                    UpdateDllLockStatus(_selectedAddin.DllPath);
                }
            };
            timer.Start();
        });
    }

    private void OnOneNoteProcessStopped(object sender, EventArrivedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            UpdateOneNoteStatus();
            StatusText.Text = "OneNote process stopped";

            // Update DLL lock status immediately
            if (_selectedAddin?.DllPath != null)
            {
                UpdateDllLockStatus(_selectedAddin.DllPath);
            }
        });
    }

    private void OnDllFileChanged(object sender, FileSystemEventArgs e)
    {
        // Debounce file change events (they often fire multiple times)
        var timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(500)
        };
        timer.Tick += (s, args) =>
        {
            timer.Stop();
            Dispatcher.Invoke(() =>
            {
                if (_selectedAddin?.DllPath != null &&
                    string.Equals(e.FullPath, _selectedAddin.DllPath, StringComparison.OrdinalIgnoreCase))
                {
                    UpdateDllInformation(_selectedAddin.DllPath);
                    StatusText.Text = $"DLL file updated: {Path.GetFileName(e.Name)}";
                }
            });
        };
        timer.Start();
    }

    private void OnFileWatcherError(object sender, ErrorEventArgs e)
    {
        Console.WriteLine($"File watcher error: {e.GetException().Message}");
        Dispatcher.Invoke(() =>
        {
            // Try to restart file monitoring
            if (_selectedAddin?.DllPath != null)
            {
                SetupFileMonitoring(_selectedAddin.DllPath);
            }
        });
    }

    private void StatusUpdateTimer_Tick(object? sender, EventArgs e)
    {
        // Periodic update for file lock status (fallback when events don't work)
        // But don't spam the console with repeated checks
        if (_selectedAddin?.DllPath != null)
        {
            UpdateDllLockStatusQuiet(_selectedAddin.DllPath);
        }
    }

    private void UpdateDllLockStatus(string dllPath)
    {
        if (!File.Exists(dllPath))
            return;

        try
        {
            var wasLocked = DllLockedText.Text.Contains("🔒");
            var (isLocked, lockDetails) = CheckFileLock(dllPath);

            // Only update UI if status changed
            if (isLocked != wasLocked)
            {
                if (isLocked)
                {
                    DllLockedText.Text = "🔒 Yes";
                    DllLockedText.Foreground = Brushes.Orange;
                    DllLockedText.ToolTip = lockDetails;
                }
                else
                {
                    DllLockedText.Text = "🔓 No";
                    DllLockedText.Foreground = Brushes.Green;
                    DllLockedText.ToolTip = lockDetails;
                }

                // Log lock status changes for debugging
                Console.WriteLine($"DLL lock status changed: {(isLocked ? "LOCKED" : "UNLOCKED")} - {lockDetails}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error updating DLL lock status: {ex.Message}");
        }
    }

    private void UpdateDllLockStatusQuiet(string dllPath)
    {
        if (!File.Exists(dllPath))
            return;

        try
        {
            var wasLocked = DllLockedText.Text.Contains("🔒");
            var (isLocked, lockDetails) = CheckFileLock(dllPath);

            // Only update UI if status changed
            if (isLocked != wasLocked)
            {
                if (isLocked)
                {
                    DllLockedText.Text = "🔒 Yes";
                    DllLockedText.Foreground = Brushes.Orange;
                    DllLockedText.ToolTip = lockDetails;
                }
                else
                {
                    DllLockedText.Text = "🔓 No";
                    DllLockedText.Foreground = Brushes.Green;
                    DllLockedText.ToolTip = lockDetails;
                }

                // Log lock status changes for debugging
                Console.WriteLine($"DLL lock status changed: {(isLocked ? "LOCKED" : "UNLOCKED")} - {lockDetails}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error updating DLL lock status: {ex.Message}");
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        // Cleanup monitoring resources
        try
        {
            _processStartWatcher?.Stop();
            _processStartWatcher?.Dispose();

            _processStopWatcher?.Stop();
            _processStopWatcher?.Dispose();

            _fileWatcher?.Dispose();
            _statusUpdateTimer?.Stop();

            Console.WriteLine("Live monitoring cleaned up");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error cleaning up monitoring: {ex.Message}");
        }

        base.OnClosed(e);
    }


    private void CopyRegistryPathButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedAddin == null)
        {
            MessageBox.Show("No add-in is currently selected.", "Information",
                           MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            var registryPath = _selectedAddin.OfficeAddinRegistryPath;
            Clipboard.SetText(registryPath);
            StatusText.Text = $"Copied to clipboard: {registryPath}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error copying to clipboard: {ex.Message}", "Error",
                           MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}