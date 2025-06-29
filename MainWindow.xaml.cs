using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OneNoteAddinManager.Models;
using OneNoteAddinManager.Services;
using OneNoteAddinManager.ViewModels;

namespace OneNoteAddinManager;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly AddinManager _addinManager;
    private readonly ObservableCollection<AddinInfo> _addins;
    private AddinInfo? _selectedAddin;
    
    public ICommand ShowLoadBehaviorInfoCommand { get; private set; }


    public MainWindow()
    {
        Console.WriteLine("Starting MainWindow initialization...");

        // Initialize fields first
        _addinManager = new AddinManager();
        _addins = new ObservableCollection<AddinInfo>();
        
        // Initialize commands
        ShowLoadBehaviorInfoCommand = new RelayCommand<string>(ShowLoadBehaviorInfo);

        try
        {
            InitializeComponent();
            Console.WriteLine("InitializeComponent completed");
            
            // Set DataContext for command binding
            DataContext = this;

            AddinsListBox.ItemsSource = _addins;
            
            // Wire up OneNote control event
            HeaderOneNoteControl.StatusChanged += OneNoteControl_StatusChanged;
            Console.WriteLine("ListBox bound");

            CheckAdminPrivileges();
            Console.WriteLine("Admin privileges checked");

            LoadAddins();
            Console.WriteLine("Add-ins loaded");

            // Load initial documentation when window is loaded
            Loaded += MainWindow_Loaded;


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
            DocumentationViewer.DocumentId = "welcome";
        }
    }

    private void ShowAddinDetails(AddinInfo addin)
    {
        WelcomeText.Visibility = Visibility.Collapsed;
        DetailsTabControl.Visibility = Visibility.Visible;

        // Update Add-in Information Control
        AddInInfoControl.RegistryPath = addin.OfficeAddinRegistryPath;
        
        // Wire up events for the new ViewModel
        WireUpAddInInfoControlEvents();

        // Populate LoadBehavior information
        LoadBehaviorText.Text = $"LoadBehavior: {addin.LoadBehavior}";
        LoadBehaviorExplanationText.Text = addin.LoadBehaviorExplanation;

        // Populate registry keys
        RegistryKeysItemsControl.ItemsSource = addin.RegistryKeys;

        // Update DLL information
        DllInfoControl.DllPath = addin.DllPath;

    }

    private void HideAddinDetails()
    {
        WelcomeText.Visibility = Visibility.Visible;
        DetailsTabControl.Visibility = Visibility.Collapsed;
        
        // Clear the Add-in Information Control
        AddInInfoControl.RegistryPath = null;
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
        DocumentationViewer.DocumentId = "welcome";
    }


    private void DocumentationViewer_DocumentChanged(object sender, string documentId)
    {
        // Handle document change if needed
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
            DocumentationViewer.DocumentId = "Addin-Information";
        }
        else if (DetailsTabControl?.SelectedItem == COMDetailsTab)
        {
            DocumentationViewer.DocumentId = "COM-Information";
        }
        else
        {
            // Fallback when no tab is selected or details are hidden
            DocumentationViewer.DocumentId = "welcome";
        }
    }


    private void WireUpAddInInfoControlEvents()
    {
        if (AddInInfoControl.DataContext is OneNoteAddinManager.ViewModels.AddInInfoViewModel viewModel)
        {
            viewModel.InfoRequested += AddInInfoViewModel_InfoRequested;
        }
    }

    private void AddInInfoViewModel_InfoRequested(object? sender, OneNoteAddinManager.ViewModels.InfoRequestedEventArgs e)
    {
        string documentationTopic = e.FieldName.ToLower() switch
        {
            "name" => "field-name",
            "friendlyname" => "field-friendlyname", 
            "status" => "field-status",
            "guid" => "field-guid",
            "dllpath" => "field-dllpath",
            "registrypath" => "field-registrypath",
            _ => "fields"
        };
        
        DocumentationViewer.DocumentId = documentationTopic;
    }

    private void OneNoteControl_StatusChanged(object? sender, string status)
    {
        StatusText.Text = status;
    }

    private void ShowLoadBehaviorInfo(string? documentationTopic)
    {
        DocumentationViewer.DocumentId = documentationTopic ?? "field-loadbehavior";
    }
    
    public event PropertyChangedEventHandler? PropertyChanged;
    
    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}