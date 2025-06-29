using System;
using System.Windows;
using System.Windows.Controls;
using OneNoteAddinManager.ViewModels;

namespace OneNoteAddinManager.Controls
{
    /// <summary>
    /// AddInDetailsPanel - MVVM UserControl for displaying complete add-in details
    /// </summary>
    public partial class AddInDetailsPanel : UserControl
    {
        private AddInDetailsPanelViewModel? _viewModel;

        public AddInDetailsPanel()
        {
            InitializeComponent();
            
            // Start with no add-in selected (null)
            UpdateRegistryPath(null);
            
            // Clean up ViewModel when control is unloaded
            this.Unloaded += AddInDetailsPanel_Unloaded;
        }

        private void UpdateRegistryPath(string? registryPath)
        {
            // Dispose old ViewModel
            _viewModel?.Dispose();
            
            // Create new ViewModel for the new registry path
            _viewModel = new AddInDetailsPanelViewModel(registryPath);
            this.DataContext = _viewModel;  // Set internal DataContext for LoadBehavior bindings
            
            // Update child controls directly (they don't use binding to our ViewModel)
            AddInInfoControl.RegistryPath = registryPath;
            DllInfoControl.DllPath = _viewModel.DllPath;
        }

        // ONLY public interface - everything else is handled by ViewModel
        public static readonly DependencyProperty RegistryPathProperty =
            DependencyProperty.Register("RegistryPath", typeof(string), typeof(AddInDetailsPanel),
                new PropertyMetadata(null, OnRegistryPathChanged));

        public string? RegistryPath
        {
            get => (string?)GetValue(RegistryPathProperty);
            set => SetValue(RegistryPathProperty, value);
        }

        private static void OnRegistryPathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (AddInDetailsPanel)d;
            var newRegistryPath = e.NewValue as string;
            System.Diagnostics.Debug.WriteLine($"[AddInDetailsPanel] OnRegistryPathChanged called with: {newRegistryPath}");
            control.UpdateRegistryPath(newRegistryPath);
        }

        // Event for requesting documentation (for LoadBehavior info button)
        public event EventHandler<InfoRequestedEventArgs>? InfoRequested;

        private void LoadBehaviorInfoButton_Click(object sender, RoutedEventArgs e)
        {
            // Forward to documentation viewer using the established pattern
            InfoRequested?.Invoke(this, new InfoRequestedEventArgs("loadbehavior"));
        }

        private void AddInDetailsPanel_Unloaded(object sender, RoutedEventArgs e)
        {
            // Clean up ViewModel
            _viewModel?.Dispose();
        }

        private void AddInInfoControl_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}