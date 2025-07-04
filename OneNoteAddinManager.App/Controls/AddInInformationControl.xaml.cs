using System;
using System.Windows;
using System.Windows.Controls;
using OneNoteAddinManager.App.ViewModels;

namespace OneNoteAddinManager.App.Controls
{
    /// <summary>
    /// Pure View - contains NO business logic, only UI structure and ViewModel binding
    /// </summary>
    public partial class AddInInformationControl : UserControl
    {
        private AddInInfoViewModel? _viewModel;

        public AddInInformationControl()
        {
            InitializeComponent();
            
            // Start with no add-in selected (null)
            UpdateViewModel(null);
            
            // Clean up ViewModel when control is unloaded
            this.Unloaded += AddInInformationControl_Unloaded;
        }

        private void UpdateViewModel(string? registryPath)
        {
            // Dispose old ViewModel
            _viewModel?.Dispose();
            
            // Create new ViewModel for the new registry path (null = no add-in)
            _viewModel = new AddInInfoViewModel(registryPath);
            this.DataContext = _viewModel;
        }

        // ONLY public interface - everything else is handled by ViewModel
        public static readonly DependencyProperty RegistryPathProperty =
            DependencyProperty.Register("RegistryPath", typeof(string), typeof(AddInInformationControl),
                new PropertyMetadata(null, OnRegistryPathChanged));

        public string? RegistryPath
        {
            get => (string?)GetValue(RegistryPathProperty);
            set => SetValue(RegistryPathProperty, value);
        }

        private static void OnRegistryPathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (AddInInformationControl)d;
            // Create new ViewModel instance for new registry path (handle null properly)
            var newPath = e.NewValue as string;
            control.UpdateViewModel(string.IsNullOrEmpty(newPath) ? null : newPath);
        }

        private void AddInInformationControl_Unloaded(object sender, RoutedEventArgs e)
        {
            // Clean up ViewModel
            _viewModel?.Dispose();
        }
    }
}