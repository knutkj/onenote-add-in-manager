using System;
using System.Windows;
using System.Windows.Controls;
using OneNoteAddinManager.ViewModels;

namespace OneNoteAddinManager.Controls
{
    /// <summary>
    /// Pure View - contains NO business logic, only UI structure and ViewModel binding
    /// </summary>
    public partial class DllInformationControl : UserControl
    {
        private DllInfoViewModel? _viewModel;

        public DllInformationControl()
        {
            InitializeComponent();
            
            // Start with no file selected (null)
            UpdateViewModel(null);
            
            // Clean up ViewModel when control is unloaded
            this.Unloaded += DllInformationControl_Unloaded;
        }

        private void UpdateViewModel(string? dllPath)
        {
            // Dispose old ViewModel
            _viewModel?.Dispose();
            
            // Create new ViewModel for the new path (null = no file)
            _viewModel = new DllInfoViewModel(dllPath);
            this.DataContext = _viewModel;
        }

        // ONLY public interface - everything else is handled by ViewModel
        public static readonly DependencyProperty DllPathProperty =
            DependencyProperty.Register("DllPath", typeof(string), typeof(DllInformationControl),
                new PropertyMetadata(null, OnDllPathChanged));

        public string? DllPath
        {
            get => (string?)GetValue(DllPathProperty);
            set => SetValue(DllPathProperty, value);
        }

        private static void OnDllPathChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (DllInformationControl)d;
            // Create new ViewModel instance for new path (handle null properly)
            var newPath = e.NewValue as string;
            control.UpdateViewModel(string.IsNullOrEmpty(newPath) ? null : newPath);
        }

        private void DllInformationControl_Unloaded(object sender, RoutedEventArgs e)
        {
            // Clean up ViewModel
            _viewModel?.Dispose();
        }
    }
}