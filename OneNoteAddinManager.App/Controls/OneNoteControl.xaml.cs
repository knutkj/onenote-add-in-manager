using System;
using System.Windows;
using System.Windows.Controls;
using OneNoteAddinManager.App.ViewModels;

namespace OneNoteAddinManager.App.Controls
{
    /// <summary>
    /// Pure View - contains NO business logic, only UI structure and ViewModel binding
    /// OneNote control that totally owns all OneNote functionality
    /// </summary>
    public partial class OneNoteControl : UserControl
    {
        private OneNoteViewModel? _viewModel;

        public OneNoteControl()
        {
            InitializeComponent();
            
            // Create and set the ViewModel
            _viewModel = new OneNoteViewModel();
            this.DataContext = _viewModel;
            
            // Forward ViewModel events to the parent
            _viewModel.StatusChanged += OnStatusChanged;
            
            // Clean up ViewModel when control is unloaded
            this.Unloaded += OneNoteControl_Unloaded;
        }

        // Event to notify parent about status changes
        public event EventHandler<string>? StatusChanged;

        private void OnStatusChanged(object? sender, string status)
        {
            StatusChanged?.Invoke(this, status);
        }

        // Public method for manual status updates (called by external process monitoring)
        public void UpdateOneNoteStatus()
        {
            _viewModel?.UpdateOneNoteStatus();
        }

        private void OneNoteControl_Unloaded(object sender, RoutedEventArgs e)
        {
            // Clean up ViewModel
            if (_viewModel != null)
            {
                _viewModel.StatusChanged -= OnStatusChanged;
                _viewModel.Dispose();
                _viewModel = null;
            }
        }
    }
}