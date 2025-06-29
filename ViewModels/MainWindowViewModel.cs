using System.ComponentModel;

namespace OneNoteAddinManager.ViewModels
{
    /// <summary>
    /// ViewModel for MainWindow - minimal implementation for AddInDetailsPanel binding
    /// </summary>
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private string? _selectedAddinRegistryPath;

        /// <summary>
        /// Registry path of the currently selected add-in (for binding to AddInDetailsPanel)
        /// </summary>
        public string? SelectedAddinRegistryPath
        {
            get => _selectedAddinRegistryPath;
            set
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindowViewModel] SelectedAddinRegistryPath setter called with: {value}");
                _selectedAddinRegistryPath = value;
                OnPropertyChanged(nameof(SelectedAddinRegistryPath));
                System.Diagnostics.Debug.WriteLine($"[MainWindowViewModel] PropertyChanged event fired for SelectedAddinRegistryPath");
            }
        }

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}