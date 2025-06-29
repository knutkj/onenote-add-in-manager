using System;
using System.ComponentModel;
using System.Linq;
using OneNoteAddinManager.Models;
using OneNoteAddinManager.Services;

namespace OneNoteAddinManager.ViewModels
{
    /// <summary>
    /// ViewModel for AddInDetailsPanel - manages complete add-in details display
    /// Following established MVVM patterns from AddInInfoViewModel
    /// </summary>
    public class AddInDetailsPanelViewModel : INotifyPropertyChanged, IDisposable
    {
        private AddinInfo? _addinInfo;
        private bool _disposed = false;

        public AddInDetailsPanelViewModel(string? registryPath)
        {
            LoadAddinInfo(registryPath);
        }

        private void LoadAddinInfo(string? registryPath)
        {
            if (!string.IsNullOrEmpty(registryPath))
            {
                try
                {
                    // Extract add-in name from registry path and load its info
                    var addinName = ExtractAddinNameFromPath(registryPath);
                    if (!string.IsNullOrEmpty(addinName))
                    {
                        var registryManager = new RegistryManager();
                        var addins = registryManager.GetInstalledAddins();
                        _addinInfo = addins.Find(a => a.Name.Equals(addinName, StringComparison.OrdinalIgnoreCase));
                    }
                }
                catch (Exception)
                {
                    _addinInfo = null;
                }
            }
            else
            {
                _addinInfo = null;
            }
        }

        private string? ExtractAddinNameFromPath(string registryPath)
        {
            var parts = registryPath.Split('\\');
            return parts.Length > 0 ? parts[parts.Length - 1] : null;
        }

        // Properties for binding to the UI

        /// <summary>
        /// Registry path to pass to AddInInformationControl
        /// </summary>
        public string? RegistryPath => _addinInfo?.OfficeAddinRegistryPath;

        /// <summary>
        /// DLL path to pass to DllInformationControl
        /// </summary>
        public string? DllPath => _addinInfo?.DllPath;

        /// <summary>
        /// LoadBehavior text for display
        /// </summary>
        public string LoadBehaviorText => _addinInfo != null ? $"LoadBehavior: {_addinInfo.LoadBehavior}" : "LoadBehavior: Not Available";

        /// <summary>
        /// LoadBehavior explanation text for display
        /// </summary>
        public string LoadBehaviorExplanationText => _addinInfo?.LoadBehaviorExplanation ?? "Add-in information not available.";

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Clean up managed resources
                    _addinInfo = null;
                }

                _disposed = true;
            }
        }

        #endregion
    }
}