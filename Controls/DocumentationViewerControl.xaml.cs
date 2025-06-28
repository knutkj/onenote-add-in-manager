using System;
using System.Windows;
using System.Windows.Controls;
using OneNoteAddinManager.ViewModels;

namespace OneNoteAddinManager.Controls
{
    public partial class DocumentationViewerControl : UserControl
    {
        public static readonly DependencyProperty DocumentIdProperty =
            DependencyProperty.Register("DocumentId", typeof(string), typeof(DocumentationViewerControl),
                new PropertyMetadata(null, OnDocumentIdChanged));

        public event EventHandler<string>? DocumentChanged;

        private DocumentationViewerViewModel _viewModel;

        public DocumentationViewerControl()
        {
            InitializeComponent();
            _viewModel = new DocumentationViewerViewModel();
            _viewModel.DocumentChanged += (sender, documentId) => DocumentChanged?.Invoke(this, documentId);
            DataContext = _viewModel;
            Unloaded += DocumentationViewerControl_Unloaded;
        }

        public string? DocumentId
        {
            get => (string?)GetValue(DocumentIdProperty);
            set => SetValue(DocumentIdProperty, value);
        }

        private static void OnDocumentIdChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DocumentationViewerControl control && control._viewModel != null)
            {
                control._viewModel.CurrentDocumentId = e.NewValue as string;
            }
        }


        private void DocumentationViewerControl_Unloaded(object sender, RoutedEventArgs e)
        {
            _viewModel?.Dispose();
        }
    }
}