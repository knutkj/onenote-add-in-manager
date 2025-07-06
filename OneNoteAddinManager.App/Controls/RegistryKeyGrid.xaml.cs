using System.Windows;
using System.Windows.Controls;

namespace OneNoteAddinManager.App.Controls
{
    /// <summary>
    /// Interaction logic for RegistryKeyGrid.xaml
    /// </summary>
    public partial class RegistryKeyGrid : UserControl
    {
        public RegistryKeyGrid()
        {
            InitializeComponent();
        }
        
        // Observable HelpId Property for interactive help system
        public static readonly DependencyProperty HelpIdProperty =
            DependencyProperty.Register("HelpId", typeof(string), typeof(RegistryKeyGrid), 
                new PropertyMetadata(null));
        
        public string HelpId
        {
            get { return (string)GetValue(HelpIdProperty); }
            set { SetValue(HelpIdProperty, value); }
        }
        
        // Registry Path Property
        public static readonly DependencyProperty RegistryPathProperty =
            DependencyProperty.Register("RegistryPath", typeof(string), typeof(RegistryKeyGrid), 
                new PropertyMetadata(string.Empty));
        
        public string RegistryPath
        {
            get { return (string)GetValue(RegistryPathProperty); }
            set { SetValue(RegistryPathProperty, value); }
        }
        
        // Description Property
        public static readonly DependencyProperty DescriptionProperty =
            DependencyProperty.Register("Description", typeof(string), typeof(RegistryKeyGrid), 
                new PropertyMetadata(string.Empty));
        
        public string Description
        {
            get { return (string)GetValue(DescriptionProperty); }
            set { SetValue(DescriptionProperty, value); }
        }
    }
}