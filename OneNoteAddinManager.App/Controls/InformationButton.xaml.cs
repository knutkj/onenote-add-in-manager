using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace OneNoteAddinManager.App.Controls
{
    public partial class InformationButton : UserControl
    {
        public static readonly DependencyProperty DocumentationTopicProperty =
            DependencyProperty.Register(nameof(DocumentationTopic), typeof(string), typeof(InformationButton), new PropertyMetadata(string.Empty));

        public static readonly DependencyProperty ToolTipTextProperty =
            DependencyProperty.Register(nameof(ToolTipText), typeof(string), typeof(InformationButton), new PropertyMetadata(string.Empty));

        public static readonly DependencyProperty ShowInfoCommandProperty =
            DependencyProperty.Register(nameof(ShowInfoCommand), typeof(ICommand), typeof(InformationButton), new PropertyMetadata(null));

        public string DocumentationTopic
        {
            get { return (string)GetValue(DocumentationTopicProperty); }
            set { SetValue(DocumentationTopicProperty, value); }
        }

        public string ToolTipText
        {
            get { return (string)GetValue(ToolTipTextProperty); }
            set { SetValue(ToolTipTextProperty, value); }
        }

        public ICommand ShowInfoCommand
        {
            get { return (ICommand)GetValue(ShowInfoCommandProperty); }
            set { SetValue(ShowInfoCommandProperty, value); }
        }

        public InformationButton()
        {
            InitializeComponent();
        }
    }
}