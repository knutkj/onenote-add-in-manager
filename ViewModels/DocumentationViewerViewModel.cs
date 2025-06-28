using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace OneNoteAddinManager.ViewModels
{
    public class DocumentationViewerViewModel : INotifyPropertyChanged, IDisposable
    {
        private string? _currentDocumentId;
        private ObservableCollection<UIElement> _markdownContent;

        public event PropertyChangedEventHandler? PropertyChanged;
        public event EventHandler<string>? DocumentChanged;

        public DocumentationViewerViewModel()
        {
            _markdownContent = new ObservableCollection<UIElement>();
        }

        public string? CurrentDocumentId
        {
            get => _currentDocumentId;
            set
            {
                if (_currentDocumentId != value)
                {
                    _currentDocumentId = value;
                    OnPropertyChanged();
                    LoadDocumentation(value);
                    if (value != null)
                    {
                        DocumentChanged?.Invoke(this, value);
                    }
                }
            }
        }

        public ObservableCollection<UIElement> MarkdownContent
        {
            get => _markdownContent;
            private set
            {
                _markdownContent = value;
                OnPropertyChanged();
            }
        }

        private void LoadDocumentation(string? topic)
        {
            if (string.IsNullOrEmpty(topic))
            {
                MarkdownContent.Clear();
                return;
            }

            try
            {
                var resourceName = $"OneNoteAddinManager.Documentation.{topic}.md";
                var assembly = Assembly.GetExecutingAssembly();

                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null)
                {
                    // Fallback to file system
                    var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documentation", $"{topic}.md");
                    if (File.Exists(filePath))
                    {
                        var content = File.ReadAllText(filePath);
                        DisplaySimpleMarkdown(content);
                    }
                    else
                    {
                        ShowErrorMessage($"Documentation for '{topic}' not found.");
                    }
                    return;
                }

                using var reader = new StreamReader(stream);
                var markdownContent = reader.ReadToEnd();
                DisplaySimpleMarkdown(markdownContent);
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"Error loading documentation: {ex.Message}");
            }
        }

        private void ShowErrorMessage(string message)
        {
            MarkdownContent.Clear();
            MarkdownContent.Add(new TextBlock
            {
                Text = message,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                Foreground = Brushes.Red
            });
        }

        private void DisplaySimpleMarkdown(string markdownContent)
        {
            MarkdownContent.Clear();

            var lines = markdownContent.Split('\n');
            var currentParagraph = new StringBuilder();

            foreach (var line in lines)
            {
                var trimmedLine = line.TrimEnd();

                if (string.IsNullOrWhiteSpace(trimmedLine))
                {
                    // End current paragraph if it has content
                    if (currentParagraph.Length > 0)
                    {
                        AddParagraph(currentParagraph.ToString());
                        currentParagraph.Clear();
                    }
                    continue;
                }

                // Headers
                if (trimmedLine.StartsWith("# "))
                {
                    // Finish current paragraph first
                    if (currentParagraph.Length > 0)
                    {
                        AddParagraph(currentParagraph.ToString());
                        currentParagraph.Clear();
                    }

                    AddHeader1(trimmedLine.Substring(2));
                    continue;
                }

                if (trimmedLine.StartsWith("## "))
                {
                    if (currentParagraph.Length > 0)
                    {
                        AddParagraph(currentParagraph.ToString());
                        currentParagraph.Clear();
                    }

                    AddHeader2(trimmedLine.Substring(3));
                    continue;
                }

                if (trimmedLine.StartsWith("### "))
                {
                    if (currentParagraph.Length > 0)
                    {
                        AddParagraph(currentParagraph.ToString());
                        currentParagraph.Clear();
                    }

                    AddHeader3(trimmedLine.Substring(4));
                    continue;
                }

                // List items
                if (trimmedLine.StartsWith("- ") || trimmedLine.StartsWith("* "))
                {
                    if (currentParagraph.Length > 0)
                    {
                        AddParagraph(currentParagraph.ToString());
                        currentParagraph.Clear();
                    }

                    AddListItem(trimmedLine.Substring(2));
                    continue;
                }

                // Code blocks (simple detection)
                if (trimmedLine.StartsWith("```"))
                {
                    continue; // Skip markdown code fences
                }

                // Regular paragraph text
                if (currentParagraph.Length > 0)
                {
                    currentParagraph.Append(" ");
                }
                currentParagraph.Append(ProcessBoldText(trimmedLine));
            }

            // Add final paragraph if exists
            if (currentParagraph.Length > 0)
            {
                AddParagraph(currentParagraph.ToString());
            }
        }

        private void AddHeader1(string text)
        {
            var textBlock = new TextBlock
            {
                Text = text,
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x2B, 0x57, 0x9A)),
                Margin = new Thickness(0, 15, 0, 10),
                TextWrapping = TextWrapping.Wrap
            };
            MarkdownContent.Add(textBlock);
        }

        private void AddHeader2(string text)
        {
            var textBlock = new TextBlock
            {
                Text = text,
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x2B, 0x57, 0x9A)),
                Margin = new Thickness(0, 12, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            MarkdownContent.Add(textBlock);
        }

        private void AddHeader3(string text)
        {
            var textBlock = new TextBlock
            {
                Text = text,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
                Margin = new Thickness(0, 10, 0, 6),
                TextWrapping = TextWrapping.Wrap
            };
            MarkdownContent.Add(textBlock);
        }

        private void AddListItem(string text)
        {
            var textBlock = new TextBlock
            {
                Text = "• " + ProcessBoldText(text),
                FontSize = 12,
                Margin = new Thickness(20, 2, 0, 2),
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Segoe UI")
            };
            MarkdownContent.Add(textBlock);
        }

        private void AddParagraph(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            var textBlock = new TextBlock
            {
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Segoe UI"),
                LineHeight = 18
            };

            // Process bold formatting
            ProcessBoldInlines(textBlock, text);

            MarkdownContent.Add(textBlock);
        }

        private string ProcessBoldText(string text)
        {
            // Simple bold removal for plain text scenarios
            return text.Replace("**", "");
        }

        private void ProcessBoldInlines(TextBlock textBlock, string text)
        {
            var parts = text.Split(new[] { "**" }, StringSplitOptions.None);
            bool isBold = false;

            foreach (var part in parts)
            {
                if (string.IsNullOrEmpty(part))
                {
                    isBold = !isBold;
                    continue;
                }

                var run = new Run(part);
                if (isBold)
                {
                    run.FontWeight = FontWeights.Bold;
                }

                textBlock.Inlines.Add(run);
                isBold = !isBold;
            }
        }

        protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            MarkdownContent?.Clear();
        }
    }
}