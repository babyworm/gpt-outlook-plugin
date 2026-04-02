using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using GptOutlookPlugin.Core;

namespace GptOutlookPlugin.UI
{
    public class DiffView : UserControl
    {
        private readonly RichTextBox _richTextBox;

        private static readonly SolidColorBrush DeletedBg = new SolidColorBrush(Color.FromRgb(255, 235, 235));
        private static readonly SolidColorBrush DeletedFg = new SolidColorBrush(Color.FromRgb(180, 30, 30));
        private static readonly SolidColorBrush InsertedBg = new SolidColorBrush(Color.FromRgb(230, 255, 235));
        private static readonly SolidColorBrush InsertedFg = new SolidColorBrush(Color.FromRgb(20, 130, 40));
        private static readonly SolidColorBrush ModifiedBg = new SolidColorBrush(Color.FromRgb(255, 248, 220));
        private static readonly SolidColorBrush ModifiedFg = new SolidColorBrush(Color.FromRgb(160, 110, 0));
        private static readonly SolidColorBrush UnchangedFg = new SolidColorBrush(Color.FromRgb(100, 100, 100));
        private static readonly SolidColorBrush PrefixDeletedFg = new SolidColorBrush(Color.FromRgb(220, 60, 60));
        private static readonly SolidColorBrush PrefixInsertedFg = new SolidColorBrush(Color.FromRgb(40, 160, 60));
        private static readonly FontFamily MonoFont = new FontFamily("Consolas");

        public DiffView()
        {
            _richTextBox = new RichTextBox
            {
                IsReadOnly = true,
                BorderThickness = new Thickness(0),
                Background = new SolidColorBrush(Color.FromRgb(252, 252, 252)),
                FontFamily = MonoFont,
                FontSize = 12,
                Padding = new Thickness(4),
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            Content = _richTextBox;
        }

        public void SetTexts(string original, string modified)
        {
            var doc = new FlowDocument { PagePadding = new Thickness(0) };

            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(modified))
            {
                var p = new Paragraph();
                p.Inlines.Add(new Run(modified ?? ""));
                doc.Blocks.Add(p);
                _richTextBox.Document = doc;
                return;
            }

            var diffs = DiffEngine.ComputeSentenceDiff(original, modified);

            foreach (var diff in diffs)
            {
                var para = new Paragraph
                {
                    Margin = new Thickness(0, 0, 0, 0),
                    Padding = new Thickness(6, 2, 6, 2),
                    LineHeight = 18
                };

                switch (diff.Type)
                {
                    case DiffType.Deleted:
                        para.Background = DeletedBg;
                        para.Inlines.Add(new Run("\u2212 ") { Foreground = PrefixDeletedFg, FontWeight = FontWeights.Bold });
                        para.Inlines.Add(new Run(diff.Text)
                        {
                            Foreground = DeletedFg,
                            TextDecorations = TextDecorations.Strikethrough
                        });
                        break;

                    case DiffType.Inserted:
                        para.Background = InsertedBg;
                        para.Inlines.Add(new Run("+ ") { Foreground = PrefixInsertedFg, FontWeight = FontWeights.Bold });
                        para.Inlines.Add(new Run(diff.Text) { Foreground = InsertedFg });
                        break;

                    case DiffType.Modified:
                        para.Background = ModifiedBg;
                        para.Inlines.Add(new Run("~ ") { Foreground = ModifiedFg, FontWeight = FontWeights.Bold });
                        para.Inlines.Add(new Run(diff.Text) { Foreground = ModifiedFg });
                        break;

                    default:
                        para.Inlines.Add(new Run("  "));
                        para.Inlines.Add(new Run(diff.Text) { Foreground = UnchangedFg });
                        break;
                }

                doc.Blocks.Add(para);
            }

            _richTextBox.Document = doc;
        }
    }
}
