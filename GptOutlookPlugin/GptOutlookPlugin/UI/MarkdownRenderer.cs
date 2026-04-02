using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace GptOutlookPlugin.UI
{
    /// <summary>
    /// Simple Markdown to WPF FlowDocument converter.
    /// Handles: headers, bold, italic, code, code blocks, lists, line breaks.
    /// </summary>
    public static class MarkdownRenderer
    {
        private static readonly SolidColorBrush CodeBg = new SolidColorBrush(Color.FromRgb(245, 245, 245));
        private static readonly SolidColorBrush CodeFg = new SolidColorBrush(Color.FromRgb(200, 50, 50));
        private static readonly SolidColorBrush CodeBlockBg = new SolidColorBrush(Color.FromRgb(40, 44, 52));
        private static readonly SolidColorBrush CodeBlockFg = new SolidColorBrush(Color.FromRgb(220, 220, 220));
        private static readonly FontFamily MonoFont = new FontFamily("Consolas");
        private static readonly SolidColorBrush TextColor = new SolidColorBrush(Color.FromRgb(40, 40, 40));

        public static RichTextBox Render(string markdown)
        {
            var rtb = new RichTextBox
            {
                IsReadOnly = true,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(0),
                FontSize = 13,
                Foreground = TextColor
            };

            var doc = new FlowDocument
            {
                PagePadding = new Thickness(0),
                LineHeight = 1
            };

            if (string.IsNullOrEmpty(markdown))
            {
                rtb.Document = doc;
                return rtb;
            }

            var lines = markdown.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            bool inCodeBlock = false;
            string codeBlockContent = "";

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];

                // Code block start/end
                if (line.TrimStart().StartsWith("```"))
                {
                    if (inCodeBlock)
                    {
                        // End code block
                        var codePara = new Paragraph
                        {
                            Background = CodeBlockBg,
                            Padding = new Thickness(10, 8, 10, 8),
                            Margin = new Thickness(0, 4, 0, 4),
                            FontFamily = MonoFont,
                            FontSize = 12
                        };
                        codePara.Inlines.Add(new Run(codeBlockContent.TrimEnd()) { Foreground = CodeBlockFg });
                        doc.Blocks.Add(codePara);
                        codeBlockContent = "";
                        inCodeBlock = false;
                    }
                    else
                    {
                        inCodeBlock = true;
                    }
                    continue;
                }

                if (inCodeBlock)
                {
                    codeBlockContent += line + "\n";
                    continue;
                }

                // Empty line = paragraph break
                if (string.IsNullOrWhiteSpace(line))
                {
                    doc.Blocks.Add(new Paragraph(new Run("")) { Margin = new Thickness(0, 2, 0, 2), FontSize = 4 });
                    continue;
                }

                // Headers
                if (line.StartsWith("### "))
                {
                    var p = new Paragraph { Margin = new Thickness(0, 6, 0, 3) };
                    p.Inlines.Add(new Run(line.Substring(4)) { FontWeight = FontWeights.Bold, FontSize = 14 });
                    doc.Blocks.Add(p);
                    continue;
                }
                if (line.StartsWith("## "))
                {
                    var p = new Paragraph { Margin = new Thickness(0, 8, 0, 3) };
                    p.Inlines.Add(new Run(line.Substring(3)) { FontWeight = FontWeights.Bold, FontSize = 15 });
                    doc.Blocks.Add(p);
                    continue;
                }
                if (line.StartsWith("# "))
                {
                    var p = new Paragraph { Margin = new Thickness(0, 10, 0, 4) };
                    p.Inlines.Add(new Run(line.Substring(2)) { FontWeight = FontWeights.Bold, FontSize = 16 });
                    doc.Blocks.Add(p);
                    continue;
                }

                // List items
                var listMatch = Regex.Match(line, @"^(\s*)[-*]\s+(.+)$");
                if (listMatch.Success)
                {
                    var indent = listMatch.Groups[1].Value.Length;
                    var p = new Paragraph { Margin = new Thickness(10 + indent * 8, 1, 0, 1) };
                    p.Inlines.Add(new Run("\u2022  ") { Foreground = new SolidColorBrush(Color.FromRgb(33, 120, 216)) });
                    AddInlineFormatted(p, listMatch.Groups[2].Value);
                    doc.Blocks.Add(p);
                    continue;
                }

                // Numbered list
                var numMatch = Regex.Match(line, @"^\s*(\d+)[.)]\s+(.+)$");
                if (numMatch.Success)
                {
                    var p = new Paragraph { Margin = new Thickness(10, 1, 0, 1) };
                    p.Inlines.Add(new Run(numMatch.Groups[1].Value + ".  ")
                    {
                        Foreground = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                        FontWeight = FontWeights.SemiBold
                    });
                    AddInlineFormatted(p, numMatch.Groups[2].Value);
                    doc.Blocks.Add(p);
                    continue;
                }

                // Regular paragraph
                var para = new Paragraph { Margin = new Thickness(0, 1, 0, 1) };
                AddInlineFormatted(para, line);
                doc.Blocks.Add(para);
            }

            rtb.Document = doc;
            return rtb;
        }

        /// <summary>
        /// Parses inline markdown: **bold**, *italic*, `code`
        /// </summary>
        private static void AddInlineFormatted(Paragraph para, string text)
        {
            // Pattern: **bold**, *italic*, `code`
            var pattern = @"(\*\*(.+?)\*\*)|(\*(.+?)\*)|(`(.+?)`)";
            int lastIndex = 0;

            foreach (Match match in Regex.Matches(text, pattern))
            {
                // Add text before the match
                if (match.Index > lastIndex)
                    para.Inlines.Add(new Run(text.Substring(lastIndex, match.Index - lastIndex)));

                if (match.Groups[2].Success) // **bold**
                {
                    para.Inlines.Add(new Run(match.Groups[2].Value) { FontWeight = FontWeights.Bold });
                }
                else if (match.Groups[4].Success) // *italic*
                {
                    para.Inlines.Add(new Run(match.Groups[4].Value) { FontStyle = FontStyles.Italic });
                }
                else if (match.Groups[6].Success) // `code`
                {
                    para.Inlines.Add(new Run(match.Groups[6].Value)
                    {
                        FontFamily = MonoFont,
                        Background = CodeBg,
                        Foreground = CodeFg,
                        FontSize = 12
                    });
                }

                lastIndex = match.Index + match.Length;
            }

            // Add remaining text
            if (lastIndex < text.Length)
                para.Inlines.Add(new Run(text.Substring(lastIndex)));
        }
    }
}
