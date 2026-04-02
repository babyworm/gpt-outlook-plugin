using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.UI
{
    public class ChatTaskPane : UserControl
    {
        private readonly ScrollViewer _scrollViewer;
        private readonly ItemsControl _messagesControl;
        private readonly TextBox _inputBox;

        public ChatTaskPane()
        {
            MinWidth = 320;
            Background = new SolidColorBrush(Color.FromRgb(250, 250, 250));

            var rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // === Header ===
            var header = new Border
            {
                Background = new LinearGradientBrush(
                    Color.FromRgb(33, 120, 216),
                    Color.FromRgb(25, 90, 170),
                    90),
                Padding = new Thickness(14, 10, 14, 10)
            };
            var headerGrid = new Grid();
            var modeLabel = new TextBlock
            {
                Foreground = Brushes.White,
                FontSize = 15,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };
            modeLabel.SetBinding(TextBlock.TextProperty, new Binding("ModeDisplayName"));

            var headerBtnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var rerunBtn = new Button
            {
                Content = "\u21BB Rerun",
                Padding = new Thickness(8, 3, 8, 3),
                Margin = new Thickness(0, 0, 4, 0),
                FontSize = 11,
                Background = new SolidColorBrush(Color.FromArgb(40, 255, 255, 255)),
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromArgb(80, 255, 255, 255)),
                Cursor = Cursors.Hand
            };
            rerunBtn.SetBinding(Button.CommandProperty, new Binding("RerunCommand"));

            var clearBtn = new Button
            {
                Content = "Clear",
                Padding = new Thickness(8, 3, 8, 3),
                FontSize = 11,
                Background = new SolidColorBrush(Color.FromArgb(40, 255, 255, 255)),
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromArgb(80, 255, 255, 255)),
                Cursor = Cursors.Hand
            };
            clearBtn.SetBinding(Button.CommandProperty, new Binding("ClearCommand"));

            headerBtnPanel.Children.Add(rerunBtn);
            headerBtnPanel.Children.Add(clearBtn);

            headerGrid.Children.Add(modeLabel);
            headerGrid.Children.Add(headerBtnPanel);
            header.Child = headerGrid;
            Grid.SetRow(header, 0);
            rootGrid.Children.Add(header);

            // === Messages ===
            _messagesControl = new ItemsControl();
            _messagesControl.ItemsPanel = new ItemsPanelTemplate(new FrameworkElementFactory(typeof(StackPanel)));
            _messagesControl.ItemTemplate = CreateMessageTemplate();
            _messagesControl.SetBinding(ItemsControl.ItemsSourceProperty, new Binding("Messages"));

            _scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(4, 8, 4, 8),
                Content = _messagesControl
            };
            Grid.SetRow(_scrollViewer, 1);
            rootGrid.Children.Add(_scrollViewer);

            // === Loading indicator ===
            var loadingBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(232, 240, 254)),
                Padding = new Thickness(14, 6, 14, 6),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 220, 245)),
                BorderThickness = new Thickness(0, 1, 0, 0)
            };
            var loadingPanel = new StackPanel { Orientation = Orientation.Horizontal };
            loadingPanel.Children.Add(new TextBlock
            {
                Text = "\u2022 \u2022 \u2022",
                FontSize = 16,
                Foreground = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                Margin = new Thickness(0, 0, 8, 0)
            });
            loadingPanel.Children.Add(new TextBlock
            {
                Text = "AI is thinking...",
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(80, 80, 80)),
                FontStyle = FontStyles.Italic,
                VerticalAlignment = VerticalAlignment.Center
            });
            loadingBorder.Child = loadingPanel;
            loadingBorder.SetBinding(VisibilityProperty,
                new Binding("IsLoading") { Converter = new BoolToVisibilityConverter() });
            Grid.SetRow(loadingBorder, 2);
            rootGrid.Children.Add(loadingBorder);

            // === Status bar ===
            var statusBar = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                Padding = new Thickness(10, 3, 10, 3),
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220)),
                BorderThickness = new Thickness(0, 1, 0, 0)
            };
            var statusText = new TextBlock
            {
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120))
            };
            statusText.SetBinding(TextBlock.TextProperty, new Binding("StatusText"));
            statusBar.Child = statusText;
            Grid.SetRow(statusBar, 3);
            rootGrid.Children.Add(statusBar);

            // === Input area ===
            var inputBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(8, 8, 8, 8)
            };
            var inputGrid = new Grid();
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _inputBox = new TextBox
            {
                Padding = new Thickness(10, 8, 10, 8),
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(Color.FromRgb(252, 252, 252))
            };
            _inputBox.SetBinding(TextBox.TextProperty,
                new Binding("InputText") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
            _inputBox.KeyDown += InputBox_KeyDown;
            Grid.SetColumn(_inputBox, 0);

            var sendBtn = new Button
            {
                Content = "Send",
                Margin = new Thickness(6, 0, 0, 0),
                Padding = new Thickness(16, 8, 16, 8),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            sendBtn.SetBinding(Button.CommandProperty, new Binding("SendCommand"));
            Grid.SetColumn(sendBtn, 1);

            inputGrid.Children.Add(_inputBox);
            inputGrid.Children.Add(sendBtn);
            inputBorder.Child = inputGrid;
            Grid.SetRow(inputBorder, 4);
            rootGrid.Children.Add(inputBorder);

            Content = rootGrid;
        }

        private DataTemplate CreateMessageTemplate()
        {
            var template = new DataTemplate(typeof(ChatMessageViewModel));
            var factory = new FrameworkElementFactory(typeof(ChatMessageBubble));
            template.VisualTree = factory;
            return template;
        }

        private void InputBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var vm = DataContext as ChatTaskPaneViewModel;
                if (vm?.SendCommand?.CanExecute(null) == true)
                {
                    vm.SendCommand.Execute(null);
                    e.Handled = true;
                }
            }
        }
    }

    public class ChatMessageBubble : Border
    {
        public ChatMessageBubble()
        {
            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // ReviewCorrectionViewModel 인 경우 교정 카드 렌더링
            if (DataContext is ReviewCorrectionViewModel rcvm)
            {
                RenderCorrectionCard(rcvm);
                return;
            }

            var msg = DataContext as ChatMessageViewModel;
            if (msg == null) return;

            CornerRadius = new CornerRadius(10);
            Padding = new Thickness(12);
            MaxWidth = 320;

            var panel = new StackPanel();

            if (msg.IsUser)
            {
                Background = new SolidColorBrush(Color.FromRgb(33, 120, 216));
                Margin = new Thickness(50, 3, 8, 3);
                HorizontalAlignment = HorizontalAlignment.Right;

                panel.Children.Add(new TextBlock
                {
                    Text = msg.Content,
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = 13,
                    Foreground = Brushes.White,
                    LineHeight = 20
                });
            }
            else if (msg.IsError)
            {
                Background = new SolidColorBrush(Color.FromRgb(253, 237, 237));
                BorderBrush = new SolidColorBrush(Color.FromRgb(239, 154, 154));
                BorderThickness = new Thickness(1);
                Margin = new Thickness(8, 3, 8, 3);

                panel.Children.Add(new TextBlock
                {
                    Text = msg.Content,
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.FromRgb(183, 28, 28)),
                    LineHeight = 18
                });
            }
            else // Assistant
            {
                Background = Brushes.White;
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230));
                BorderThickness = new Thickness(1);
                Margin = new Thickness(8, 3, 50, 3);
                HorizontalAlignment = HorizontalAlignment.Left;

                if (msg.ShowDiff && !string.IsNullOrEmpty(msg.OriginalText))
                {
                    var diffView = new DiffView();
                    diffView.SetTexts(msg.OriginalText, msg.Content);
                    diffView.MaxHeight = 400;
                    panel.Children.Add(diffView);
                }
                else
                {
                    var rendered = MarkdownRenderer.Render(msg.Content);
                    rendered.MaxHeight = 500;
                    panel.Children.Add(rendered);
                }

                // 버튼 패널
                var btnPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 8, 0, 0)
                };

                if (msg.ShowApplyButton)
                {
                    var applyBtn = new Button
                    {
                        Content = "\u2713  Apply",
                        Padding = new Thickness(10, 4, 10, 4),
                        FontSize = 11,
                        Background = new SolidColorBrush(Color.FromRgb(240, 245, 250)),
                        BorderBrush = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                        BorderThickness = new Thickness(1),
                        Foreground = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                        Cursor = Cursors.Hand,
                        Tag = msg.Content
                    };
                    applyBtn.Click += ApplyButton_Click;
                    btnPanel.Children.Add(applyBtn);
                }

                if (msg.ShowCopyButton || msg.ShowApplyButton)
                {
                    var copyBtn = new Button
                    {
                        Content = "\uD83D\uDCCB  Copy",
                        Padding = new Thickness(10, 4, 10, 4),
                        Margin = new Thickness(msg.ShowApplyButton ? 6 : 0, 0, 0, 0),
                        FontSize = 11,
                        Background = new SolidColorBrush(Color.FromRgb(245, 245, 245)),
                        BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                        BorderThickness = new Thickness(1),
                        Foreground = new SolidColorBrush(Color.FromRgb(80, 80, 80)),
                        Cursor = Cursors.Hand,
                        Tag = msg.Content
                    };
                    copyBtn.Click += CopyButton_Click;
                    btnPanel.Children.Add(copyBtn);
                }

                if (btnPanel.Children.Count > 0)
                    panel.Children.Add(btnPanel);
            }

            Child = panel;
        }

        private void RenderCorrectionCard(ReviewCorrectionViewModel rcvm)
        {
            CornerRadius = new CornerRadius(8);
            Margin = new Thickness(8, 4, 8, 4);
            Padding = new Thickness(0);
            BorderBrush = new SolidColorBrush(Color.FromRgb(200, 210, 230));
            BorderThickness = new Thickness(1);
            Background = Brushes.White;
            MaxWidth = 340;
            HorizontalAlignment = HorizontalAlignment.Left;

            var card = new StackPanel();

            // 삭제 영역 (원문)
            var origBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(255, 240, 240)),
                Padding = new Thickness(10, 6, 10, 6)
            };
            var origStack = new StackPanel();
            origStack.Children.Add(new TextBlock
            {
                Text = "Original",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(160, 80, 80)),
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 2)
            });
            origStack.Children.Add(new TextBlock
            {
                Text = rcvm.OriginalSnippet,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(140, 40, 40))
            });
            origBorder.Child = origStack;
            card.Children.Add(origBorder);

            // 교정문 영역
            var corrBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(235, 255, 240)),
                Padding = new Thickness(10, 6, 10, 6)
            };
            var corrStack = new StackPanel();
            corrStack.Children.Add(new TextBlock
            {
                Text = "Corrected",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 120, 50)),
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 2)
            });
            corrStack.Children.Add(new TextBlock
            {
                Text = rcvm.CorrectedSnippet,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(20, 130, 40))
            });
            corrBorder.Child = corrStack;
            card.Children.Add(corrBorder);

            // 이유
            var reasonBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                Padding = new Thickness(10, 4, 10, 4),
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
                BorderThickness = new Thickness(0, 1, 0, 0)
            };
            reasonBorder.Child = new TextBlock
            {
                Text = rcvm.ReasonText,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 11,
                FontStyle = FontStyles.Italic,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 100, 100))
            };
            card.Children.Add(reasonBorder);

            // Accept / Skip 버튼
            var btnBar = new Border
            {
                Padding = new Thickness(8, 6, 8, 6),
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
                BorderThickness = new Thickness(0, 1, 0, 0)
            };
            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal };

            var acceptBtn = new Button
            {
                Content = "\u2713 Accept",
                Padding = new Thickness(10, 3, 10, 3),
                FontSize = 11,
                Background = new SolidColorBrush(Color.FromRgb(230, 248, 235)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(40, 160, 60)),
                BorderThickness = new Thickness(1),
                Foreground = new SolidColorBrush(Color.FromRgb(20, 130, 40)),
                Cursor = Cursors.Hand,
                Tag = rcvm
            };
            acceptBtn.Click += AcceptCorrection_Click;

            var skipBtn = new Button
            {
                Content = "\u2717 Skip",
                Padding = new Thickness(10, 3, 10, 3),
                Margin = new Thickness(6, 0, 0, 0),
                FontSize = 11,
                Background = new SolidColorBrush(Color.FromRgb(245, 245, 245)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                BorderThickness = new Thickness(1),
                Foreground = new SolidColorBrush(Color.FromRgb(100, 100, 100)),
                Cursor = Cursors.Hand,
                Tag = rcvm
            };
            skipBtn.Click += SkipCorrection_Click;

            btnPanel.Children.Add(acceptBtn);
            btnPanel.Children.Add(skipBtn);
            btnBar.Child = btnPanel;
            card.Children.Add(btnBar);

            Child = card;
        }

        private void AcceptCorrection_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is ReviewCorrectionViewModel rcvm)
            {
                var element = this as FrameworkElement;
                while (element != null)
                {
                    if (element is ChatTaskPane pane && pane.DataContext is ChatTaskPaneViewModel vm)
                    {
                        vm.ApplyCorrection(rcvm.Correction);
                        // 카드를 "적용됨" 상태로 변경
                        Background = new SolidColorBrush(Color.FromRgb(235, 248, 240));
                        Opacity = 0.7;
                        IsEnabled = false;
                        break;
                    }
                    element = VisualTreeHelper.GetParent(element) as FrameworkElement;
                }
            }
        }

        private void SkipCorrection_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is ReviewCorrectionViewModel rcvm)
            {
                rcvm.Correction.Skipped = true;
                Background = new SolidColorBrush(Color.FromRgb(245, 245, 245));
                Opacity = 0.5;
                IsEnabled = false;
            }
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string content)
            {
                try
                {
                    System.Windows.Clipboard.SetText(content);
                    btn.Content = "\u2713  Copied!";
                }
                catch { }
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string content)
            {
                var element = this as FrameworkElement;
                while (element != null)
                {
                    if (element is ChatTaskPane pane && pane.DataContext is ChatTaskPaneViewModel vm)
                    {
                        vm.ApplyResult(content);
                        break;
                    }
                    element = VisualTreeHelper.GetParent(element) as FrameworkElement;
                }
            }
        }
    }

    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b && b ? Visibility.Visible : Visibility.Collapsed;

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }
}
