using System;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using GptOutlookPlugin.Core;

namespace GptOutlookPlugin.UI
{
    public class SettingsWindow : Window
    {
        private readonly SettingsManager _settingsManager;
        private readonly RadioButton _rbCodex;
        private readonly RadioButton _rbOpenAi;
        private readonly Slider _sliderTimeout;
        private readonly Label _lblTimeout;
        private readonly PasswordBox _txtApiKey;
        private readonly ComboBox _cmbModel;
        private readonly Slider _sliderHistory;
        private readonly Label _lblHistory;
        private readonly ComboBox _cmbLanguage;
        private readonly ComboBox _cmbTone;
        private readonly TextBox _txtCustomTone;
        private readonly TextBlock _txtCodexStatus;

        public SettingsWindow(SettingsManager settingsManager)
        {
            _settingsManager = settingsManager;

            Title = "GPT Email Assistant - Settings";
            Width = 480;
            Height = 560;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(248, 248, 248));

            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(20)
            };
            var root = new StackPanel();

            // === Title ===
            root.Children.Add(new TextBlock
            {
                Text = "GPT Email Assistant",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                Margin = new Thickness(0, 0, 0, 16)
            });

            // === AI Provider ===
            root.Children.Add(CreateSectionHeader("AI Provider"));
            var providerPanel = new StackPanel { Margin = new Thickness(8, 4, 0, 12) };
            _rbCodex = new RadioButton
            {
                Content = "Codex CLI (WSL) — recommended",
                Margin = new Thickness(0, 4, 0, 4),
                FontSize = 13
            };
            _rbOpenAi = new RadioButton
            {
                Content = "OpenAI API (direct)",
                Margin = new Thickness(0, 4, 0, 4),
                FontSize = 13
            };
            providerPanel.Children.Add(_rbCodex);
            providerPanel.Children.Add(_rbOpenAi);
            root.Children.Add(providerPanel);

            // === Codex CLI Status (read-only) ===
            root.Children.Add(CreateSectionHeader("Codex CLI Status"));
            var codexStatusBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(240, 244, 248)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 215, 230)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 4, 0, 12)
            };
            _txtCodexStatus = new TextBlock
            {
                Text = "Detecting...",
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 60))
            };
            codexStatusBorder.Child = _txtCodexStatus;
            root.Children.Add(codexStatusBorder);

            // === Codex Timeout ===
            root.Children.Add(CreateSectionHeader("Codex CLI Timeout"));
            var timeoutPanel = new Grid { Margin = new Thickness(0, 4, 0, 12) };
            timeoutPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            timeoutPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });

            _sliderTimeout = new Slider
            {
                Minimum = 30, Maximum = 300, TickFrequency = 30,
                IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center
            };
            _sliderTimeout.ValueChanged += (s, e) =>
                _lblTimeout.Content = $"{(int)_sliderTimeout.Value}s";
            Grid.SetColumn(_sliderTimeout, 0);

            _lblTimeout = new Label
            {
                HorizontalContentAlignment = HorizontalAlignment.Right,
                FontWeight = FontWeights.SemiBold,
                FontSize = 13
            };
            Grid.SetColumn(_lblTimeout, 1);

            timeoutPanel.Children.Add(_sliderTimeout);
            timeoutPanel.Children.Add(_lblTimeout);
            root.Children.Add(timeoutPanel);

            // === OpenAI API ===
            root.Children.Add(CreateSectionHeader("OpenAI API (Fallback)"));
            var apiGrid = CreateLabeledGrid();
            _txtApiKey = new PasswordBox { Padding = new Thickness(6), FontSize = 13 };
            _cmbModel = new ComboBox { Padding = new Thickness(4), FontSize = 13 };
            _cmbModel.Items.Add(new ComboBoxItem { Content = "gpt-4o", Tag = "gpt-4o" });
            _cmbModel.Items.Add(new ComboBoxItem { Content = "gpt-4o-mini", Tag = "gpt-4o-mini" });
            _cmbModel.Items.Add(new ComboBoxItem { Content = "gpt-4.1", Tag = "gpt-4.1" });
            _cmbModel.Items.Add(new ComboBoxItem { Content = "gpt-4.1-mini", Tag = "gpt-4.1-mini" });
            _cmbModel.Items.Add(new ComboBoxItem { Content = "o3-mini", Tag = "o3-mini" });
            AddRow(apiGrid, 0, "API Key:", _txtApiKey);
            AddRow(apiGrid, 1, "Model:", _cmbModel);
            root.Children.Add(apiGrid);

            // === Context / History ===
            root.Children.Add(CreateSectionHeader("Conversation History"));
            var histPanel = new Grid { Margin = new Thickness(0, 4, 0, 12) };
            histPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            histPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });

            _sliderHistory = new Slider
            {
                Minimum = 3, Maximum = 30, TickFrequency = 1,
                IsSnapToTickEnabled = true, VerticalAlignment = VerticalAlignment.Center
            };
            _sliderHistory.ValueChanged += (s, e) =>
                _lblHistory.Content = $"{(int)_sliderHistory.Value} msgs";
            Grid.SetColumn(_sliderHistory, 0);

            _lblHistory = new Label
            {
                HorizontalContentAlignment = HorizontalAlignment.Right,
                FontWeight = FontWeights.SemiBold,
                FontSize = 13
            };
            Grid.SetColumn(_lblHistory, 1);

            histPanel.Children.Add(_sliderHistory);
            histPanel.Children.Add(_lblHistory);
            root.Children.Add(histPanel);

            // === Translation Default ===
            root.Children.Add(CreateSectionHeader("Default Translation Language"));
            _cmbLanguage = new ComboBox
            {
                Padding = new Thickness(4), FontSize = 13,
                Margin = new Thickness(0, 4, 0, 12)
            };
            _cmbLanguage.Items.Add(new ComboBoxItem { Content = "Korean (한국어)", Tag = "Korean" });
            _cmbLanguage.Items.Add(new ComboBoxItem { Content = "English", Tag = "English" });
            _cmbLanguage.Items.Add(new ComboBoxItem { Content = "Japanese (日本語)", Tag = "Japanese" });
            _cmbLanguage.Items.Add(new ComboBoxItem { Content = "Chinese (中文)", Tag = "Chinese" });
            _cmbLanguage.Items.Add(new ComboBoxItem { Content = "Spanish (Español)", Tag = "Spanish" });
            root.Children.Add(_cmbLanguage);

            // === Writing Tone ===
            root.Children.Add(CreateSectionHeader("Writing Tone"));
            _cmbTone = new ComboBox
            {
                Padding = new Thickness(4), FontSize = 13,
                Margin = new Thickness(0, 4, 0, 4)
            };
            _cmbTone.Items.Add(new ComboBoxItem { Content = "Professional and polite (정중하고 격식 있는)", Tag = "Professional and polite" });
            _cmbTone.Items.Add(new ComboBoxItem { Content = "Friendly but professional (친근하지만 프로)", Tag = "Friendly but professional" });
            _cmbTone.Items.Add(new ComboBoxItem { Content = "Direct and concise (직접적이고 간결한)", Tag = "Direct and concise" });
            _cmbTone.Items.Add(new ComboBoxItem { Content = "Casual (캐주얼한)", Tag = "Casual" });
            _cmbTone.Items.Add(new ComboBoxItem { Content = "Formal (격식체)", Tag = "Formal" });
            _cmbTone.Items.Add(new ComboBoxItem { Content = "Custom (직접 입력)", Tag = "Custom" });
            _cmbTone.SelectionChanged += (s, ev) =>
            {
                var selected = _cmbTone.SelectedItem as ComboBoxItem;
                _txtCustomTone.Visibility = selected?.Tag?.ToString() == "Custom"
                    ? Visibility.Visible : Visibility.Collapsed;
            };
            root.Children.Add(_cmbTone);

            _txtCustomTone = new TextBox
            {
                FontSize = 13, Padding = new Thickness(6),
                Height = 60, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12),
                Visibility = Visibility.Collapsed,
                ToolTip = "예: \"기술적이면서도 이해하기 쉽게, 한국어 존대말 사용\""
            };
            root.Children.Add(_txtCustomTone);

            // === Buttons ===
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };
            var saveBtn = new Button
            {
                Content = "Save",
                Width = 90, Padding = new Thickness(6),
                FontSize = 13, FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(Color.FromRgb(33, 120, 216)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Margin = new Thickness(0, 0, 8, 0)
            };
            saveBtn.Click += SaveButton_Click;
            var cancelBtn = new Button
            {
                Content = "Cancel",
                Width = 90, Padding = new Thickness(6),
                FontSize = 13
            };
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            btnPanel.Children.Add(saveBtn);
            btnPanel.Children.Add(cancelBtn);
            root.Children.Add(btnPanel);

            scroll.Content = root;
            Content = scroll;

            LoadSettings();
            LoadCodexStatus();
        }

        private void LoadSettings()
        {
            var s = _settingsManager.Current;
            _rbCodex.IsChecked = s.AiProvider == "codex-cli";
            _rbOpenAi.IsChecked = s.AiProvider == "openai-api";
            _sliderTimeout.Value = s.CodexCli.TimeoutSeconds;
            _lblTimeout.Content = $"{s.CodexCli.TimeoutSeconds}s";
            _txtApiKey.Password = s.OpenAiApi.ApiKey;
            _sliderHistory.Value = s.Context.MaxHistoryMessages;
            _lblHistory.Content = $"{s.Context.MaxHistoryMessages} msgs";

            SelectComboByTag(_cmbModel, s.OpenAiApi.Model);
            SelectComboByTag(_cmbLanguage, s.DefaultTranslateTarget);

            if (!string.IsNullOrEmpty(s.CustomTonePrompt))
            {
                SelectComboByTag(_cmbTone, "Custom");
                _txtCustomTone.Text = s.CustomTonePrompt;
                _txtCustomTone.Visibility = Visibility.Visible;
            }
            else
            {
                SelectComboByTag(_cmbTone, s.DefaultTone);
            }
        }

        private void LoadCodexStatus()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "wsl.exe",
                    Arguments = "bash -c \"grep -E '^(model |model_reasoning)' ~/.codex/config.toml 2>/dev/null\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };
                using (var proc = Process.Start(psi))
                {
                    var output = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit(3000);

                    if (!string.IsNullOrWhiteSpace(output))
                        _txtCodexStatus.Text = output.Trim().Replace("\"", "");
                    else
                        _txtCodexStatus.Text = "Codex CLI config not found";
                }
            }
            catch
            {
                _txtCodexStatus.Text = "Could not read Codex config";
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var s = _settingsManager.Current;
            s.AiProvider = _rbOpenAi.IsChecked == true ? "openai-api" : "codex-cli";
            s.CodexCli.TimeoutSeconds = (int)_sliderTimeout.Value;
            s.OpenAiApi.ApiKey = _txtApiKey.Password;
            s.Context.MaxHistoryMessages = (int)_sliderHistory.Value;

            if (_cmbModel.SelectedItem is ComboBoxItem modelItem)
                s.OpenAiApi.Model = modelItem.Tag?.ToString() ?? "gpt-4o";
            if (_cmbLanguage.SelectedItem is ComboBoxItem langItem)
                s.DefaultTranslateTarget = langItem.Tag?.ToString() ?? "Korean";

            if (_cmbTone.SelectedItem is ComboBoxItem toneItem)
            {
                var tag = toneItem.Tag?.ToString() ?? "Professional and polite";
                if (tag == "Custom")
                {
                    s.DefaultTone = "Custom";
                    s.CustomTonePrompt = _txtCustomTone.Text.Trim();
                }
                else
                {
                    s.DefaultTone = tag;
                    s.CustomTonePrompt = "";
                }
            }

            _settingsManager.Save();
            DialogResult = true;
            Close();
        }

        private static void SelectComboByTag(ComboBox combo, string tag)
        {
            foreach (ComboBoxItem item in combo.Items)
            {
                if (item.Tag?.ToString() == tag)
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
            if (combo.Items.Count > 0)
                combo.SelectedIndex = 0;
        }

        private static TextBlock CreateSectionHeader(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(80, 80, 80)),
                Margin = new Thickness(0, 4, 0, 2)
            };
        }

        private static Grid CreateLabeledGrid()
        {
            var grid = new Grid { Margin = new Thickness(0, 4, 0, 12) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            return grid;
        }

        private static void AddRow(Grid grid, int row, string label, UIElement control)
        {
            var lbl = new TextBlock
            {
                Text = label,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 13,
                Margin = new Thickness(0, 4, 8, 4)
            };
            Grid.SetRow(lbl, row);
            Grid.SetColumn(lbl, 0);
            Grid.SetRow(control as FrameworkElement, row);
            Grid.SetColumn(control as FrameworkElement, 1);
            (control as FrameworkElement).Margin = new Thickness(0, 4, 0, 4);
            grid.Children.Add(lbl);
            grid.Children.Add(control);
        }
    }
}
