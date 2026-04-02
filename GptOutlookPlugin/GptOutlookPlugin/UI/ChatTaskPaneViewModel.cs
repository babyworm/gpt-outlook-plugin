using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;

namespace GptOutlookPlugin.UI
{
    public class ChatTaskPaneViewModel : INotifyPropertyChanged
    {
        private readonly ContextManager _contextManager;
        private readonly AiServiceManager _aiService;
        private readonly OutlookInterop _outlook;
        private ConversationSession _currentSession;

        public ObservableCollection<ChatMessageViewModel> Messages { get; }
            = new ObservableCollection<ChatMessageViewModel>();

        private string _inputText = "";
        public string InputText
        {
            get => _inputText;
            set { _inputText = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSend)); }
        }

        private bool _isLoading;
        public bool IsLoading
        {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSend)); }
        }

        private FeatureMode _currentMode = FeatureMode.Review;
        public FeatureMode CurrentMode
        {
            get => _currentMode;
            set { _currentMode = value; OnPropertyChanged(); OnPropertyChanged(nameof(ModeDisplayName)); }
        }

        public string ModeDisplayName
        {
            get
            {
                switch (CurrentMode)
                {
                    case FeatureMode.Review: return "Review / Proofread";
                    case FeatureMode.Compose: return "Compose";
                    case FeatureMode.AutoReply: return "Auto Reply";
                    case FeatureMode.Translate: return "Translate";
                    case FeatureMode.Summarize: return "Summarize";
                    default: return CurrentMode.ToString();
                }
            }
        }

        private string _statusText = "Ready";
        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(); }
        }

        public bool CanSend => !string.IsNullOrWhiteSpace(InputText) && !IsLoading;

        public ICommand SendCommand { get; }
        public ICommand ClearCommand { get; }

        private string _userLanguage = "Korean";
        private string _userName = "";
        private string _userEmail = "";
        private string _tone = "Professional and polite";

        public ChatTaskPaneViewModel(ContextManager contextManager, AiServiceManager aiService, OutlookInterop outlook)
        {
            _contextManager = contextManager;
            _aiService = aiService;
            _outlook = outlook;

            SendCommand = new RelayCommand(async _ => await SendMessageAsync(), _ => CanSend);
            ClearCommand = new RelayCommand(_ => ClearConversation());

            _aiService.OnProviderSwitch += msg => StatusText = msg;

            // Office locale 및 사용자 정보 감지
            try { _userLanguage = _outlook.GetUserLanguage(); } catch { }
            try { _userName = _outlook.GetUserDisplayName(); } catch { }
            try { _userEmail = _outlook.GetUserEmailAddress(); } catch { }
        }

        /// <summary>
        /// Settings에서 톤 설정이 변경되면 호출.
        /// </summary>
        public void UpdateTone(string tone)
        {
            _tone = tone;
        }

        public void StartMode(FeatureMode mode, string initialPrompt = null)
        {
            CurrentMode = mode;

            var emailKey = _outlook.GetCurrentEntryIdOrTemp();
            var emailCtx = _outlook.GetCurrentEmailContext();

            // 이메일 + 모드별 독립 세션
            _currentSession = _contextManager.GetOrCreateSession(emailKey, mode);
            _currentSession.EmailContext = emailCtx;

            RefreshMessages();

            if (initialPrompt != null)
                _ = SendWithTextAsync(initialPrompt);
            else if (_currentSession.Messages.Count == 0)
                _ = SendAutoPromptAsync(mode);
        }

        public void StartWithSelection(FeatureMode mode, string selectedText)
        {
            CurrentMode = mode;

            var emailKey = _outlook.GetCurrentEntryIdOrTemp();
            var emailCtx = _outlook.GetCurrentEmailContext() ?? new EmailContext();
            emailCtx.Body = selectedText;

            _currentSession = _contextManager.GetOrCreateSession(emailKey + "-sel", mode);
            _currentSession.EmailContext = emailCtx;
            _currentSession.Clear();

            RefreshMessages();
            _ = SendAutoPromptAsync(mode);
        }

        private async Task SendAutoPromptAsync(FeatureMode mode)
        {
            // AutoReply: 즉시 자동 답장
            if (mode == FeatureMode.AutoReply)
            {
                await SendWithTextAsync("이 이메일에 대한 답장 초안을 작성해 주세요.");
                return;
            }

            // Compose 모드: 자동 실행 대신 사용자에게 먼저 질문
            if (mode == FeatureMode.Compose)
            {
                var isReply = _currentSession.EmailContext != null
                    && !string.IsNullOrWhiteSpace(_currentSession.EmailContext.Body);

                var guideMsg = isReply
                    ? "답장을 작성합니다. 어떤 내용으로 답장할까요?\n(예: \"미팅 수락\", \"정중하게 거절\", \"일정 조율 요청\")"
                    : "새 이메일을 작성합니다. 어떤 내용으로 만들까요?\n(예: \"프로젝트 진행 상황 공유\", \"미팅 요청\")";

                Messages.Add(new ChatMessageViewModel(ChatRole.Assistant, guideMsg));
                return;
            }

            string prompt;
            switch (mode)
            {
                case FeatureMode.Review:
                    prompt = "이 이메일을 리뷰하고 교정해 주세요.";
                    break;
                case FeatureMode.Translate:
                    prompt = $"이 이메일을 {_userLanguage}로 번역해 주세요. 번역문만 출력하세요.";
                    break;
                case FeatureMode.Summarize:
                    // 스레드 여부에 따라 가이드 메시지 표시
                    var hasThread = _currentSession.EmailContext != null
                        && _currentSession.EmailContext.Body.Contains("From:");
                    if (hasThread)
                    {
                        Messages.Add(new ChatMessageViewModel(ChatRole.Assistant,
                            "무엇을 요약할까요?\n\n"
                            + "1️⃣ 이 메일만 요약\n"
                            + "2️⃣ 전체 스레드 요약\n\n"
                            + "번호를 입력하거나 직접 요청해 주세요."));
                        return;
                    }
                    prompt = "이 이메일을 요약해 주세요.";
                    break;
                default:
                    return;
            }
            await SendWithTextAsync(prompt);
        }

        private async Task SendWithTextAsync(string text)
        {
            InputText = text;
            await SendMessageAsync();
        }

        private async Task SendMessageAsync()
        {
            if (_currentSession == null || string.IsNullOrWhiteSpace(InputText)) return;

            var dispatcher = Dispatcher.CurrentDispatcher;
            var userText = InputText.Trim();
            InputText = "";

            _currentSession.AddMessage(ChatRole.User, userText);
            Messages.Add(new ChatMessageViewModel(ChatRole.User, userText));

            IsLoading = true;
            StatusText = "AI processing...";

            try
            {
                var allMessages = _contextManager.BuildMessages(
                    _currentSession, _userLanguage, _userLanguage, _userName, _userEmail, _tone);
                var response = await Task.Run(() =>
                    _aiService.SendAsync(allMessages, CancellationToken.None));

                dispatcher.Invoke(() =>
                {
                    _currentSession.AddMessage(ChatRole.Assistant, response);

                    var showDiff = CurrentMode == FeatureMode.Review;
                    Messages.Add(new ChatMessageViewModel(ChatRole.Assistant, response,
                        showDiff: showDiff,
                        originalText: _currentSession.EmailContext?.Body,
                        mode: CurrentMode));

                    StatusText = "Ready";
                });
            }
            catch (Exception ex)
            {
                dispatcher.Invoke(() =>
                {
                    Messages.Add(new ChatMessageViewModel(ChatRole.System, $"Error: {ex.Message}"));
                    StatusText = $"Error: {ex.Message}";
                });
            }
            finally
            {
                dispatcher.Invoke(() => IsLoading = false);
            }
        }

        public void ApplyResult(string resultText)
        {
            switch (CurrentMode)
            {
                case FeatureMode.Review:
                case FeatureMode.Translate:
                    _outlook.ApplyToBody(resultText);
                    StatusText = "Applied to email body.";
                    break;
                case FeatureMode.Compose:
                case FeatureMode.AutoReply:
                    // "Subject: ..." 패턴에서 제목과 본문 분리
                    string subject = null;
                    string body = resultText;

                    if (resultText.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                    {
                        var firstNewline = resultText.IndexOf('\n');
                        if (firstNewline > 0)
                        {
                            subject = resultText.Substring(8, firstNewline - 8).Trim();
                            body = resultText.Substring(firstNewline).Trim();
                        }
                    }

                    _outlook.ApplyComposeDraft(body, subject);
                    StatusText = "Draft applied.";
                    break;
            }
        }

        private void ClearConversation()
        {
            if (_currentSession != null)
            {
                _contextManager.ClearSession(_currentSession.SessionKey);
                _currentSession = null;
            }
            Messages.Clear();
            StatusText = "Conversation cleared.";
        }

        private void RefreshMessages()
        {
            Messages.Clear();
            if (_currentSession != null)
            {
                foreach (var msg in _currentSession.Messages)
                    Messages.Add(new ChatMessageViewModel(msg.Role, msg.Content));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class ChatMessageViewModel
    {
        public ChatRole Role { get; }
        public string Content { get; }
        public bool IsUser => Role == ChatRole.User;
        public bool IsAssistant => Role == ChatRole.Assistant;
        public bool IsError => Role == ChatRole.System;
        public bool ShowDiff { get; }
        public string OriginalText { get; }
        public FeatureMode Mode { get; }

        /// <summary>
        /// Apply 버튼: Compose/AutoReply/Review에서만 표시
        /// </summary>
        public bool ShowApplyButton => IsAssistant && !IsError
            && (Mode == FeatureMode.Compose || Mode == FeatureMode.AutoReply || Mode == FeatureMode.Review);

        /// <summary>
        /// Copy 버튼: Translate/Summarize에서 표시
        /// </summary>
        public bool ShowCopyButton => IsAssistant && !IsError
            && (Mode == FeatureMode.Translate || Mode == FeatureMode.Summarize);

        public ChatMessageViewModel(ChatRole role, string content, bool showDiff = false,
            string originalText = null, FeatureMode mode = FeatureMode.Review)
        {
            Role = role;
            Content = content;
            ShowDiff = showDiff;
            OriginalText = originalText;
            Mode = mode;
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object parameter) => _execute(parameter);
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
