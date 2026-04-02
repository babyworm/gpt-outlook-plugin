using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;
using GptOutlookPlugin.UI;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace GptOutlookPlugin
{
    public partial class ThisAddIn
    {
        private SettingsManager _settingsManager;
        private ContextManager _contextManager;
        private AiServiceManager _aiService;
        private GptRibbon _ribbon;

        // Explorer (메인 창) 용 TaskPane
        private CustomTaskPane _explorerTaskPane;
        private ChatTaskPaneViewModel _explorerViewModel;

        // Inspector (작성/읽기 창) 별 TaskPane 관리
        private Dictionary<Outlook.Inspector, CustomTaskPane> _inspectorPanes
            = new Dictionary<Outlook.Inspector, CustomTaskPane>();
        private Dictionary<Outlook.Inspector, ChatTaskPaneViewModel> _inspectorViewModels
            = new Dictionary<Outlook.Inspector, ChatTaskPaneViewModel>();

        private Outlook.Inspectors _inspectors;

        public OutlookInterop OutlookInterop { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _settingsManager = new SettingsManager();
            var settings = _settingsManager.Current;

            OutlookInterop = new OutlookInterop(Application);
            _contextManager = new ContextManager(settings.Context.MaxHistoryMessages);

            var codexProvider = new CodexCliProvider(settings.CodexCli);
            var openAiProvider = new OpenAiApiProvider(settings.OpenAiApi);

            _aiService = settings.AiProvider == "openai-api"
                ? new AiServiceManager(openAiProvider, codexProvider)
                : new AiServiceManager(codexProvider, openAiProvider);

            // Explorer TaskPane
            _explorerViewModel = new ChatTaskPaneViewModel(_contextManager, _aiService, OutlookInterop);
            var explorerHost = new TaskPaneHost(_explorerViewModel);
            _explorerTaskPane = CustomTaskPanes.Add(explorerHost, "GPT Email Assistant");
            _explorerTaskPane.Width = 380;
            _explorerTaskPane.Visible = false;

            // Inspector 이벤트 감시
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += OnNewInspector;

            // Ribbon 이벤트 연결
            if (_ribbon != null)
                WireRibbonEvents();
        }

        private void WireRibbonEvents()
        {
            _ribbon.OnModeRequested += mode =>
            {
                var (vm, pane) = GetActiveViewModelAndPane();
                vm.StartMode(mode);
                pane.Visible = true;
            };
            _ribbon.OnSelectionModeRequested += (mode, text) =>
            {
                var (vm, pane) = GetActiveViewModelAndPane();
                vm.StartWithSelection(mode, text);
                pane.Visible = true;
            };
            _ribbon.OnSettingsRequested += () =>
            {
                var settingsWindow = new SettingsWindow(_settingsManager);
                if (settingsWindow.ShowDialog() == true)
                {
                    var s = _settingsManager.Current;
                    var tone = !string.IsNullOrEmpty(s.CustomTonePrompt) ? s.CustomTonePrompt : s.DefaultTone;
                    _explorerViewModel.UpdateTone(tone);
                    foreach (var vm in _inspectorViewModels.Values)
                        vm.UpdateTone(tone);
                }
            };
        }

        /// <summary>
        /// 현재 활성 창(Inspector 또는 Explorer)에 맞는 ViewModel과 TaskPane을 반환.
        /// </summary>
        private (ChatTaskPaneViewModel vm, CustomTaskPane pane) GetActiveViewModelAndPane()
        {
            var inspector = Application.ActiveInspector();
            if (inspector != null && _inspectorPanes.TryGetValue(inspector, out var pane))
            {
                return (_inspectorViewModels[inspector], pane);
            }
            return (_explorerViewModel, _explorerTaskPane);
        }

        private void OnNewInspector(Outlook.Inspector inspector)
        {
            // Inspector가 열릴 때 TaskPane 생성
            ((Outlook.InspectorEvents_10_Event)inspector).Close += () => OnInspectorClose(inspector);

            var vm = new ChatTaskPaneViewModel(_contextManager, _aiService, OutlookInterop);
            var host = new TaskPaneHost(vm);
            var pane = CustomTaskPanes.Add(host, "GPT Email Assistant", inspector);
            pane.Width = 380;
            pane.Visible = false;

            _inspectorPanes[inspector] = pane;
            _inspectorViewModels[inspector] = vm;
        }

        private void OnInspectorClose(Outlook.Inspector inspector)
        {
            if (_inspectorPanes.TryGetValue(inspector, out var pane))
            {
                try { CustomTaskPanes.Remove(pane); } catch { }
                _inspectorPanes.Remove(inspector);
                _inspectorViewModels.Remove(inspector);
            }
        }

        protected override object RequestService(Guid serviceGuid)
        {
            var ribbonGuid = typeof(Office.IRibbonExtensibility).GUID;
            if (serviceGuid == ribbonGuid)
            {
                if (_ribbon == null)
                    _ribbon = new GptRibbon();
                return _ribbon;
            }
            return base.RequestService(serviceGuid);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _inspectorPanes.Clear();
            _inspectorViewModels.Clear();
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
