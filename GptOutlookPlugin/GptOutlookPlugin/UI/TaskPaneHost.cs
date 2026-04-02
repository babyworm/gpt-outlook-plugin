using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GptOutlookPlugin.UI
{
    /// <summary>
    /// WinForms UserControl that hosts the WPF ChatTaskPane via ElementHost.
    /// Required because VSTO's CustomTaskPane only accepts WinForms controls.
    /// </summary>
    public class TaskPaneHost : UserControl
    {
        private readonly ElementHost _elementHost;
        private readonly ChatTaskPane _chatPane;

        public ChatTaskPaneViewModel ViewModel { get; }

        public TaskPaneHost(ChatTaskPaneViewModel viewModel)
        {
            ViewModel = viewModel;

            _chatPane = new ChatTaskPane
            {
                DataContext = viewModel
            };

            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _chatPane
            };

            Controls.Add(_elementHost);
            Dock = DockStyle.Fill;
        }
    }
}
