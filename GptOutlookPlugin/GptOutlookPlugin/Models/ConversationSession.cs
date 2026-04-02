using System.Collections.Generic;

namespace GptOutlookPlugin.Models
{
    public class ConversationSession
    {
        public string SessionKey { get; }
        public FeatureMode CurrentMode { get; private set; }
        public EmailContext EmailContext { get; set; }
        public List<ChatMessage> Messages { get; } = new List<ChatMessage>();

        private readonly int _maxHistory;

        public ConversationSession(string sessionKey, FeatureMode mode, int maxHistory = 10)
        {
            SessionKey = sessionKey;
            CurrentMode = mode;
            _maxHistory = maxHistory;
        }

        public void AddMessage(ChatRole role, string content)
        {
            Messages.Add(new ChatMessage(role, content));
            TrimHistory();
        }

        public void SwitchMode(FeatureMode newMode)
        {
            CurrentMode = newMode;
        }

        public void Clear()
        {
            Messages.Clear();
        }

        private void TrimHistory()
        {
            while (Messages.Count > _maxHistory)
                Messages.RemoveAt(0);
        }
    }
}
