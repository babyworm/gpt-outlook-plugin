using System.Collections.Generic;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public class ContextManager
    {
        private readonly Dictionary<string, ConversationSession> _sessions
            = new Dictionary<string, ConversationSession>();
        private readonly int _maxHistory;

        public ContextManager(int maxHistory = 10)
        {
            _maxHistory = maxHistory;
        }

        /// <summary>
        /// 이메일 + 모드별로 독립 세션 관리.
        /// 같은 이메일이라도 Review와 Translate는 별도 세션.
        /// </summary>
        public ConversationSession GetOrCreateSession(string emailKey, FeatureMode mode)
        {
            var sessionKey = $"{emailKey}:{mode}";

            if (_sessions.TryGetValue(sessionKey, out var existing))
                return existing;

            var session = new ConversationSession(sessionKey, mode, _maxHistory);
            _sessions[sessionKey] = session;
            return session;
        }

        public void ClearSession(string sessionKey)
        {
            _sessions.Remove(sessionKey);
        }

        public List<ChatMessage> BuildMessages(ConversationSession session,
            string targetLanguage = "Korean", string userLanguage = "Korean",
            string userName = "", string userEmail = "",
            string tone = "Professional and polite",
            string reviewSensitivity = "Medium")
        {
            var messages = new List<ChatMessage>();

            var systemPrompt = PromptTemplates.GetSystemPrompt(
                session.CurrentMode,
                session.EmailContext ?? new EmailContext(),
                targetLanguage,
                userLanguage,
                userName,
                userEmail,
                tone,
                reviewSensitivity);
            messages.Add(new ChatMessage(ChatRole.System, systemPrompt));

            messages.AddRange(session.Messages);

            return messages;
        }
    }
}
