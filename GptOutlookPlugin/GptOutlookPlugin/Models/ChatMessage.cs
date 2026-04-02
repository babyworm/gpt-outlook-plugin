using System;

namespace GptOutlookPlugin.Models
{
    public enum ChatRole
    {
        System,
        User,
        Assistant
    }

    public class ChatMessage
    {
        public ChatRole Role { get; }
        public string Content { get; }
        public DateTime Timestamp { get; }

        public ChatMessage(ChatRole role, string content)
        {
            Role = role;
            Content = content;
            Timestamp = DateTime.UtcNow;
        }

        public string RoleString => Role.ToString().ToLower();
    }
}
