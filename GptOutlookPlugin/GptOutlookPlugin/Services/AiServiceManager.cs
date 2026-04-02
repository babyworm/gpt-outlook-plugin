using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Services
{
    public class AiServiceManager
    {
        private readonly IAiProvider _primary;
        private readonly IAiProvider _fallback;

        public event Action<string> OnProviderSwitch;

        public AiServiceManager(IAiProvider primary, IAiProvider fallback)
        {
            _primary = primary;
            _fallback = fallback;
        }

        public async Task<string> SendAsync(List<ChatMessage> messages, CancellationToken ct)
        {
            if (_primary.IsAvailable())
            {
                try
                {
                    return await _primary.SendAsync(messages, ct);
                }
                catch (Exception)
                {
                    OnProviderSwitch?.Invoke($"Switched from {_primary.Name} to {_fallback.Name}");
                }
            }

            if (_fallback.IsAvailable())
                return await _fallback.SendAsync(messages, ct);

            throw new InvalidOperationException(
                "No AI provider available. Check Codex CLI (WSL) or configure an OpenAI API key.");
        }
    }
}
