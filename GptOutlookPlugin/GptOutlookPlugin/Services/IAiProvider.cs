using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Services
{
    public interface IAiProvider
    {
        Task<string> SendAsync(List<ChatMessage> messages, CancellationToken ct);
        bool IsAvailable();
        string Name { get; }
    }
}
