using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Services
{
    public class OpenAiApiProvider : IAiProvider
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private readonly OpenAiApiSettings _settings;

        public string Name => "OpenAI API";

        public OpenAiApiProvider(OpenAiApiSettings settings)
        {
            _settings = settings;
        }

        public bool IsAvailable()
        {
            return !string.IsNullOrWhiteSpace(_settings.ApiKey);
        }

        public async Task<string> SendAsync(List<ChatMessage> messages, CancellationToken ct)
        {
            var requestBody = new
            {
                model = _settings.Model,
                max_tokens = _settings.MaxTokens,
                messages = messages.Select(m => new
                {
                    role = m.RoleString,
                    content = m.Content
                }).ToArray()
            };

            var json = JsonConvert.SerializeObject(requestBody);
            var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions")
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
            request.Headers.Add("Authorization", $"Bearer {_settings.ApiKey}");

            var response = await HttpClient.SendAsync(request, ct);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException($"OpenAI API error ({response.StatusCode}): {responseBody}");

            var parsed = JObject.Parse(responseBody);
            return parsed["choices"]?[0]?["message"]?["content"]?.ToString()?.Trim()
                   ?? throw new InvalidOperationException("Empty response from OpenAI API.");
        }
    }
}
