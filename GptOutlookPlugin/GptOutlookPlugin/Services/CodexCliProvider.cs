using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Services
{
    public class CodexCliProvider : IAiProvider
    {
        private readonly CodexCliSettings _settings;
        private readonly string _scriptPath;

        public string Name => "Codex CLI (WSL)";

        public CodexCliProvider(CodexCliSettings settings)
        {
            _settings = settings;
            // scripts/codex-run.sh 의 WSL 경로
            _scriptPath = "/mnt/c/Works/gpt_outlook_plugin/scripts/codex-run.sh";
        }

        public bool IsAvailable()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "wsl.exe",
                    Arguments = $"bash {_scriptPath} \"--version-check\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var proc = Process.Start(psi))
                {
                    proc.WaitForExit(10000);
                    return true; // script exists and runs
                }
            }
            catch
            {
                return false;
            }
        }

        public async Task<string> SendAsync(List<ChatMessage> messages, CancellationToken ct)
        {
            // 프롬프트를 임시 파일에 저장 (쉘 이스케이프 문제 방지)
            var tempFile = Path.GetTempFileName();
            try
            {
                var prompt = BuildPromptArgument(messages);
                File.WriteAllText(tempFile, prompt, Encoding.UTF8);

                // Windows 경로를 WSL 경로로 변환
                var wslTempPath = tempFile.Replace("\\", "/").Replace("C:", "/mnt/c");

                var psi = new ProcessStartInfo
                {
                    FileName = "wsl.exe",
                    Arguments = $"bash {_scriptPath} \"$(cat {wslTempPath})\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                using (var proc = Process.Start(psi))
                {
                    var outputTask = proc.StandardOutput.ReadToEndAsync();

                    var completed = await Task.Run(() => proc.WaitForExit(_settings.TimeoutSeconds * 1000));
                    if (!completed)
                    {
                        try { proc.Kill(); } catch { }
                        throw new TaskCanceledException("Codex CLI timed out.");
                    }

                    ct.ThrowIfCancellationRequested();

                    var output = await outputTask;

                    if (proc.ExitCode != 0)
                        throw new InvalidOperationException($"Codex CLI failed (exit {proc.ExitCode})");

                    return ParseResponse(output);
                }
            }
            finally
            {
                try { File.Delete(tempFile); } catch { }
            }
        }

        public string BuildPromptArgument(List<ChatMessage> messages)
        {
            var sb = new StringBuilder();
            foreach (var msg in messages)
            {
                if (msg.Role == ChatRole.System)
                    sb.AppendLine($"[System]: {msg.Content}");
                else if (msg.Role == ChatRole.User)
                    sb.AppendLine($"[User]: {msg.Content}");
                else
                    sb.AppendLine($"[Assistant]: {msg.Content}");
            }
            return sb.ToString().Trim();
        }

        public string ParseResponse(string raw)
        {
            return raw?.Trim() ?? "";
        }
    }
}
