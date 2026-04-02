namespace GptOutlookPlugin.Models
{
    public class AppSettings
    {
        public string AiProvider { get; set; } = "codex-cli";
        public CodexCliSettings CodexCli { get; set; } = new CodexCliSettings();
        public OpenAiApiSettings OpenAiApi { get; set; } = new OpenAiApiSettings();
        public ContextSettings Context { get; set; } = new ContextSettings();
        public string DefaultTranslateTarget { get; set; } = "Korean";
        public string DefaultTone { get; set; } = "Professional and polite";
        public string CustomTonePrompt { get; set; } = "";
        public string ReviewSensitivity { get; set; } = "Medium";
    }

    public class CodexCliSettings
    {
        public string Command { get; set; } = "wsl.exe";
        public string Args { get; set; } = "codex exec -s read-only";
        public int TimeoutSeconds { get; set; } = 120;
    }

    public class OpenAiApiSettings
    {
        public string ApiKey { get; set; } = "";
        public string Model { get; set; } = "gpt-4o";
        public int MaxTokens { get; set; } = 4096;
    }

    public class ContextSettings
    {
        public int MaxHistoryMessages { get; set; } = 10;
    }
}
