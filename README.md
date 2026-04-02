# GPT Email Assistant — Outlook Plugin

A VSTO Add-in for Outlook Classic (Windows Desktop) that integrates AI-powered email assistance directly into your workflow.

## Features

| Feature | Description |
|---------|-------------|
| **Review** | Proofread and correct emails for grammar, tone, and clarity |
| **Compose** | Guided email composition with interactive prompts |
| **Auto Reply** | One-click automatic reply draft generation |
| **Translate** | Translate emails based on Office locale |
| **Summarize** | Summarize single emails or entire threads |
| **Settings** | Configure AI provider, tone, timeout, and language |

## Architecture

```
Outlook Classic
├── Ribbon Tab (GPT Email Assistant)
├── CustomTaskPane (WPF chat panel via ElementHost)
│   ├── Markdown rendering
│   ├── Diff view (Review mode)
│   └── Context-aware Apply/Copy buttons
└── VSTO Add-in Core (C#)
    ├── OutlookInterop — email read/write, locale detection
    ├── ContextManager — per-email, per-mode conversation sessions
    ├── PromptTemplates — locale & tone-aware system prompts
    └── AI Service — Codex CLI (WSL) primary, OpenAI API fallback
```

## AI Backend

- **Primary:** [Codex CLI](https://github.com/openai/codex) via WSL — uses your existing configuration (`~/.codex/config.toml`)
- **Fallback:** OpenAI API (gpt-4o / gpt-4.1) — requires API key in Settings

## Requirements

- Windows 10/11 with Outlook Classic (Desktop)
- Visual Studio 2022 with "Office/SharePoint development" workload
- .NET Framework 4.8
- WSL with Codex CLI installed (for primary AI backend)

## Quick Start

1. Clone this repo
2. Open `GptOutlookPlugin/GptOutlookPlugin.slnx` in Visual Studio 2022
3. Restore NuGet packages
4. Press F5 — Outlook opens with the plugin loaded
5. Click any button in the "GPT Email Assistant" ribbon tab

## Tech Stack

- C# / .NET Framework 4.8
- VSTO (Visual Studio Tools for Office)
- WPF (code-only, no XAML — for VSTO compatibility)
- Newtonsoft.Json, DiffPlex
- MSTest, Moq

## License

MIT License — see [LICENSE](LICENSE)

## Author

Hyun-Gyu (Ethan) Kim — babyworm@gmail.com
