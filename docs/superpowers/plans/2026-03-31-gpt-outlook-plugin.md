# GPT Outlook Plugin Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Outlook Classic용 VSTO Add-in으로 ChatGPT 기반 이메일 리뷰/교정, 작성 보조, 번역, 요약 기능을 제공하는 플러그인을 구현한다.

**Architecture:** C# VSTO Add-in with WPF CustomTaskPane (via ElementHost). AI backend은 WSL Codex CLI (1순위) + OpenAI API (fallback) 이중 구조. MVVM 패턴으로 채팅형 사이드 패널 구현. 이메일별 대화 컨텍스트 유지.

**Tech Stack:** C# / .NET Framework 4.8, VSTO, WPF, Newtonsoft.Json, DiffPlex, MSTest, Moq

**Spec:** `docs/superpowers/specs/2026-03-31-gpt-outlook-plugin-design.md`

---

## File Map

### Main Project: `GptOutlookPlugin/`

| File | Responsibility |
|------|---------------|
| `ThisAddIn.cs` | VSTO entry point, 컴포넌트 초기화 및 배선 |
| `GptRibbon.xml` | 리본 탭 XML 정의 |
| `GptRibbon.cs` | 리본 이벤트 핸들러 |
| `Models/ChatMessage.cs` | 채팅 메시지 모델 (Role, Content, Timestamp) |
| `Models/EmailContext.cs` | 이메일 컨텍스트 (Subject, Body, Recipients) |
| `Models/ConversationSession.cs` | 대화 세션 (히스토리, 모드, 이메일 컨텍스트) |
| `Models/FeatureMode.cs` | 기능 모드 enum |
| `Models/AppSettings.cs` | 설정 모델 |
| `Services/IAiProvider.cs` | AI 프로바이더 인터페이스 |
| `Services/CodexCliProvider.cs` | WSL Codex CLI 호출 |
| `Services/OpenAiApiProvider.cs` | OpenAI API 호출 |
| `Services/AiServiceManager.cs` | Provider 선택 & fallback |
| `Core/ContextManager.cs` | 이메일별 대화 세션 관리 |
| `Core/PromptTemplates.cs` | 기능별 시스템 프롬프트 |
| `Core/OutlookInterop.cs` | Outlook MailItem 읽기/쓰기 헬퍼 |
| `Core/SettingsManager.cs` | JSON 설정 파일 로드/저장 |
| `Core/DiffEngine.cs` | 문장 단위 diff 계산 |
| `UI/ChatTaskPaneViewModel.cs` | 채팅 패널 ViewModel (MVVM) |
| `UI/ChatTaskPane.xaml` | 채팅 패널 WPF UI |
| `UI/ChatTaskPane.xaml.cs` | 채팅 패널 code-behind (최소) |
| `UI/DiffView.xaml` | Diff 표시 WPF UserControl |
| `UI/DiffView.xaml.cs` | Diff View code-behind |
| `UI/TaskPaneHost.cs` | WPF → WinForms ElementHost 브릿지 |
| `UI/SettingsWindow.xaml` | 설정 다이얼로그 |
| `UI/SettingsWindow.xaml.cs` | 설정 다이얼로그 code-behind |
| `Config/appsettings.json` | 기본 설정 파일 |

### Test Project: `GptOutlookPlugin.Tests/`

| File | Tests |
|------|-------|
| `Models/ConversationSessionTests.cs` | 세션 생성, 메시지 추가, 히스토리 트림 |
| `Services/CodexCliProviderTests.cs` | CLI 호출, 타임아웃, 파싱 |
| `Services/OpenAiApiProviderTests.cs` | API 호출, 응답 파싱 |
| `Services/AiServiceManagerTests.cs` | Fallback 로직, provider 전환 |
| `Core/ContextManagerTests.cs` | 세션 생성/조회, 키 전략 |
| `Core/DiffEngineTests.cs` | 문장 diff, 변경/추가/삭제 감지 |
| `Core/PromptTemplatesTests.cs` | 프롬프트 생성, 변수 치환 |

---

## Task Dependencies

```
Task 1 (Setup)
  └──► Task 2 (Models) ──► Task 3 (Config)
         │                      │
         ▼                      ▼
       Task 4 (CodexCli) ──► Task 6 (AiServiceManager)
       Task 5 (OpenAiApi) ──┘     │
         │                        ▼
         ▼                  Task 7 (ContextManager)
       Task 8 (Interop)          │
       Task 9 (DiffEngine)       │
         │                        │
         ▼                        ▼
       Task 10 (ChatTaskPane + DiffView)
         │
         ▼
       Task 11 (Ribbon + Context Menu)
         │
         ▼
       Task 12 (ThisAddIn Wiring)
         │
         ▼
       Task 13 (Settings Dialog)
```

---

## Task 1: Project Setup

**Files:**
- Create: `GptOutlookPlugin.sln`
- Create: `GptOutlookPlugin/GptOutlookPlugin.csproj`
- Create: `GptOutlookPlugin.Tests/GptOutlookPlugin.Tests.csproj`
- Create: `GptOutlookPlugin/Config/appsettings.json`

**Prerequisites:** Visual Studio 2022 with "Office/SharePoint development" workload installed.

- [ ] **Step 1: Create VSTO Outlook Add-in project in Visual Studio**

Open Visual Studio 2022 on Windows:
1. File → New → Project
2. Search "Outlook VSTO Add-in"
3. Select "Outlook VSTO Add-in" (C#)
4. Project name: `GptOutlookPlugin`
5. Location: navigate to the `gpt_outlook_plugin` directory (WSL path accessible via `\\wsl$\Ubuntu\home\babyworm\work\gpt_outlook_plugin\`)
6. Framework: `.NET Framework 4.8`
7. Click Create

This generates `ThisAddIn.cs`, `ThisAddIn.Designer.cs`, and project infrastructure.

- [ ] **Step 2: Add test project**

In Visual Studio:
1. Right-click Solution → Add → New Project
2. Search "MSTest Test Project (.NET Framework)"
3. Name: `GptOutlookPlugin.Tests`
4. Framework: `.NET Framework 4.8`
5. Right-click `GptOutlookPlugin.Tests` → Add → Project Reference → check `GptOutlookPlugin`

- [ ] **Step 3: Install NuGet packages**

In Package Manager Console:

```powershell
# Main project
Install-Package Newtonsoft.Json -ProjectName GptOutlookPlugin
Install-Package DiffPlex -ProjectName GptOutlookPlugin

# Test project
Install-Package Moq -ProjectName GptOutlookPlugin.Tests
Install-Package Newtonsoft.Json -ProjectName GptOutlookPlugin.Tests
```

- [ ] **Step 4: Add WPF references to main project**

Right-click `GptOutlookPlugin` → Add Reference:
- Assemblies → Framework → check:
  - `PresentationCore`
  - `PresentationFramework`
  - `WindowsBase`
  - `WindowsFormsIntegration` (for ElementHost)
  - `System.Xaml`

- [ ] **Step 5: Create directory structure and config file**

In the project, create folders: `Models`, `Services`, `Core`, `UI`, `Config`.

Create `Config/appsettings.json` (set Build Action: "Embedded Resource", Copy to Output: "Copy if newer"):

```json
{
  "aiProvider": "codex-cli",
  "codexCli": {
    "command": "wsl.exe",
    "args": "codex exec -s read-only",
    "timeoutSeconds": 60
  },
  "openAiApi": {
    "apiKey": "",
    "model": "gpt-4o",
    "maxTokens": 4096
  },
  "context": {
    "maxHistoryMessages": 10
  },
  "defaultTranslateTarget": "ko"
}
```

- [ ] **Step 6: Verify build**

Build → Build Solution (Ctrl+Shift+B). Should compile with zero errors.

- [ ] **Step 7: Commit**

```bash
git init
git add -A
git commit -m "chore: initialize VSTO Outlook Add-in project with test project and NuGet packages"
```

---

## Task 2: Core Models & Enums

**Files:**
- Create: `GptOutlookPlugin/Models/FeatureMode.cs`
- Create: `GptOutlookPlugin/Models/ChatMessage.cs`
- Create: `GptOutlookPlugin/Models/EmailContext.cs`
- Create: `GptOutlookPlugin/Models/ConversationSession.cs`
- Create: `GptOutlookPlugin/Models/AppSettings.cs`
- Test: `GptOutlookPlugin.Tests/Models/ConversationSessionTests.cs`

- [ ] **Step 1: Write ConversationSession tests**

Create `GptOutlookPlugin.Tests/Models/ConversationSessionTests.cs`:

```csharp
using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Tests.Models
{
    [TestClass]
    public class ConversationSessionTests
    {
        [TestMethod]
        public void NewSession_HasEmptyHistory()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review);
            Assert.AreEqual(0, session.Messages.Count);
            Assert.AreEqual(FeatureMode.Review, session.CurrentMode);
        }

        [TestMethod]
        public void AddMessage_AppendsToHistory()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review);
            session.AddMessage(ChatRole.User, "Review this email");
            session.AddMessage(ChatRole.Assistant, "Here is the review...");

            Assert.AreEqual(2, session.Messages.Count);
            Assert.AreEqual(ChatRole.User, session.Messages[0].Role);
            Assert.AreEqual(ChatRole.Assistant, session.Messages[1].Role);
        }

        [TestMethod]
        public void TrimHistory_KeepsMaxMessages()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review, maxHistory: 3);

            for (int i = 0; i < 5; i++)
                session.AddMessage(ChatRole.User, $"Message {i}");

            Assert.AreEqual(3, session.Messages.Count);
            Assert.AreEqual("Message 2", session.Messages[0].Content);
            Assert.AreEqual("Message 4", session.Messages[2].Content);
        }

        [TestMethod]
        public void SetEmailContext_UpdatesContext()
        {
            var session = new ConversationSession("test-id", FeatureMode.Translate);
            var ctx = new EmailContext
            {
                Subject = "Meeting Tomorrow",
                Body = "Let's meet at 3pm.",
                Recipients = "bob@example.com"
            };

            session.EmailContext = ctx;

            Assert.AreEqual("Meeting Tomorrow", session.EmailContext.Subject);
        }

        [TestMethod]
        public void SwitchMode_ChangesCurrentMode()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review);
            session.SwitchMode(FeatureMode.Translate);

            Assert.AreEqual(FeatureMode.Translate, session.CurrentMode);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: Build → Run All Tests (Ctrl+R, A)
Expected: Compilation errors — `ConversationSession`, `ChatRole`, `FeatureMode`, `EmailContext` do not exist yet.

- [ ] **Step 3: Implement FeatureMode enum**

Create `GptOutlookPlugin/Models/FeatureMode.cs`:

```csharp
namespace GptOutlookPlugin.Models
{
    public enum FeatureMode
    {
        Review,
        Compose,
        Translate,
        Summarize
    }
}
```

- [ ] **Step 4: Implement ChatMessage**

Create `GptOutlookPlugin/Models/ChatMessage.cs`:

```csharp
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
```

- [ ] **Step 5: Implement EmailContext**

Create `GptOutlookPlugin/Models/EmailContext.cs`:

```csharp
namespace GptOutlookPlugin.Models
{
    public class EmailContext
    {
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
        public string BodyHtml { get; set; } = "";
        public string Recipients { get; set; } = "";
        public string SenderEmail { get; set; } = "";
        public bool IsComposing { get; set; }
    }
}
```

- [ ] **Step 6: Implement ConversationSession**

Create `GptOutlookPlugin/Models/ConversationSession.cs`:

```csharp
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
```

- [ ] **Step 7: Implement AppSettings**

Create `GptOutlookPlugin/Models/AppSettings.cs`:

```csharp
namespace GptOutlookPlugin.Models
{
    public class AppSettings
    {
        public string AiProvider { get; set; } = "codex-cli";
        public CodexCliSettings CodexCli { get; set; } = new CodexCliSettings();
        public OpenAiApiSettings OpenAiApi { get; set; } = new OpenAiApiSettings();
        public ContextSettings Context { get; set; } = new ContextSettings();
        public string DefaultTranslateTarget { get; set; } = "ko";
    }

    public class CodexCliSettings
    {
        public string Command { get; set; } = "wsl.exe";
        public string Args { get; set; } = "codex exec -s read-only";
        public int TimeoutSeconds { get; set; } = 60;
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
```

- [ ] **Step 8: Run tests to verify they pass**

Run: Build → Run All Tests
Expected: All 5 tests PASS.

- [ ] **Step 9: Commit**

```bash
git add GptOutlookPlugin/Models/ GptOutlookPlugin.Tests/Models/
git commit -m "feat: add core models — ChatMessage, EmailContext, ConversationSession, AppSettings"
```

---

## Task 3: Configuration & Prompt Templates

**Files:**
- Create: `GptOutlookPlugin/Core/SettingsManager.cs`
- Create: `GptOutlookPlugin/Core/PromptTemplates.cs`
- Test: `GptOutlookPlugin.Tests/Core/PromptTemplatesTests.cs`

- [ ] **Step 1: Write PromptTemplates tests**

Create `GptOutlookPlugin.Tests/Core/PromptTemplatesTests.cs`:

```csharp
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Tests.Core
{
    [TestClass]
    public class PromptTemplatesTests
    {
        [TestMethod]
        public void GetSystemPrompt_Review_ContainsEmailBody()
        {
            var ctx = new EmailContext { Subject = "Test", Body = "Hello world" };
            var prompt = PromptTemplates.GetSystemPrompt(FeatureMode.Review, ctx);

            StringAssert.Contains(prompt, "Hello world");
            StringAssert.Contains(prompt, "proofread");
        }

        [TestMethod]
        public void GetSystemPrompt_Translate_ContainsTargetLanguage()
        {
            var ctx = new EmailContext { Body = "Hello" };
            var prompt = PromptTemplates.GetSystemPrompt(FeatureMode.Translate, ctx, targetLanguage: "ko");

            StringAssert.Contains(prompt, "ko");
        }

        [TestMethod]
        public void GetSystemPrompt_Compose_ContainsRecipient()
        {
            var ctx = new EmailContext
            {
                Body = "Original email",
                Recipients = "boss@company.com"
            };
            var prompt = PromptTemplates.GetSystemPrompt(FeatureMode.Compose, ctx);

            StringAssert.Contains(prompt, "boss@company.com");
        }

        [TestMethod]
        public void GetSystemPrompt_AllModes_ReturnNonEmpty()
        {
            var ctx = new EmailContext { Body = "test" };

            foreach (FeatureMode mode in System.Enum.GetValues(typeof(FeatureMode)))
            {
                var prompt = PromptTemplates.GetSystemPrompt(mode, ctx);
                Assert.IsFalse(string.IsNullOrWhiteSpace(prompt),
                    $"Prompt for {mode} should not be empty");
            }
        }
    }
}
```

- [ ] **Step 2: Run tests — expect failure**

Run: Build → Run All Tests
Expected: FAIL — `PromptTemplates` class not found.

- [ ] **Step 3: Implement PromptTemplates**

Create `GptOutlookPlugin/Core/PromptTemplates.cs`:

```csharp
using System;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public static class PromptTemplates
    {
        public static string GetSystemPrompt(FeatureMode mode, EmailContext ctx, string targetLanguage = "ko")
        {
            string emailSection = $"\n\n---\nEmail Subject: {ctx.Subject}\nRecipients: {ctx.Recipients}\nEmail Body:\n{ctx.Body}\n---";

            switch (mode)
            {
                case FeatureMode.Review:
                    return "You are an expert email proofreader. "
                         + "Review the following email for grammar, tone, clarity, and professionalism. "
                         + "Present your corrections in a structured before/after format. "
                         + "For each change, briefly explain why."
                         + emailSection;

                case FeatureMode.Compose:
                    return "You are an expert business email writer. "
                         + $"Draft an appropriate reply to the following email. Consider the recipient ({ctx.Recipients}) and context. "
                         + "Match the formality level of the original email. "
                         + "Provide the full draft reply."
                         + emailSection;

                case FeatureMode.Translate:
                    return "You are a professional translator specializing in business communication. "
                         + $"Translate the following email to {targetLanguage}. "
                         + "Maintain the original tone, formality, and formatting. "
                         + "Present the result as: original text, then translated text."
                         + emailSection;

                case FeatureMode.Summarize:
                    return "You are an expert email analyst. "
                         + "Summarize the following email concisely. Include: "
                         + "1) Key points, 2) Action items or requests, 3) Deadlines if any."
                         + emailSection;

                default:
                    throw new ArgumentOutOfRangeException(nameof(mode));
            }
        }
    }
}
```

- [ ] **Step 4: Implement SettingsManager**

Create `GptOutlookPlugin/Core/SettingsManager.cs`:

```csharp
using System;
using System.IO;
using Newtonsoft.Json;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public class SettingsManager
    {
        private readonly string _settingsPath;
        public AppSettings Current { get; private set; }

        public SettingsManager()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(appData, "GptOutlookPlugin");
            Directory.CreateDirectory(dir);
            _settingsPath = Path.Combine(dir, "appsettings.json");
            Load();
        }

        public void Load()
        {
            if (File.Exists(_settingsPath))
            {
                var json = File.ReadAllText(_settingsPath);
                Current = JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
            }
            else
            {
                Current = new AppSettings();
                Save();
            }
        }

        public void Save()
        {
            var json = JsonConvert.SerializeObject(Current, Formatting.Indented);
            File.WriteAllText(_settingsPath, json);
        }
    }
}
```

- [ ] **Step 5: Run tests — expect pass**

Run: Build → Run All Tests
Expected: All PromptTemplates tests PASS.

- [ ] **Step 6: Commit**

```bash
git add GptOutlookPlugin/Core/PromptTemplates.cs GptOutlookPlugin/Core/SettingsManager.cs GptOutlookPlugin.Tests/Core/
git commit -m "feat: add PromptTemplates and SettingsManager with JSON config persistence"
```

---

## Task 4: AI Provider Interface & CodexCliProvider

**Files:**
- Create: `GptOutlookPlugin/Services/IAiProvider.cs`
- Create: `GptOutlookPlugin/Services/CodexCliProvider.cs`
- Test: `GptOutlookPlugin.Tests/Services/CodexCliProviderTests.cs`

- [ ] **Step 1: Write CodexCliProvider tests**

Create `GptOutlookPlugin.Tests/Services/CodexCliProviderTests.cs`:

```csharp
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;
using System.Collections.Generic;

namespace GptOutlookPlugin.Tests.Services
{
    [TestClass]
    public class CodexCliProviderTests
    {
        [TestMethod]
        public void BuildPromptArg_FormatsMessagesCorrectly()
        {
            var provider = new CodexCliProvider(new CodexCliSettings());
            var messages = new List<ChatMessage>
            {
                new ChatMessage(ChatRole.System, "You are helpful."),
                new ChatMessage(ChatRole.User, "Hello")
            };

            var arg = provider.BuildPromptArgument(messages);

            StringAssert.Contains(arg, "You are helpful.");
            StringAssert.Contains(arg, "Hello");
        }

        [TestMethod]
        public void ParseResponse_ExtractsContent()
        {
            var provider = new CodexCliProvider(new CodexCliSettings());
            var raw = "Some preamble\n\nHere is the actual response content.\nWith multiple lines.";

            var result = provider.ParseResponse(raw);

            Assert.IsFalse(string.IsNullOrWhiteSpace(result));
        }

        [TestMethod]
        public async Task SendAsync_TimesOut_ThrowsException()
        {
            var settings = new CodexCliSettings
            {
                Command = "wsl.exe",
                Args = "sleep 100",
                TimeoutSeconds = 1
            };
            var provider = new CodexCliProvider(settings);
            var messages = new List<ChatMessage>
            {
                new ChatMessage(ChatRole.User, "test")
            };

            await Assert.ThrowsExceptionAsync<TaskCanceledException>(
                () => provider.SendAsync(messages, CancellationToken.None));
        }
    }
}
```

- [ ] **Step 2: Run tests — expect failure**

Run: Build → Run All Tests
Expected: FAIL — `IAiProvider`, `CodexCliProvider` not found.

- [ ] **Step 3: Implement IAiProvider interface**

Create `GptOutlookPlugin/Services/IAiProvider.cs`:

```csharp
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
```

- [ ] **Step 4: Implement CodexCliProvider**

Create `GptOutlookPlugin/Services/CodexCliProvider.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Services
{
    public class CodexCliProvider : IAiProvider
    {
        private readonly CodexCliSettings _settings;

        public string Name => "Codex CLI (WSL)";

        public CodexCliProvider(CodexCliSettings settings)
        {
            _settings = settings;
        }

        public bool IsAvailable()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = _settings.Command,
                    Arguments = "echo ok",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var proc = Process.Start(psi))
                {
                    proc.WaitForExit(5000);
                    return proc.ExitCode == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        public async Task<string> SendAsync(List<ChatMessage> messages, CancellationToken ct)
        {
            var promptArg = BuildPromptArgument(messages);
            var fullArgs = $"{_settings.Args} \"{EscapeForShell(promptArg)}\"";

            var psi = new ProcessStartInfo
            {
                FileName = _settings.Command,
                Arguments = fullArgs,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8
            };

            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(ct))
            {
                cts.CancelAfter(TimeSpan.FromSeconds(_settings.TimeoutSeconds));

                using (var proc = Process.Start(psi))
                {
                    var outputTask = proc.StandardOutput.ReadToEndAsync();
                    var errorTask = proc.StandardError.ReadToEndAsync();

                    var completed = await Task.Run(() => proc.WaitForExit(_settings.TimeoutSeconds * 1000));
                    if (!completed)
                    {
                        try { proc.Kill(); } catch { }
                        throw new TaskCanceledException("Codex CLI timed out.");
                    }

                    cts.Token.ThrowIfCancellationRequested();

                    var output = await outputTask;
                    var error = await errorTask;

                    if (proc.ExitCode != 0)
                        throw new InvalidOperationException($"Codex CLI failed (exit {proc.ExitCode}): {error}");

                    return ParseResponse(output);
                }
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

        private static string EscapeForShell(string input)
        {
            return input.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("$", "\\$");
        }
    }
}
```

- [ ] **Step 5: Run tests — expect pass (except timeout test which needs WSL)**

Run: Build → Run All Tests
Expected: `BuildPromptArg` and `ParseResponse` tests PASS. Timeout test may PASS or SKIP depending on WSL availability.

- [ ] **Step 6: Commit**

```bash
git add GptOutlookPlugin/Services/IAiProvider.cs GptOutlookPlugin/Services/CodexCliProvider.cs GptOutlookPlugin.Tests/Services/
git commit -m "feat: add IAiProvider interface and CodexCliProvider with WSL subprocess execution"
```

---

## Task 5: OpenAiApiProvider & AiServiceManager

**Files:**
- Create: `GptOutlookPlugin/Services/OpenAiApiProvider.cs`
- Create: `GptOutlookPlugin/Services/AiServiceManager.cs`
- Test: `GptOutlookPlugin.Tests/Services/AiServiceManagerTests.cs`

- [ ] **Step 1: Write AiServiceManager tests**

Create `GptOutlookPlugin.Tests/Services/AiServiceManagerTests.cs`:

```csharp
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;

namespace GptOutlookPlugin.Tests.Services
{
    [TestClass]
    public class AiServiceManagerTests
    {
        [TestMethod]
        public async Task SendAsync_UsesPrimaryWhenAvailable()
        {
            var primary = new Mock<IAiProvider>();
            primary.Setup(p => p.IsAvailable()).Returns(true);
            primary.Setup(p => p.Name).Returns("Primary");
            primary.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                   .ReturnsAsync("primary response");

            var fallback = new Mock<IAiProvider>();

            var manager = new AiServiceManager(primary.Object, fallback.Object);
            var messages = new List<ChatMessage> { new ChatMessage(ChatRole.User, "test") };

            var result = await manager.SendAsync(messages, CancellationToken.None);

            Assert.AreEqual("primary response", result);
            fallback.Verify(f => f.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()), Times.Never);
        }

        [TestMethod]
        public async Task SendAsync_FallsBackWhenPrimaryFails()
        {
            var primary = new Mock<IAiProvider>();
            primary.Setup(p => p.IsAvailable()).Returns(true);
            primary.Setup(p => p.Name).Returns("Primary");
            primary.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                   .ThrowsAsync(new System.Exception("primary failed"));

            var fallback = new Mock<IAiProvider>();
            fallback.Setup(p => p.IsAvailable()).Returns(true);
            fallback.Setup(p => p.Name).Returns("Fallback");
            fallback.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync("fallback response");

            var manager = new AiServiceManager(primary.Object, fallback.Object);
            var messages = new List<ChatMessage> { new ChatMessage(ChatRole.User, "test") };

            var result = await manager.SendAsync(messages, CancellationToken.None);

            Assert.AreEqual("fallback response", result);
        }

        [TestMethod]
        public async Task SendAsync_SkipsPrimaryWhenUnavailable()
        {
            var primary = new Mock<IAiProvider>();
            primary.Setup(p => p.IsAvailable()).Returns(false);
            primary.Setup(p => p.Name).Returns("Primary");

            var fallback = new Mock<IAiProvider>();
            fallback.Setup(p => p.IsAvailable()).Returns(true);
            fallback.Setup(p => p.Name).Returns("Fallback");
            fallback.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync("fallback response");

            var manager = new AiServiceManager(primary.Object, fallback.Object);
            var messages = new List<ChatMessage> { new ChatMessage(ChatRole.User, "test") };

            var result = await manager.SendAsync(messages, CancellationToken.None);

            Assert.AreEqual("fallback response", result);
        }
    }
}
```

- [ ] **Step 2: Run tests — expect failure**

Run: Build → Run All Tests
Expected: FAIL — `AiServiceManager` not found.

- [ ] **Step 3: Implement OpenAiApiProvider**

Create `GptOutlookPlugin/Services/OpenAiApiProvider.cs`:

```csharp
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
```

- [ ] **Step 4: Implement AiServiceManager**

Create `GptOutlookPlugin/Services/AiServiceManager.cs`:

```csharp
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
```

- [ ] **Step 5: Run tests — expect pass**

Run: Build → Run All Tests
Expected: All 3 AiServiceManager tests PASS.

- [ ] **Step 6: Commit**

```bash
git add GptOutlookPlugin/Services/OpenAiApiProvider.cs GptOutlookPlugin/Services/AiServiceManager.cs GptOutlookPlugin.Tests/Services/AiServiceManagerTests.cs
git commit -m "feat: add OpenAiApiProvider and AiServiceManager with automatic fallback"
```

---

## Task 6: Context Manager

**Files:**
- Create: `GptOutlookPlugin/Core/ContextManager.cs`
- Test: `GptOutlookPlugin.Tests/Core/ContextManagerTests.cs`

- [ ] **Step 1: Write ContextManager tests**

Create `GptOutlookPlugin.Tests/Core/ContextManagerTests.cs`:

```csharp
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Tests.Core
{
    [TestClass]
    public class ContextManagerTests
    {
        [TestMethod]
        public void GetOrCreateSession_CreatesNewSession()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var session = mgr.GetOrCreateSession("email-001", FeatureMode.Review);

            Assert.IsNotNull(session);
            Assert.AreEqual("email-001", session.SessionKey);
            Assert.AreEqual(FeatureMode.Review, session.CurrentMode);
        }

        [TestMethod]
        public void GetOrCreateSession_ReturnsSameSession()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var s1 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            var s2 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);

            Assert.AreSame(s1, s2);
        }

        [TestMethod]
        public void GetOrCreateSession_DifferentEmailsGetDifferentSessions()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var s1 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            var s2 = mgr.GetOrCreateSession("email-002", FeatureMode.Translate);

            Assert.AreNotSame(s1, s2);
        }

        [TestMethod]
        public void GetOrCreateSession_SameEmailDifferentMode_SwitchesMode()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var s1 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            s1.AddMessage(ChatRole.User, "Review this");

            var s2 = mgr.GetOrCreateSession("email-001", FeatureMode.Translate);

            Assert.AreSame(s1, s2);
            Assert.AreEqual(FeatureMode.Translate, s2.CurrentMode);
            Assert.AreEqual(1, s2.Messages.Count); // history preserved
        }

        [TestMethod]
        public void ClearSession_RemovesSession()
        {
            var mgr = new ContextManager(maxHistory: 10);
            mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            mgr.ClearSession("email-001");

            var session = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            Assert.AreEqual(0, session.Messages.Count);
        }

        [TestMethod]
        public void BuildMessages_IncludesSystemPromptAndHistory()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var session = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            session.EmailContext = new EmailContext { Subject = "Test", Body = "Hello" };
            session.AddMessage(ChatRole.User, "Review this");
            session.AddMessage(ChatRole.Assistant, "Looks good");

            var messages = mgr.BuildMessages(session);

            Assert.AreEqual(ChatRole.System, messages[0].Role);
            Assert.AreEqual(3, messages.Count); // system + 2 history
        }
    }
}
```

- [ ] **Step 2: Run tests — expect failure**

Run: Build → Run All Tests
Expected: FAIL — `ContextManager` not found.

- [ ] **Step 3: Implement ContextManager**

Create `GptOutlookPlugin/Core/ContextManager.cs`:

```csharp
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

        public ConversationSession GetOrCreateSession(string emailKey, FeatureMode mode)
        {
            if (_sessions.TryGetValue(emailKey, out var existing))
            {
                if (existing.CurrentMode != mode)
                    existing.SwitchMode(mode);
                return existing;
            }

            var session = new ConversationSession(emailKey, mode, _maxHistory);
            _sessions[emailKey] = session;
            return session;
        }

        public void ClearSession(string emailKey)
        {
            _sessions.Remove(emailKey);
        }

        public List<ChatMessage> BuildMessages(ConversationSession session, string targetLanguage = "ko")
        {
            var messages = new List<ChatMessage>();

            var systemPrompt = PromptTemplates.GetSystemPrompt(
                session.CurrentMode,
                session.EmailContext ?? new EmailContext(),
                targetLanguage);
            messages.Add(new ChatMessage(ChatRole.System, systemPrompt));

            messages.AddRange(session.Messages);

            return messages;
        }
    }
}
```

- [ ] **Step 4: Run tests — expect pass**

Run: Build → Run All Tests
Expected: All 6 ContextManager tests PASS.

- [ ] **Step 5: Commit**

```bash
git add GptOutlookPlugin/Core/ContextManager.cs GptOutlookPlugin.Tests/Core/ContextManagerTests.cs
git commit -m "feat: add ContextManager for email-keyed conversation session management"
```

---

## Task 7: Outlook Interop Helper

**Files:**
- Create: `GptOutlookPlugin/Core/OutlookInterop.cs`

Note: No unit tests for this — it directly depends on Outlook COM interop which requires a running Outlook instance. Tested manually via debug mode.

- [ ] **Step 1: Implement OutlookInterop**

Create `GptOutlookPlugin/Core/OutlookInterop.cs`:

```csharp
using System;
using Microsoft.Office.Interop.Outlook;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public class OutlookInterop
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _app;

        public OutlookInterop(Microsoft.Office.Interop.Outlook.Application app)
        {
            _app = app;
        }

        public EmailContext GetCurrentEmailContext()
        {
            var explorer = _app.ActiveExplorer();
            if (explorer?.Selection?.Count > 0)
            {
                var item = explorer.Selection[1];
                if (item is MailItem mail)
                    return ExtractContext(mail, isComposing: false);
            }

            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem composing)
                return ExtractContext(composing, isComposing: true);

            return null;
        }

        public string GetCurrentEntryIdOrTemp()
        {
            var explorer = _app.ActiveExplorer();
            if (explorer?.Selection?.Count > 0)
            {
                var item = explorer.Selection[1];
                if (item is MailItem mail && !string.IsNullOrEmpty(mail.EntryID))
                    return mail.EntryID;
            }

            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem composing)
            {
                if (!string.IsNullOrEmpty(composing.EntryID))
                    return composing.EntryID;
            }

            return "temp-" + Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        public string GetSelectedText()
        {
            var inspector = _app.ActiveInspector();
            if (inspector == null) return null;

            var doc = inspector.WordEditor as Microsoft.Office.Interop.Word.Document;
            var selection = doc?.Application?.Selection;
            if (selection != null && !string.IsNullOrWhiteSpace(selection.Text))
                return selection.Text.Trim();

            return null;
        }

        public void ApplyToBody(string newBody)
        {
            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem mail)
            {
                mail.Body = newBody;
            }
        }

        public void ApplyToHtmlBody(string newHtmlBody)
        {
            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem mail)
            {
                mail.HTMLBody = newHtmlBody;
            }
        }

        public void InsertReplyDraft(string replyBody)
        {
            var explorer = _app.ActiveExplorer();
            if (explorer?.Selection?.Count > 0)
            {
                var item = explorer.Selection[1];
                if (item is MailItem original)
                {
                    var reply = original.Reply();
                    reply.Body = replyBody + "\n\n" + reply.Body;
                    reply.Display(false);
                }
            }
        }

        private EmailContext ExtractContext(MailItem mail, bool isComposing)
        {
            return new EmailContext
            {
                Subject = mail.Subject ?? "",
                Body = mail.Body ?? "",
                BodyHtml = mail.HTMLBody ?? "",
                Recipients = GetRecipients(mail),
                SenderEmail = mail.SenderEmailAddress ?? "",
                IsComposing = isComposing
            };
        }

        private string GetRecipients(MailItem mail)
        {
            if (mail.Recipients == null) return "";
            var list = new System.Collections.Generic.List<string>();
            foreach (Recipient r in mail.Recipients)
                list.Add(r.Address ?? r.Name ?? "");
            return string.Join("; ", list);
        }
    }
}
```

- [ ] **Step 2: Verify build**

Build → Build Solution. Should compile if Outlook interop references are present (from VSTO template).

- [ ] **Step 3: Commit**

```bash
git add GptOutlookPlugin/Core/OutlookInterop.cs
git commit -m "feat: add OutlookInterop helper for MailItem read/write and selection"
```

---

## Task 8: Diff Engine

**Files:**
- Create: `GptOutlookPlugin/Core/DiffEngine.cs`
- Test: `GptOutlookPlugin.Tests/Core/DiffEngineTests.cs`

- [ ] **Step 1: Write DiffEngine tests**

Create `GptOutlookPlugin.Tests/Core/DiffEngineTests.cs`:

```csharp
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Core;

namespace GptOutlookPlugin.Tests.Core
{
    [TestClass]
    public class DiffEngineTests
    {
        [TestMethod]
        public void ComputeDiff_IdenticalTexts_AllUnchanged()
        {
            var result = DiffEngine.ComputeSentenceDiff("Hello world.", "Hello world.");

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(DiffType.Unchanged, result[0].Type);
        }

        [TestMethod]
        public void ComputeDiff_ModifiedSentence_ShowsDeleteAndInsert()
        {
            var original = "I want to talk about the project.";
            var modified = "I would like to discuss the project.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Any(r => r.Type == DiffType.Deleted));
            Assert.IsTrue(result.Any(r => r.Type == DiffType.Inserted));
        }

        [TestMethod]
        public void ComputeDiff_AddedSentence_ShowsInsert()
        {
            var original = "First sentence.";
            var modified = "First sentence. Second sentence added.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Any(r => r.Type == DiffType.Inserted));
            Assert.IsTrue(result.Any(r => r.Type == DiffType.Unchanged && r.Text.Contains("First")));
        }

        [TestMethod]
        public void ComputeDiff_RemovedSentence_ShowsDelete()
        {
            var original = "Keep this. Remove this.";
            var modified = "Keep this.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Any(r => r.Type == DiffType.Deleted && r.Text.Contains("Remove")));
        }

        [TestMethod]
        public void ComputeDiff_MultiParagraph_PreservesLineBreaks()
        {
            var original = "Paragraph one.\n\nParagraph two.";
            var modified = "Paragraph one.\n\nParagraph two modified.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Count > 0);
        }
    }
}
```

- [ ] **Step 2: Run tests — expect failure**

Run: Build → Run All Tests
Expected: FAIL — `DiffEngine`, `DiffType` not found.

- [ ] **Step 3: Implement DiffEngine**

Create `GptOutlookPlugin/Core/DiffEngine.cs`:

```csharp
using System.Collections.Generic;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;

namespace GptOutlookPlugin.Core
{
    public enum DiffType
    {
        Unchanged,
        Inserted,
        Deleted,
        Modified
    }

    public class DiffLine
    {
        public DiffType Type { get; set; }
        public string Text { get; set; }
    }

    public static class DiffEngine
    {
        public static List<DiffLine> ComputeSentenceDiff(string original, string modified)
        {
            var differ = new Differ();
            var builder = new InlineDiffBuilder(differ);
            var diff = builder.BuildDiffModel(original, modified, ignoreWhitespace: false);

            var result = new List<DiffLine>();

            foreach (var line in diff.Lines)
            {
                switch (line.Type)
                {
                    case ChangeType.Unchanged:
                        result.Add(new DiffLine { Type = DiffType.Unchanged, Text = line.Text });
                        break;
                    case ChangeType.Deleted:
                        result.Add(new DiffLine { Type = DiffType.Deleted, Text = line.Text });
                        break;
                    case ChangeType.Inserted:
                        result.Add(new DiffLine { Type = DiffType.Inserted, Text = line.Text });
                        break;
                    case ChangeType.Modified:
                        result.Add(new DiffLine { Type = DiffType.Modified, Text = line.Text });
                        break;
                    case ChangeType.Imaginary:
                        break;
                }
            }

            return result;
        }
    }
}
```

- [ ] **Step 4: Run tests — expect pass**

Run: Build → Run All Tests
Expected: All 5 DiffEngine tests PASS.

- [ ] **Step 5: Commit**

```bash
git add GptOutlookPlugin/Core/DiffEngine.cs GptOutlookPlugin.Tests/Core/DiffEngineTests.cs
git commit -m "feat: add DiffEngine using DiffPlex for sentence-level diff display"
```

---

## Task 9: ChatTaskPane WPF UI

**Files:**
- Create: `GptOutlookPlugin/UI/ChatTaskPaneViewModel.cs`
- Create: `GptOutlookPlugin/UI/ChatTaskPane.xaml`
- Create: `GptOutlookPlugin/UI/ChatTaskPane.xaml.cs`
- Create: `GptOutlookPlugin/UI/DiffView.xaml`
- Create: `GptOutlookPlugin/UI/DiffView.xaml.cs`

- [ ] **Step 1: Implement ChatTaskPaneViewModel**

Create `GptOutlookPlugin/UI/ChatTaskPaneViewModel.cs`:

```csharp
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;

namespace GptOutlookPlugin.UI
{
    public class ChatTaskPaneViewModel : INotifyPropertyChanged
    {
        private readonly ContextManager _contextManager;
        private readonly AiServiceManager _aiService;
        private readonly OutlookInterop _outlook;
        private ConversationSession _currentSession;

        public ObservableCollection<ChatMessageViewModel> Messages { get; } = new ObservableCollection<ChatMessageViewModel>();

        private string _inputText = "";
        public string InputText
        {
            get => _inputText;
            set { _inputText = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSend)); }
        }

        private bool _isLoading;
        public bool IsLoading
        {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanSend)); }
        }

        private FeatureMode _currentMode = FeatureMode.Review;
        public FeatureMode CurrentMode
        {
            get => _currentMode;
            set { _currentMode = value; OnPropertyChanged(); OnPropertyChanged(nameof(ModeDisplayName)); }
        }

        public string ModeDisplayName
        {
            get
            {
                switch (CurrentMode)
                {
                    case FeatureMode.Review: return "Review / Proofread";
                    case FeatureMode.Compose: return "Compose Reply";
                    case FeatureMode.Translate: return "Translate";
                    case FeatureMode.Summarize: return "Summarize";
                    default: return CurrentMode.ToString();
                }
            }
        }

        private string _statusText = "Ready";
        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(); }
        }

        public bool CanSend => !string.IsNullOrWhiteSpace(InputText) && !IsLoading;

        public ICommand SendCommand { get; }
        public ICommand ClearCommand { get; }

        public ChatTaskPaneViewModel(ContextManager contextManager, AiServiceManager aiService, OutlookInterop outlook)
        {
            _contextManager = contextManager;
            _aiService = aiService;
            _outlook = outlook;

            SendCommand = new RelayCommand(async _ => await SendMessageAsync(), _ => CanSend);
            ClearCommand = new RelayCommand(_ => ClearConversation());

            _aiService.OnProviderSwitch += msg => StatusText = msg;
        }

        public void StartMode(FeatureMode mode, string initialPrompt = null)
        {
            CurrentMode = mode;

            var emailKey = _outlook.GetCurrentEntryIdOrTemp();
            var emailCtx = _outlook.GetCurrentEmailContext();

            _currentSession = _contextManager.GetOrCreateSession(emailKey, mode);
            _currentSession.EmailContext = emailCtx;

            RefreshMessages();

            if (initialPrompt != null)
                _ = SendWithTextAsync(initialPrompt);
            else if (_currentSession.Messages.Count == 0)
                _ = SendAutoPromptAsync(mode, emailCtx);
        }

        public void StartWithSelection(FeatureMode mode, string selectedText)
        {
            CurrentMode = mode;

            var emailKey = _outlook.GetCurrentEntryIdOrTemp();
            var emailCtx = _outlook.GetCurrentEmailContext() ?? new EmailContext();
            emailCtx.Body = selectedText;

            _currentSession = _contextManager.GetOrCreateSession(emailKey + "-sel", mode);
            _currentSession.EmailContext = emailCtx;
            _currentSession.Clear();

            RefreshMessages();
            _ = SendAutoPromptAsync(mode, emailCtx);
        }

        private async Task SendAutoPromptAsync(FeatureMode mode, EmailContext ctx)
        {
            string prompt;
            switch (mode)
            {
                case FeatureMode.Review:
                    prompt = "Please review and proofread this email.";
                    break;
                case FeatureMode.Compose:
                    prompt = "Please draft a reply to this email.";
                    break;
                case FeatureMode.Translate:
                    prompt = "Please translate this email.";
                    break;
                case FeatureMode.Summarize:
                    prompt = "Please summarize this email.";
                    break;
                default:
                    return;
            }
            await SendWithTextAsync(prompt);
        }

        private async Task SendWithTextAsync(string text)
        {
            InputText = text;
            await SendMessageAsync();
        }

        private async Task SendMessageAsync()
        {
            if (_currentSession == null || string.IsNullOrWhiteSpace(InputText)) return;

            var userText = InputText.Trim();
            InputText = "";

            _currentSession.AddMessage(ChatRole.User, userText);
            Messages.Add(new ChatMessageViewModel(ChatRole.User, userText));

            IsLoading = true;
            StatusText = "AI processing...";

            try
            {
                var allMessages = _contextManager.BuildMessages(_currentSession);
                var response = await _aiService.SendAsync(allMessages, CancellationToken.None);

                _currentSession.AddMessage(ChatRole.Assistant, response);

                var showDiff = CurrentMode == FeatureMode.Review || CurrentMode == FeatureMode.Translate;
                Messages.Add(new ChatMessageViewModel(ChatRole.Assistant, response,
                    showDiff: showDiff,
                    originalText: _currentSession.EmailContext?.Body));

                StatusText = "Ready";
            }
            catch (Exception ex)
            {
                Messages.Add(new ChatMessageViewModel(ChatRole.System, $"Error: {ex.Message}"));
                StatusText = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void ApplyResult(string resultText)
        {
            switch (CurrentMode)
            {
                case FeatureMode.Review:
                case FeatureMode.Translate:
                    _outlook.ApplyToBody(resultText);
                    StatusText = "Applied to email body.";
                    break;
                case FeatureMode.Compose:
                    _outlook.InsertReplyDraft(resultText);
                    StatusText = "Reply draft created.";
                    break;
            }
        }

        private void ClearConversation()
        {
            if (_currentSession != null)
            {
                _contextManager.ClearSession(_currentSession.SessionKey);
                _currentSession = null;
            }
            Messages.Clear();
            StatusText = "Conversation cleared.";
        }

        private void RefreshMessages()
        {
            Messages.Clear();
            if (_currentSession != null)
            {
                foreach (var msg in _currentSession.Messages)
                    Messages.Add(new ChatMessageViewModel(msg.Role, msg.Content));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class ChatMessageViewModel
    {
        public ChatRole Role { get; }
        public string Content { get; }
        public bool IsUser => Role == ChatRole.User;
        public bool IsAssistant => Role == ChatRole.Assistant;
        public bool IsError => Role == ChatRole.System;
        public bool ShowDiff { get; }
        public string OriginalText { get; }
        public bool ShowApplyButton => IsAssistant && !IsError;

        public ChatMessageViewModel(ChatRole role, string content, bool showDiff = false, string originalText = null)
        {
            Role = role;
            Content = content;
            ShowDiff = showDiff;
            OriginalText = originalText;
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object parameter) => _execute(parameter);
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
```

- [ ] **Step 2: Create ChatTaskPane XAML**

Create `GptOutlookPlugin/UI/ChatTaskPane.xaml`:

```xml
<UserControl x:Class="GptOutlookPlugin.UI.ChatTaskPane"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns:local="clr-namespace:GptOutlookPlugin.UI"
             Background="#F5F5F5"
             MinWidth="300">

    <UserControl.Resources>
        <BooleanToVisibilityConverter x:Key="BoolToVis"/>
        <local:InverseBoolToVisibilityConverter x:Key="InverseBoolToVis"/>

        <!-- User message style -->
        <DataTemplate x:Key="UserMessageTemplate">
            <Border Background="#DCF8C6" CornerRadius="8" Padding="10" Margin="40,4,8,4"
                    HorizontalAlignment="Right" MaxWidth="280">
                <TextBlock Text="{Binding Content}" TextWrapping="Wrap" FontSize="13"/>
            </Border>
        </DataTemplate>

        <!-- Assistant message style -->
        <DataTemplate x:Key="AssistantMessageTemplate">
            <Border Background="White" CornerRadius="8" Padding="10" Margin="8,4,40,4"
                    HorizontalAlignment="Left" MaxWidth="280" BorderBrush="#E0E0E0" BorderThickness="1">
                <StackPanel>
                    <!-- Plain text (default) -->
                    <TextBlock Text="{Binding Content}" TextWrapping="Wrap" FontSize="13"
                               Visibility="{Binding ShowDiff, Converter={StaticResource InverseBoolToVis}}"/>
                    <!-- Diff view (Review/Translate modes) -->
                    <local:DiffView OriginalText="{Binding OriginalText}"
                                    ModifiedText="{Binding Content}"
                                    Visibility="{Binding ShowDiff, Converter={StaticResource BoolToVis}}"
                                    MaxHeight="300"/>
                    <Button Content="Apply" Margin="0,6,0,0" Padding="8,4"
                            HorizontalAlignment="Left" FontSize="11"
                            Visibility="{Binding ShowApplyButton, Converter={StaticResource BoolToVis}}"
                            Click="ApplyButton_Click" Tag="{Binding Content}"/>
                </StackPanel>
            </Border>
        </DataTemplate>

        <!-- Error message style -->
        <DataTemplate x:Key="ErrorMessageTemplate">
            <Border Background="#FFEBEE" CornerRadius="8" Padding="10" Margin="8,4,8,4">
                <TextBlock Text="{Binding Content}" TextWrapping="Wrap" FontSize="12"
                           Foreground="#C62828"/>
            </Border>
        </DataTemplate>

        <local:MessageTemplateSelector x:Key="MessageSelector"
            UserTemplate="{StaticResource UserMessageTemplate}"
            AssistantTemplate="{StaticResource AssistantMessageTemplate}"
            ErrorTemplate="{StaticResource ErrorMessageTemplate}"/>
    </UserControl.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header: Mode display -->
        <Border Grid.Row="0" Background="#1976D2" Padding="10,8">
            <Grid>
                <TextBlock Text="{Binding ModeDisplayName}" Foreground="White"
                           FontSize="14" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <Button Content="Clear" HorizontalAlignment="Right" Padding="8,2"
                        FontSize="11" Command="{Binding ClearCommand}"
                        Background="Transparent" Foreground="White" BorderBrush="White"/>
            </Grid>
        </Border>

        <!-- Messages list -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" x:Name="MessagesScroller">
            <ItemsControl ItemsSource="{Binding Messages}"
                          ItemTemplateSelector="{StaticResource MessageSelector}">
                <ItemsControl.ItemsPanel>
                    <ItemsPanelTemplate>
                        <StackPanel/>
                    </ItemsPanelTemplate>
                </ItemsControl.ItemsPanel>
            </ItemsControl>
        </ScrollViewer>

        <!-- Loading indicator -->
        <Border Grid.Row="2" Background="#E3F2FD" Padding="8,4"
                Visibility="{Binding IsLoading, Converter={StaticResource BoolToVis}}">
            <TextBlock Text="AI is thinking..." FontSize="12" Foreground="#1565C0" FontStyle="Italic"/>
        </Border>

        <!-- Input area -->
        <Grid Grid.Row="3" Margin="8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" Text="{Binding InputText, UpdateSourceTrigger=PropertyChanged}"
                     Padding="8,6" FontSize="13" VerticalContentAlignment="Center"
                     AcceptsReturn="False" KeyDown="InputBox_KeyDown"
                     x:Name="InputBox"/>
            <Button Grid.Column="1" Content="Send" Margin="4,0,0,0" Padding="12,6"
                    Command="{Binding SendCommand}" FontSize="13"
                    Background="#1976D2" Foreground="White"/>
        </Grid>
    </Grid>
</UserControl>
```

- [ ] **Step 3: Create ChatTaskPane code-behind**

Create `GptOutlookPlugin/UI/ChatTaskPane.xaml.cs`:

```csharp
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.UI
{
    public partial class ChatTaskPane : UserControl
    {
        private ChatTaskPaneViewModel ViewModel => DataContext as ChatTaskPaneViewModel;

        public ChatTaskPane()
        {
            InitializeComponent();
        }

        private void InputBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && ViewModel?.SendCommand?.CanExecute(null) == true)
            {
                ViewModel.SendCommand.Execute(null);
                e.Handled = true;
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string content)
            {
                ViewModel?.ApplyResult(content);
            }
        }
    }

    public class InverseBoolToVisibilityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b && b ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

    public class MessageTemplateSelector : DataTemplateSelector
    {
        public DataTemplate UserTemplate { get; set; }
        public DataTemplate AssistantTemplate { get; set; }
        public DataTemplate ErrorTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item is ChatMessageViewModel msg)
            {
                if (msg.IsError) return ErrorTemplate;
                if (msg.IsUser) return UserTemplate;
                return AssistantTemplate;
            }
            return base.SelectTemplate(item, container);
        }
    }
}
```

- [ ] **Step 4: Create DiffView UserControl**

Create `GptOutlookPlugin/UI/DiffView.xaml`:

```xml
<UserControl x:Class="GptOutlookPlugin.UI.DiffView"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <RichTextBox x:Name="DiffDocument" IsReadOnly="True" BorderThickness="0"
                 Background="Transparent" FontFamily="Consolas" FontSize="12"/>
</UserControl>
```

Create `GptOutlookPlugin/UI/DiffView.xaml.cs`:

```csharp
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using GptOutlookPlugin.Core;

namespace GptOutlookPlugin.UI
{
    public partial class DiffView : UserControl
    {
        public static readonly DependencyProperty OriginalTextProperty =
            DependencyProperty.Register("OriginalText", typeof(string), typeof(DiffView),
                new PropertyMetadata("", OnTextsChanged));

        public static readonly DependencyProperty ModifiedTextProperty =
            DependencyProperty.Register("ModifiedText", typeof(string), typeof(DiffView),
                new PropertyMetadata("", OnTextsChanged));

        public string OriginalText
        {
            get => (string)GetValue(OriginalTextProperty);
            set => SetValue(OriginalTextProperty, value);
        }

        public string ModifiedText
        {
            get => (string)GetValue(ModifiedTextProperty);
            set => SetValue(ModifiedTextProperty, value);
        }

        public DiffView()
        {
            InitializeComponent();
        }

        private static void OnTextsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DiffView view)
                view.RenderDiff();
        }

        private void RenderDiff()
        {
            var doc = new FlowDocument();
            var paragraph = new Paragraph();

            if (string.IsNullOrEmpty(OriginalText) || string.IsNullOrEmpty(ModifiedText))
            {
                paragraph.Inlines.Add(new Run(ModifiedText ?? ""));
                doc.Blocks.Add(paragraph);
                DiffDocument.Document = doc;
                return;
            }

            var diffs = DiffEngine.ComputeSentenceDiff(OriginalText, ModifiedText);

            foreach (var diff in diffs)
            {
                Run run;
                switch (diff.Type)
                {
                    case DiffType.Deleted:
                        run = new Run(diff.Text + "\n")
                        {
                            Background = new SolidColorBrush(Color.FromRgb(255, 220, 220)),
                            TextDecorations = TextDecorations.Strikethrough,
                            Foreground = new SolidColorBrush(Color.FromRgb(180, 0, 0))
                        };
                        paragraph.Inlines.Add(new Run("- ") { Foreground = Brushes.Red });
                        paragraph.Inlines.Add(run);
                        break;

                    case DiffType.Inserted:
                        run = new Run(diff.Text + "\n")
                        {
                            Background = new SolidColorBrush(Color.FromRgb(220, 255, 220)),
                            Foreground = new SolidColorBrush(Color.FromRgb(0, 120, 0))
                        };
                        paragraph.Inlines.Add(new Run("+ ") { Foreground = Brushes.Green });
                        paragraph.Inlines.Add(run);
                        break;

                    default:
                        run = new Run(diff.Text + "\n");
                        paragraph.Inlines.Add(new Run("  "));
                        paragraph.Inlines.Add(run);
                        break;
                }
            }

            doc.Blocks.Add(paragraph);
            DiffDocument.Document = doc;
        }
    }
}
```

- [ ] **Step 5: Verify build**

Build → Build Solution. Should compile with WPF references in place.

- [ ] **Step 6: Commit**

```bash
git add GptOutlookPlugin/UI/
git commit -m "feat: add ChatTaskPane WPF UI with MVVM, DiffView, and message templates"
```

---

## Task 10: TaskPaneHost, Ribbon & Context Menu

**Files:**
- Create: `GptOutlookPlugin/UI/TaskPaneHost.cs`
- Create: `GptOutlookPlugin/GptRibbon.xml`
- Create: `GptOutlookPlugin/GptRibbon.cs`

- [ ] **Step 1: Implement TaskPaneHost**

Create `GptOutlookPlugin/UI/TaskPaneHost.cs`:

```csharp
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GptOutlookPlugin.UI
{
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
```

- [ ] **Step 2: Create Ribbon XML**

Create `GptOutlookPlugin/GptRibbon.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"
          onLoad="Ribbon_Load">
  <ribbon>
    <tabs>
      <tab id="GptTab" label="GPT Email Assistant">
        <group id="GptMainGroup" label="AI Actions">
          <button id="btnReview"
                  label="Review"
                  imageMso="SpellingMenu"
                  size="large"
                  onAction="OnReviewClick"
                  screentip="Review and proofread the current email"/>
          <button id="btnCompose"
                  label="Compose"
                  imageMso="MailMergeStartLetters"
                  size="large"
                  onAction="OnComposeClick"
                  screentip="Draft a reply to the current email"/>
          <button id="btnTranslate"
                  label="Translate"
                  imageMso="TranslateMenu"
                  size="large"
                  onAction="OnTranslateClick"
                  screentip="Translate the current email"/>
          <button id="btnSummarize"
                  label="Summarize"
                  imageMso="OutlineDemoteToBodyText"
                  size="large"
                  onAction="OnSummarizeClick"
                  screentip="Summarize the current email"/>
        </group>
        <group id="GptSettingsGroup" label="Settings">
          <button id="btnSettings"
                  label="Settings"
                  imageMso="AdvancedFileProperties"
                  size="normal"
                  onAction="OnSettingsClick"
                  screentip="Configure AI provider and preferences"/>
        </group>
      </tab>
    </tabs>

    <contextMenus>
      <contextMenu idMso="ContextMenuText">
        <menuSeparator id="GptSep"/>
        <menu id="GptContextMenu" label="GPT Assistant" imageMso="SpellingMenu">
          <button id="ctxReview"
                  label="Review Selection"
                  onAction="OnContextReviewClick"/>
          <button id="ctxTranslate"
                  label="Translate Selection"
                  onAction="OnContextTranslateClick"/>
          <button id="ctxCompose"
                  label="Compose from Selection"
                  onAction="OnContextComposeClick"/>
        </menu>
      </contextMenu>
    </contextMenus>
  </ribbon>
</customUI>
```

Set Build Action to "Embedded Resource" in file properties.

- [ ] **Step 3: Implement GptRibbon**

Create `GptOutlookPlugin/GptRibbon.cs`:

```csharp
using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin
{
    public class GptRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        public event Action<FeatureMode> OnModeRequested;
        public event Action<FeatureMode, string> OnSelectionModeRequested;
        public event Action OnSettingsRequested;

        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("GptOutlookPlugin.GptRibbon.xml"))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnReviewClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Review);
        }

        public void OnComposeClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Compose);
        }

        public void OnTranslateClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Translate);
        }

        public void OnSummarizeClick(IRibbonControl control)
        {
            OnModeRequested?.Invoke(FeatureMode.Summarize);
        }

        public void OnSettingsClick(IRibbonControl control)
        {
            OnSettingsRequested?.Invoke();
        }

        public void OnContextReviewClick(IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.OutlookInterop.GetSelectedText();
            if (!string.IsNullOrEmpty(selection))
                OnSelectionModeRequested?.Invoke(FeatureMode.Review, selection);
        }

        public void OnContextTranslateClick(IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.OutlookInterop.GetSelectedText();
            if (!string.IsNullOrEmpty(selection))
                OnSelectionModeRequested?.Invoke(FeatureMode.Translate, selection);
        }

        public void OnContextComposeClick(IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.OutlookInterop.GetSelectedText();
            if (!string.IsNullOrEmpty(selection))
                OnSelectionModeRequested?.Invoke(FeatureMode.Compose, selection);
        }
    }
}
```

- [ ] **Step 4: Verify build**

Build → Build Solution. Should compile cleanly.

- [ ] **Step 5: Commit**

```bash
git add GptOutlookPlugin/UI/TaskPaneHost.cs GptOutlookPlugin/GptRibbon.xml GptOutlookPlugin/GptRibbon.cs
git commit -m "feat: add TaskPaneHost bridge, Ribbon tab, and context menu definitions"
```

---

## Task 11: ThisAddIn Wiring

**Files:**
- Modify: `GptOutlookPlugin/ThisAddIn.cs`

- [ ] **Step 1: Wire all components in ThisAddIn**

Replace the content of `GptOutlookPlugin/ThisAddIn.cs`:

```csharp
using System;
using Microsoft.Office.Tools;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;
using GptOutlookPlugin.UI;

namespace GptOutlookPlugin
{
    public partial class ThisAddIn
    {
        private SettingsManager _settingsManager;
        private ContextManager _contextManager;
        private AiServiceManager _aiService;
        private ChatTaskPaneViewModel _viewModel;
        private CustomTaskPane _taskPane;
        private GptRibbon _ribbon;

        public OutlookInterop OutlookInterop { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _settingsManager = new SettingsManager();
            var settings = _settingsManager.Current;

            OutlookInterop = new OutlookInterop(Application);
            _contextManager = new ContextManager(settings.Context.MaxHistoryMessages);

            var codexProvider = new CodexCliProvider(settings.CodexCli);
            var openAiProvider = new OpenAiApiProvider(settings.OpenAiApi);

            _aiService = settings.AiProvider == "openai-api"
                ? new AiServiceManager(openAiProvider, codexProvider)
                : new AiServiceManager(codexProvider, openAiProvider);

            _viewModel = new ChatTaskPaneViewModel(_contextManager, _aiService, OutlookInterop);

            var host = new TaskPaneHost(_viewModel);
            _taskPane = CustomTaskPanes.Add(host, "GPT Email Assistant");
            _taskPane.Width = 380;
            _taskPane.Visible = false;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new GptRibbon();
            _ribbon.OnModeRequested += mode =>
            {
                _viewModel.StartMode(mode);
                _taskPane.Visible = true;
            };
            _ribbon.OnSelectionModeRequested += (mode, text) =>
            {
                _viewModel.StartWithSelection(mode, text);
                _taskPane.Visible = true;
            };
            _ribbon.OnSettingsRequested += () =>
            {
                var settingsWindow = new SettingsWindow(_settingsManager);
                settingsWindow.ShowDialog();
            };
            return _ribbon;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
```

- [ ] **Step 2: Verify build**

Build → Build Solution. May have errors if SettingsWindow doesn't exist yet — that's Task 12.

- [ ] **Step 3: Commit**

```bash
git add GptOutlookPlugin/ThisAddIn.cs
git commit -m "feat: wire all components in ThisAddIn — AI service, context manager, ribbon, task pane"
```

---

## Task 12: Settings Dialog

**Files:**
- Create: `GptOutlookPlugin/UI/SettingsWindow.xaml`
- Create: `GptOutlookPlugin/UI/SettingsWindow.xaml.cs`

- [ ] **Step 1: Create SettingsWindow XAML**

Create `GptOutlookPlugin/UI/SettingsWindow.xaml`:

```xml
<Window x:Class="GptOutlookPlugin.UI.SettingsWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GPT Email Assistant - Settings"
        Width="450" Height="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- AI Provider Selection -->
        <GroupBox Grid.Row="0" Header="AI Provider" Margin="0,0,0,12">
            <StackPanel Margin="8">
                <RadioButton x:Name="RbCodex" Content="Codex CLI (WSL)" Margin="0,4"
                             IsChecked="True"/>
                <RadioButton x:Name="RbOpenAi" Content="OpenAI API" Margin="0,4"/>
            </StackPanel>
        </GroupBox>

        <!-- Codex CLI Settings -->
        <GroupBox Grid.Row="1" Header="Codex CLI Settings" Margin="0,0,0,12">
            <Grid Margin="8">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Command:" VerticalAlignment="Center"/>
                <TextBox Grid.Row="0" Grid.Column="1" x:Name="TxtCodexCommand" Margin="0,2" Padding="4"/>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Timeout (sec):" VerticalAlignment="Center"/>
                <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtCodexTimeout" Margin="0,2" Padding="4"/>
            </Grid>
        </GroupBox>

        <!-- OpenAI API Settings -->
        <GroupBox Grid.Row="2" Header="OpenAI API Settings" Margin="0,0,0,12">
            <Grid Margin="8">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="0" Text="API Key:" VerticalAlignment="Center"/>
                <PasswordBox Grid.Row="0" Grid.Column="1" x:Name="TxtApiKey" Margin="0,2" Padding="4"/>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Model:" VerticalAlignment="Center"/>
                <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtModel" Margin="0,2" Padding="4"/>
            </Grid>
        </GroupBox>

        <!-- Translation Default -->
        <GroupBox Grid.Row="3" Header="Translation" Margin="0,0,0,12">
            <Grid Margin="8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Target Language:" VerticalAlignment="Center"/>
                <ComboBox Grid.Column="1" x:Name="CmbLanguage" Padding="4">
                    <ComboBoxItem Content="ko" Tag="ko"/>
                    <ComboBoxItem Content="en" Tag="en"/>
                    <ComboBoxItem Content="ja" Tag="ja"/>
                    <ComboBoxItem Content="zh" Tag="zh"/>
                </ComboBox>
            </Grid>
        </GroupBox>

        <!-- Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Content="Save" Width="80" Margin="0,0,8,0" Padding="4"
                    Click="SaveButton_Click"/>
            <Button Content="Cancel" Width="80" Padding="4"
                    Click="CancelButton_Click"/>
        </StackPanel>
    </Grid>
</Window>
```

- [ ] **Step 2: Implement SettingsWindow code-behind**

Create `GptOutlookPlugin/UI/SettingsWindow.xaml.cs`:

```csharp
using System.Windows;
using GptOutlookPlugin.Core;

namespace GptOutlookPlugin.UI
{
    public partial class SettingsWindow : Window
    {
        private readonly SettingsManager _settingsManager;

        public SettingsWindow(SettingsManager settingsManager)
        {
            InitializeComponent();
            _settingsManager = settingsManager;
            LoadSettings();
        }

        private void LoadSettings()
        {
            var s = _settingsManager.Current;

            RbCodex.IsChecked = s.AiProvider == "codex-cli";
            RbOpenAi.IsChecked = s.AiProvider == "openai-api";

            TxtCodexCommand.Text = s.CodexCli.Command;
            TxtCodexTimeout.Text = s.CodexCli.TimeoutSeconds.ToString();

            TxtApiKey.Password = s.OpenAiApi.ApiKey;
            TxtModel.Text = s.OpenAiApi.Model;

            foreach (System.Windows.Controls.ComboBoxItem item in CmbLanguage.Items)
            {
                if (item.Tag?.ToString() == s.DefaultTranslateTarget)
                {
                    CmbLanguage.SelectedItem = item;
                    break;
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var s = _settingsManager.Current;

            s.AiProvider = RbOpenAi.IsChecked == true ? "openai-api" : "codex-cli";
            s.CodexCli.Command = TxtCodexCommand.Text;
            int.TryParse(TxtCodexTimeout.Text, out var timeout);
            s.CodexCli.TimeoutSeconds = timeout > 0 ? timeout : 60;

            s.OpenAiApi.ApiKey = TxtApiKey.Password;
            s.OpenAiApi.Model = TxtModel.Text;

            if (CmbLanguage.SelectedItem is System.Windows.Controls.ComboBoxItem selected)
                s.DefaultTranslateTarget = selected.Tag?.ToString() ?? "ko";

            _settingsManager.Save();
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
```

- [ ] **Step 3: Verify full build**

Build → Build Solution. All projects should compile with zero errors.

- [ ] **Step 4: Commit**

```bash
git add GptOutlookPlugin/UI/SettingsWindow.xaml GptOutlookPlugin/UI/SettingsWindow.xaml.cs
git commit -m "feat: add Settings dialog for AI provider, API key, and translation preferences"
```

---

## Task 13: Integration Test (Manual)

- [ ] **Step 1: Run all unit tests**

Run: Build → Run All Tests
Expected: All tests PASS (ConversationSession: 5, PromptTemplates: 4, AiServiceManager: 3, ContextManager: 6, DiffEngine: 5, CodexCliProvider: 3 = ~26 tests).

- [ ] **Step 2: Debug test with Outlook**

1. Set `GptOutlookPlugin` as startup project
2. Press F5 — Outlook opens with the add-in loaded
3. Verify "GPT Email Assistant" tab appears in Ribbon
4. Open any email → click "Summarize" → TaskPane should open
5. Verify chat messages appear in the TaskPane

- [ ] **Step 3: Test context menu**

1. Open an email in a separate window (double-click)
2. Select some text in the body
3. Right-click → "GPT Assistant" → "Translate Selection"
4. Verify TaskPane opens with the selection context

- [ ] **Step 4: Test Settings dialog**

1. Click "Settings" button in Ribbon
2. Verify settings load correctly
3. Change a setting, save, verify it persists (check `%APPDATA%\GptOutlookPlugin\appsettings.json`)

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "chore: complete integration — all unit tests passing, manual testing verified"
```
