# GPT Outlook Plugin - Design Specification

**Date:** 2026-03-31
**Status:** Approved

## Overview

Outlook Classic (Windows Desktop)용 VSTO Add-in으로, ChatGPT를 활용하여 이메일 리뷰/교정, 작성 보조, 번역, 요약 기능을 제공한다. 사이드 패널에서 대화형 상호작용을 지원하며, 이메일별 컨텍스트를 유지한다.

## Requirements

### 기능 (우선순위 순)

1. **이메일 리뷰/교정** - 문법, 톤, 명확성 검토 및 수정 제안
2. **이메일 작성 보조** - 답장 초안 생성, 적절한 표현 제시
3. **이메일 번역** - 이메일을 한국어/영어 등 대상 언어로 번역
4. **이메일 요약** - 긴 이메일이나 스레드의 핵심 요약

### 비기능 요구사항

- 개인 사용 전용 (Visual Studio 디버그 모드 실행)
- AI 백엔드: WSL Codex CLI (1순위), OpenAI API (fallback)
- 대화 컨텍스트 유지: 동일 이메일에 대해 추가 요청 가능

## Architecture

### Technology Stack

- **Plugin Type:** VSTO Add-in (C# / .NET Framework)
- **UI:** WPF (CustomTaskPane via ElementHost)
- **AI Primary:** Codex CLI via `wsl.exe` subprocess
- **AI Fallback:** OpenAI API (gpt-4o) via HttpClient
- **Pattern:** MVVM for WPF UI

### Component Diagram

```
+-----------------------------------------------------------+
|                    Outlook Classic                          |
|  +-----------+  +--------------------------------------+   |
|  | Ribbon Tab |  | CustomTaskPane (WPF via ElementHost) |   |
|  | [Review]  |  | +----------------------------------+ |   |
|  | [Compose] |  | | Chat Messages (ItemsControl)     | |   |
|  | [Translate]| | |   AI: Review result...           | |   |
|  | [Summarize]| | |   User: Make it softer           | |   |
|  | [Settings]|  | |   AI: Revised result...          | |   |
|  +-----------+  | +----------------------------------+ |   |
|  Context Menu   | | [Input box]            [Send]    | |   |
|                 | +----------------------------------+ |   |
|                 +--------------------------------------+   |
|                                                           |
|  +-----------------------------------------------------+  |
|  |              VSTO Add-in Core (C#)                   |  |
|  |  +------------+ +------------+ +------------------+  |  |
|  |  | Outlook    | | Context    | | AI Service       |  |  |
|  |  | Interop    | | Manager    | | (CodexCli /      |  |  |
|  |  | Layer      | |            | |  OpenAI API)     |  |  |
|  |  +------------+ +------------+ +------------------+  |  |
|  +-----------------------------------------------------+  |
+-----------------------------------------------------------+
```

### Core Components

#### 1. Outlook Interop Layer (`OutlookInterop.cs`)

현재 이메일의 정보를 읽고 수정 결과를 반영하는 레이어.

- 현재 선택된/작성 중인 `MailItem` 접근
- Subject, Body (HTML/Plain), Recipients, SenderEmailAddress 추출
- 이메일 본문에 AI 결과 반영 (사용자 확인 후)

#### 2. Context Manager (`ContextManager.cs`, `ConversationSession.cs`)

이메일별 대화 세션을 관리.

```
Dictionary<string, ConversationSession>
Key = MailItem.EntryID (이메일 고유 ID)
```

**ConversationSession 모델:**

| Field | Type | Description |
|-------|------|-------------|
| SystemPrompt | string | 기능별 프롬프트 템플릿 + 이메일 컨텍스트 |
| EmailContext | EmailContext | 제목, 본문, 수신자, 감지된 언어 |
| Messages | List\<ChatMessage\> | 대화 히스토리 (role + content) |
| CurrentMode | FeatureMode enum | Review, Compose, Translate, Summarize |

**세션 키 전략:**
- 기존 이메일: `MailItem.EntryID` 사용
- 새 이메일 작성 중 (EntryID 미존재): 임시 GUID 할당, 저장 후 EntryID로 교체

**컨텍스트 윈도우 전략:**
- 최근 10개 메시지 유지
- 이메일 본문은 항상 시스템 프롬프트에 포함
- 토큰 초과 시 오래된 메시지부터 제거

#### 3. AI Service (`IAiProvider.cs`, `AiServiceManager.cs`)

이중 백엔드 전략으로 안정성 확보.

**IAiProvider 인터페이스:**

```csharp
public interface IAiProvider
{
    Task<string> SendAsync(List<ChatMessage> messages, CancellationToken ct);
    bool IsAvailable();
}
```

**CodexCliProvider (1순위):**
- `Process.Start("wsl.exe", "codex exec -s read-only \"<prompt>\"")`
- stdout 파싱하여 응답 추출
- 타임아웃: 60초
- 실패 시 OpenAiApiProvider로 자동 전환

**OpenAiApiProvider (fallback):**
- HttpClient -> `api.openai.com/v1/chat/completions`
- 모델: gpt-4o
- API 키: 로컬 설정 파일에서 로드
- Streaming 지원으로 사이드 패널에 실시간 표시

### Prompt Templates

| Feature | System Prompt Core |
|---------|-------------------|
| Review | "You are an email proofreading expert. Review grammar, tone, and clarity. Present corrections in a before/after comparison format." |
| Compose | "You are a business email writing expert. Draft an appropriate reply considering the recipient and context." |
| Translate | "You are a professional translator. Translate to {target_lang} while maintaining the email's tone and formality." |
| Summarize | "You are an email summarization expert. Concisely organize key points, requests, and required actions." |

## UI Design

### Ribbon Tab: "GPT Email Assistant"

| Button | Icon | Action |
|--------|------|--------|
| Review | pencil | 현재 이메일 교정 시작, TaskPane 열기 |
| Compose | edit | 답장 초안 생성, TaskPane 열기 |
| Translate | globe | 번역 (언어 드롭다운 포함), TaskPane 열기 |
| Summarize | list | 이메일/스레드 요약, TaskPane 열기 |
| Settings | gear | 설정 다이얼로그 열기 |

### Context Menu (텍스트 선택 시)

- "GPT: Review Selection"
- "GPT: Translate Selection"
- "GPT: Compose Reply from Selection"

### ChatTaskPane (WPF)

- 상단: 현재 모드 표시 + 모드 전환 탭
- 중앙: 채팅 메시지 리스트 (ItemsControl + DataTemplate)
  - AI 메시지: 왼쪽 정렬, "적용" 버튼 포함
  - 사용자 메시지: 오른쪽 정렬
  - **Diff View**: Review/Translate 결과는 변경 사항을 diff 형태로 표시
- 하단: 텍스트 입력 + 전송 버튼

### Diff View (변경 사항 표시)

Review, Translate 등 원문 대비 변경이 발생하는 기능에서 결과를 diff view로 표시한다.

**표시 방식: Inline Diff**

```
  Dear Mr. Kim,
- I want to talk about the project timeline.
+ I would like to discuss the project timeline.
  As we discussed last week,
- the deadline is too short for us.
+ the current deadline presents a challenge for our team.
```

- 삭제된 텍스트: 빨간 배경 + 취소선
- 추가된 텍스트: 초록 배경
- 변경 없는 텍스트: 기본 배경 (컨텍스트 라인)

**구현:**
- AI 응답에서 원문과 수정문을 파싱
- 문장 단위 diff 알고리즘 적용 (문자 단위는 너무 세밀하여 가독성 저하)
- WPF `RichTextBox` 또는 커스텀 `FlowDocument`로 색상 표시
- 토글 버튼으로 Diff View ↔ 수정문만 보기 전환 가능

### User Flow

```
이메일 열기/작성
  ├── 리본 버튼 클릭 → 해당 모드로 TaskPane 열기
  ├── 텍스트 선택 + 우클릭 → 컨텍스트 메뉴에서 기능 선택
  └── TaskPane에서 직접 입력 → 추가 요청 전송
       │
       ▼
  Context Manager: 세션 생성 또는 기존 세션 로드
       │
       ▼
  AI Service: 프롬프트 구성 & 전송
       │
       ▼
  응답을 TaskPane 채팅에 표시
       │
       ├── [적용] → 이메일 본문에 결과 반영 (모드별 동작 참조)
       ├── [추가 요청] → 대화 계속
       └── [모드 전환] → 새 모드로 전환 (세션 유지)
```

### "적용" 동작 (모드별)

| Mode | 적용 동작 |
|------|----------|
| Review | 이메일 본문을 교정된 텍스트로 교체 (선택 영역만 교정한 경우 해당 부분만) |
| Compose | 새 답장 창에 초안 삽입, 또는 현재 작성 중인 본문에 삽입 |
| Translate | 이메일 본문을 번역 결과로 교체 |
| Summarize | "적용" 없음 — 읽기 전용 결과 표시만 (클립보드 복사 버튼 제공) |
```

## Project Structure

```
GptOutlookPlugin/
├── GptOutlookPlugin.sln
├── GptOutlookPlugin/
│   ├── ThisAddIn.cs                  -- VSTO entry point
│   ├── GptRibbon.xml                 -- Ribbon XML definition
│   ├── GptRibbon.cs                  -- Ribbon event handlers
│   │
│   ├── UI/
│   │   ├── ChatTaskPane.xaml         -- WPF chat panel
│   │   ├── ChatTaskPane.xaml.cs      -- Code-behind
│   │   ├── ChatMessage.cs            -- Message model (User/Assistant)
│   │   ├── SettingsWindow.xaml       -- Settings dialog
│   │   └── TaskPaneHost.cs           -- WPF -> WinForms hosting adapter
│   │
│   ├── Services/
│   │   ├── IAiProvider.cs            -- AI provider interface
│   │   ├── CodexCliProvider.cs       -- WSL Codex CLI invocation
│   │   ├── OpenAiApiProvider.cs      -- OpenAI API invocation
│   │   └── AiServiceManager.cs      -- Provider selection & fallback
│   │
│   ├── Core/
│   │   ├── ContextManager.cs         -- Conversation session & history
│   │   ├── ConversationSession.cs    -- Session model
│   │   ├── PromptTemplates.cs        -- Feature-specific system prompts
│   │   └── OutlookInterop.cs         -- Email read/write helpers
│   │
│   └── Config/
│       ├── Settings.cs               -- User settings model
│       └── appsettings.json          -- API key, defaults, timeouts
│
└── docs/
    └── superpowers/
        └── specs/
            └── 2026-03-31-gpt-outlook-plugin-design.md
```

## Configuration

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

## Development Environment

- **IDE:** Visual Studio 2022 (Windows)
- **Target Framework:** .NET Framework 4.8 (VSTO compatibility)
- **Outlook:** Outlook Classic (Desktop, Windows)
- **Build/Run:** F5 in Visual Studio (auto-launches Outlook with add-in)
- **Distribution:** Personal use only (debug mode)

## Constraints & Decisions

1. **VSTO over Office Web Add-in:** 단일 프로세스로 단순한 아키텍처, CLI 직접 호출 가능
2. **WPF over WinForms for UI:** 데이터 바인딩, MVVM 패턴으로 채팅 UI 구현이 자연스러움
3. **EntryID as session key:** Outlook 내 이메일 고유 식별, 이메일 전환 시 자동 세션 관리
4. **Dual AI backend:** Codex CLI 장애 시에도 OpenAI API로 서비스 연속성 보장
