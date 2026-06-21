# 🔬 Architecture

## Overview

```
┌─────────────────────────────────────────────────────┐
│                  Microsoft Office                    │
│                                                      │
│  ┌──────────────────────────────────────────────┐   │
│  │         Office Task Pane (WebView2)          │   │
│  │                                              │   │
│  │   https://localhost:3000/taskpane.html       │   │
│  │                                              │   │
│  │  ┌─────────────────────────────────────┐    │   │
│  │  │        React App (Office.js)        │    │   │
│  │  │                                     │    │   │
│  │  │  ┌──────────┐  ┌────────────────┐  │    │   │
│  │  │  │   Chat   │  │ Text Transform │  │    │   │
│  │  │  └──────────┘  └────────────────┘  │    │   │
│  │  │  ┌──────────┐  ┌────────────────┐  │    │   │
│  │  │  │Summarize │  │ Formula Helper │  │    │   │
│  │  │  └──────────┘  └────────────────┘  │    │   │
│  │  │                                     │    │   │
│  │  │  LMStudioClient (fetch + SSE)       │    │   │
│  │  └──────────────┬──────────────────────┘    │   │
│  └─────────────────│──────────────────────────┘   │
└────────────────────│────────────────────────────────┘
                     │ HTTP (localhost only)
                     ▼
┌─────────────────────────────────────────────────────┐
│            LMStudio Local Server :1234/v1           │
│                                                      │
│   GET  /v1/models                                   │
│   POST /v1/chat/completions  (stream: true/false)   │
└─────────────────────────────────────────────────────┘
```

## Component Breakdown

### `src/api/lmstudio.ts`
- `LMStudioClient` class — configurable base URL
- `getModels()` — populates model dropdowns
- `chat()` — single-shot completion
- `chatStream()` — SSE streaming generator

### `src/taskpane/components/App.tsx`
- Detects Office host via `Office.context.host`
- Shows Formula Helper tab for Excel only
- Settings panel with configurable server URL

### `src/taskpane/components/ChatPanel.tsx`
- Message history with real-time streaming cursor
- `Enter` to send, `Shift+Enter` for newline

### `src/taskpane/components/TextTransform.tsx`
- Reads selection via `Word.run()` or `Excel.run()`
- 8 transformation prompts
- Inserts result back via `selection.insertText()`

### `src/taskpane/components/Summarizer.tsx`
- Word: `document.body.text`
- Outlook: `mailbox.item.body.getAsync()`
- PowerPoint: shape text iteration
- Truncates at 12,000 chars for context window safety

### `src/taskpane/components/FormulaHelper.tsx`
- System prompt forces formula-first output
- Parses `=` prefixed line as the formula
- Inserts into selected cell via `range.formulas`

## Security Model

- Task pane runs in a sandboxed WebView2 context
- Only calls `localhost:1234` — no internet traffic
- Certificates issued by local CA (mkcert), trusted only on this machine
- No tokens, API keys, or user data leave the device
