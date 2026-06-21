# 🤖 Office + LMStudio Plugin

> Chat with your **local LMStudio AI models** directly inside Microsoft Office — Word, Excel, Outlook, and PowerPoint.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Office Add-in](https://img.shields.io/badge/Office-Add--in-orange?logo=microsoft)](https://learn.microsoft.com/en-us/office/dev/add-ins/)
[![LMStudio](https://img.shields.io/badge/LMStudio-Local%20AI-purple)](https://lmstudio.ai/)
[![Node.js](https://img.shields.io/badge/Node.js-18%2B-green?logo=node.js)](https://nodejs.org/)

---

## ✨ Features

| Feature | Word | Excel | Outlook | PowerPoint |
|---|:---:|:---:|:---:|:---:|
| 💬 **Chat** — full conversational AI | ✅ | ✅ | ✅ | ✅ |
| ✨ **Text Transform** — rewrite, summarize, translate, improve | ✅ | ✅ | ✅ | ✅ |
| 📄 **Document Summarizer** — summarize full content | ✅ | — | ✅ | ✅ |
| ⚡ **Formula Helper** — natural language → Excel formula | — | ✅ | — | — |

All features run **100% locally** — no data leaves your machine.

---

## 🏗️ Architecture

This project ships two delivery formats:

```
Office_LMStudio_Plugin/
├── office-addin/        # Web-based Office Add-in (React + TypeScript)
│   ├── manifest.xml         # Word, Excel, PowerPoint manifest
│   ├── manifest-outlook.xml # Outlook manifest
│   └── src/
│       ├── api/lmstudio.ts         # LMStudio OpenAI-compatible API client
│       └── taskpane/components/
│           ├── App.tsx             # Root: host detection, settings, tabs
│           ├── ChatPanel.tsx       # Streaming chat with model selector
│           ├── TextTransform.tsx   # 8 text transformation modes
│           ├── Summarizer.tsx      # Full document / email summarization
│           └── FormulaHelper.tsx   # NL → Excel formula generator
└── vsto/                # VSTO Add-in (C# .NET, Windows Desktop only)
    └── LMStudioAddin/
        ├── LMStudioClient.cs   # HttpClient + SSE streaming
        ├── TaskPaneControl.cs  # WinForms task pane: Chat, Transform, Summarize
        ├── Ribbon.cs           # Word ribbon button
        └── ThisAddIn.cs        # Entry point + task pane toggle
```

### LMStudio API Integration

Connects to LMStudio's **OpenAI-compatible local REST API**:

| Endpoint | Usage |
|---|---|
| `GET /v1/models` | Populate model dropdown at startup |
| `POST /v1/chat/completions` | All AI features (streaming SSE supported) |

Default base URL: `http://localhost:1234/v1` (configurable via ⚙️ in the task pane).

---

## 🚀 Quick Start

### Prerequisites

- [LMStudio](https://lmstudio.ai/) installed with **Local Server** enabled (port 1234)
- [Node.js 18+](https://nodejs.org/) and npm
- Microsoft Office 2016 or later (Desktop)

### 1. Generate project files

```bash
node setup.js
```

### 2. Install dependencies & certificates

```bash
cd office-addin
npm install
```

Certificates (mkcert) are generated automatically. Install mkcert CA once:

```bash
# Install mkcert via winget (Windows)
winget install FiloSottile.mkcert
mkcert -install
mkcert -key-file certs/localhost-key.pem -cert-file certs/localhost.pem localhost 127.0.0.1
```

### 3. Start the dev server

```bash
npm run dev-server
# Serves at https://localhost:3000
```

### 4. Sideload into Office

Open the provided launcher document on your Desktop:

```
C:\Users\<You>\Desktop\LMStudio-AI.docx
```

Or sideload manually — see [SIDELOADING.md](docs/SIDELOADING.md).

---

## 📖 Documentation

| Document | Description |
|---|---|
| [docs/SIDELOADING.md](docs/SIDELOADING.md) | How to load the add-in into each Office app |
| [docs/VSTO.md](docs/VSTO.md) | Building and deploying the VSTO add-in |
| [docs/API.md](docs/API.md) | LMStudio API integration details |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Component and data flow diagrams |

---

## 🔧 Configuration

| Setting | Default | Description |
|---|---|---|
| Server URL | `http://localhost:1234/v1` | LMStudio local server URL |
| Dev server port | `3000` | Webpack dev server port |

Change the server URL at runtime via the **⚙️ Settings** button in the task pane header.

---

## 🛠️ Development

```bash
# Development build (watch mode via dev-server)
npm run dev-server

# Production build
npm run build

# Validate manifest
npm run validate
```

### Tech Stack

- **Office Add-in**: React 18, TypeScript 5, Office.js, Webpack 5
- **VSTO**: C# .NET Framework 4.7.2, Microsoft.Office.Interop.Word
- **Certificates**: mkcert (locally trusted CA)
- **AI Backend**: LMStudio OpenAI-compatible REST API

---

## 📋 Text Transformation Modes

The **Transform** tab supports 8 modes:

| Mode | Description |
|---|---|
| **Improve** | Fix grammar, clarity, and style |
| **Rewrite** | Rewrite more clearly and concisely |
| **Summarize** | Condense to 2-3 sentences |
| **Translate** | Auto-detect and translate to English |
| **Formal** | Rewrite in a professional tone |
| **Casual** | Rewrite in a friendly, conversational tone |
| **Bullet Points** | Convert to a bullet-point list |
| **Expand** | Add more detail and context |

---

## 🔒 Privacy

- All AI processing happens **on your local machine** via LMStudio
- No data is sent to any cloud service
- The add-in communicates only with `localhost:1234`

---

## 📄 License

[MIT](LICENSE) — © 2025 Networkconsultant
