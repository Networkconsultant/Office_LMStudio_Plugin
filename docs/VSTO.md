# 🏗️ VSTO Add-in (C# .NET)

The VSTO add-in provides a **Windows Desktop-only** alternative with deeper Office integration.

## Prerequisites

- Windows 10/11
- Visual Studio 2022 with **"Office/SharePoint Development"** workload
- Microsoft Office 2016+ Desktop
- .NET Framework 4.7.2

## Project Structure

```
vsto/
├── LMStudioAddin.sln
└── LMStudioAddin/
    ├── LMStudioAddin.csproj
    ├── Properties/AssemblyInfo.cs
    ├── ThisAddIn.cs           # Entry point + task pane management
    ├── ThisAddIn.Ribbon.cs    # Ribbon extensibility override
    ├── Ribbon.cs              # CustomUI XML + button handler
    ├── LMStudioClient.cs      # HttpClient + SSE streaming
    └── TaskPaneControl.cs     # WinForms: Chat, Transform, Summarize tabs
```

## Running

1. Open `vsto/LMStudioAddin.sln` in Visual Studio 2022
2. Press **F5** — builds, deploys, and launches Word
3. **"LMStudio AI"** group appears in the **Home** ribbon

## LMStudioClient.cs

```csharp
using var client = new LMStudioClient("http://localhost:1234/v1");

// List models
var models = await client.GetModelsAsync();

// Chat
var reply = await client.ChatAsync(new ChatRequest {
    Model = "llama3",
    Messages = new List<ChatMessage> {
        new ChatMessage { Role = "user", Content = "Hello!" }
    }
});

// Streaming
var full = await client.ChatStreamAsync(request, chunk => AppendToUI(chunk));
```

## Deployment

1. **Build → Publish** in Visual Studio (ClickOnce)
2. Distribute the installer — no VS required on target machine
3. Production deployment requires a code-signing certificate
