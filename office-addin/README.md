# 📦 Office Add-in — office-addin/

This folder contains the **React + TypeScript** Office Add-in that runs as a web app served over HTTPS and loaded into Office via the task pane.

## Structure

```
office-addin/
├── manifest.xml             # Sideload manifest: Word, Excel, PowerPoint
├── manifest-outlook.xml     # Sideload manifest: Outlook
├── package.json
├── tsconfig.json
├── webpack.config.js
├── certs/                   # mkcert-generated localhost certs (gitignored)
│   ├── localhost.pem
│   └── localhost-key.pem
└── src/
    ├── api/
    │   └── lmstudio.ts      # LMStudio API client
    └── taskpane/
        ├── index.html
        ├── taskpane.tsx     # Entry point
        ├── components/
        │   ├── App.tsx
        │   ├── ChatPanel.tsx
        │   ├── TextTransform.tsx
        │   ├── Summarizer.tsx
        │   └── FormulaHelper.tsx
        └── styles/
            └── taskpane.css
```

## Commands

| Command | Description |
|---|---|
| `npm install` | Install dependencies |
| `npm run dev-server` | Start HTTPS dev server on port 3000 |
| `npm run build` | Production build → `dist/` |
| `npm run build:dev` | Development build → `dist/` |
| `npm run validate` | Validate `manifest.xml` |
| `npm run certs` | Install office-addin-dev-certs (alternative to mkcert) |

## LMStudio API Client (`src/api/lmstudio.ts`)

```typescript
const client = new LMStudioClient("http://localhost:1234/v1");

// List models
const models = await client.getModels();

// Single-shot chat
const reply = await client.chat({ model, messages });

// Streaming chat (SSE)
for await (const chunk of client.chatStream({ model, messages, stream: true })) {
  process.stdout.write(chunk);
}
```
