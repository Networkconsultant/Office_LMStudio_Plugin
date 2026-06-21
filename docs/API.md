# ⚡ LMStudio API Reference

## Base URL

```
http://localhost:1234/v1
```

Configurable via the ⚙️ settings button in the task pane.

## Endpoints Used

### `GET /v1/models`

Lists all models currently loaded in LMStudio.

**Response**
```json
{
  "object": "list",
  "data": [
    {
      "id": "lmstudio-community/Meta-Llama-3-8B-Instruct-GGUF",
      "object": "model",
      "owned_by": "organization-owner"
    }
  ]
}
```

---

### `POST /v1/chat/completions`

**Request (non-streaming)**
```json
{
  "model": "your-model-id",
  "messages": [
    { "role": "system", "content": "You are a helpful assistant." },
    { "role": "user",   "content": "Hello!" }
  ],
  "temperature": 0.7,
  "stream": false
}
```

**Request (streaming)**
```json
{ "model": "your-model-id", "messages": [...], "stream": true }
```

**Streaming Response (SSE)**
```
data: {"choices":[{"delta":{"content":"Hello"},"index":0}]}
data: {"choices":[{"delta":{"content":"!"},"index":0}]}
data: [DONE]
```

---

## TypeScript Client Usage

```typescript
import { LMStudioClient } from "./api/lmstudio";

const client = new LMStudioClient("http://localhost:1234/v1");

// List models
const models = await client.getModels();

// Non-streaming
const reply = await client.chat({ model, messages });

// Streaming
for await (const chunk of client.chatStream({ model, messages, stream: true })) {
  appendToUI(chunk);
}
```

## Enabling the LMStudio Local Server

1. Open LMStudio
2. Click **Local Server** icon in left sidebar
3. Click **Start Server** (default port: **1234**)
4. Enable **CORS** if testing from browser devtools
