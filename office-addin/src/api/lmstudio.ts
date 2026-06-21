export interface LMStudioModel {
  id: string;
  object: string;
  created: number;
  owned_by: string;
}

export interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

export interface ChatCompletionRequest {
  model: string;
  messages: ChatMessage[];
  temperature?: number;
  max_tokens?: number;
  stream?: boolean;
}

export class LMStudioClient {
  private baseUrl: string;

  constructor(baseUrl = "http://localhost:1234/v1") {
    this.baseUrl = baseUrl.replace(/\/$/, "");
  }

  setBaseUrl(url: string): void {
    this.baseUrl = url.replace(/\/$/, "");
  }

  /** Fetch available models from LMStudio. */
  async getModels(): Promise<LMStudioModel[]> {
    const res = await fetch(`${this.baseUrl}/models`);
    if (!res.ok) throw new Error(`Failed to fetch models: ${res.statusText}`);
    const data = await res.json();
    return data.data ?? [];
  }

  /** Single-shot chat completion (non-streaming). */
  async chat(req: ChatCompletionRequest): Promise<string> {
    const res = await fetch(`${this.baseUrl}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...req, stream: false }),
    });
    if (!res.ok) throw new Error(`Chat request failed: ${res.statusText}`);
    const data = await res.json();
    return data.choices?.[0]?.message?.content ?? "";
  }

  /** Streaming chat completion — yields text chunks as they arrive (SSE). */
  async *chatStream(req: ChatCompletionRequest): AsyncGenerator<string> {
    const res = await fetch(`${this.baseUrl}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...req, stream: true }),
    });
    if (!res.ok) throw new Error(`Stream request failed: ${res.statusText}`);

    const reader = res.body!.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop()!;
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed === "data: [DONE]") continue;
        if (trimmed.startsWith("data: ")) {
          try {
            const json = JSON.parse(trimmed.slice(6));
            const content: string | undefined = json.choices?.[0]?.delta?.content;
            if (content) yield content;
          } catch {
            // skip malformed SSE chunks
          }
        }
      }
    }
  }
}

/** Singleton client — components may create their own or use this default. */
export const defaultClient = new LMStudioClient();
