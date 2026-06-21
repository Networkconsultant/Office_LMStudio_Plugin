import * as React from "react";
import { LMStudioClient, ChatMessage, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

export default function ChatPanel({ baseUrl }: Props) {
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [streamingText, setStreamingText] = React.useState("");
  const [error, setError] = React.useState("");
  const bottomRef = React.useRef<HTMLDivElement>(null);

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  /* Load models whenever baseUrl changes */
  React.useEffect(() => {
    setError("");
    client
      .getModels()
      .then((m) => {
        setModels(m);
        if (m.length > 0 && !selectedModel) setSelectedModel(m[0].id);
      })
      .catch(() => setError("⚠️ Cannot connect to LMStudio. Is the local server running?"));
  }, [client]);

  /* Auto-scroll to latest message */
  React.useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, streamingText]);

  const send = async () => {
    if (!input.trim() || loading) return;
    const userMsg: ChatMessage = { role: "user", content: input.trim() };
    const history = [...messages, userMsg];
    setMessages(history);
    setInput("");
    setLoading(true);
    setStreamingText("");
    setError("");

    try {
      let full = "";
      for await (const chunk of client.chatStream({
        model: selectedModel,
        messages: history,
        stream: true,
      })) {
        full += chunk;
        setStreamingText(full);
      }
      setMessages([...history, { role: "assistant", content: full }]);
    } catch (e: any) {
      setError(`Error: ${e.message}`);
    } finally {
      setLoading(false);
      setStreamingText("");
    }
  };

  return (
    <div className="panel chat-panel">
      {/* Model selector */}
      <div className="model-row">
        <select
          value={selectedModel}
          onChange={(e) => setSelectedModel(e.target.value)}
          disabled={models.length === 0}
        >
          {models.length === 0 && <option>— no models found —</option>}
          {models.map((m) => (
            <option key={m.id} value={m.id}>
              {m.id}
            </option>
          ))}
        </select>
        <button
          className="icon-btn"
          onClick={() => setMessages([])}
          title="Clear chat"
          disabled={loading}
        >
          🗑️
        </button>
      </div>

      {error && <p className="error-msg">{error}</p>}

      {/* Message history */}
      <div className="chat-log">
        {messages.length === 0 && (
          <p className="empty-hint">Send a message to start chatting with your local AI model.</p>
        )}
        {messages.map((msg, i) => (
          <div key={i} className={`chat-bubble ${msg.role}`}>
            <span className="bubble-label">{msg.role === "user" ? "You" : "AI"}</span>
            <p>{msg.content}</p>
          </div>
        ))}
        {loading && (
          <div className="chat-bubble assistant streaming">
            <span className="bubble-label">AI</span>
            <p>
              {streamingText || "Thinking…"}
              {streamingText && <span className="cursor">▋</span>}
            </p>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {/* Input area */}
      <div className="chat-input-row">
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Type a message… (Enter to send, Shift+Enter for newline)"
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
          rows={2}
          disabled={loading}
        />
        <button
          className="send-btn"
          onClick={send}
          disabled={loading || !input.trim() || !selectedModel}
          title="Send"
        >
          ➤
        </button>
      </div>
    </div>
  );
}
