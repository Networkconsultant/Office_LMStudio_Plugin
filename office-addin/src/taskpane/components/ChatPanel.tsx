import * as React from "react";
import { LMStudioClient, ChatMessage, LMStudioModel } from "../../api/lmstudio";
import { readDocumentText, readSelectionText, wordInsertAtEnd, insertTextViaOfficeJs, getHostApp } from "../../api/doc";

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
  const [docContext, setDocContext] = React.useState("");
  const [docContextLabel, setDocContextLabel] = React.useState("");
  const bottomRef = React.useRef<HTMLDivElement>(null);
  const host = React.useMemo(getHostApp, []);

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

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

  React.useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, streamingText]);

  const attachDocument = async () => {
    try {
      const text = await readDocumentText(10000);
      setDocContext(text);
      setDocContextLabel(`📄 Full document (${text.length.toLocaleString()} chars)`);
    } catch (e: any) {
      setError(`Could not read document: ${e.message}`);
    }
  };

  const attachSelection = async () => {
    try {
      const text = await readSelectionText();
      if (!text.trim()) { setError("Nothing selected"); return; }
      setDocContext(text);
      setDocContextLabel(`📋 Selection (${text.length.toLocaleString()} chars)`);
    } catch (e: any) {
      setError(`Could not read selection: ${e.message}`);
    }
  };

  const send = async () => {
    if (!input.trim() || loading) return;
    // Prepend doc context to first message if attached
    let content = input.trim();
    if (docContext && messages.length === 0) {
      content = `Document context:\n\n${docContext}\n\n---\n\n${content}`;
    }
    const userMsg: ChatMessage = { role: "user", content: input.trim() };
    const sendMsg: ChatMessage = { role: "user", content };
    const history = [...messages, userMsg];
    const sendHistory = [...messages.slice(0, -0), sendMsg];
    setMessages(history);
    setInput("");
    setLoading(true);
    setStreamingText("");
    setError("");

    try {
      let full = "";
      for await (const chunk of client.chatStream({
        model: selectedModel,
        messages: sendHistory,
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

  const insertLastResponse = async () => {
    const lastAI = [...messages].reverse().find((m) => m.role === "assistant");
    if (!lastAI) return;
    try {
      if (host === "Word") {
        await wordInsertAtEnd(lastAI.content);
      } else {
        await insertTextViaOfficeJs(lastAI.content);
      }
      setError("✅ Inserted into document");
    } catch (e: any) {
      setError(`Insert error: ${e.message}`);
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
            <option key={m.id} value={m.id}>{m.id}</option>
          ))}
        </select>
        <button className="icon-btn" onClick={() => { setMessages([]); setDocContext(""); setDocContextLabel(""); }}
          title="Clear chat" disabled={loading}>🗑️</button>
      </div>

      {/* Document context bar */}
      <div className="doc-context-bar">
        <button className="small-btn" onClick={attachDocument} title="Attach full document as context">
          📄 Attach Doc
        </button>
        <button className="small-btn" onClick={attachSelection} title="Attach selected text as context">
          📋 Selection
        </button>
        {messages.length > 0 && (
          <button className="small-btn" onClick={insertLastResponse} title="Insert last AI response into document">
            📝 Insert to Doc
          </button>
        )}
        {docContextLabel && (
          <span className="context-badge" title={docContext}>
            {docContextLabel}
            <button onClick={() => { setDocContext(""); setDocContextLabel(""); }}
              style={{ marginLeft: 4, background: "none", border: "none", cursor: "pointer", fontSize: 10 }}>✕</button>
          </span>
        )}
      </div>

      {error && <p className={error.startsWith("✅") ? "status-msg" : "error-msg"}>{error}</p>}

      {/* Message history */}
      <div className="chat-log">
        {messages.length === 0 && (
          <p className="empty-hint">
            Send a message to chat with your local AI.{"\n"}
            Use "Attach Doc" to give the AI context from your {host} document.
          </p>
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
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); }
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
