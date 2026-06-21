import * as React from "react";
import ChatPanel from "./ChatPanel";
import TextTransform from "./TextTransform";
import Summarizer from "./Summarizer";
import FormulaHelper from "./FormulaHelper";
import FileManager from "./FileManager";

type Tab = "chat" | "transform" | "summarize" | "formula" | "files";

function getHostType(): string {
  if (typeof Office === "undefined" || !Office.context?.host) return "unknown";
  return Office.HostType[Office.context.host] ?? "unknown";
}

export default function App() {
  const [tab, setTab] = React.useState<Tab>("chat");
  const [baseUrl, setBaseUrl] = React.useState("http://localhost:1234/v1");
  const [showSettings, setShowSettings] = React.useState(false);
  const hostType = React.useMemo(getHostType, []);
  const isExcel = hostType === "Excel";

  const tabs: Array<{ id: Tab; label: string }> = [
    { id: "chat", label: "💬 Chat" },
    { id: "transform", label: "✨ Transform" },
    { id: "summarize", label: "📄 Summarize" },
    ...(isExcel ? [{ id: "formula" as Tab, label: "⚡ Formula" }] : []),
    { id: "files", label: "📁 Files" },
  ];

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>🤖 LMStudio AI</h1>
        <button
          className="settings-btn"
          onClick={() => setShowSettings(!showSettings)}
          title="Settings"
        >
          ⚙️
        </button>
      </header>

      {showSettings && (
        <div className="settings-panel">
          <label>LMStudio Server URL</label>
          <input
            type="text"
            value={baseUrl}
            onChange={(e) => setBaseUrl(e.target.value)}
            placeholder="http://localhost:1234/v1"
          />
          <small>Default: http://localhost:1234/v1</small>
        </div>
      )}

      <nav className="tab-nav">
        {tabs.map((t) => (
          <button
            key={t.id}
            className={`tab-btn${tab === t.id ? " active" : ""}`}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </nav>

      <main className="tab-content">
        {tab === "chat" && <ChatPanel baseUrl={baseUrl} />}
        {tab === "transform" && <TextTransform baseUrl={baseUrl} />}
        {tab === "summarize" && <Summarizer baseUrl={baseUrl} />}
        {tab === "formula" && isExcel && <FormulaHelper baseUrl={baseUrl} />}
        {tab === "files" && <FileManager baseUrl={baseUrl} />}
      </main>
    </div>
  );
}
