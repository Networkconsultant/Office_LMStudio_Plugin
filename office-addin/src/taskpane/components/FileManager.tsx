import * as React from "react";
import { bridgeClient, FileEntry, DirListing } from "../../api/bridge";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

type AIAction = "summarize" | "rewrite" | "qa" | "extract-key-points" | "translate";

const ACTION_PROMPTS: Record<AIAction, (text: string, extra?: string) => string> = {
  summarize: (t) => `Summarize the following document content concisely:\n\n${t}`,
  rewrite:   (t) => `Rewrite the following content for clarity and professionalism:\n\n${t}`,
  qa:        (t, q) => `Using the following document as context, answer this question: "${q}"\n\nDocument:\n${t}`,
  "extract-key-points": (t) => `Extract the key points from the following document as a numbered list:\n\n${t}`,
  translate: (t) => `Translate the following document to English (detect the source language automatically):\n\n${t}`,
};

const SUPPORTED_EXTS = new Set([".txt",".md",".csv",".json",".xml",".html",".htm",
  ".js",".ts",".tsx",".jsx",".py",".cs",".css",".log",".docx",".xlsx",".xls"]);

const MAX_CHARS = 14000;

/** Format bytes as KB/MB */
function fmtSize(bytes: number | null): string {
  if (bytes === null) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Insert text into current Word document as formatted paragraphs */
async function insertIntoDocument(text: string, heading?: string): Promise<void> {
  await Word.run(async (ctx) => {
    const body = ctx.document.body;
    if (heading) {
      const h = body.insertParagraph(heading, Word.InsertLocation.end);
      h.style = "Heading 2";
    }
    // Split on double newlines to create separate paragraphs
    const paragraphs = text.split(/\n{2,}/);
    for (const para of paragraphs) {
      const trimmed = para.trim();
      if (!trimmed) continue;
      // Detect bullet lines
      if (trimmed.match(/^[-*•]\s/m)) {
        const lines = trimmed.split("\n").filter((l) => l.trim());
        for (const line of lines) {
          const p = body.insertParagraph(line.replace(/^[-*•]\s*/, ""), Word.InsertLocation.end);
          p.style = "List Paragraph";
        }
      } else {
        body.insertParagraph(trimmed, Word.InsertLocation.end);
      }
    }
    await ctx.sync();
  });
}

export default function FileManager({ baseUrl }: Props) {
  const [bridgeAvailable, setBridgeAvailable] = React.useState<boolean | null>(null);
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");

  // File browser state
  const [currentDir, setCurrentDir] = React.useState("");
  const [listing, setListing] = React.useState<DirListing | null>(null);
  const [loadingDir, setLoadingDir] = React.useState(false);
  const [breadcrumbs, setBreadcrumbs] = React.useState<string[]>([]);

  // File content state
  const [selectedFile, setSelectedFile] = React.useState<FileEntry | null>(null);
  const [fileContent, setFileContent] = React.useState("");
  const [loadingFile, setLoadingFile] = React.useState(false);

  // AI state
  const [action, setAction] = React.useState<AIAction>("summarize");
  const [qaQuestion, setQaQuestion] = React.useState("");
  const [result, setResult] = React.useState("");
  const [processing, setProcessing] = React.useState(false);

  // Edit state
  const [editMode, setEditMode] = React.useState(false);
  const [editContent, setEditContent] = React.useState("");

  const [status, setStatus] = React.useState("");

  const lmClient = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  // ── Init ──────────────────────────────────────────────────────────────────
  React.useEffect(() => {
    bridgeClient.isAvailable().then(async (ok) => {
      setBridgeAvailable(ok);
      if (ok) {
        const home = await bridgeClient.getHomeDirs();
        navigateTo(home.documents);
      }
    });
    lmClient.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, []);

  // ── Navigation ────────────────────────────────────────────────────────────
  const navigateTo = async (dir: string) => {
    setLoadingDir(true);
    setSelectedFile(null);
    setFileContent("");
    setResult("");
    try {
      const data = await bridgeClient.listDir(dir);
      setListing(data);
      setCurrentDir(dir);
      // Build breadcrumbs from path
      const parts = dir.replace(/\\/g, "/").split("/").filter(Boolean);
      setBreadcrumbs(parts.map((_, i) =>
        (dir.startsWith("\\\\") ? "\\\\" : "") +
        dir.replace(/\\/g, "/").split("/").filter(Boolean).slice(0, i + 1).join("/")
      ));
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    } finally {
      setLoadingDir(false);
    }
  };

  const navigateUp = () => {
    if (!currentDir) return;
    const parts = currentDir.replace(/\\/g, "/").split("/").filter(Boolean);
    if (parts.length <= 1) return;
    const parent = currentDir.replace(/\\/g, "/")
      .split("/").filter(Boolean).slice(0, -1).join("/");
    navigateTo(parent.includes(":") ? parent : "/" + parent);
  };

  // ── Load file ─────────────────────────────────────────────────────────────
  const loadFile = async (entry: FileEntry) => {
    if (entry.type === "dir") { navigateTo(entry.path); return; }
    if (!entry.ext || !SUPPORTED_EXTS.has(entry.ext)) {
      setStatus(`Unsupported file type: ${entry.ext}`); return;
    }
    setLoadingFile(true);
    setSelectedFile(entry);
    setResult("");
    setEditMode(false);
    try {
      const data = await bridgeClient.readFile(entry.path);
      const truncated = data.text.length > MAX_CHARS
        ? data.text.slice(0, MAX_CHARS) + `\n\n[Truncated — ${data.text.length.toLocaleString()} total chars]`
        : data.text;
      setFileContent(truncated);
      setEditContent(data.text);
      setStatus("");
    } catch (e: any) {
      setStatus(`Error reading file: ${e.message}`);
    } finally {
      setLoadingFile(false);
    }
  };

  // ── AI processing ─────────────────────────────────────────────────────────
  const runAI = async () => {
    if (!fileContent || !selectedModel) return;
    setProcessing(true);
    setResult("");
    setStatus("Processing with AI…");
    try {
      const prompt = ACTION_PROMPTS[action](fileContent, action === "qa" ? qaQuestion : undefined);
      const response = await lmClient.chat({
        model: selectedModel,
        messages: [{ role: "user", content: prompt }],
      });
      setResult(response);
      setStatus("");
    } catch (e: any) {
      setStatus(`AI error: ${e.message}`);
    } finally {
      setProcessing(false);
    }
  };

  // ── Save file ─────────────────────────────────────────────────────────────
  const saveFile = async () => {
    if (!selectedFile) return;
    try {
      await bridgeClient.writeFile(selectedFile.path, editContent);
      setFileContent(editContent);
      setEditMode(false);
      setStatus(`✅ Saved: ${selectedFile.name}`);
      // Refresh listing
      await navigateTo(currentDir);
    } catch (e: any) {
      setStatus(`Save error: ${e.message}`);
    }
  };

  const saveResultAsFile = async () => {
    if (!result) return;
    try {
      const base = selectedFile?.name.replace(/\.[^.]+$/, "") || "output";
      const res = await bridgeClient.saveAs(result, `${base}-ai-output.txt`, currentDir);
      setStatus(`✅ Saved: ${res.path}`);
      await navigateTo(currentDir);
    } catch (e: any) {
      setStatus(`Save error: ${e.message}`);
    }
  };

  const insertResultIntoDoc = async () => {
    if (!result) return;
    try {
      const heading = selectedFile ? `AI Output — ${selectedFile.name}` : "AI Output";
      await insertIntoDocument(result, heading);
      setStatus("✅ Inserted into document");
    } catch (e: any) {
      setStatus(`Insert error: ${e.message}. (Only supported in Word)`);
    }
  };

  // ── Render: bridge not running ────────────────────────────────────────────
  if (bridgeAvailable === false) {
    return (
      <div className="panel">
        <div className="bridge-warning">
          <h3>📁 File Bridge Not Running</h3>
          <p>The file bridge server is required for local file access.</p>
          <p>Start it in a terminal:</p>
          <pre>cd office-addin{"\n"}npm run bridge</pre>
          <p>Then refresh this panel.</p>
          <button className="action-btn" onClick={() =>
            bridgeClient.isAvailable().then(ok => setBridgeAvailable(ok))
          }>
            🔄 Retry Connection
          </button>
        </div>
      </div>
    );
  }

  if (bridgeAvailable === null) {
    return <div className="panel"><p className="empty-hint">Connecting to file bridge…</p></div>;
  }

  return (
    <div className="panel file-manager-panel">
      {/* Model selector */}
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      {/* File browser */}
      <div className="file-browser">
        <div className="browser-toolbar">
          <button className="small-btn" onClick={navigateUp} title="Go up">⬆️</button>
          <span className="current-path" title={currentDir}>
            {currentDir.split(/[\\/]/).pop() || currentDir}
          </span>
          <button className="small-btn" onClick={() => navigateTo(currentDir)} title="Refresh">🔄</button>
        </div>

        {loadingDir ? (
          <p className="empty-hint">Loading…</p>
        ) : (
          <div className="file-list">
            {listing?.entries.length === 0 && <p className="empty-hint">Empty folder</p>}
            {listing?.entries.map((entry) => (
              <div
                key={entry.path}
                className={`file-item${selectedFile?.path === entry.path ? " selected" : ""}${
                  !SUPPORTED_EXTS.has(entry.ext || "") && entry.type === "file" ? " unsupported" : ""
                }`}
                onClick={() => loadFile(entry)}
                title={entry.path}
              >
                <span className="file-icon">
                  {entry.type === "dir" ? "📁" :
                   entry.ext === ".docx" ? "📝" :
                   entry.ext === ".xlsx" || entry.ext === ".xls" ? "📊" :
                   entry.ext === ".txt" || entry.ext === ".md" ? "📄" :
                   "📃"}
                </span>
                <span className="file-name">{entry.name}</span>
                {entry.size !== null && (
                  <span className="file-size">{fmtSize(entry.size)}</span>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* File content + AI panel */}
      {selectedFile && (
        <div className="file-content-panel">
          <div className="field-header">
            <label>{selectedFile.name}</label>
            <span>
              <button className="small-btn" onClick={() => setEditMode(!editMode)}>
                {editMode ? "👁 View" : "✏️ Edit"}
              </button>
            </span>
          </div>

          {editMode ? (
            <>
              <textarea
                value={editContent}
                onChange={(e) => setEditContent(e.target.value)}
                rows={7}
                style={{ fontFamily: "monospace", fontSize: "11px" }}
              />
              <button className="action-btn" onClick={saveFile}>💾 Save Changes</button>
            </>
          ) : (
            <div className="result-box" style={{ fontSize: "11px", maxHeight: "120px" }}>
              {loadingFile ? "Loading…" : fileContent || "(empty)"}
            </div>
          )}

          {/* AI actions */}
          {!editMode && (
            <div className="ai-actions">
              <div className="field-group">
                <label>AI Action</label>
                <select value={action} onChange={(e) => setAction(e.target.value as AIAction)}>
                  <option value="summarize">📄 Summarize</option>
                  <option value="rewrite">✨ Rewrite</option>
                  <option value="extract-key-points">🔑 Extract Key Points</option>
                  <option value="translate">🌐 Translate to English</option>
                  <option value="qa">❓ Ask a Question</option>
                </select>
              </div>

              {action === "qa" && (
                <div className="field-group">
                  <input
                    type="text"
                    value={qaQuestion}
                    onChange={(e) => setQaQuestion(e.target.value)}
                    placeholder="What is the main topic of this document?"
                    style={{ width: "100%", padding: "5px 8px", fontSize: "12px",
                             border: "1px solid #d1d5db", borderRadius: "4px" }}
                    onKeyDown={(e) => e.key === "Enter" && runAI()}
                  />
                </div>
              )}

              <button className="action-btn" onClick={runAI}
                disabled={processing || !fileContent || !selectedModel || (action === "qa" && !qaQuestion)}>
                {processing ? "⏳ Processing…" : "🤖 Run AI"}
              </button>
            </div>
          )}
        </div>
      )}

      {/* Result */}
      {result && (
        <div className="field-group result-group">
          <div className="field-header">
            <label>Result</label>
            <span>
              <button className="small-btn" onClick={insertResultIntoDoc} title="Insert into Word document">
                📝 Insert
              </button>
              <button className="small-btn" onClick={saveResultAsFile} title="Save as new file">
                💾 Save
              </button>
              <button className="small-btn" onClick={() =>
                navigator.clipboard.writeText(result).then(() => setStatus("📋 Copied"))
              }>📋 Copy</button>
            </span>
          </div>
          <div className="result-box">{result}</div>
        </div>
      )}

      {status && <p className="status-msg">{status}</p>}
    </div>
  );
}
