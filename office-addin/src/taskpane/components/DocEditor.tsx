import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";
import {
  getHostApp, OfficeHost,
  // Word
  wordGetBody, wordGetBodyStructured, wordGetSelection, wordGetDocInfo,
  wordReplaceSelection, wordInsertAtCursor, wordInsertAtEnd, wordInsertSection,
  wordFindAndReplace, wordBoldSelection, wordSetSelectionStyle,
  // Excel
  excelGetSelection, excelGetActiveSheetData, excelWriteFormulaToSelection,
  excelWriteValuesToSelection, parseTableText, excelGetSheetNames,
  // PPT
  pptGetAllText, pptGetSlideCount, pptAddSlide,
  // Generic
  readDocumentText, readSelectionText, insertTextViaOfficeJs,
} from "../../api/doc";

interface Props { baseUrl: string; }

type Action =
  | "read-doc" | "read-selection"
  | "rewrite-selection" | "improve-selection" | "expand-selection"
  | "insert-at-cursor" | "insert-at-end" | "insert-section"
  | "find-replace"
  | "summarize-doc" | "analyze-doc" | "extract-headings" | "generate-outline"
  | "excel-analyze" | "excel-table" | "excel-formula" | "excel-append"
  | "ppt-summary" | "ppt-add-slide" | "ppt-outline";

interface ActionDef { label: string; description: string; host: OfficeHost | "All"; requiresContent: boolean; }
const ACTIONS: Record<Action, ActionDef> = {
  "read-doc":           { label: "📄 Read Full Document",        description: "Load document text into the editor",                      host: "All",          requiresContent: false },
  "read-selection":     { label: "📋 Read Selection",            description: "Load selected text into the editor",                     host: "All",          requiresContent: false },
  "rewrite-selection":  { label: "✏️ Rewrite Selection",         description: "Rewrite selected text for clarity and flow",             host: "Word",         requiresContent: true  },
  "improve-selection":  { label: "✨ Improve Selection",         description: "Fix grammar, tone, and style of selected text",          host: "Word",         requiresContent: true  },
  "expand-selection":   { label: "📝 Expand Selection",          description: "Add detail and context to selected text",                host: "Word",         requiresContent: true  },
  "insert-at-cursor":   { label: "➕ Insert at Cursor",          description: "Insert AI output at current cursor position",            host: "Word",         requiresContent: true  },
  "insert-at-end":      { label: "⬇️ Append to Document",        description: "Append AI output at the end of the document",           host: "Word",         requiresContent: true  },
  "insert-section":     { label: "📌 Insert as New Section",     description: "Add AI output as a titled section in the document",     host: "Word",         requiresContent: true  },
  "find-replace":       { label: "🔍 Find & Replace",            description: "Find text in the document and replace it",              host: "Word",         requiresContent: false },
  "summarize-doc":      { label: "📄 Summarize Document",        description: "Generate a summary of the whole document",              host: "All",          requiresContent: false },
  "analyze-doc":        { label: "🔬 Analyze Document",          description: "AI analysis: tone, structure, key themes",              host: "All",          requiresContent: false },
  "extract-headings":   { label: "📑 Extract Headings",          description: "List all headings and structure of the document",       host: "Word",         requiresContent: false },
  "generate-outline":   { label: "🗂️ Generate Outline",          description: "Create a structured outline from the document",        host: "All",          requiresContent: false },
  "excel-analyze":      { label: "📊 Analyze Sheet Data",        description: "AI analysis of the active spreadsheet",                 host: "Excel",        requiresContent: false },
  "excel-table":        { label: "📋 Generate Table",            description: "Ask AI to create a table, paste into sheet",            host: "Excel",        requiresContent: true  },
  "excel-formula":      { label: "⚡ Generate Formula",          description: "Describe a formula in plain English, insert into cell", host: "Excel",        requiresContent: true  },
  "excel-append":       { label: "➕ Append AI Rows",            description: "Add AI-generated rows below existing data",             host: "Excel",        requiresContent: true  },
  "ppt-summary":        { label: "📊 Summarize Slides",          description: "Get a summary of all slide content",                    host: "PowerPoint",   requiresContent: false },
  "ppt-add-slide":      { label: "➕ Add New Slide",             description: "AI generates title + content for a new slide",          host: "PowerPoint",   requiresContent: true  },
  "ppt-outline":        { label: "🗂️ Generate Slide Outline",   description: "Create slide-by-slide outline from a topic",            host: "PowerPoint",   requiresContent: true  },
};

// System prompts per action
const getPrompt = (action: Action, content: string, extra?: string): string => {
  const prompts: Partial<Record<Action, string>> = {
    "rewrite-selection":  `Rewrite the following text for clarity and better flow. Return ONLY the rewritten text, no explanations:\n\n${content}`,
    "improve-selection":  `Fix the grammar, punctuation, and style of the following text. Preserve the original meaning. Return ONLY the corrected text:\n\n${content}`,
    "expand-selection":   `Expand the following text with more detail and context. Keep the same tone. Return ONLY the expanded text:\n\n${content}`,
    "insert-at-cursor":   `${extra}\n\nContext from document:\n${content}`,
    "insert-at-end":      `${extra}\n\nContext from document:\n${content}`,
    "insert-section":     `Write a new section titled "${extra}" for a document about the following content:\n\n${content}`,
    "summarize-doc":      `Summarize this document in 3-5 bullet points:\n\n${content}`,
    "analyze-doc":        `Analyze this document. Cover: main purpose, tone, structure, key themes, and suggestions for improvement:\n\n${content}`,
    "extract-headings":   `List every heading and subheading in this document as a numbered outline:\n\n${content}`,
    "generate-outline":   `Generate a structured outline for this document with main sections and sub-points:\n\n${content}`,
    "excel-analyze":      `Analyze this spreadsheet data (CSV format). Describe what it contains, key patterns, and insights:\n\n${content}`,
    "excel-table":        `Create a table in CSV format (first row = headers) for: ${extra}. Return ONLY the CSV, nothing else.`,
    "excel-formula":      `Write an Excel formula for: ${content}. Return ONLY the formula (starting with =) on the first line, then a blank line, then a brief explanation.`,
    "excel-append":       `Generate additional rows in CSV format (no header) matching this existing spreadsheet data:\n\n${content}\n\nInstruction: ${extra}`,
    "ppt-summary":        `Summarize these presentation slides concisely, one sentence per slide:\n\n${content}`,
    "ppt-add-slide":      `Write content for a new slide about: ${extra}. Format:\nTITLE: <title>\nCONTENT: <bullet points>`,
    "ppt-outline":        `Create a complete slide deck outline for the topic: ${extra}. List 6-10 slides with title and 3 bullet points each.`,
  };
  return prompts[action] ?? `${extra || "Process the following"}\n\n${content}`;
};

export default function DocEditor({ baseUrl }: Props) {
  const host = React.useMemo(getHostApp, []);
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [docContent, setDocContent] = React.useState("");
  const [result, setResult] = React.useState("");
  const [action, setAction] = React.useState<Action>("read-doc");
  const [extraInput, setExtraInput] = React.useState("");
  const [findText, setFindText] = React.useState("");
  const [replaceText, setReplaceText] = React.useState("");
  const [docInfo, setDocInfo] = React.useState<{ title: string; wordCount: number; paraCount?: number } | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [insertMode, setInsertMode] = React.useState<"replace" | "cursor" | "end" | "section">("end");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
    // Auto-load doc info
    if (host === "Word") {
      wordGetDocInfo().then(setDocInfo).catch(() => {});
    }
  }, [client]);

  // Filter actions by host
  const availableActions = (Object.entries(ACTIONS) as [Action, ActionDef][])
    .filter(([, def]) => def.host === "All" || def.host === host);

  const currentActionDef = ACTIONS[action];

  // ── Helpers ──────────────────────────────────────────────────────────────

  const needsExtraInput = ["insert-at-cursor", "insert-at-end", "insert-section",
    "excel-table", "excel-formula", "excel-append", "ppt-add-slide", "ppt-outline"].includes(action);

  const needsFindReplace = action === "find-replace";

  // ── Read document content ────────────────────────────────────────────────

  const handleReadDoc = async () => {
    setStatus("📄 Reading document…");
    try {
      if (host === "Word") {
        const info = await wordGetDocInfo();
        setDocInfo(info);
        const body = await wordGetBody();
        setDocContent(body);
        setStatus(`✅ Loaded ${info.wordCount.toLocaleString()} words, ${info.paraCount} paragraphs`);
      } else if (host === "Excel") {
        const { csv, sheetName, address } = await excelGetActiveSheetData(300);
        setDocContent(csv);
        setStatus(`✅ Loaded sheet "${sheetName}" (${address})`);
      } else if (host === "PowerPoint") {
        const { slides, totalText } = await pptGetAllText();
        setDocContent(totalText);
        setStatus(`✅ Loaded ${slides.length} slides`);
      } else {
        const text = await readDocumentText();
        setDocContent(text);
        setStatus("✅ Document loaded");
      }
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    }
  };

  const handleReadSelection = async () => {
    setStatus("📋 Reading selection…");
    try {
      if (host === "Excel") {
        const data = await excelGetSelection();
        const text = `Selection: ${data.address} (${data.rowCount}×${data.columnCount})\n` +
          data.values.map((r) => r.join("\t")).join("\n");
        setDocContent(text);
        setStatus(`✅ Read range ${data.address}`);
      } else {
        const text = await readSelectionText();
        setDocContent(text || "(Nothing selected)");
        setStatus(text ? `✅ Read ${text.length} characters` : "⚠️ Nothing selected");
      }
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    }
  };

  // ── Run AI ────────────────────────────────────────────────────────────────

  const handleRunAI = async () => {
    if (!selectedModel) { setStatus("⚠️ No model selected"); return; }
    setLoading(true);
    setResult("");
    setStatus("🤖 Processing…");

    try {
      // Read document if needed and not already loaded
      let content = docContent;
      const needsDocRead = ["summarize-doc", "analyze-doc", "extract-headings", "generate-outline",
        "excel-analyze", "ppt-summary"].includes(action);
      if (needsDocRead && !content) {
        content = await readDocumentText();
        setDocContent(content);
      }

      const prompt = getPrompt(action, content, extraInput);
      const response = await client.chat({
        model: selectedModel,
        messages: [{ role: "user", content: prompt }],
      });
      setResult(response);
      setStatus("✅ Done");
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  // ── Write to document ─────────────────────────────────────────────────────

  const handleWriteToDoc = async () => {
    if (!result) return;
    setStatus("Writing to document…");
    try {
      if (host === "Word") {
        const mode = insertMode;
        if (mode === "replace") {
          await wordReplaceSelection(result);
          setStatus("✅ Selection replaced");
        } else if (mode === "cursor") {
          await wordInsertAtCursor(result);
          setStatus("✅ Inserted at cursor");
        } else if (mode === "section") {
          const title = extraInput || "AI Output";
          await wordInsertSection(title, result);
          setStatus("✅ New section added");
        } else {
          await wordInsertAtEnd(result, extraInput || undefined);
          setStatus("✅ Appended to document");
        }
      } else if (host === "Excel") {
        if (action === "excel-formula") {
          const addr = await excelWriteFormulaToSelection(result.split("\n")[0].trim());
          setStatus(`✅ Formula inserted into ${addr}`);
        } else if (action === "excel-append") {
          const rows = parseTableText(result);
          if (rows.length > 0) {
            const { excelAppendRows } = await import("../../api/doc");
            const addr = await excelAppendRows(rows);
            setStatus(`✅ Appended ${rows.length} rows at ${addr}`);
          }
        } else {
          const rows = parseTableText(result);
          if (rows.length > 0) {
            const addr = await excelWriteValuesToSelection(rows);
            setStatus(`✅ Written to ${addr}`);
          } else {
            await insertTextViaOfficeJs(result);
            setStatus("✅ Inserted as text");
          }
        }
      } else if (host === "PowerPoint") {
        if (action === "ppt-add-slide") {
          const titleMatch = result.match(/TITLE:\s*(.+)/i);
          const contentMatch = result.match(/CONTENT:\s*([\s\S]+)/i);
          const title = titleMatch?.[1]?.trim() ?? extraInput ?? "New Slide";
          const content = contentMatch?.[1]?.trim() ?? result;
          await pptAddSlide(title, content);
          setStatus(`✅ Slide added: "${title}"`);
        } else {
          await insertTextViaOfficeJs(result);
          setStatus("✅ Inserted");
        }
      } else {
        await insertTextViaOfficeJs(result);
        setStatus("✅ Inserted");
      }
    } catch (e: any) {
      setStatus(`Write error: ${e.message}`);
    }
  };

  // ── Find & Replace ────────────────────────────────────────────────────────

  const handleFindReplace = async () => {
    if (!findText) return;
    try {
      setStatus("Searching…");
      const count = await wordFindAndReplace(findText, replaceText);
      setStatus(`✅ Replaced ${count} occurrence${count !== 1 ? "s" : ""}`);
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    }
  };

  // ── Quick action shortcuts ────────────────────────────────────────────────

  const handleQuickAction = async (quickAction: Action) => {
    setAction(quickAction);
    setResult("");
    setStatus("");
    // Auto-run for read actions
    if (quickAction === "read-doc") { await handleReadDoc(); return; }
    if (quickAction === "read-selection") { await handleReadSelection(); return; }
  };

  // ── Host badge ────────────────────────────────────────────────────────────
  const hostEmoji: Record<OfficeHost, string> = {
    Word: "📝", Excel: "📊", PowerPoint: "📊", Outlook: "📧", Unknown: "🖥️"
  };
  const hostColor: Record<OfficeHost, string> = {
    Word: "#2b5eb8", Excel: "#217346", PowerPoint: "#c43e1c", Outlook: "#0072c6", Unknown: "#6b7280"
  };

  return (
    <div className="panel doc-editor-panel">
      {/* Host indicator + doc info */}
      <div className="doc-host-bar" style={{ background: hostColor[host] }}>
        <span>{hostEmoji[host]} {host}</span>
        {docInfo && (
          <span className="doc-info-badge">
            {docInfo.title !== "(Untitled)" ? docInfo.title : ""} · {docInfo.wordCount.toLocaleString()} words
          </span>
        )}
      </div>

      {/* Model selector */}
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      {/* Quick action bar */}
      <div className="quick-actions">
        <button className="quick-btn" onClick={() => handleQuickAction("read-doc")}>📄 Read Doc</button>
        <button className="quick-btn" onClick={() => handleQuickAction("read-selection")}>📋 Selection</button>
        {host === "Word" && (
          <button className="quick-btn" onClick={async () => {
            const text = await readSelectionText().catch(() => docContent);
            if (!text || !selectedModel) return;
            setLoading(true); setResult(""); setStatus("Rewriting…");
            try {
              const r = await client.chat({ model: selectedModel,
                messages: [{ role: "user", content: `Improve and rewrite for clarity:\n\n${text}` }] });
              setResult(r);
              setStatus("✅ Click 'Write to Doc → Replace Selection' to apply");
            } catch (e: any) { setStatus(`Error: ${e.message}`); }
            finally { setLoading(false); }
          }}>✨ Quick Rewrite</button>
        )}
      </div>

      {/* Action selector */}
      <div className="field-group">
        <label>Action</label>
        <select value={action} onChange={(e) => { setAction(e.target.value as Action); setResult(""); }}>
          {availableActions.map(([id, def]) => (
            <option key={id} value={id}>{def.label}</option>
          ))}
        </select>
        <small style={{ color: "#6b7280", fontSize: "11px", marginTop: 2 }}>
          {currentActionDef?.description}
        </small>
      </div>

      {/* Extra input for certain actions */}
      {needsExtraInput && (
        <div className="field-group">
          <label>
            {action === "insert-section" ? "Section title" :
             action === "ppt-add-slide" ? "Slide topic" :
             action === "excel-formula" ? "Describe the formula" :
             action === "ppt-outline" ? "Presentation topic" :
             action === "excel-table" ? "Describe the table" :
             "Instruction / prompt"}
          </label>
          <input
            type="text"
            value={extraInput}
            onChange={(e) => setExtraInput(e.target.value)}
            placeholder={
              action === "excel-formula" ? 'e.g. "Sum B2:B100 where A equals Sales"' :
              action === "ppt-add-slide" ? "e.g. Q3 Financial Results" :
              "Enter your instruction…"
            }
            style={{ width: "100%", padding: "5px 8px", border: "1px solid #d1d5db",
                     borderRadius: "4px", fontSize: "12px" }}
          />
        </div>
      )}

      {/* Find & Replace */}
      {needsFindReplace && (
        <div className="field-group">
          <label>Find text</label>
          <input type="text" value={findText} onChange={(e) => setFindText(e.target.value)}
            style={{ width: "100%", padding: "5px 8px", border: "1px solid #d1d5db",
                     borderRadius: "4px", fontSize: "12px", marginBottom: 4 }} />
          <label>Replace with</label>
          <input type="text" value={replaceText} onChange={(e) => setReplaceText(e.target.value)}
            style={{ width: "100%", padding: "5px 8px", border: "1px solid #d1d5db",
                     borderRadius: "4px", fontSize: "12px" }} />
          <button className="action-btn" style={{ marginTop: 6 }}
            onClick={handleFindReplace} disabled={!findText}>
            🔍 Find & Replace
          </button>
        </div>
      )}

      {/* Document content preview */}
      {docContent && !needsFindReplace && (
        <div className="field-group">
          <div className="field-header">
            <label>Document content ({docContent.length.toLocaleString()} chars)</label>
            <button className="small-btn" onClick={() => setDocContent("")}>✕ Clear</button>
          </div>
          <div className="doc-preview">{docContent.slice(0, 400)}{docContent.length > 400 ? "…" : ""}</div>
        </div>
      )}

      {/* Run button */}
      {!needsFindReplace && (
        <button className="action-btn" onClick={handleRunAI}
          disabled={loading || !selectedModel ||
            (currentActionDef?.requiresContent && !docContent && !extraInput)}>
          {loading ? "⏳ Processing…" : "🤖 Run AI"}
        </button>
      )}

      {/* Result */}
      {result && (
        <div className="field-group result-group">
          <div className="field-header">
            <label>Result</label>
            <span>
              <button className="small-btn"
                onClick={() => navigator.clipboard.writeText(result).then(() => setStatus("📋 Copied"))}>
                📋
              </button>
            </span>
          </div>
          <div className="result-box">{result}</div>

          {/* Write to doc controls */}
          <div className="write-controls">
            {host === "Word" && !["summarize-doc","analyze-doc","extract-headings","generate-outline",
              "find-replace","excel-analyze","ppt-summary"].includes(action) && (
              <div className="write-mode-row">
                <select value={insertMode}
                  onChange={(e) => setInsertMode(e.target.value as typeof insertMode)}
                  style={{ fontSize: "11px", padding: "3px 6px", border: "1px solid #d1d5db",
                           borderRadius: "4px", flex: 1 }}>
                  <option value="replace">Replace selection</option>
                  <option value="cursor">Insert at cursor</option>
                  <option value="end">Append to end</option>
                  <option value="section">New section</option>
                </select>
                <button className="action-btn write-btn" onClick={handleWriteToDoc}>
                  📝 Write to Doc
                </button>
              </div>
            )}
            {(host === "Excel" || host === "PowerPoint") && (
              <button className="action-btn write-btn" onClick={handleWriteToDoc}>
                {host === "Excel" ? "📥 Write to Sheet" : "➕ Add to Presentation"}
              </button>
            )}
            {host === "Word" && ["summarize-doc","analyze-doc","extract-headings","generate-outline"].includes(action) && (
              <button className="action-btn write-btn" onClick={async () => {
                await wordInsertAtEnd(result, action.replace(/-/g, " ").replace(/\b\w/g, l => l.toUpperCase()));
                setStatus("✅ Appended to document");
              }}>
                ⬇️ Append to Doc
              </button>
            )}
          </div>
        </div>
      )}

      {status && <p className={`status-msg${status.startsWith("Error") ? " error-msg" : ""}`}>{status}</p>}
    </div>
  );
}
