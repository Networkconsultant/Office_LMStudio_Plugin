import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

type TransformType = "improve" | "rewrite" | "summarize" | "translate" | "formal" | "casual" | "bulletpoints" | "expand";

const TRANSFORMS: Record<TransformType, string> = {
  improve: "Improve the grammar, clarity, and style of the following text (preserve meaning):",
  rewrite: "Rewrite the following text more clearly and concisely:",
  summarize: "Summarize the following text in 2-3 sentences:",
  translate: "Translate the following text to English (detect source language automatically):",
  formal: "Rewrite the following text in a formal, professional tone:",
  casual: "Rewrite the following text in a casual, friendly, conversational tone:",
  bulletpoints: "Convert the following text into a concise bullet-point list:",
  expand: "Expand the following text with more detail and context while staying on topic:",
};

async function readSelectionText(): Promise<string> {
  /* Try Word */
  try {
    return await new Promise<string>((resolve, reject) =>
      Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        sel.text ? resolve(sel.text) : reject(new Error("empty"));
      }).catch(reject)
    );
  } catch {}
  /* Try Excel */
  try {
    return await new Promise<string>((resolve, reject) =>
      Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("text");
        await ctx.sync();
        const flat = range.text.flat().join(" ").trim();
        flat ? resolve(flat) : reject(new Error("empty"));
      }).catch(reject)
    );
  } catch {}
  throw new Error("Could not read selection — paste text manually.");
}

async function replaceSelectionText(text: string): Promise<void> {
  try {
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
    });
  } catch {
    throw new Error("Insert not supported in this app — copy the result manually.");
  }
}

export default function TextTransform({ baseUrl }: Props) {
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [transform, setTransform] = React.useState<TransformType>("improve");
  const [sourceText, setSourceText] = React.useState("");
  const [result, setResult] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, [client]);

  const handleGetSelection = async () => {
    try {
      const text = await readSelectionText();
      setSourceText(text);
      setStatus("");
    } catch (e: any) {
      setStatus(e.message);
    }
  };

  const handleRun = async () => {
    if (!sourceText.trim() || loading) return;
    setLoading(true);
    setResult("");
    setStatus("");
    try {
      const prompt = `${TRANSFORMS[transform]}\n\n${sourceText}`;
      const response = await client.chat({
        model: selectedModel,
        messages: [{ role: "user", content: prompt }],
      });
      setResult(response);
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleInsert = async () => {
    if (!result) return;
    try {
      await replaceSelectionText(result);
      setStatus("✅ Inserted into document");
    } catch (e: any) {
      setStatus(e.message);
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(result).then(() => setStatus("📋 Copied to clipboard"));
  };

  return (
    <div className="panel transform-panel">
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      <div className="field-group">
        <label>Transformation</label>
        <select value={transform} onChange={(e) => setTransform(e.target.value as TransformType)}>
          {(Object.keys(TRANSFORMS) as TransformType[]).map((k) => (
            <option key={k} value={k}>{k.charAt(0).toUpperCase() + k.slice(1)}</option>
          ))}
        </select>
      </div>

      <div className="field-group">
        <div className="field-header">
          <label>Source Text</label>
          <button className="small-btn" onClick={handleGetSelection}>📋 Get Selection</button>
        </div>
        <textarea
          value={sourceText}
          onChange={(e) => setSourceText(e.target.value)}
          placeholder="Select text in your document or paste here…"
          rows={5}
        />
      </div>

      <button className="action-btn" onClick={handleRun} disabled={loading || !sourceText.trim()}>
        {loading ? "⏳ Processing…" : "✨ Transform"}
      </button>

      {result && (
        <div className="field-group result-group">
          <div className="field-header">
            <label>Result</label>
            <span>
              <button className="small-btn" onClick={handleInsert}>📝 Insert</button>
              <button className="small-btn" onClick={handleCopy}>📋 Copy</button>
            </span>
          </div>
          <div className="result-box">{result}</div>
        </div>
      )}

      {status && <p className="status-msg">{status}</p>}
    </div>
  );
}
