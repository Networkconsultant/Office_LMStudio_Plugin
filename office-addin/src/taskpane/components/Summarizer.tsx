import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

type SummaryLength = "short" | "medium" | "detailed";
const LENGTH_INSTRUCTIONS: Record<SummaryLength, string> = {
  short: "in 2-3 sentences",
  medium: "in one concise paragraph",
  detailed: "in detail using a bullet-point list of key points",
};

async function getDocumentText(): Promise<string> {
  /* Word */
  try {
    return await new Promise<string>((resolve, reject) =>
      Word.run(async (ctx) => {
        const body = ctx.document.body;
        body.load("text");
        await ctx.sync();
        resolve(body.text);
      }).catch(reject)
    );
  } catch {}

  /* Outlook */
  try {
    const item = Office.context?.mailbox?.item as any;
    if (item?.body) {
      return await new Promise<string>((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, (result: any) => {
          result.status === Office.AsyncResultStatus.Succeeded
            ? resolve(result.value)
            : reject(new Error(result.error?.message));
        });
      });
    }
  } catch {}

  /* PowerPoint — collect text from all shapes on all slides */
  try {
    return await new Promise<string>((resolve, reject) =>
      PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const texts: string[] = [];
        for (const slide of slides.items) {
          slide.shapes.load("items");
          await ctx.sync();
          for (const shape of slide.shapes.items) {
            if ((shape as any).textFrame) {
              (shape as any).textFrame.load("text");
              await ctx.sync();
              texts.push((shape as any).textFrame.text);
            }
          }
        }
        resolve(texts.filter(Boolean).join("\n\n"));
      }).catch(reject)
    );
  } catch {}

  throw new Error("Unable to read document content from this Office application.");
}

const MAX_CHARS = 12000; // stay within typical context windows

export default function Summarizer({ baseUrl }: Props) {
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [summary, setSummary] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [length, setLength] = React.useState<SummaryLength>("medium");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, [client]);

  const handleSummarize = async () => {
    setLoading(true);
    setSummary("");
    setStatus("📄 Reading document…");
    try {
      const raw = await getDocumentText();
      if (!raw.trim()) {
        setStatus("The document appears to be empty.");
        return;
      }
      const truncated = raw.length > MAX_CHARS ? raw.slice(0, MAX_CHARS) + "…" : raw;
      setStatus("🤖 Summarizing…");
      const prompt = `Summarize the following document ${LENGTH_INSTRUCTIONS[length]}:\n\n${truncated}`;
      const response = await client.chat({
        model: selectedModel,
        messages: [{ role: "user", content: prompt }],
      });
      setSummary(response);
      setStatus("");
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(summary).then(() => setStatus("📋 Copied to clipboard"));
  };

  return (
    <div className="panel summarizer-panel">
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      <div className="field-group">
        <label>Summary Length</label>
        <select value={length} onChange={(e) => setLength(e.target.value as SummaryLength)}>
          <option value="short">Short (2-3 sentences)</option>
          <option value="medium">Medium (1 paragraph)</option>
          <option value="detailed">Detailed (bullet points)</option>
        </select>
      </div>

      <button className="action-btn" onClick={handleSummarize} disabled={loading}>
        {loading ? "⏳ Working…" : "📄 Summarize Document"}
      </button>

      {status && <p className="status-msg">{status}</p>}

      {summary && (
        <div className="field-group result-group">
          <div className="field-header">
            <label>Summary</label>
            <button className="small-btn" onClick={handleCopy}>📋 Copy</button>
          </div>
          <div className="result-box">{summary}</div>
        </div>
      )}
    </div>
  );
}
