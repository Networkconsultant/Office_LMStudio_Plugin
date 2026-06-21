import * as React from "react";
import { LMStudioClient, LMStudioModel } from "../../api/lmstudio";

interface Props {
  baseUrl: string;
}

const SYSTEM_PROMPT =
  "You are an Excel formula expert. When given a natural language description, " +
  "respond with ONLY the Excel formula on the first line (starting with =), " +
  "followed by a blank line, then a concise plain-English explanation of how it works.";

export default function FormulaHelper({ baseUrl }: Props) {
  const [models, setModels] = React.useState<LMStudioModel[]>([]);
  const [selectedModel, setSelectedModel] = React.useState("");
  const [nlInput, setNlInput] = React.useState("");
  const [formula, setFormula] = React.useState("");
  const [explanation, setExplanation] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("");

  const client = React.useMemo(() => new LMStudioClient(baseUrl), [baseUrl]);

  React.useEffect(() => {
    client.getModels().then((m) => {
      setModels(m);
      if (m.length > 0) setSelectedModel(m[0].id);
    }).catch(() => {});
  }, [client]);

  const handleGenerate = async () => {
    if (!nlInput.trim() || loading) return;
    setLoading(true);
    setFormula("");
    setExplanation("");
    setStatus("");
    try {
      const response = await client.chat({
        model: selectedModel,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: nlInput },
        ],
      });
      const lines = response.trim().split("\n");
      const formulaLine = lines.find((l) => l.trim().startsWith("=")) ?? lines[0];
      const rest = lines.filter((l) => l.trim() !== formulaLine.trim()).join("\n").trim();
      setFormula(formulaLine.trim());
      setExplanation(rest);
    } catch (e: any) {
      setStatus(`Error: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleInsert = async () => {
    if (!formula) return;
    try {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("address");
        await ctx.sync();
        range.formulas = [[formula]];
        await ctx.sync();
        setStatus(`✅ Formula inserted into ${range.address}`);
      });
    } catch (e: any) {
      setStatus(`Insert error: ${e.message}`);
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(formula).then(() => setStatus("📋 Copied to clipboard"));
  };

  const examples = [
    'Sum column B where column A equals "Sales"',
    "Average of C2:C100 ignoring blank cells",
    "Find the last non-empty row in column D",
    "VLOOKUP: find product price from a table on Sheet2",
    "Count unique values in the range A2:A500",
  ];

  return (
    <div className="panel formula-panel">
      <div className="model-row">
        <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)}>
          {models.map((m) => <option key={m.id} value={m.id}>{m.id}</option>)}
        </select>
      </div>

      <div className="field-group">
        <label>Describe what you want the formula to do</label>
        <textarea
          value={nlInput}
          onChange={(e) => setNlInput(e.target.value)}
          placeholder='e.g. "Sum values in B2:B100 only if column A equals Sales"'
          rows={3}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleGenerate(); }
          }}
        />
        <div className="examples">
          {examples.map((ex) => (
            <button key={ex} className="example-chip" onClick={() => setNlInput(ex)}>
              {ex}
            </button>
          ))}
        </div>
      </div>

      <button className="action-btn" onClick={handleGenerate} disabled={loading || !nlInput.trim()}>
        {loading ? "⏳ Generating…" : "⚡ Generate Formula"}
      </button>

      {formula && (
        <div className="formula-result">
          <div className="formula-box">
            <code>{formula}</code>
            <span>
              <button className="small-btn" onClick={handleInsert} title="Insert into selected cell">📥 Insert</button>
              <button className="small-btn" onClick={handleCopy} title="Copy to clipboard">📋 Copy</button>
            </span>
          </div>
          {explanation && <p className="explanation">{explanation}</p>}
        </div>
      )}

      {status && <p className="status-msg">{status}</p>}
    </div>
  );
}
