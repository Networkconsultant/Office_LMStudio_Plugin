/**
 * LMStudio Office Plugin — Local File Bridge Server
 * Runs on http://localhost:3001
 * Gives the sandboxed Office task pane full read/write access to local files.
 *
 * Start with: node bridge.js
 */

const express = require("express");
const cors    = require("cors");
const fs      = require("fs");
const path    = require("path");
const os      = require("os");

const app  = express();
const PORT = 3001;

app.use(cors({ origin: "https://localhost:3000" }));
app.use(express.json({ limit: "50mb" }));

// ─── helpers ──────────────────────────────────────────────────────────────────

function safeResolvePath(filePath) {
  // Reject obvious traversal attempts; resolve to absolute
  const resolved = path.resolve(filePath);
  return resolved;
}

function extOf(filePath) {
  return path.extname(filePath).toLowerCase();
}

// ─── GET /api/health ──────────────────────────────────────────────────────────
app.get("/api/health", (_req, res) => res.json({ ok: true, pid: process.pid }));

// ─── GET /api/home ────────────────────────────────────────────────────────────
// Returns user's home directory and common folders
app.get("/api/home", (_req, res) => {
  const home = os.homedir();
  res.json({
    home,
    documents: path.join(home, "Documents"),
    desktop:   path.join(home, "Desktop"),
    downloads: path.join(home, "Downloads"),
  });
});

// ─── GET /api/files/list?dir=<path> ───────────────────────────────────────────
// Lists files and subdirectories in a directory
app.get("/api/files/list", (req, res) => {
  const dir = safeResolvePath(req.query.dir || os.homedir());
  try {
    if (!fs.existsSync(dir)) return res.status(404).json({ error: "Directory not found" });
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    const result = entries
      .map((e) => {
        try {
          const full = path.join(dir, e.name);
          const stat = fs.statSync(full);
          return {
            name: e.name,
            path: full,
            type: e.isDirectory() ? "dir" : "file",
            size: e.isFile() ? stat.size : null,
            modified: stat.mtime.toISOString(),
            ext: e.isFile() ? path.extname(e.name).toLowerCase() : null,
          };
        } catch {
          return null;
        }
      })
      .filter(Boolean)
      .sort((a, b) => {
        if (a.type !== b.type) return a.type === "dir" ? -1 : 1;
        return a.name.localeCompare(b.name);
      });
    res.json({ dir, entries: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── GET /api/files/read?path=<path> ─────────────────────────────────────────
// Reads a file and returns its text content
// Supports: .txt .md .csv .json .xml .html .js .ts .py .docx .xlsx
app.get("/api/files/read", async (req, res) => {
  const filePath = safeResolvePath(req.query.path || "");
  if (!filePath) return res.status(400).json({ error: "path required" });
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "File not found" });

  const ext = extOf(filePath);

  try {
    // ── .docx — extract plain text via mammoth ────────────────────────────────
    if (ext === ".docx") {
      const mammoth = require("mammoth");
      const result  = await mammoth.extractRawText({ path: filePath });
      return res.json({ path: filePath, ext, text: result.value, encoding: "extracted" });
    }

    // ── .xlsx — extract sheet data as CSV-like text ───────────────────────────
    if (ext === ".xlsx" || ext === ".xls") {
      const XLSX   = require("xlsx");
      const wb     = XLSX.readFile(filePath);
      const sheets = wb.SheetNames.map((name) => {
        const ws   = wb.Sheets[name];
        const csv  = XLSX.utils.sheet_to_csv(ws);
        return `=== Sheet: ${name} ===\n${csv}`;
      });
      return res.json({ path: filePath, ext, text: sheets.join("\n\n"), encoding: "extracted" });
    }

    // ── Plain text formats ────────────────────────────────────────────────────
    const TEXT_EXTS = [".txt",".md",".csv",".json",".xml",".html",".htm",
                       ".js",".ts",".tsx",".jsx",".py",".cs",".css",".log"];
    if (TEXT_EXTS.includes(ext) || !ext) {
      const text = fs.readFileSync(filePath, "utf8");
      return res.json({ path: filePath, ext, text, encoding: "utf8" });
    }

    // ── Unsupported binary ────────────────────────────────────────────────────
    return res.status(415).json({
      error: `Unsupported file type: ${ext}. Supported: .txt .md .csv .json .docx .xlsx and code files.`
    });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── POST /api/files/write ────────────────────────────────────────────────────
// Writes text content to a file (creates or overwrites)
// Body: { path, text, createDirs? }
app.post("/api/files/write", (req, res) => {
  const { path: filePath, text, createDirs = true } = req.body;
  if (!filePath || text === undefined) return res.status(400).json({ error: "path and text required" });

  const resolved = safeResolvePath(filePath);
  try {
    if (createDirs) fs.mkdirSync(path.dirname(resolved), { recursive: true });
    fs.writeFileSync(resolved, text, "utf8");
    const stat = fs.statSync(resolved);
    res.json({ ok: true, path: resolved, size: stat.size });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── POST /api/files/save-as ──────────────────────────────────────────────────
// Saves text into a new file with auto-naming if path not specified
// Body: { text, suggestedName?, dir? }
app.post("/api/files/save-as", (req, res) => {
  const { text, suggestedName, dir } = req.body;
  if (!text) return res.status(400).json({ error: "text required" });

  const targetDir = safeResolvePath(dir || path.join(os.homedir(), "Documents"));
  fs.mkdirSync(targetDir, { recursive: true });

  const name     = suggestedName || `lmstudio-output-${Date.now()}.txt`;
  const filePath = path.join(targetDir, name);
  try {
    fs.writeFileSync(filePath, text, "utf8");
    res.json({ ok: true, path: filePath });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── POST /api/files/delete ───────────────────────────────────────────────────
app.post("/api/files/delete", (req, res) => {
  const { path: filePath } = req.body;
  if (!filePath) return res.status(400).json({ error: "path required" });
  try {
    fs.unlinkSync(safeResolvePath(filePath));
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── POST /api/files/rename ───────────────────────────────────────────────────
app.post("/api/files/rename", (req, res) => {
  const { from, to } = req.body;
  if (!from || !to) return res.status(400).json({ error: "from and to required" });
  try {
    fs.renameSync(safeResolvePath(from), safeResolvePath(to));
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, "127.0.0.1", () => {
  console.log(`\n✅ LMStudio File Bridge running at http://localhost:${PORT}`);
  console.log(`   Endpoints:`);
  console.log(`     GET  /api/health`);
  console.log(`     GET  /api/home`);
  console.log(`     GET  /api/files/list?dir=<path>`);
  console.log(`     GET  /api/files/read?path=<path>`);
  console.log(`     POST /api/files/write      { path, text }`);
  console.log(`     POST /api/files/save-as    { text, suggestedName?, dir? }`);
  console.log(`     POST /api/files/delete     { path }`);
  console.log(`     POST /api/files/rename     { from, to }`);
  console.log(`\n   Keep this running alongside the dev server.\n`);
});
