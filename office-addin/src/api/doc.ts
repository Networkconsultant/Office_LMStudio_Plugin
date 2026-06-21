/**
 * doc.ts — Unified document API for Word, Excel, and PowerPoint
 * All functions are host-aware: they detect which Office app is running
 * and call the appropriate Office.js API.
 */

// ─────────────────────────────────────────────────────────────────────────────
// Host detection
// ─────────────────────────────────────────────────────────────────────────────

export type OfficeHost = "Word" | "Excel" | "PowerPoint" | "Outlook" | "Unknown";

export function getHostApp(): OfficeHost {
  if (typeof Office === "undefined" || !Office.context?.host) return "Unknown";
  const h = Office.context.host;
  if (h === Office.HostType.Word) return "Word";
  if (h === Office.HostType.Excel) return "Excel";
  if (h === Office.HostType.PowerPoint) return "PowerPoint";
  if (h === Office.HostType.Outlook) return "Outlook";
  return "Unknown";
}

// ─────────────────────────────────────────────────────────────────────────────
// Shared helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Format multi-line text into formatted Word paragraphs. */
async function insertFormattedParagraphs(
  ctx: Word.RequestContext,
  insertTarget: Word.Body | Word.Range,
  text: string,
  heading?: string
): Promise<void> {
  const body = insertTarget as Word.Body;
  if (heading) {
    const h = body.insertParagraph(heading, Word.InsertLocation.end);
    h.style = "Heading 2";
  }
  const paras = text.split(/\n{2,}/);
  for (const para of paras) {
    const trimmed = para.trim();
    if (!trimmed) continue;
    if (trimmed.match(/^[-*•]\s/m)) {
      for (const line of trimmed.split("\n").filter((l) => l.trim())) {
        const p = body.insertParagraph(line.replace(/^[-*•]\s*/, ""), Word.InsertLocation.end);
        p.style = "List Paragraph";
      }
    } else {
      body.insertParagraph(trimmed, Word.InsertLocation.end);
    }
  }
  await ctx.sync();
}

// ─────────────────────────────────────────────────────────────────────────────
// WORD operations
// ─────────────────────────────────────────────────────────────────────────────

/** Read the currently selected text in Word. */
export async function wordGetSelection(): Promise<string> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text");
    await ctx.sync();
    return sel.text;
  });
}

/** Read the full body text of the active Word document. */
export async function wordGetBody(): Promise<string> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load("text");
    await ctx.sync();
    return body.text;
  });
}

/** Read body text with paragraph structure preserved. */
export async function wordGetBodyStructured(): Promise<{ text: string; paraCount: number; wordCount: number }> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load("text");
    const paras = body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    const text = body.text;
    const wordCount = text.split(/\s+/).filter(Boolean).length;
    return { text, paraCount: paras.items.length, wordCount };
  });
}

/** Replace the current selection with new text. */
export async function wordReplaceSelection(text: string): Promise<void> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.insertText(text, Word.InsertLocation.replace);
    await ctx.sync();
  });
}

/** Insert text at the current cursor / end of selection. */
export async function wordInsertAtCursor(text: string): Promise<void> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.insertText(text, Word.InsertLocation.after);
    await ctx.sync();
  });
}

/** Append formatted text to the end of the document. */
export async function wordInsertAtEnd(text: string, heading?: string): Promise<void> {
  return Word.run(async (ctx) => {
    await insertFormattedParagraphs(ctx, ctx.document.body, text, heading);
  });
}

/** Insert a horizontal rule then append text (for clearly demarcated AI output). */
export async function wordInsertSection(title: string, text: string): Promise<void> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    const divider = body.insertParagraph("─".repeat(40), Word.InsertLocation.end);
    divider.style = "Normal";
    divider.font.color = "#9ca3af";
    const h = body.insertParagraph(title, Word.InsertLocation.end);
    h.style = "Heading 3";
    const paras = text.split(/\n{2,}/);
    for (const para of paras) {
      const trimmed = para.trim();
      if (!trimmed) continue;
      body.insertParagraph(trimmed, Word.InsertLocation.end);
    }
    await ctx.sync();
  });
}

/** Apply bold formatting to the current selection. */
export async function wordBoldSelection(bold: boolean): Promise<void> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.font.bold = bold;
    await ctx.sync();
  });
}

/** Apply heading style to the current selection paragraph. */
export async function wordSetSelectionStyle(style: string): Promise<void> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.style = style;
    await ctx.sync();
  });
}

/** Search and replace text throughout the document. */
export async function wordFindAndReplace(find: string, replace: string): Promise<number> {
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(find, { matchCase: false, matchWholeWord: false });
    results.load("items/text");
    await ctx.sync();
    for (const r of results.items) {
      r.insertText(replace, Word.InsertLocation.replace);
    }
    await ctx.sync();
    return results.items.length;
  });
}

/** Get document properties (title, word count, paragraph count). */
export async function wordGetDocInfo(): Promise<{ title: string; wordCount: number; paraCount: number }> {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    const props = ctx.document.properties;
    body.load("text");
    const paras = body.paragraphs;
    paras.load("items");
    props.load("title");
    await ctx.sync();
    const text = body.text;
    return {
      title: props.title || "(Untitled)",
      wordCount: text.split(/\s+/).filter(Boolean).length,
      paraCount: paras.items.length,
    };
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// EXCEL operations
// ─────────────────────────────────────────────────────────────────────────────

export interface ExcelRangeData {
  address: string;
  values: (string | number | boolean)[][];
  formulas: string[][];
  rowCount: number;
  columnCount: number;
}

/** Read the currently selected range (values + formulas). */
export async function excelGetSelection(): Promise<ExcelRangeData> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load(["address", "values", "formulas", "rowCount", "columnCount"]);
    await ctx.sync();
    return {
      address: range.address,
      values: range.values as (string | number | boolean)[][],
      formulas: range.formulas as string[][],
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
  });
}

/** Read the used range of the active sheet as a CSV-like string. */
export async function excelGetActiveSheetData(maxRows = 200): Promise<{ csv: string; address: string; sheetName: string }> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRange();
    sheet.load("name");
    used.load(["address", "values", "rowCount", "columnCount"]);
    await ctx.sync();
    const rows = (used.values as string[][]).slice(0, maxRows);
    const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    return { csv, address: used.address, sheetName: sheet.name };
  });
}

/** Write values to a range by address (e.g. "A1:C3"). */
export async function excelWriteRange(address: string, values: (string | number | boolean)[][]): Promise<void> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.values = values;
    await ctx.sync();
  });
}

/** Write a formula to the currently selected cell. */
export async function excelWriteFormulaToSelection(formula: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("address");
    await ctx.sync();
    range.formulas = [[formula]];
    await ctx.sync();
    return range.address;
  });
}

/** Write values starting at the selected cell. */
export async function excelWriteValuesToSelection(values: (string | number | boolean)[][]): Promise<string> {
  return Excel.run(async (ctx) => {
    const sel = ctx.workbook.getSelectedRange();
    sel.load(["address", "rowIndex", "columnIndex"]);
    await ctx.sync();
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const startCell = sheet.getCell(sel.rowIndex, sel.columnIndex);
    const writeRange = startCell.getResizedRange(values.length - 1, values[0].length - 1);
    writeRange.load("address");
    await ctx.sync();
    writeRange.values = values;
    await ctx.sync();
    return writeRange.address;
  });
}

/** Parse AI-generated table text (markdown or CSV) into a 2D values array. */
export function parseTableText(text: string): (string | number | boolean)[][] {
  const lines = text.split("\n").map((l) => l.trim()).filter((l) => l && !l.match(/^\|?[-:]+\|/));
  return lines.map((line) => {
    const cols = line.startsWith("|")
      ? line.split("|").filter((_, i, a) => i > 0 && i < a.length - 1).map((c) => c.trim())
      : line.split(",").map((c) => c.replace(/^"|"$/g, "").trim());
    return cols.map((c) => {
      const n = Number(c);
      return isNaN(n) || c === "" ? c : n;
    });
  });
}

/** Get the names of all worksheets. */
export async function excelGetSheetNames(): Promise<string[]> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    return sheets.items.map((s) => s.name);
  });
}

/** Append AI-generated rows to the first empty row below used range. */
export async function excelAppendRows(values: (string | number | boolean)[][]): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRange(true);
    used.load(["rowCount", "rowIndex", "columnIndex"]);
    await ctx.sync();
    const nextRow = used.rowIndex + used.rowCount;
    const range = sheet.getRangeByIndexes(nextRow, used.columnIndex, values.length, values[0].length);
    range.load("address");
    await ctx.sync();
    range.values = values;
    await ctx.sync();
    return range.address;
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// POWERPOINT operations
// ─────────────────────────────────────────────────────────────────────────────

export interface PptSlideInfo {
  index: number;
  id: string;
  textContent: string;
}

/** Read text content from all slides. */
export async function pptGetAllText(): Promise<{ slides: PptSlideInfo[]; totalText: string }> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    const slideInfos: PptSlideInfo[] = [];
    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.load("id");
      const shapes = slide.shapes;
      shapes.load("items");
      await ctx.sync();
      const texts: string[] = [];
      for (const shape of shapes.items) {
        try {
          const tf = (shape as any).textFrame;
          if (tf) {
            tf.load("text");
            await ctx.sync();
            if (tf.text) texts.push(tf.text);
          }
        } catch {}
      }
      slideInfos.push({ index: i + 1, id: slide.id, textContent: texts.join(" | ") });
    }
    return {
      slides: slideInfos,
      totalText: slideInfos.map((s) => `[Slide ${s.index}] ${s.textContent}`).join("\n"),
    };
  });
}

/** Get the number of slides in the presentation. */
export async function pptGetSlideCount(): Promise<number> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    return slides.items.length;
  });
}

/** Add a new slide with a title and body text. */
export async function pptAddSlide(title: string, body: string): Promise<void> {
  return PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.load("items");
    await ctx.sync();
    ctx.presentation.slides.add();
    await ctx.sync();
    // Get the last (newly added) slide
    ctx.presentation.slides.load("items");
    await ctx.sync();
    const slides = ctx.presentation.slides.items;
    const newSlide = slides[slides.length - 1];
    newSlide.shapes.load("items");
    await ctx.sync();
    // Try to set title/content in the first two shapes
    const shapes = newSlide.shapes.items;
    if (shapes.length > 0) {
      try {
        const tf0 = (shapes[0] as any).textFrame;
        if (tf0) { tf0.text = title; }
      } catch {}
    }
    if (shapes.length > 1) {
      try {
        const tf1 = (shapes[1] as any).textFrame;
        if (tf1) { tf1.text = body; }
      } catch {}
    }
    await ctx.sync();
  });
}

/** Insert text into the Office.js selection (cross-host fallback). */
export async function insertTextViaOfficeJs(text: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message ?? "Insert failed"));
      }
    );
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Generic read (used by Summarizer, Chat context)
// ─────────────────────────────────────────────────────────────────────────────

/** Read the active document's text content from any supported host. */
export async function readDocumentText(maxChars = 14000): Promise<string> {
  const host = getHostApp();
  let text = "";
  if (host === "Word") {
    text = await wordGetBody();
  } else if (host === "Excel") {
    const { csv } = await excelGetActiveSheetData(300);
    text = csv;
  } else if (host === "PowerPoint") {
    const { totalText } = await pptGetAllText();
    text = totalText;
  } else if (host === "Outlook") {
    text = await new Promise<string>((resolve, reject) => {
      const item = Office.context?.mailbox?.item as any;
      if (!item?.body) return reject(new Error("No mail item"));
      item.body.getAsync(Office.CoercionType.Text, (res: any) => {
        res.status === Office.AsyncResultStatus.Succeeded
          ? resolve(res.value)
          : reject(new Error(res.error?.message));
      });
    });
  } else {
    throw new Error("No supported Office host detected.");
  }
  return text.length > maxChars ? text.slice(0, maxChars) + `\n\n[… truncated at ${maxChars.toLocaleString()} chars]` : text;
}

/** Read the current selection from any supported host. */
export async function readSelectionText(): Promise<string> {
  const host = getHostApp();
  if (host === "Word") return wordGetSelection();
  if (host === "Excel") {
    const data = await excelGetSelection();
    return data.values.map((r) => r.join("\t")).join("\n");
  }
  // Fallback: Office.js generic selection
  return new Promise<string>((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value));
        else reject(new Error(res.error?.message ?? "Could not read selection"));
      }
    );
  });
}
