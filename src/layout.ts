import type { PptxIR, PptxShape, XlsxSheet } from "./types";

// ── PPTX: scale a slide to fit a target width while preserving aspect ratio ──

export interface SlideScale {
  scale: number;
  rendered_width_px: number;
  rendered_height_px: number;
}

export function computeSlideScale(ir: PptxIR, targetWidthPx: number): SlideScale {
  const native = ir.slide_size.width_px || 960;
  const aspect = (ir.slide_size.height_px || 540) / native;
  const scale = targetWidthPx / native;
  return {
    scale,
    rendered_width_px: targetWidthPx,
    rendered_height_px: targetWidthPx * aspect,
  };
}

// Normalise PowerPoint/OOXML alignment codes ("l", "ctr", "r", "just") to
// CSS / React Native textAlign values. Also accepts already-normalised values
// so it's safe to call on any string from the IR.
export type TextAlign = "left" | "right" | "center" | "justify";
export function normaliseAlign(a?: string | null): TextAlign {
  switch (a) {
    case "ctr":
    case "center":
      return "center";
    case "r":
    case "right":
      return "right";
    case "just":
    case "justify":
      return "justify";
    default:
      return "left";
  }
}

// Merge adjacent text runs that share identical visual styling into a single
// run with concatenated text. The DOCX / PPTX IR routinely emits dozens of
// short runs per paragraph (one per OOXML <w:r> element) even when nothing
// visually distinguishes them — Word stores spacing / hidden language tags as
// separate runs. Each becomes its own native <Text> on RN, and the per-run
// font-metric measurement is what dominates the 6+ second freeze on long
// documents. Merging keeps the rendered output identical (since the styles
// are equal) while collapsing the component count by 3-10×.
import type { Run } from "./types";

const STYLE_KEYS: (keyof Run)[] = [
  "bold", "italic", "underline", "strike", "font", "size_px",
  "color", "highlight", "superscript", "subscript",
];

function runsSameStyle(a: Run, b: Run): boolean {
  for (const k of STYLE_KEYS) {
    if (a[k] !== b[k]) return false;
  }
  // Tabs and breaks must remain on their own run — they affect layout, not
  // styling, but merging them in would change visible output.
  if (a.tab || b.tab || a.break || b.break) return false;
  return true;
}

export function flattenRuns(runs: Run[] | undefined): Run[] {
  if (!runs || runs.length <= 1) return runs ?? [];
  const out: Run[] = [];
  for (const r of runs) {
    const last = out[out.length - 1];
    if (last && runsSameStyle(last, r)) {
      last.text = (last.text ?? "") + (r.text ?? "");
    } else {
      // Shallow clone so we can append to .text without mutating the IR.
      out.push({ ...r });
    }
  }
  return out;
}

// Resolve OOXML <w:vMerge> markers into a (rowSpan, skip) decision per
// cell. Cells with `v_merge: "restart"` get their row_span calculated by
// counting how many subsequent rows have a `v_merge: "continue"` cell at
// the SAME grid column. Continue cells are flagged `skip: true` so the
// renderer can drop their borders / text. Cells without v_merge keep
// `row_span: 1, skip: false`.
//
// Grid columns are tracked via col_span (a cell with col_span=3 advances
// the column cursor by 3) — matches the renderer's own column accounting.
import type { DocxIR } from "./types";

export type TableCell = NonNullable<DocxIR["paragraphs"][number]["rows"]>[number][number];

export interface ResolvedCell {
  rowSpan: number; // 1 if no merge, >1 if this cell's region spans down.
  skip: boolean;   // true if this cell is a continuation of a region above.
}

export function resolveVerticalMerges(
  rows: NonNullable<DocxIR["paragraphs"][number]["rows"]>,
): ResolvedCell[][] {
  const out: ResolvedCell[][] = rows.map((r) => r.map(() => ({ rowSpan: 1, skip: false })));

  // For each row, walk cells while tracking the grid column. When we see
  // a "restart", scan downward for matching "continue" cells; when we see
  // a "continue", we'll be flagged as skip when its parent restart found
  // us, so do nothing here.
  rows.forEach((row, ri) => {
    let col = 0;
    row.forEach((cell, ci) => {
      const span = cell.col_span ?? 1;
      if (cell.v_merge === "restart") {
        // Count downward.
        let extra = 0;
        for (let rj = ri + 1; rj < rows.length; rj++) {
          // Find the cell at grid column `col` in row rj.
          let cj = 0, found = -1;
          for (let k = 0; k < rows[rj].length; k++) {
            const s = rows[rj][k].col_span ?? 1;
            if (cj === col) { found = k; break; }
            cj += s;
            if (cj > col) break;
          }
          if (found < 0) break;
          if (rows[rj][found].v_merge === "continue") {
            extra += 1;
            out[rj][found].skip = true;
          } else {
            break;
          }
        }
        out[ri][ci].rowSpan = 1 + extra;
      }
      col += span;
    });
  });

  return out;
}

// Walk a PptxIR and return the set of image_rid keys actually referenced by
// any slide background or shape (including nested grouped shapes). Used to
// prune `ir.images` post-parse — PPT files routinely ship unreferenced media
// (slide masters, hidden layouts) that can easily double the IR's memory
// footprint when each image is inlined as base64.
export function collectUsedPptxImageRids(ir: import("./types").PptxIR): Set<string> {
  const used = new Set<string>();
  const walk = (shape: PptxShape) => {
    if (shape.image_rid) used.add(shape.image_rid);
    if (shape.shapes) shape.shapes.forEach(walk);
  };
  for (const slide of ir.slides) {
    if (slide.background?.image_rid) used.add(slide.background.image_rid);
    slide.shapes.forEach(walk);
  }
  return used;
}

// Apply a scale factor to a shape's px coordinates. Returns a new object — does
// not mutate. Useful for emitting CSS top/left/width/height OR RN style props.
export function scaleShape(shape: PptxShape, scale: number) {
  return {
    left: shape.x_px * scale,
    top: shape.y_px * scale,
    width: shape.width_px * scale,
    height: shape.height_px * scale,
    rotation: shape.rotation ?? 0,
  };
}

// PowerPoint scheme-color aliases that the backend (resolve_scheme_color in
// pptx/reader.ex) maps onto the actual theme keys. Mirror that mapping here
// in case the IR ever emits an unresolved scheme name (defensive fallback).
const SCHEME_ALIASES: Record<string, string> = {
  tx1: "dk1",
  tx2: "dk2",
  bg1: "lt1",
  bg2: "lt2",
  phClr: "accent1",
};

// Resolve a theme color reference (e.g. "accent1") to a hex value, falling back
// to the raw string if it's already a literal color or not in the theme map.
// The backend emits hex strings WITHOUT a leading "#" (e.g. "FFFFFF" or "FF6633");
// this function prepends "#" so the output is a valid CSS / RN color.
//
// Alpha (0..100) is applied via #RRGGBBAA when provided; 100 (or omitted) means
// fully opaque, 0 means fully transparent.
export function resolveThemeColor(
  themeColors: Record<string, string>,
  ref?: string | null,
  alpha?: number,
): string | undefined {
  if (ref == null || ref === "") return undefined;

  // Try direct lookup, then alias map (tx1 → dk1 etc.).
  let v = themeColors[ref];
  if (v == null) {
    const aliased = SCHEME_ALIASES[ref];
    if (aliased) v = themeColors[aliased];
  }
  // Last resort: treat the ref itself as a literal color (e.g. "FFFFFF").
  if (v == null) v = ref;

  if (!v || typeof v !== "string") return undefined;
  if (v === "auto" || v === "none") return undefined;

  let normalised: string | null = null;
  if (v.startsWith("#") || v.startsWith("rgb") || v.startsWith("hsl")) {
    normalised = v;
  } else if (/^[0-9a-fA-F]{6}$/.test(v) || /^[0-9a-fA-F]{8}$/.test(v) || /^[0-9a-fA-F]{3}$/.test(v)) {
    normalised = `#${v}`;
  }
  if (!normalised) return undefined;

  // Apply alpha by appending hex AA — only valid for #RRGGBB form.
  if (alpha != null && alpha >= 0 && alpha < 100 && /^#[0-9a-fA-F]{6}$/.test(normalised)) {
    const aa = Math.round((alpha / 100) * 255).toString(16).padStart(2, "0");
    return `${normalised}${aa}`;
  }
  return normalised;
}

// ── XLSX: convert column letter ("A", "Z", "AA") to zero-based index ────────

export function colLetterToIndex(letter: string): number {
  let n = 0;
  for (let i = 0; i < letter.length; i++) {
    n = n * 26 + (letter.charCodeAt(i) - 64);
  }
  return n - 1;
}

export function indexToColLetter(index: number): string {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Parse an A1-style cell ref ("B12") into (col, row) zero-based indices.
export function parseCellRef(ref: string): { col: number; row: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  return { col: colLetterToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

// Resolve a column's width in pixels with fallback to a sensible default.
export function getColWidthPx(sheet: XlsxSheet, colIndex: number, defaultPx = 80): number {
  const letter = indexToColLetter(colIndex);
  return sheet.col_widths?.[letter] ?? defaultPx;
}

export function getRowHeightPx(sheet: XlsxSheet, rowIndex: number, defaultPx = 22): number {
  return sheet.row_heights?.[String(rowIndex + 1)] ?? defaultPx;
}
