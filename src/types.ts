// JSON IR shapes returned by GET /api/office/view?path=...
// Backend generator: lib/ex_office/json_ir.ex
// Keep this file in lockstep with the backend output — any change here must
// be reflected on the backend and vice-versa.

// Renderer-side constants used by both the web and mobile autofit
// passes. Keeping them here ensures the two platforms can't drift
// — if 0.5 ever stops being readable enough, this is the only
// place to bump.
export const AUTOFIT_FLOOR = 0.5;
// CSS `line-height` multiplier matching PPT's "single line spacing"
// (typeface ascent + descent ≈ 1.2× font-size for Latin fonts). CJK
// titles in Sarasa Mono and similar may need 1.3× — flag here when
// switching, don't hardcode in renderers.
export const PPT_LINE_HEIGHT_FACTOR = 1.2;

export interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string | boolean | null;
  strike?: boolean;
  font?: string;
  size_px?: number;
  color?: string;
  highlight?: string;
  superscript?: boolean;
  subscript?: boolean;
  tab?: boolean;
  break?: string | null;
  // OOXML `<a:rPr spc>` / `<a:defRPr spc>` letter spacing, in CSS px
  // @ 96 DPI. Pre-computed on the backend so renderers can drop it
  // straight into `letterSpacing` without re-doing pt→px math.
  letter_spacing_px?: number;
}

// ── DOCX ────────────────────────────────────────────────────────────────────

export interface DocxParaImage {
  r_id?: string;
  name?: string;
  // "inline" | "anchor" — anchor images have a position/wrapping mode; the
  // mobile renderer treats both as block-level for now.
  type?: string;
  width_px?: number;
  height_px?: number;
  // Already a `data:image/...;base64,...` URI when the backend resolved it,
  // null when the image couldn't be loaded from the source archive.
  data_url?: string | null;
  alt?: string;
}

export interface DocxParagraph {
  type: string;
  style_id?: string;
  level?: number;
  alignment?: string;
  spacing?: { before_px: number; after_px: number; line: number };
  indent?: { left_px: number; right_px: number; first_line_px: number };
  runs?: Run[];
  rows?: Array<Array<{
    text: string;
    runs?: Run[];
    col_span?: number;
    row_span?: number;
    width_px?: number | null;
    alignment?: string | null;
    vertical_alignment?: string | null;
    // Vertical merge marker from <w:vMerge>. "restart" = top cell of a
    // merged region; "continue" = cell that should be skipped (the region
    // from above covers it). Frontend resolves into a proper row_span by
    // counting continues in subsequent rows at the same grid column.
    v_merge?: "restart" | "continue" | null;
  }>>;
  // Per-column widths from <w:gridCol>, in CSS pixels. Sum should equal the
  // table's total width. Present on `type: "table"` paragraphs only.
  grid_cols_px?: number[];
  // Inline / anchored images attached to this paragraph. Each entry already
  // carries its own data_url so the renderer doesn't need to cross-reference
  // the top-level `ir.images` map for DOCX.
  images?: DocxParaImage[];
}

export interface DocxImage {
  data_url: string;
  width_px: number;
  height_px: number;
}

export interface DocxPageInfo {
  width_px: number;
  height_px: number;
  margin_top_px: number;
  margin_bottom_px: number;
  margin_left_px: number;
  margin_right_px: number;
  orientation: string;
}

// One logical Word section. A docx can mix page sizes (e.g., A4 portrait
// for the cover, A3 landscape for a wide report table); each section
// carries its own page setup. Backend splits paragraphs at sectPr / body-
// level section breaks. Renderers should iterate sections when present
// and fall back to a single section assembled from top-level page +
// paragraphs.
export interface DocxSection {
  page: DocxPageInfo;
  paragraphs: DocxParagraph[];
}

export interface DocxIR {
  type: "docx";
  title: string;
  page: DocxPageInfo;
  // Optional: backend emits this when the doc has explicit section
  // breaks with their own page setup. Older docs / fallback use the
  // top-level page + paragraphs.
  sections?: DocxSection[];
  default_font: { family: string; size_px: number };
  paragraphs: DocxParagraph[];
  images: Record<string, DocxImage>;
}

// ── PPTX ────────────────────────────────────────────────────────────────────

export interface PptxFill {
  type: string;
  color?: string;
  // Backend may emit alpha as 0..100 (0 = transparent, 100 = opaque).
  // Omitted entirely when fully opaque.
  alpha?: number;
}

export interface PptxShape {
  id?: number;
  type: string;
  x_px: number;
  y_px: number;
  width_px: number;
  height_px: number;
  rotation?: number;
  fill?: PptxFill;
  line?: { color?: string; width_px?: number };
  text_body?: {
    paragraphs: Array<{
      alignment?: string;
      runs: Run[];
      // Bullet spec inherited from layout / master; nil/absent means
      // "no bullet" (the backend handles `<a:buNone/>` overrides
      // already). Renderer just checks truthy and prepends "•".
      bullet?: boolean | { type?: string; char?: string };
      level?: number;
      // Paragraph spacing — hundredths of a point per OOXML.
      space_before_pts?: number;
      space_after_pts?: number;
      // Line spacing percent — 1000ths (140000 = 140%). Falls back
      // to PPT default (~120%) when absent.
      line_spacing_pct?: number;
    }>;
    // `<a:bodyPr>` autofit mode: "norm" / "sp" / "none".
    autofit?: string;
    // `<a:normAutofit fontScale="...">` author-applied shrink, as a
    // ratio (1.0 = no shrink). Backend converts the 1000ths-of-%
    // OOXML value before exposing.
    font_scale?: number;
  };
  image_rid?: string;
  shapes?: PptxShape[];
}

export interface PptxBackground {
  type: string; // "solid" | "gradient" | "image"
  color?: string;
  alpha?: number;
  // Set for picture-fill backgrounds. Matches the composite key used in
  // PptxIR.images (ex_office/json_ir.ex builds "{slide_idx}.{rid}").
  image_rid?: string;
}

export interface PptxSlide {
  index: number;
  type: string;
  background?: PptxBackground;
  shapes: PptxShape[];
  notes?: string;
  title?: string;
  bullets?: string[];
  subtitle?: string;
}

export interface PptxImage {
  data_url: string;
  width_px: number;
  height_px: number;
}

export interface PptxIR {
  type: "pptx";
  title: string;
  slide_size: { width_px: number; height_px: number };
  theme: { colors: Record<string, string>; font: string };
  slides: PptxSlide[];
  images: Record<string, PptxImage>;
  /**
   * TTFs embedded in the source pptx (`ppt/fonts/*.fntdata`). Each
   * entry pairs a font family name with base64-encoded TTF bytes.
   * The WebGL viewer decodes and registers these with the renderer
   * on load — same bytes pretext_rs and PowerPoint use, so layout /
   * autofit is byte-identical across all three.
   */
  embedded_fonts?: Array<{ name: string; bytes_base64: string }>;
}

// ── XLSX ────────────────────────────────────────────────────────────────────

export interface XlsxCellStyle {
  font_family?: string;
  font_size_px?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  bg_color?: string;
  align_h?: string;
  align_v?: string;
  wrap?: boolean;
  number_format?: string;
  borders?: {
    top?: string | null;
    right?: string | null;
    bottom?: string | null;
    left?: string | null;
  };
}

export interface XlsxCell {
  row: number;
  col: number;
  ref: string;
  value: string | number | boolean | null;
  formula?: string;
  type: string;
  style?: XlsxCellStyle;
}

export interface XlsxSheet {
  name: string;
  max_row: number;
  max_col: number;
  cells: Record<string, XlsxCell>;
  merge_ranges: string[];
  col_widths: Record<string, number>;
  row_heights: Record<string, number>;
  freeze_pane?: string;
  tab_color?: string;
}

export interface XlsxIR {
  type: "xlsx";
  sheet_order: string[];
  active_sheet: string;
  sheets: Record<string, XlsxSheet>;
}

export type OfficeIR = DocxIR | PptxIR | XlsxIR;

export type OfficeViewResponse =
  | { ok: true; ir: OfficeIR; path: string; upgraded: boolean }
  | { ok: false; error: string };
