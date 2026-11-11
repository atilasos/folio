// Fonts and sizes for text boxes. pdf.js's text box editor sizes its text but always writes it in Helvetica, so Folio
// keeps the chosen font on each editor, shows it on the page, saves it with the mark, and redraws those boxes in that
// font in saved copies (see drawTextFonts).
import { AnnotationEditorParamsType } from "pdfjs-dist";
import type { PDFDocument, PDFFont } from "pdf-lib";

export const TEXT_FONTS = [
  // pdf.js's own look: its editor shows sans-serif and its saved copies use Helvetica.
  { id: "sans", name: "Sans serif", css: "" },
  { id: "serif", name: "Serif", css: '"Times New Roman", Times, "Liberation Serif", serif' },
  { id: "mono", name: "Monospace", css: '"Courier New", Courier, "Liberation Mono", monospace' },
  { id: "easyjet", name: "easyJet Rounded Book", css: '"easyJet Rounded Book", sans-serif' },
] as const;
export type TextFont = (typeof TEXT_FONTS)[number]["id"];
export const DEFAULT_TEXT_FONT: TextFont = "sans";
/** Any size in this range can be typed, to a hundredth of a point; these are offered as suggestions. */
export const MIN_TEXT_SIZE = 1;
export const MAX_TEXT_SIZE = 999;
export const TEXT_SIZES = [4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72];
export const DEFAULT_TEXT_SIZE = 10;

export const isTextFont = (value: unknown): value is TextFont => TEXT_FONTS.some(f => f.id === value);
export const isTextSize = (value: unknown): value is number =>
  typeof value === "number" && value >= MIN_TEXT_SIZE && value <= MAX_TEXT_SIZE;
/** A typed size as a number of points, or null if it isn't one Folio accepts. Accepts a decimal comma. */
export function parseTextSize(text: string) {
  const size = Math.round(Number(text.trim().replace(",", ".")) * 100) / 100;
  return text.trim() && isTextSize(size) ? size : null;
}
export const fontCss = (font: TextFont | undefined) => TEXT_FONTS.find(f => f.id === font)?.css ?? "";

/* The parts of pdf.js's FreeTextEditor used here (not part of its typed public API). */
export interface TextBoxEditor {
  constructor: { _type?: string };
  editorDiv?: HTMLElement | null;
  div?: HTMLElement | null;
  annotationElementId?: string | null;
  isSelected?: boolean;
  isAttachedToDOM?: boolean;
  deleted?: boolean;
  pageIndex: number;
  rotation: number;
  parentRotation: number;
  parentDimensions: [number, number];
  width: number;
  height: number;
  _initialData?: Record<string, unknown> | null;
  _uiManager?: { updateUI(editor: TextBoxEditor): void };
  propertiesToUpdate: [number, unknown][];
  updateParams(type: number, value: unknown): void;
  addCommands(command: { cmd: () => void; undo: () => void; post?: () => void; mustExec: boolean }): void;
  fixAndSetPosition(): void;
  serialize(copying?: boolean): Record<string, unknown> | null;
  __folioFont?: TextFont;
}
interface TextBoxClass {
  _type: string;
  prototype: {
    render(): HTMLElement;
    serialize(copying?: boolean): Record<string, unknown> | null;
    renderAnnotationElement(annotation: unknown): HTMLElement | null;
  };
  deserialize(data: unknown, ...rest: unknown[]): Promise<TextBoxEditor | null>;
  updateDefaultParams(type: number, value: unknown): void;
}

let textBoxClass: TextBoxClass | null = null;

/** The text box editor class, once Folio has seen one of its editors. */
export const textBoxEditors = () => textBoxClass;

export const isTextBox = <T extends { constructor: { _type?: string } }>(editor: T): editor is T & TextBoxEditor =>
  editor.constructor._type === "freetext";

/**
 * Teaches pdf.js's text box editor class, reached through one of its editors, to draw, copy and save a font.
 * pdf.js doesn't export the class, so this runs for every editor Folio sees and patches the class once.
 */
export function prepareTextBoxes(editor: { constructor: { _type?: string } }) {
  if (!isTextBox(editor) || textBoxClass) return;
  const type = editor.constructor as unknown as TextBoxClass;
  textBoxClass = type;
  const { render, serialize, renderAnnotationElement } = type.prototype;
  type.prototype.render = function (this: TextBoxEditor) {
    const div = render.call(this);
    if (this.editorDiv) this.editorDiv.style.fontFamily = fontCss(this.__folioFont);
    return div;
  };
  type.prototype.serialize = function (this: TextBoxEditor, copying?: boolean) {
    const fonted = !!this.__folioFont && this.__folioFont !== DEFAULT_TEXT_FONT;
    // pdf.js leaves a box that came with the PDF out of saved copies unless its own properties changed. A new font
    // is a change too: compare against a value the box can't have, then restore the original.
    const initial = this._initialData;
    if (fonted && this.annotationElementId && initial && !copying) this._initialData = { ...initial, value: null };
    let data;
    try { data = serialize.call(this, copying); } finally { this._initialData = initial; }
    if (data && !data.deleted && fonted) data.fontFamily = this.__folioFont;
    return data;
  };
  // How pdf.js shows an edited box from the PDF while reading.
  type.prototype.renderAnnotationElement = function (this: TextBoxEditor, annotation: unknown) {
    const content = renderAnnotationElement.call(this, annotation);
    if (content) content.style.fontFamily = fontCss(this.__folioFont);
    return content;
  };
  // Restored marks and pasted copies come back through deserialize.
  const deserialize = type.deserialize;
  type.deserialize = async function (this: TextBoxClass, data: unknown, ...rest: unknown[]) {
    const created = await deserialize.call(this, data, ...rest);
    const font = (data as { fontFamily?: unknown } | null)?.fontFamily;
    if (created) created.__folioFont = isTextFont(font) ? font : DEFAULT_TEXT_FONT;
    return created;
  };
}

export const textBoxSize = (editor: TextBoxEditor) => {
  const size = editor.propertiesToUpdate.find(([type]) => type === AnnotationEditorParamsType.FREETEXT_SIZE)?.[1];
  return typeof size === "number" ? size : DEFAULT_TEXT_SIZE;
};

/** Changes a box's font as one step in pdf.js's Undo history. */
export function setTextBoxFont(editor: TextBoxEditor, font: TextFont) {
  const previous = editor.__folioFont ?? DEFAULT_TEXT_FONT;
  if (previous === font) return;
  const apply = (next: TextFont) => {
    editor.__folioFont = next;
    if (!editor.editorDiv) return;
    editor.editorDiv.style.fontFamily = fontCss(next);
    remeasure(editor);
  };
  editor.addCommands({ cmd: () => apply(font), undo: () => apply(previous), post: () => editor._uiManager?.updateUI(editor), mustExec: true });
}

/** Fits the box to its text again, as pdf.js's private FreeTextEditor.#setEditorDimensions does. */
function remeasure(editor: TextBoxEditor) {
  if (!editor.isAttachedToDOM || !editor.div) return;
  const [parentWidth, parentHeight] = editor.parentDimensions;
  const rect = editor.div.getBoundingClientRect();
  const [width, height] = editor.rotation % 180 === editor.parentRotation % 180 ? [rect.width, rect.height] : [rect.height, rect.width];
  editor.width = width / parentWidth;
  editor.height = height / parentHeight;
  editor.fixAndSetPosition();
}

/** A saved text box in a font other than Helvetica, as pdf.js writes it into a PDF. */
export interface FontedTextBox {
  font: Exclude<TextFont, "sans">;
  /** Zero-based page in the open document. */
  page: number;
  value: string;
  rect: number[];
  fontSize: number;
  color: number[];
  rotation: number;
}

export function fontedTextBox(data: Record<string, unknown>): FontedTextBox | null {
  const { fontFamily: font, pageIndex: page, value, rect, fontSize, color, rotation } = data;
  if (!isTextFont(font) || font === "sans" || typeof page !== "number" || typeof value !== "string" || !value
    || !Array.isArray(rect) || rect.length !== 4 || !rect.every(Number.isFinite)
    || typeof fontSize !== "number" || !(fontSize > 0)) return null;
  const rgb = Array.isArray(color) && color.length >= 3 ? color.slice(0, 3).map(Number) : [0, 0, 0];
  return { font, page, value, rect: rect.map(Number), fontSize, color: rgb, rotation: Number(rotation) || 0 };
}

const fmt = (n: number) => String(Math.round(n * 1000) / 1000);

/**
 * Redraws text boxes pdf.js saved in Helvetica in their chosen fonts, embedding each font so other readers show it.
 * A box is found by its page, text and rectangle; `sourcePages` gives each page of `bytes` as a zero-based page of the
 * open document when they differ (an export of some pages, or reordered ones). Boxes a font can't write keep pdf.js's
 * drawing, and PDFs pdf-lib can't rewrite faithfully (encrypted ones) come back unchanged.
 */
export async function drawTextFonts(bytes: Uint8Array, boxes: FontedTextBox[], sourcePages?: number[]): Promise<Uint8Array> {
  if (!boxes.length) return bytes;
  const [{ PDFArray, PDFDict, PDFDocument: Document, PDFHexString, PDFName, PDFNumber, PDFString, StandardFonts }, kit] =
    await Promise.all([import("pdf-lib"), import("@pdf-lib/fontkit")]);
  let pdf: PDFDocument;
  try {
    pdf = await Document.load(bytes, { updateMetadata: false });
  } catch {
    return bytes;
  }
  // Its ES build has only a default export.
  pdf.registerFontkit(kit.default);
  const fonts = new Map<FontedTextBox["font"], Promise<{ font: PDFFont; chars: Set<number> }>>();
  const fontFor = (id: FontedTextBox["font"]) => {
    let font = fonts.get(id);
    if (!font) {
      font = (async () => {
        let embedded;
        if (id === "serif") embedded = await pdf.embedFont(StandardFonts.TimesRoman);
        else if (id === "mono") embedded = await pdf.embedFont(StandardFonts.Courier);
        else {
          const response = await fetch(new URL("/fonts/easyJetRoundedBook.ttf", import.meta.url));
          if (!response.ok) throw new Error("Couldn't load easyJet Rounded Book.");
          embedded = await pdf.embedFont(new Uint8Array(await response.arrayBuffer()));
        }
        return { font: embedded, chars: new Set(embedded.getCharacterSet()) };
      })();
      fonts.set(id, font);
    }
    return font;
  };
  const FreeText = PDFName.of("FreeText");
  let changed = false;
  for (const [index, page] of pdf.getPages().entries()) {
    const annots = page.node.Annots();
    const source = sourcePages ? sourcePages[index] : index;
    const unmatched = boxes.filter(b => b.page === source);
    if (!annots || !unmatched.length) continue;
    for (let i = 0; i < annots.size(); i++) {
      // pdf-lib throws on unexpected object types. Such an annotation isn't one of Folio's, and the file still saves.
      try {
        const annot = annots.lookupMaybe(i, PDFDict);
        if (!annot || annot.get(PDFName.of("Subtype")) !== FreeText) continue;
        const value = annot.lookupMaybe(PDFName.of("Contents"), PDFString, PDFHexString)?.decodeText();
        const rect = annot.lookupMaybe(PDFName.of("Rect"), PDFArray)?.asArray()
          .map((n: unknown) => n instanceof PDFNumber ? n.asNumber() : NaN);
        if (value === undefined || rect?.length !== 4) continue;
        const at = unmatched.findIndex(b => b.value === value && b.rect.every((n, k) => Math.abs(n - rect[k]) < 0.05));
        if (at < 0) continue;
        const [box] = unmatched.splice(at, 1);
        const { font, chars } = await fontFor(box.font);
        if ([...value.replace(/\n/g, "")].some(c => !chars.has(c.codePointAt(0)!))) continue;
        annot.set(PDFName.of("AP"), pdf.context.obj({ N: pdf.context.register(appearance(box, rect, font)) }));
        changed = true;
      } catch {
        continue;
      }
    }
  }
  if (!changed) return bytes;
  return pdf.save({ updateFieldAppearances: false });

  // Lays the text out the way pdf.js does for its Helvetica appearance (FreeTextAnnotation.createNewAppearanceStream),
  // shrinking it to fit the box when the new font runs wider.
  function appearance(box: FontedTextBox, rect: number[], font: PDFFont) {
    const [x1, y1, x2, y2] = rect;
    let [w, h] = [x2 - x1, y2 - y1];
    if (box.rotation % 180 !== 0) [w, h] = [h, w];
    const lines = box.value.split("\n");
    const widest = Math.max(...lines.map(line => font.widthOfTextAtSize(line, box.fontSize)));
    const scale = Math.min(1, widest > w ? w / widest : 1, 1.35 * box.fontSize * lines.length > h ? h / (1.35 * box.fontSize * lines.length) : 1);
    const size = box.fontSize * scale;
    const lineHeight = 1.35 * size;
    const ascent = size;
    const turns: Record<number, [number[], number[], number[]]> = {
      0: [[1, 0, 0, 1], [x1, y1, w, h], [x1, y2 - ascent]],
      90: [[0, 1, -1, 0], [y1, -x2, w, h], [y1, -x1 - ascent]],
      180: [[-1, 0, 0, -1], [-x2, -y2, w, h], [-x2, -y1 - ascent]],
      270: [[0, -1, 1, 0], [-y2, x1, w, h], [-y2, x2 - ascent]],
    };
    const [matrix, clip, first] = turns[((box.rotation % 360) + 360) % 360] ?? turns[0];
    const [r, g, b] = box.color.map(c => fmt(Math.min(255, Math.max(0, c)) / 255));
    const ops = ["q", `${matrix.join(" ")} 0 0 cm`, `${clip.map(fmt).join(" ")} re W n`, "BT", `${r} ${g} ${b} rg`,
      `0 Tc /FolioText ${fmt(size)} Tf`, `${first.map(fmt).join(" ")} Td ${font.encodeText(lines[0]).toString()} Tj`];
    for (const line of lines.slice(1)) ops.push(`0 -${fmt(lineHeight)} Td ${font.encodeText(line).toString()} Tj`);
    ops.push("ET", "Q");
    return pdf.context.stream(ops.join("\n"), {
      Type: "XObject", Subtype: "Form", FormType: 1, BBox: rect, Matrix: [1, 0, 0, 1, -x1, -y1],
      Resources: { Font: { FolioText: font.ref } },
    });
  }
}
