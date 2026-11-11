// Highlights, drawings, text boxes and notes as first-class, auto-saved document data.
// pdf.js owns the on-page editors; this keeps a saved copy per document and recreates the
// editors whenever the document is opened again.
import { AnnotationEditorParamsType, AnnotationEditorType } from "pdfjs-dist";
import type { EventBus, PDFViewer } from "./viewer";
import type { PDFDocumentProxy } from "./pdf";
import { getNotes, saveNotes, type DocNotes, type StoredMark } from "./recent";
import {
  DEFAULT_TEXT_FONT, DEFAULT_TEXT_SIZE, fontedTextBox, isTextBox, isTextFont, prepareTextBoxes, setTextBoxFont,
  textBoxEditors, textBoxSize, type FontedTextBox, type TextFont,
} from "./textFonts";

export type Mark = StoredMark;

const SAVED_TYPES: Record<string, Mark["type"]> = { highlight: "highlight", ink: "ink", freetext: "freetext" };

/* Minimal views of pdf.js internals used here (they're not part of its typed public API). */
interface Editor {
  id: string;
  pageIndex: number;
  div?: HTMLElement | null;
  deleted?: boolean;
  annotationElementId?: string | null;
  isClone?: boolean;
  x: number;
  y: number;
  pageDimensions: [number, number];
  comment: { text?: string; deleted?: boolean } | null;
  getPDFRect(): number[];
  getRectInCurrentCoords(rect: number[], pageHeight: number): number[];
  fixAndSetPosition(): void;
  isEmpty(): boolean;
  serialize(isForCopying?: boolean): Record<string, unknown> | null;
  remove(): void;
  constructor: { _type?: string };
  __folioMark?: string;
  __folioGeneration?: number;
  __folioText?: string;
  __folioShape?: Mark["shape"];
}
interface EditorLayer {
  div: HTMLElement;
  deserialize(data: Record<string, unknown>): Promise<Editor | null>;
  add(editor: Editor): void;
}
interface UIManager {
  addEditor(editor: Editor): void;
  getMode(): number;
  readonly hasSelection: boolean;
  endCurrentEditing(): void;
  unselectAll(): void;
  getEditors(pageIndex: number): Iterable<Editor>;
  getEditor(id: string): Editor | undefined;
  getLayer(pageIndex: number): EditorLayer | undefined;
  highlightSelection(methodOfCreation?: string): void;
}

const toHex = (color: unknown) =>
  Array.isArray(color) && color.length >= 3
    ? "#" + color.slice(0, 3).map((c) => Math.round(Number(c)).toString(16).padStart(2, "0")).join("")
    : "#ffd84d";

const newId = () => crypto.randomUUID();

export interface TextStyle { font: TextFont; size: number }

export interface MarkupState {
  marks: Mark[];
  note: string;
  bookmarkedPages: number[] | null;
}

export class MarkupController {
  #bus: EventBus;
  #viewer: PDFViewer;
  #ui: UIManager | null = null;
  #docId: string | null = null;
  #generation = 0;
  #storageGeneration = 0;
  #persist = false;
  #note = "";
  #bookmarkedPages: number[] | null = null;
  /** Saved marks whose page hasn't rendered yet, keyed by page index. */
  #pending = new Map<number, Mark[]>();
  /** Last known record for every mark (text, note, created time survive re-serialization). */
  #known = new Map<string, Mark>();
  #restoring = new Set<number>();
  #selectionText = "";
  #penWidth = 1;
  /** Font and size for the next text box. */
  #textStyle: TextStyle = { font: DEFAULT_TEXT_FONT, size: DEFAULT_TEXT_SIZE };
  #syncTimer = 0;
  #saveTimer = 0;
  #listener: (state: MarkupState) => void;
  /** True while Folio itself switches pdf.js modes for a one-click highlight. */
  quickHighlighting = false;

  constructor(bus: EventBus, viewer: PDFViewer, listener: (state: MarkupState) => void) {
    this.#bus = bus;
    this.#viewer = viewer;
    this.#listener = listener;
    bus.on("annotationeditoruimanager", ({ uiManager }: { uiManager: UIManager }) => {
      this.#ui = uiManager;
      // Every editor passes through here before it's drawn: new text boxes take the chosen font.
      const addEditor = uiManager.addEditor.bind(uiManager);
      uiManager.addEditor = (editor) => {
        prepareTextBoxes(editor);
        if (isTextBox(editor) && !editor.__folioFont) editor.__folioFont = editor.annotationElementId ? DEFAULT_TEXT_FONT : this.#textStyle.font;
        addEditor(editor);
      };
    });
    // pdf.js asks the host app to enter highlight mode when highlighting a selection.
    bus.on("showannotationeditorui", ({ source, mode }: { source: UIManager; mode: number }) => {
      if (source === this.#ui) this.#viewer.annotationEditorMode = { mode };
    });
    document.addEventListener("selectionchange", () => {
      if (!this.#docId || this.quickHighlighting || this.#ui?.getMode() !== AnnotationEditorType.HIGHLIGHT) return;
      const selection = document.getSelection();
      if (!selection) return;
      if (selection.isCollapsed && this.#selectionText) {
        // pdf.js clears the range when it creates the mark. Record its passage before the next drag.
        this.sync();
      }
      const anchor = selection.anchorNode instanceof Element ? selection.anchorNode : selection.anchorNode?.parentElement;
      this.#selectionText = anchor?.closest(".textLayer") && this.#viewer.viewer?.contains(anchor)
        ? selection.toString().replace(/\s+/g, " ").trim() : "";
    });
    bus.on("annotationeditorlayerrendered", ({ source, pageNumber }: { source: unknown; pageNumber: number }) => {
      if (source === this.#viewer.getPageView(pageNumber - 1)) void this.#restorePage(pageNumber - 1);
    });
    for (const name of ["editingstateschanged", "annotationeditormodechanged", "editingaction"]) {
      bus.on(name, () => this.scheduleSync());
    }
    bus.on("annotationeditormodechanged", ({ mode }: { mode: number }) => {
      if (mode === AnnotationEditorType.INK) this.#applyPenWidth();
      if (mode === AnnotationEditorType.FREETEXT) this.#applyTextSize();
    });
  }

  get state(): MarkupState {
    return { marks: this.#sorted(), note: this.#note, bookmarkedPages: this.#bookmarkedPages };
  }

  clearEditorSelection() {
    this.#ui?.unselectAll();
  }

  get penWidth() { return this.#penWidth; }

  setPenWidth(width: number) {
    // PDF.js groups consecutive strokes. Finish that group before changing the next stroke's width.
    this.#ui?.endCurrentEditing();
    this.clearEditorSelection();
    this.#penWidth = width;
    this.#applyPenWidth();
  }

  #applyPenWidth() {
    if (this.#ui && !this.#ui.hasSelection) {
      this.#bus.dispatch("switchannotationeditorparams", { source: this, type: AnnotationEditorParamsType.INK_THICKNESS, value: this.#penWidth });
    }
  }

  /** The selected text box's font and size, or those the next text box gets. */
  get textStyle(): TextStyle {
    const [box] = this.#selectedTextBoxes();
    return box ? { font: box.__folioFont ?? DEFAULT_TEXT_FONT, size: textBoxSize(box) } : { ...this.#textStyle };
  }

  /** Changes the selected text boxes, or the next one when none is selected. Returns whether the default changed. */
  setTextStyle(style: Partial<TextStyle>) {
    const boxes = this.#selectedTextBoxes();
    if (boxes.length) {
      for (const box of boxes) {
        if (style.font) setTextBoxFont(box, style.font);
        if (style.size) box.updateParams(AnnotationEditorParamsType.FREETEXT_SIZE, style.size);
      }
      this.scheduleSync();
      return false;
    }
    this.#textStyle = { ...this.#textStyle, ...style };
    if (style.size) this.#applyTextSize();
    return true;
  }

  #applyTextSize() {
    const editors = textBoxEditors();
    if (editors) editors.updateDefaultParams(AnnotationEditorParamsType.FREETEXT_SIZE, this.#textStyle.size);
    else if (this.#ui && !this.#ui.hasSelection) {
      this.#bus.dispatch("switchannotationeditorparams", { source: this, type: AnnotationEditorParamsType.FREETEXT_SIZE, value: this.#textStyle.size });
    }
  }

  /** Every text box on the drawn pages, including ones that came with the PDF. */
  #textBoxes() {
    const ui = this.#ui;
    const boxes = [];
    for (let page = 0; ui && page < this.#viewer.pagesCount; page++) {
      for (const editor of ui.getEditors(page)) if (isTextBox(editor) && !editor.deleted) boxes.push(editor);
    }
    return boxes;
  }

  #selectedTextBoxes() {
    return this.#ui?.hasSelection ? this.#textBoxes().filter(e => e.isSelected) : [];
  }

  /** Text boxes in a font other than Helvetica, as saved copies will hold them (see drawTextFonts). */
  fontedTextBoxes(): FontedTextBox[] {
    const boxes: FontedTextBox[] = [];
    for (const editor of this.#textBoxes()) {
      let data: Record<string, unknown> | null = null;
      try { data = editor.serialize(false); } catch { data = null; }
      const box = data && fontedTextBox(data);
      if (box) boxes.push(box);
    }
    for (const marks of this.#pending.values()) {
      for (const mark of marks) {
        const box = mark.type === "freetext" ? fontedTextBox({ ...mark.data, pageIndex: mark.page }) : null;
        if (box) boxes.push(box);
      }
    }
    return boxes;
  }

  /** Loads saved markup for a newly opened document. Password-protected documents aren't persisted. */
  async load(docId: string, persist: boolean, reading?: Pick<MarkupState, "note" | "bookmarkedPages">) {
    const readingGeneration = this.#storageGeneration;
    this.unload();
    const generation = this.#generation;
    this.#docId = docId;
    const record = persist ? await getNotes(docId) : undefined;
    if (generation !== this.#generation) return;
    const cleared = record && record.generation < this.#storageGeneration;
    if (cleared || (record && readingGeneration < record.generation)) reading = undefined;
    const saved = cleared ? undefined : record?.value;
    if (record) this.#storageGeneration = Math.max(this.#storageGeneration, record.generation);
    this.#persist = persist;
    this.#note = reading?.note ?? saved?.note ?? "";
    this.#bookmarkedPages = reading?.bookmarkedPages ?? saved?.bookmarkedPages ?? [];
    for (const mark of saved?.marks ?? []) {
      this.#known.set(mark.id, mark);
      const list = this.#pending.get(mark.page) ?? [];
      list.push(mark);
      this.#pending.set(mark.page, list);
    }
    this.#emit();
    if (reading) this.#scheduleSave();
    // Pages that rendered before the saved data arrived.
    for (const page of [...this.#pending.keys()]) if (this.#ui?.getLayer(page)) void this.#restorePage(page);
  }

  unload() {
    // Read the live editors first: an edit committed in place (retyping a text box, then Escape) fires no event
    // that schedules a sync, so it would otherwise be lost.
    this.sync();
    this.flush();
    this.#generation++;
    this.#persist = false;
    this.#ui = null;
    this.quickHighlighting = false;
    this.#restoring.clear();
    this.#selectionText = "";
    const selection = document.getSelection();
    if (selection?.anchorNode && this.#viewer.viewer?.contains(selection.anchorNode)) selection.removeAllRanges();
    this.#docId = null;
    this.#pending.clear();
    this.#known.clear();
    this.#note = "";
    this.#bookmarkedPages = null;
    this.#emit();
  }

  clearSaved(generation: number) {
    if (generation <= this.#storageGeneration) return;
    this.#storageGeneration = generation;
    if (!this.#persist) return;
    this.#generation++;
    this.quickHighlighting = false;
    window.clearTimeout(this.#saveTimer);
    this.#restoring.clear();
    this.#pending.clear();
    this.#known.clear();
    this.#note = "";
    this.#bookmarkedPages = [];
    for (const editor of this.#liveEditors()) editor.remove();
    this.#emit();
  }

  async #restorePage(page: number) {
    const marks = this.#pending.get(page);
    const layer = this.#ui?.getLayer(page);
    if (!marks?.length || !layer || this.#restoring.has(page)) return;
    this.#restoring.add(page);
    const generation = this.#generation;
    try {
      let added = false;
      // Marks pdf.js couldn't bring back stay pending, which keeps their saved records (and Notes entries) intact.
      const missed: Mark[] = [];
      for (const mark of marks) {
        // pdf.js only puts a text box's saved text on the page when the box is recreated as a copy; without it
        // the box comes back empty and the next sync drops it. isClone stops the copy being offset like a paste
        // (the same steps pdf.js takes when it duplicates a page).
        const textBox = mark.type === "freetext";
        const editor = await layer.deserialize({ ...mark.data, deleted: false, annotationElementId: null, isCopy: textBox }).catch(() => null);
        if (generation !== this.#generation) return;
        if (!editor) {
          missed.push(mark);
          continue;
        }
        if (textBox) {
          prepareTextBoxes(editor);
          (editor as Editor & { __folioFont?: TextFont }).__folioFont = isTextFont(mark.data.fontFamily) ? mark.data.fontFamily : DEFAULT_TEXT_FONT;
        }
        editor.__folioMark = mark.id;
        editor.__folioGeneration = this.#storageGeneration;
        if (mark.type === "highlight") editor.__folioText = mark.text;
        editor.__folioShape = mark.shape;
        if (textBox) editor.isClone = true;
        layer.add(editor);
        // An empty editor would be thrown away by pdf.js (and its mark with it) at the next tool change.
        if (editor.isEmpty()) {
          editor.remove();
          missed.push(mark);
          continue;
        }
        if (textBox) this.#moveTo(editor, mark.data.rect);
        if (mark.note) this.#setEditorNote(editor, mark.note);
        added = true;
      }
      // pdf.js hides an editor layer that was empty when its page rendered, and saved marks arrive after that.
      if (added) layer.div.hidden = false;
      if (missed.length) this.#pending.set(page, missed);
      else this.#pending.delete(page);
    } finally {
      if (generation === this.#generation) this.#restoring.delete(page);
    }
    this.scheduleSync();
  }

  /** Puts a restored text box back on its saved rectangle: pdf.js nudges a new box by its font size when it's first drawn. */
  #moveTo(editor: Editor, rect: unknown) {
    if (!Array.isArray(rect)) return;
    try {
      const [pageWidth, pageHeight] = editor.pageDimensions;
      const [wantX, wantY] = editor.getRectInCurrentCoords(rect, pageHeight);
      const [haveX, haveY] = editor.getRectInCurrentCoords(editor.getPDFRect(), pageHeight);
      editor.x += (wantX - haveX) / pageWidth;
      editor.y += (wantY - haveY) / pageHeight;
      editor.fixAndSetPosition();
    } catch {
      /* the text is restored either way */
    }
  }

  #setEditorNote(editor: Editor, note: string) {
    try {
      // Stored on the pdf.js editor too, so saved copies of the PDF carry the note as a pop-up.
      (editor as unknown as { comment: string }).comment = note;
    } catch {
      /* the note still lives in Folio's own record */
    }
  }

  #liveEditors(): Editor[] {
    const ui = this.#ui;
    const count = this.#viewer.pagesCount;
    if (!ui || !count) return [];
    const editors: Editor[] = [];
    for (let page = 0; page < count; page++) {
      for (const editor of ui.getEditors(page)) {
        if (SAVED_TYPES[editor.constructor._type ?? ""] && !editor.annotationElementId && !editor.deleted) editors.push(editor);
      }
    }
    return editors;
  }

  async addShape(page: number, shape: NonNullable<Mark["shape"]>, data: Record<string, unknown>) {
    const layer = this.#ui?.getLayer(page);
    if (!layer) throw new Error("The page is not ready.");
    const editor = await layer.deserialize(data);
    if (this.#ui?.getLayer(page) !== layer) return;
    if (!editor) throw new Error("The drawing could not be created.");
    editor.__folioShape = shape;
    layer.add(editor);
    this.sync();
  }

  scheduleSync() {
    window.clearTimeout(this.#syncTimer);
    this.#syncTimer = window.setTimeout(() => this.sync(), 250);
  }

  /** Reads every live editor back into records, then saves. Returns ids of newly seen marks. */
  sync(): string[] {
    window.clearTimeout(this.#syncTimer);
    if (!this.#docId) return [];
    const fresh: string[] = [];
    const live = new Map<string, Mark>();
    for (const editor of this.#liveEditors()) {
      // pdf.js can reattach a removed editor through its Undo history.
      if (this.#persist && editor.__folioGeneration !== undefined && editor.__folioGeneration < this.#storageGeneration) {
        editor.remove();
        continue;
      }
      editor.__folioGeneration = this.#storageGeneration;
      if (editor.isEmpty()) {
        // A saved mark whose editor reads empty (a text box cleared but not yet committed) keeps its record until
        // the edit is committed. Removing a mark removes its editor, so deletes still work.
        const previous = editor.__folioMark ? this.#known.get(editor.__folioMark) : undefined;
        if (previous) live.set(previous.id, previous);
        continue;
      }
      let data: Record<string, unknown> | null = null;
      try { data = editor.serialize(false); } catch { data = null; }
      if (!data) continue;
      if (!editor.__folioMark) {
        editor.__folioMark = newId();
        fresh.push(editor.__folioMark);
        if (data.quadPoints) editor.__folioText = this.#selectionText;
      }
      const previous = this.#known.get(editor.__folioMark);
      const type = SAVED_TYPES[editor.constructor._type!];
      const comment = editor.comment;
      const text = type === "highlight" ? previous?.text || editor.__folioText || ""
        : type === "freetext" ? String(data.value ?? "") : "";
      const mark: Mark = {
        id: editor.__folioMark,
        type,
        ...(editor.__folioShape && { shape: editor.__folioShape }),
        page: editor.pageIndex,
        text,
        note: comment && !comment.deleted ? comment.text ?? "" : previous?.note ?? "",
        color: toHex(data.color),
        createdAt: previous?.createdAt ?? Date.now(),
        data: { ...data, id: undefined, popupRef: undefined },
      };
      live.set(mark.id, mark);
    }
    // Marks on pages that haven't rendered yet are kept as saved.
    for (const marks of this.#pending.values()) for (const mark of marks) live.set(mark.id, mark);
    const changed = live.size !== this.#known.size
      || [...live.values()].some((mark) => JSON.stringify(mark) !== JSON.stringify(this.#known.get(mark.id)));
    this.#known = live;
    if (changed) {
      this.#emit();
      this.#scheduleSave();
    }
    return fresh;
  }

  #sorted() {
    return [...this.#known.values()].sort((a, b) => {
      if (a.page !== b.page) return a.page - b.page;
      const ra = a.data.rect as number[] | undefined;
      const rb = b.data.rect as number[] | undefined;
      // PDF y grows upward: higher on the page first.
      return (rb?.[3] ?? 0) - (ra?.[3] ?? 0) || a.createdAt - b.createdAt;
    });
  }

  #emit() {
    this.#listener(this.state);
  }

  #scheduleSave() {
    if (!this.#persist) return;
    window.clearTimeout(this.#saveTimer);
    this.#saveTimer = window.setTimeout(() => this.flush(), 400);
  }

  flush() {
    window.clearTimeout(this.#saveTimer);
    if (!this.#persist || !this.#docId) return;
    const notes: DocNotes = { id: this.#docId, note: this.#note, marks: [...this.#known.values()],
      bookmarkedPages: this.#bookmarkedPages ?? [], updatedAt: Date.now() };
    return saveNotes(notes, this.#storageGeneration);
  }

  /**
   * Saved marks on pages pdf.js hasn't drawn yet, serialized like their editors would be. pdf.js only recreates a
   * mark's editor (and adds it to its annotation storage) when the page renders, so saved copies and printouts
   * need these added to include every mark.
   */
  async unrenderedMarks(pdf: PDFDocumentProxy): Promise<Record<string, unknown>[]> {
    const generation = this.#generation;
    const marks: Record<string, unknown>[] = [];
    for (const [pageIndex, pending] of this.#pending) {
      const [, , pageRight] = (await pdf.getPage(pageIndex + 1)).view;
      if (generation !== this.#generation) throw new Error("The document changed.");
      for (const mark of pending) {
        const data: Record<string, unknown> = { ...mark.data, pageIndex, deleted: false };
        const top = (mark.data.rect as number[] | undefined)?.[3];
        // The pop-up pdf.js gives an editor's comment (see AnnotationEditor.addComment).
        if (mark.note && top !== undefined) data.popup = { contents: mark.note, deleted: false, rect: [pageRight + 1, top - 100, pageRight + 181, top] };
        marks.push(data);
      }
    }
    return marks;
  }

  /** Highlights the current text selection in one click. Resolves to the new mark's id. */
  async highlightSelection(color: string): Promise<string | null> {
    const generation = this.#generation;
    const ui = this.#ui;
    const selection = document.getSelection();
    if (!ui || !selection || selection.isCollapsed) return null;
    // A page's text can be selected before pdf.js has finished drawing the page, but its highlight layer only exists
    // afterwards, and pdf.js drops a highlight (and the selection) on a page without one. Wait for the layer.
    const anchor = selection.anchorNode instanceof Element ? selection.anchorNode : selection.anchorNode?.parentElement;
    const page = Number((anchor?.closest(".page") as HTMLElement | null)?.dataset.pageNumber);
    for (let i = 0; i < 200 && page && !ui.getLayer(page - 1); i++) {
      await new Promise((r) => setTimeout(r, 25));
      if (generation !== this.#generation) return null;
    }
    if (selection.isCollapsed) return null;
    this.sync();
    const before = new Set(this.#liveEditors().map((e) => e.id));
    this.#selectionText = selection.toString().replace(/\s+/g, " ").trim();
    this.quickHighlighting = true;
    try {
      this.#bus.dispatch("switchannotationeditorparams", { source: null, type: AnnotationEditorParamsType.HIGHLIGHT_COLOR, value: color });
      ui.highlightSelection("floating_button");
      // pdf.js clears the selection when it goes ahead; if it's still there, pdf.js found nothing to highlight.
      if (!selection.isCollapsed) return null;
      // pdf.js switches into highlight mode, creates the editor, and we return to reading mode. The switch first
      // redraws the visible pages, which can take seconds on a busy computer, and returning to reading mode before
      // it finishes cancels it and loses the highlight, so give it plenty of time.
      let created: Editor | undefined;
      for (let i = 0; i < 400 && !created; i++) {
        await new Promise((r) => setTimeout(r, 25));
        if (generation !== this.#generation) return null;
        created = this.#liveEditors().find((e) => !before.has(e.id) && !e.isEmpty());
      }
      // pdf.js focuses the new highlight. Leaving highlight mode with it focused moves focus to the current page's
      // highlight layer, and throws if that page hasn't drawn one yet, which leaves pdf.js stuck mid-switch and
      // every later highlight ignored. Nothing needs the focus, so drop it first.
      const focused = document.activeElement;
      if (focused instanceof HTMLElement && focused.closest(".annotationEditorLayer")) focused.blur();
      this.#viewer.annotationEditorMode = { mode: AnnotationEditorType.NONE };
      await new Promise((r) => setTimeout(r, 0));
      if (!created || generation !== this.#generation) return null;
      this.sync();
      return created.__folioMark ?? null;
    } finally {
      if (generation === this.#generation) this.quickHighlighting = false;
    }
  }

  #editorFor(markId: string) {
    return this.#liveEditors().find((e) => e.__folioMark === markId);
  }

  /** The highlight under a screen point (pdf.js highlights don't take clicks while reading), with its screen box. */
  markAt(clientX: number, clientY: number): { id: string; rect: { left: number; top: number; right: number; bottom: number } } | null {
    for (const editor of this.#liveEditors()) {
      if (editor.constructor._type !== "highlight" || !editor.__folioMark) continue;
      const view = this.#viewer.getPageView(editor.pageIndex) as unknown as
        { div?: HTMLElement; viewport?: { convertToViewportPoint(x: number, y: number): number[] } } | undefined;
      const surface = view?.div?.querySelector(".canvasWrapper") ?? view?.div;
      if (!surface || !view?.viewport) continue;
      const box = surface.getBoundingClientRect();
      if (clientY < box.top || clientY > box.bottom || clientX < box.left || clientX > box.right) continue;
      const quads = (this.#known.get(editor.__folioMark)?.data.quadPoints ?? editor.serialize(false)?.quadPoints) as ArrayLike<number> | undefined;
      if (!quads) continue;
      const q = Array.from(quads);
      const hits: { left: number; top: number; right: number; bottom: number }[] = [];
      for (let i = 0; i + 7 < q.length; i += 8) {
        const corners = [0, 2, 4, 6].map((k) => view.viewport!.convertToViewportPoint(q[i + k], q[i + k + 1]));
        const xs = corners.map((c) => box.left + c[0]);
        const ys = corners.map((c) => box.top + c[1]);
        hits.push({ left: Math.min(...xs), right: Math.max(...xs), top: Math.min(...ys), bottom: Math.max(...ys) });
      }
      if (hits.some((r) => clientX >= r.left - 2 && clientX <= r.right + 2 && clientY >= r.top - 2 && clientY <= r.bottom + 2)) {
        return {
          id: editor.__folioMark,
          rect: {
            left: Math.min(...hits.map((r) => r.left)), right: Math.max(...hits.map((r) => r.right)),
            top: Math.min(...hits.map((r) => r.top)), bottom: Math.max(...hits.map((r) => r.bottom)),
          },
        };
      }
    }
    return null;
  }

  recolor(markId: string, hex: string) {
    const editor = this.#editorFor(markId) as (Editor & { updateParams?: (type: number, value: string) => void }) | undefined;
    editor?.updateParams?.(AnnotationEditorParamsType.HIGHLIGHT_COLOR, hex);
    this.scheduleSync();
  }

  markForEditorId(editorId: string) {
    const editor = this.#ui?.getEditor(editorId);
    return editor?.__folioMark ?? null;
  }

  setNote(markId: string, note: string) {
    const editor = this.#editorFor(markId);
    if (editor) this.#setEditorNote(editor, note);
    const known = this.#known.get(markId);
    if (known) this.#known.set(markId, { ...known, note });
    for (const marks of this.#pending.values()) {
      const pending = marks.find((m) => m.id === markId);
      if (pending) pending.note = note;
    }
    this.#emit();
    this.#scheduleSave();
    if (editor) this.scheduleSync();
  }

  setDocNote(note: string) {
    this.#note = note;
    this.#emit();
    this.#scheduleSave();
  }

  toggleBookmark(page: number) {
    const pages = this.#bookmarkedPages;
    if (!pages || !Number.isInteger(page) || page < 1 || page > this.#viewer.pagesCount) return;
    this.#bookmarkedPages = pages.includes(page)
      ? pages.filter(n => n !== page)
      : [...pages, page].sort((a, b) => a - b);
    this.#emit();
    this.#scheduleSave();
  }

  remove(markId: string) {
    const editor = this.#editorFor(markId);
    if (editor) editor.remove();
    for (const [page, marks] of this.#pending) {
      const rest = marks.filter((m) => m.id !== markId);
      if (rest.length) this.#pending.set(page, rest);
      else this.#pending.delete(page);
    }
    this.#known.delete(markId);
    this.#emit();
    this.#scheduleSave();
  }

  /** Scrolls a mark into view and briefly pulses it. */
  reveal(markId: string) {
    const mark = this.#known.get(markId);
    if (!mark) return;
    const rect = mark.data.rect as number[] | undefined;
    this.#viewer.scrollPageIntoView({
      pageNumber: mark.page + 1,
      destArray: rect ? [null, { name: "XYZ" }, null, rect[3] + 48, null] : undefined,
    } as Parameters<PDFViewer["scrollPageIntoView"]>[0]);
    window.setTimeout(() => {
      const div = this.#editorFor(markId)?.div;
      if (!div) return;
      div.classList.remove("folio-pulse");
      void div.offsetWidth;
      div.classList.add("folio-pulse");
    }, 350);
  }
}
