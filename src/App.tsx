import { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react";
import {
  EventBus,
  LinkTarget,
  PDFFindController,
  PDFLinkService,
  PDFViewer,
  ScrollMode,
  SpreadMode,
} from "./viewer";
import { AnnotationEditorParamsType, AnnotationEditorType } from "pdfjs-dist";
import { loadPdf, copyPdfPages, thumbnailDataUrl, type PDFDocumentProxy } from "./pdf";
import { contentKey, getRecent, MAX_RECENT_FILE_BYTES, onStoredDataCleared, removeRecent, saveRecent, updateLastPage } from "./recent";
import { readPdfBlob, readPdfResponse } from "./pdfInput";
import { FIT_SCALES, readView, writeView, type ReadingView } from "./session";
import { readSetting, writeSetting } from "./preferences";
import { DocumentTabBar } from "./DocumentTabBar";
import { trackTitlebar } from "./titlebar";
import type { DocumentTabs, OpenSource } from "./documentTabs";
import { SearchResults } from "./SearchResults";
import { Icon, type IconName } from "./icons";
import { Welcome } from "./Welcome";
import { Sidebar, type SidebarTab } from "./Sidebar";
import { MarkupController, type Mark, type MarkupState, type TextStyle } from "./markup";
import { DEFAULT_TEXT_FONT, DEFAULT_TEXT_SIZE, TEXT_FONTS, TEXT_SIZES, drawTextFonts, isTextFont, isTextSize, type TextFont } from "./textFonts";
import { HighlightMenu, NoteBox, SelectionPopover, useTextSelection, type SelectionInfo } from "./Selection";
import { NotesPanel } from "./NotesPanel";
import { notesJson, notesMarkdown } from "./noteExport";
import { copyText } from "./clipboard";
import { imageRegions, trackPageImages } from "./nightImages";
import { recognizeScannedPages } from "./pageOcr";
import { scanLanguagesFrom } from "./scanLanguage";
import { ImageMenu, ImageViewer, type ImageTarget } from "./ImageViewer";
import { MoreMenu, type MoreItem } from "./MoreMenu";
import { IS_MAC, keys } from "./platform";
import { AiPanel, forgetClearedChats, type AiContext } from "./AiPanel";
import { ZoomMenu } from "./ZoomMenu";
import { SearchBar, type FindResult } from "./SearchBar";
import { AboutDialog, InfoDialog, OpenUrlDialog, PasswordDialog, ShortcutsDialog, parseWebUrl } from "./Dialogs";
import { SignatureDialog } from "./SignatureDialog";
import { ContentEditor, type ContentEditorHandle, type NextSelection } from "./ContentEditor";
import { VersionHistoryPanel } from "./VersionHistoryPanel";
import { saveVersion, type PdfChange } from "./versionHistory";
import { SettingsDialog } from "./SettingsDialog";
import { ChangelogDialog } from "./ChangelogDialog";
import { PrintDialog } from "./PrintDialog";
import type { Release } from "./releaseNotes";
import { detectSignatureTargets, placeSignature, type SignatureTarget } from "./signatures";
import { keepDrawingPointer } from "./drawing";
import { drawShapes, type Shape } from "./shapes";
import { enableWheelPaging } from "./wheelPaging";
import { LinkHistory, type HistoryState } from "./linkHistory";

type Tool = "select" | "highlight" | "ink" | "freetext";

const TOOL_MODES: Record<Tool, number> = {
  select: AnnotationEditorType.NONE,
  highlight: AnnotationEditorType.HIGHLIGHT,
  ink: AnnotationEditorType.INK,
  freetext: AnnotationEditorType.FREETEXT,
};

const TOOLS: { id: Tool; icon: IconName; label: string; key: string }[] = [
  { id: "select", icon: "cursor", label: "Select", key: "V" },
  { id: "highlight", icon: "highlight", label: "Highlight", key: "H" },
  { id: "ink", icon: "pen", label: "Draw", key: "D" },
  { id: "freetext", icon: "text", label: "Text", key: "T" },
];

const HIGHLIGHTS = [
  { name: "Yellow", hex: "#FFD84D" },
  { name: "Green", hex: "#5EDC9A" },
  { name: "Blue", hex: "#6CB8FF" },
  { name: "Pink", hex: "#FF8FC7" },
  { name: "Red", hex: "#FF5A4E" },
];

const INKS = [
  { name: "Black", hex: "#1C1B19" },
  { name: "Red", hex: "#E5533D" },
  { name: "Blue", hex: "#2F6FED" },
  { name: "Green", hex: "#1F9D55" },
  { name: "Purple", hex: "#8B5CF6" },
];

const PALETTE: Partial<Record<Tool, typeof INKS>> = { highlight: HIGHLIGHTS, ink: INKS, freetext: INKS };

const COLOR_PARAM: Partial<Record<Tool, number>> = {
  highlight: AnnotationEditorParamsType.HIGHLIGHT_COLOR,
  ink: AnnotationEditorParamsType.INK_COLOR,
  freetext: AnnotationEditorParamsType.FREETEXT_COLOR,
};

function storedTextStyle(): TextStyle {
  const font = readSetting("folio:text-font");
  const size = Number(readSetting("folio:text-size"));
  return { font: isTextFont(font) ? font : DEFAULT_TEXT_FONT, size: isTextSize(size) ? size : DEFAULT_TEXT_SIZE };
}

interface OpenDoc {
  /** Full content identity for Recent, reading position, annotations and chat. */
  id: string;
  name: string;
  /** The file's size; its bytes live in pdf.js's worker (pdf.getData()) rather than in a second copy here. */
  size: number;
  pdf: PDFDocumentProxy;
  /** Used only in memory when reopening an annotated snapshot for page reordering. */
  password?: string;
  /** Original library entry while content edits are held in memory. */
  sourceId?: string;
  historyAllowed: boolean;
  edits?: { original?: Blob; changes: PdfChange[] };
}

function changesForCopy(doc: OpenDoc, annotated: boolean, order?: number[]) {
  const changes = [...(doc.edits?.changes ?? [])];
  if (annotated && !changes.some(change => change.kind === "annotations")) changes.push({ kind: "annotations" });
  if (order) changes.push({ kind: "pages", order });
  return changes;
}

interface Pending {
  name: string;
  data: Uint8Array;
  error?: string;
}

// Blob copies the bytes itself, so no copy is made here. (The cast: BlobPart rules out views over a
// SharedArrayBuffer, which pdf.js never hands out, but its types don't say so.)
const pdfBlob = (bytes: Uint8Array) => new Blob([bytes as Uint8Array<ArrayBuffer>], { type: "application/pdf" });

function saveBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

export default function App({ tabs, tabKey, upgradeNotes }: { tabs: DocumentTabs; tabKey: number; upgradeNotes: Release[] }) {
  const headerRef = useRef<HTMLElement>(null);
  const active = useSyncExternalStore(tabs.subscribe, () => tabs.active === tabKey);
  const multipleTabs = useSyncExternalStore(tabs.subscribe, () => tabs.all.length > 1);
  const containerRef = useRef<HTMLDivElement>(null);
  const innerRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<PDFViewer | null>(null);
  const busRef = useRef<EventBus | null>(null);
  const linkRef = useRef<PDFLinkService | null>(null);
  const navigationRef = useRef<LinkHistory | null>(null);
  const [navigation, setNavigation] = useState<HistoryState>({ back: false, forward: false });
  const [backHint, setBackHint] = useState<string | null>(null);
  const backHintButton = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    if (!backHint) return;
    const timer = window.setTimeout(() => {
      if (document.activeElement !== backHintButton.current) setBackHint(null);
    }, 5500);
    return () => window.clearTimeout(timer);
  }, [backHint]);
  const findRef = useRef<PDFFindController | null>(null);
  const restoreRef = useRef<ReadingView | null>(null);
  const restoringRef = useRef(false);
  const preferencesRef = useRef({ sidebar: true, spread: false });
  const opening = useRef<AbortController | null>(null);
  const lastPageRef = useRef<number>(1);
  const docRef = useRef<OpenDoc | null>(null);
  const markupRef = useRef<MarkupController | null>(null);
  const [chatGeneration, setChatGeneration] = useState(0);
  const [notesGeneration, setNotesGeneration] = useState(0);
  useEffect(() => onStoredDataCleared((category, generation) => {
    if (category === "ai") {
      forgetClearedChats(generation);
      setChatGeneration(current => Math.max(current, generation));
    } else if (category === "notes") {
      markupRef.current?.clearSaved(generation);
      if (persistRef.current) {
        setNotesGeneration(current => Math.max(current, generation));
        setNoteBox(null);
      }
    }
  }), []);
  const persistRef = useRef(false);
  const exporting = useRef<PDFDocumentProxy | null>(null);
  const savingCopy = useRef(false);

  const [doc, setDoc] = useState<OpenDoc | null>(null);
  const [loading, setLoading] = useState(false);
  const [page, setPage] = useState(1);
  const [textStatus, setTextStatus] = useState<string | null>(null);
  const [scanLanguages, setScanLanguages] = useState(() => scanLanguagesFrom(readSetting("folio:scan-language")));
  const [pageInput, setPageInput] = useState("1");
  const [numPages, setNumPages] = useState(0);
  const [scale, setScale] = useState(1);
  const [wheelPaging, setWheelPaging] = useState(() => readSetting("folio:wheel-paging") === "1");
  const wheelPagingRef = useRef(wheelPaging);
  wheelPagingRef.current = wheelPaging;
  const [rotation, setRotation] = useState(0);
  const [sidebar, setSidebar] = useState(() => readSetting("folio:sidebar") !== "0" && window.innerWidth > 760);
  const [searchOpen, setSearchOpen] = useState(false);
  const [find, setFind] = useState<FindResult>({ current: 0, total: 0, state: null });
  const [theme, setTheme] = useState(document.documentElement.dataset.theme || "light");
  const [spread, setSpread] = useState(() => readSetting("folio:spread") === "1");
  const spreadRef = useRef(spread);
  spreadRef.current = spread;
  const [night, setNight] = useState(() => readSetting("folio:night") === "1");
  const [tool, setTool] = useState<Tool>("select");
  const [drawTool, setDrawTool] = useState<"Pen" | Shape>("Pen");
  const [colors, setColors] = useState<Partial<Record<Tool, string>>>({
    highlight: HIGHLIGHTS[0].hex,
    ink: INKS[0].hex,
    freetext: INKS[0].hex,
  });
  /** The selected text box's font and size, or the next box's. */
  const [textStyle, setTextStyle] = useState(storedTextStyle);
  const colorsRef = useRef(colors);
  colorsRef.current = colors;
  const pageRef = useRef(1);
  const [dirty, setDirty] = useState(false);
  const editRevision = useRef(0);
  const markDirty = useCallback(() => { editRevision.current++; setDirty(true); }, []);
  const [contentEditing, setContentEditing] = useState(false);
  const contentRef = useRef<ContentEditorHandle>(null);
  const nextSelection = useRef<NextSelection | null>(null);
  const [editorDirty, setEditorDirty] = useState(false);
  const [pageDraft, setPageDraft] = useState<{ order: number[]; copy?: Uint8Array } | null>(null);
  const pageOrder = pageDraft?.order.some((n, i) => n !== i + 1) ? pageDraft.order : undefined;
  const unsaved = !!doc?.sourceId || dirty || editorDirty || (!!pageOrder && !pageDraft?.copy);
  const [focus, setFocus] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [printCopy, setPrintCopy] = useState<{ pdf: PDFDocumentProxy; bytes: Uint8Array; page: number; name: string } | null>(null);
  const preparingPrint = useRef(false);
  useEffect(() => () => { if (printCopy) void printCopy.pdf.loadingTask.destroy(); }, [printCopy]);
  const [dialog, setDialog] = useState<"info" | "shortcuts" | "about" | "url" | "settings" | "ai-settings" | "changelog" | "upgrade" | null>(
    upgradeNotes.length ? "upgrade" : null);
  // A pasted or dropped link waits in Open from URL, so the user sees where it goes before Folio contacts it.
  const [linkDraft, setLinkDraft] = useState("");
  const openUrlDialog = (link = "") => { setLinkDraft(link); setDialog("url"); };
  const [pending, setPending] = useState<Pending | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [updateVersion, setUpdateVersion] = useState<string | null>(null);
  const [signing, setSigning] = useState(false);
  const [signatureTargets, setSignatureTargets] = useState<SignatureTarget[]>([]);
  const [signatureDismissed, setSignatureDismissed] = useState(false);
  const [markup, setMarkup] = useState<MarkupState>({ marks: [], note: "", bookmarkedPages: null });
  const [sidebarTab, setSidebarTab] = useState<SidebarTab>("pages");
  const [noteBox, setNoteBox] = useState<{ markId: string; anchor: { x: number; top: number; bottom: number }; quote: string; initial: string } | null>(null);
  const [stageWidth, setStageWidth] = useState(800);
  const [container, setContainer] = useState<HTMLDivElement | null>(null);
  const [pinned, setPinned] = useState<SelectionInfo | null>(null);
  const [highlightMenu, setHighlightMenu] = useState<{ id: string; rect: { left: number; top: number; right: number; bottom: number } } | null>(null);
  const [imageMenu, setImageMenu] = useState<{ target: ImageTarget; x: number; y: number } | null>(null);
  const [imageView, setImageView] = useState<ImageTarget | null>(null);
  const [aiOpen, setAiOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [aiContext, setAiContext] = useState<AiContext | null>(null);

  useLayoutEffect(() => {
    if (active && window.folioDesktop && !focus && !dialog && !signing && !pending && !imageView) {
      return trackTitlebar(headerRef.current!);
    }
  }, [active, doc, focus, dialog, signing, pending, imageView]);

  useEffect(() => {
    setSigning(false); setSignatureTargets([]); setSignatureDismissed(false);
    if (!doc || contentEditing) return;
    let cancelled = false;
    const timer = setTimeout(() => {
      void detectSignatureTargets(doc.pdf, page).then(targets => {
        if (!cancelled) setSignatureTargets(targets);
      }).catch(() => {});
    }, 350);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [doc, page, contentEditing]);

  docRef.current = doc;
  preferencesRef.current = { sidebar, spread };

  useLayoutEffect(() => tabs.markDirty(tabKey, unsaved), [tabs, tabKey, unsaved]);
  useEffect(() => { window.dispatchEvent(new Event("folio:tab-visibility")); }, [active]);

  const toastTimer = useRef(0);
  const flash = useCallback((msg: string) => {
    setToast(msg);
    window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(null), 2200);
  }, []);

  // ── Viewer bootstrap ────────────────────────────────────────────────
  useEffect(() => {
    const container = containerRef.current!;
    const eventBus = new EventBus();
    const linkService = new PDFLinkService({ eventBus, externalLinkTarget: LinkTarget.BLANK });
    const findController = new PDFFindController({ eventBus, linkService });
    findRef.current = findController;
    const viewer = new PDFViewer({
      container,
      viewer: innerRef.current!,
      eventBus,
      linkService,
      findController,
      annotationEditorMode: AnnotationEditorType.NONE,
      annotationEditorHighlightColors: HIGHLIGHTS.map((c) => `${c.name.toLowerCase()}=${c.hex}`).join(","),
      // Records where images are drawn (used to keep them true to color in night reading). The huge
      // minimum size turns off pdf.js's own right-click image targets, which Folio doesn't use.
      imagesRightClickMinSize: 1e9,
    });
    trackPageImages(eventBus);
    const stopDrawingPointer = keepDrawingPointer(container);
    linkService.setViewer(viewer);
    const history = new LinkHistory(viewer, (state, followed) => {
      setNavigation(state);
      if (followed && tabs.active === tabKey && tabs.takeLinkHint()) {
        setBackHint(followed.from === followed.to ? "Back to previous spot" : `Back to page ${followed.from}`);
      } else setBackHint(null);
    });
    linkService.setHistory(history);
    navigationRef.current = history;
    viewerRef.current = viewer;
    const markupController = new MarkupController(eventBus, viewer, setMarkup);
    markupRef.current = markupController;
    markupController.setTextStyle(storedTextStyle());
    for (const name of ["editingstateschanged", "annotationeditorparamschanged"]) {
      eventBus.on(name, () => setTextStyle(markupController.textStyle));
    }
    busRef.current = eventBus;
    linkRef.current = linkService;

    eventBus.on("pagesinit", () => {
      restoringRef.current = true;
      const saved = restoreRef.current;
      if (wheelPagingRef.current) viewer.scrollMode = ScrollMode.PAGE;
      viewer.spreadMode = (saved?.spread ?? spreadRef.current) ? SpreadMode.ODD : SpreadMode.NONE;
      viewer.pagesRotation = saved?.rotation ?? 0;
      viewer.currentScaleValue = saved?.scale ?? (window.innerWidth < 900 ? "page-width" : "auto");
      const restoredPage = Math.min(lastPageRef.current, viewer.pagesCount);
      viewer.currentPageNumber = restoredPage;
      if (saved && Number.isFinite(saved.top) && Number.isFinite(saved.left)) {
        viewer.scrollPageIntoView({ pageNumber: restoredPage, destArray: [null, { name: "XYZ" }, saved.left, saved.top, null] });
      }
      restoringRef.current = false;
    });
    eventBus.on("updateviewarea", ({ location }: { location: { pageNumber: number; top: number; left: number } }) => {
      const current = docRef.current;
      if (!current || restoringRef.current) return;
      writeView(current.id, { page: location.pageNumber, scale: viewer.currentScaleValue,
        rotation: viewer.pagesRotation, ...preferencesRef.current, top: location.top, left: location.left });
    });
    eventBus.on("pagechanging", ({ pageNumber }: { pageNumber: number }) => {
      pageRef.current = pageNumber;
      setPage(pageNumber);
      setPageInput(String(pageNumber));
    });
    eventBus.on("scalechanging", ({ scale }: { scale: number }) => setScale(scale));
    eventBus.on("rotationchanging", ({ pagesRotation }: { pagesRotation: number }) => setRotation(pagesRotation));
    eventBus.on("updatefindmatchescount", ({ matchesCount }: { matchesCount: { current: number; total: number } }) =>
      setFind((f) => ({ ...f, ...matchesCount })),
    );
    eventBus.on(
      "updatefindcontrolstate",
      ({ state, matchesCount }: { state: number; matchesCount: { current: number; total: number } }) =>
        setFind({ ...matchesCount, state }),
    );
    eventBus.on("annotationeditormodechanged", ({ mode }: { mode: number }) => {
      // One-click highlights pass through highlight mode; the reader stays in Select.
      if (markupRef.current?.quickHighlighting) return;
      const entry = Object.entries(TOOL_MODES).find(([, m]) => m === mode);
      const next = entry ? (entry[0] as Tool) : "select";
      setTool(next);
      // Apply the remembered colour once the editor layer for this mode is live.
      const param = COLOR_PARAM[next];
      const value = colorsRef.current[next];
      if (param && value) eventBus.dispatch("switchannotationeditorparams", { source: null, type: param, value });
    });

    // Ctrl/⌘ + wheel or trackpad pinch zooms around the cursor.
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      if (!viewer.pdfDocument) return;
      viewer.updateScale({
        drawingDelay: 300,
        scaleFactor: Math.exp(-e.deltaY / 300),
        origin: [e.clientX, e.clientY],
      });
    };
    container.addEventListener("wheel", onWheel, { passive: false });
    return () => { stopDrawingPointer(); container.removeEventListener("wheel", onWheel); };
  }, [tabs, tabKey]);

  useEffect(() => {
    const viewer = viewerRef.current!, bus = busRef.current!;
    const events = new AbortController();
    let stop: (() => void) | undefined;
    const start = () => { stop = recognizeScannedPages(bus, viewer, scanLanguages, setTextStatus); };
    const end = () => { stop?.(); stop = undefined; };
    bus.on("pagesinit", start, { signal: events.signal });
    bus.on("pagesdestroy", end, { signal: events.signal });
    if (viewer.pagesCount) start();
    return () => { events.abort(); end(); };
  }, [scanLanguages]);

  useEffect(() => {
    const viewer = viewerRef.current!;
    if (viewer.pagesCount) viewer.scrollMode = wheelPaging ? ScrollMode.PAGE : ScrollMode.VERTICAL;
    if (wheelPaging) return enableWheelPaging(containerRef.current!, viewer);
  }, [wheelPaging, doc]);

  // ── Opening documents ───────────────────────────────────────────────
  const beginOpen = useCallback(() => {
    opening.current?.abort();
    opening.current = new AbortController();
    setPending(null);
    setToast(null);
    return opening.current.signal;
  }, []);

  const openBytes = useCallback(
    async (name: string, data: Uint8Array, password?: string, signal = beginOpen(), initialPage?: number,
      edit?: { doc: OpenDoc; change: PdfChange }, exported?: OpenDoc, exportedOrder?: number[]) => {
      const revision = edit?.doc;
      const viewer = viewerRef.current!;
      setLoading(true);
      setError(null);
      let unattached: PDFDocumentProxy | undefined;
      try {
        const pdf = await loadPdf(data, signal, password);
        unattached = pdf;
        const id = await contentKey(data);
        signal.throwIfAborted();
        const existing = revision ? undefined : await getRecent(id);
        const historyAllowed = !password && data.byteLength <= MAX_RECENT_FILE_BYTES && (revision ? revision.historyAllowed :
          await pdf.getMetadata().then(({ info }) => !(info as { EncryptFilterName?: string }).EncryptFilterName).catch(() => false));
        const edits = edit && { original: edit.doc.edits?.original ?? (historyAllowed ? pdfBlob(await edit.doc.pdf.getData()) : undefined),
          changes: [...changesForCopy(edit.doc, edit.doc.pdf.annotationStorage.size > 0), edit.change] };
        signal.throwIfAborted();
        const previous = revision ?? exported;
        if (previous && docRef.current !== previous) return false;
        if (!revision && !tabs.opened(tabKey, id, name)) return false;
        const saved = previous ? readView(previous.id) ?? { page: pageRef.current, scale: viewer.currentScaleValue,
          rotation: viewer.pagesRotation, ...preferencesRef.current } : readView(id);
        restoreRef.current = initialPage && !previous ? null : saved;
        lastPageRef.current = Math.min(initialPage ?? saved?.page ?? existing?.lastPage ?? 1, pdf.numPages);
        pageRef.current = lastPageRef.current;
        if (saved) {
          setSidebar(saved.sidebar);
          setSpread(saved.spread);
          spreadRef.current = saved.spread;
          preferencesRef.current = { sidebar: saved.sidebar, spread: saved.spread };
        }

        const prev = docRef.current;
        const next: OpenDoc = { id, name, size: data.byteLength, pdf, password, sourceId: revision?.sourceId ?? revision?.id, historyAllowed, edits };
        docRef.current = next;
        // Password-protected documents aren't stored by Folio (same rule as Recent).
        const reading = previous ? markupRef.current!.state : undefined;
        if (reading?.bookmarkedPages && exportedOrder) {
          const bookmarks = new Set(reading.bookmarkedPages);
          reading.bookmarkedPages = exportedOrder.flatMap((n, i) => bookmarks.has(n) ? [i + 1] : []);
        }
        persistRef.current = !password && !revision;
        void markupRef.current!.load(id, persistRef.current, reading);
        setNoteBox(null);
        setPinned(null);
        setAiContext(null);
        viewer.setDocument(pdf);
        unattached = undefined;
        linkRef.current!.setDocument(pdf, null);
        navigationRef.current!.reset();
        prev?.pdf.loadingTask.destroy();

        // Markup is saved automatically; only documents Folio can't store count as having unsaved changes.
        (pdf.annotationStorage as { onSetModified: unknown }).onSetModified = () => {
          if (docRef.current === next && !persistRef.current) markDirty();
        };
        (pdf.annotationStorage as { onResetModified: unknown }).onResetModified = () => {
          // Serializing a subset or printout also resets pdf.js's flag. Only a full save clears ours.
          if (docRef.current === next && exporting.current !== pdf) setDirty(false);
        };

        setDoc(next);
        if (!revision) { nextSelection.current = null; setContentEditing(false); }
        setEditorDirty(false);
        setPageDraft(draft => revision && draft ? { order: draft.order } : null);
        setNumPages(pdf.numPages);
        setPage(lastPageRef.current);
        setPageInput(String(lastPageRef.current));
        setRotation(revision ? saved?.rotation ?? 0 : 0);
        setDirty(false);
        setTool("select");
        setPending(null);
        setSearchOpen(false);
        setFind({ current: 0, total: 0, state: null });
        document.title = `${name} — Folio`;

        setLoading(false);
        // Password-protected documents stay out of Recent: it would keep their bytes and a page-1
        // thumbnail unprotected on disk. Very large files are skipped too — see MAX_RECENT_FILE_BYTES.
        if (!revision && !password && data.byteLength <= MAX_RECENT_FILE_BYTES) void (async () => {
          const thumb = existing?.thumb || (await thumbnailDataUrl(pdf).catch(() => ""));
          if (signal.aborted) return;
          await saveRecent({
            id,
            name,
            size: data.byteLength,
            pages: pdf.numPages,
            openedAt: Date.now(),
            lastPage: pageRef.current,
            thumb,
            data: pdfBlob(data),
          });
        })().catch(() => {});
        return true;
      } catch (err) {
        if (signal.aborted) return;
        const e = err as { name?: string; code?: number; message?: string };
        if (e?.name === "PasswordException") {
          setPending({ name, data, error: e.code === 2 ? "Incorrect password. Try again." : undefined });
        } else {
          tabs.failed(tabKey);
          setError(`Couldn't open “${name}”. ${e?.message ?? "The file may be damaged."}`);
        }
      } finally {
        if (unattached) void unattached.loadingTask.destroy().catch(() => {});
        if (!signal.aborted) setLoading(false);
      }
    },
    [beginOpen, tabs, tabKey],
  );

  const openBlob = useCallback(async (name: string, blob: Blob) => {
    const signal = beginOpen();
    setLoading(true);
    setError(null);
    try {
      const data = await readPdfBlob(blob, signal);
      if (!signal.aborted) await openBytes(name, data, undefined, signal);
    } catch (error) {
      if (!signal.aborted) {
        setLoading(false);
        tabs.failed(tabKey);
        setError(`Couldn't open “${name}”. ${(error as Error).message}`);
        return false;
      }
    }
    return true;
  }, [beginOpen, openBytes, tabs, tabKey]);

  const openFiles = useCallback((files: FileList) => {
    tabs.open(Array.from(files, (file): OpenSource => {
      if (window.folioDesktop && /\.docx?$/i.test(file.name)) return { kind: "word", name: file.name, file };
      if (file.type && file.type !== "application/pdf" && !/\.pdf$/i.test(file.name))
        return { kind: "error", name: file.name, message: `“${file.name}” isn't a PDF.` };
      return { kind: "blob", name: file.name, blob: file };
    }));
  }, [tabs]);

  const pickFile = useCallback(() => {
    const input = document.createElement("input");
    input.type = "file";
    input.multiple = true;
    input.accept = window.folioDesktop ? "application/pdf,.pdf,.doc,.docx" : "application/pdf,.pdf";
    input.onchange = () => { if (input.files?.length) openFiles(input.files); };
    input.click();
  }, [openFiles]);

  const openUrl = useCallback(
    async (url: string) => {
      const signal = beginOpen();
      const host = new URL(url).hostname;
      setLoading(true);
      setError(null);
      setToast(`Opening ${host}…`);
      try {
        let file: { name: string; data: Uint8Array } | null;
        if (window.folioDesktop) {
          // Closing the tab or opening something else stops the download in the desktop shell too,
          // which otherwise keeps every other link waiting until it gives up.
          const desktop = window.folioDesktop;
          const request = crypto.randomUUID();
          const stop = () => { void desktop.cancelFetchPdf(request).catch(() => {}); };
          signal.addEventListener("abort", stop, { once: true });
          try { file = await desktop.fetchPdf(url, request); }
          finally { signal.removeEventListener("abort", stop); }
        } else {
          // The browser edition can only read links whose server allows cross-origin requests.
          // A minute to answer, and a minute for any pause after that: not a limit on the whole download.
          const slow = new AbortController();
          const answering = window.setTimeout(() => slow.abort(), 60_000);
          const res = await fetch(url, { signal: AbortSignal.any([signal, slow.signal]) }).catch(() => {
            throw new Error(`Couldn't download from ${host}. The site may not allow it.`);
          }).finally(() => window.clearTimeout(answering));
          const data = await readPdfResponse(res, signal, undefined, 60_000);
          const rawName = new URL(res.url || url).pathname.split("/").filter(Boolean).pop() || host;
          let name = rawName;
          try { name = decodeURIComponent(rawName); } catch {}
          file = { name: /\.pdf$/i.test(name) ? name : name + ".pdf", data };
        }
        if (signal.aborted) return;
        setToast(null);
        if (!file) { tabs.failed(tabKey); return setLoading(false); }
        await openBytes(file.name, file.data, undefined, signal);
      } catch (e: any) {
        if (signal.aborted) return;
        setToast(null);
        setLoading(false);
        tabs.failed(tabKey);
        setError(`Couldn't open that link. ${e?.message ?? ""}`.trim());
      }
    },
    [beginOpen, openBytes, tabs, tabKey],
  );

  const openSample = useCallback(async () => {
    const signal = beginOpen();
    try {
      const res = await fetch(`${import.meta.env.BASE_URL}sample.pdf`, { signal });
      const data = await readPdfResponse(res, signal);
      if (!signal.aborted) await openBytes("Folio sample.pdf", data, undefined, signal);
    } catch {
      if (!signal.aborted) { setLoading(false); tabs.failed(tabKey); setError("Couldn't open the sample document."); }
    }
  }, [beginOpen, openBytes, tabs, tabKey]);

  // Remember where the reader left off.
  useEffect(() => {
    if (!doc) return;
    const saved = readView(doc.id);
    if (saved) writeView(doc.id, { ...saved, sidebar, spread });
  }, [doc, sidebar, spread]);
  useEffect(() => {
    if (!doc) return;
    const timer = window.setTimeout(() => void updateLastPage(doc.id, page), 600);
    return () => window.clearTimeout(timer);
  }, [doc, page]);

  // ── Actions ─────────────────────────────────────────────────────────
  const goTo = useCallback((n: number) => {
    const v = viewerRef.current!;
    if (!v.pdfDocument) return;
    v.currentPageNumber = Math.min(Math.max(1, n), v.pagesCount);
  }, []);
  const zoom = (steps: number) => {
    const v = viewerRef.current!;
    if (v.pdfDocument) v.updateScale({ steps, drawingDelay: 150, origin: undefined as unknown as number[] });
  };
  const setScaleValue = (value: string) => {
    const v = viewerRef.current!;
    if (v.pdfDocument) v.currentScaleValue = value;
  };
  const toggleSpread = () => {
    const v = viewerRef.current!;
    const next = !spread;
    setSpread(next);
    writeSetting("folio:spread", next ? "1" : "0");
    if (!v.pdfDocument) return;
    const value = v.currentScaleValue;
    v.spreadMode = next ? SpreadMode.ODD : SpreadMode.NONE;
    // Re-fit so "fit width/page" accounts for two pages side by side.
    if (FIT_SCALES.includes(value)) v.currentScaleValue = value;
  };
  const rotate = () => {
    const v = viewerRef.current!;
    if (v.pdfDocument) v.pagesRotation = (v.pagesRotation + 90) % 360;
  };

  const chooseTool = async (next: Tool, editContent = false) => {
    if (contentRef.current && !(await contentRef.current.finish())) return false;
    const v = viewerRef.current!;
    if (!v.pdfDocument) return false;
    try {
      const owner = v.pdfDocument;
      if (next !== "select") await v.firstPagePromise;
      if (v.pdfDocument !== owner) return false;
      const current = Number(v.annotationEditorMode);
      if (current !== TOOL_MODES[next] && !(next === "select" && current === AnnotationEditorType.DISABLE)) {
        v.annotationEditorMode = { mode: TOOL_MODES[next] };
      }
      nextSelection.current = null;
      setPinned(null);
      setTool(next); setContentEditing(editContent); return true;
    } catch { flash("Annotations aren’t supported for this document"); return false; }
  };
  const leaveEditMode = () => chooseTool("select");
  const startSigning = async () => { if (await chooseTool("select")) setSigning(true); };

  useEffect(() => {
    if (!doc || tool !== "ink" || drawTool === "Pen") return;
    return drawShapes(containerRef.current!, viewerRef.current!, drawTool, colors.ink!, (page, data) => {
      void markupRef.current!.addShape(page, drawTool, data).catch(() => flash("Couldn't add the shape"));
    });
  }, [doc, tool, drawTool, colors.ink, scale, rotation, flash]);

  const chooseTextStyle = (style: Partial<TextStyle>) => {
    const controller = markupRef.current;
    if (!controller) return;
    // With no text box selected, the choice is the next box's and is remembered.
    if (controller.setTextStyle(style)) {
      if (style.font) writeSetting("folio:text-font", style.font);
      if (style.size) writeSetting("folio:text-size", String(style.size));
    }
    setTextStyle(controller.textStyle);
  };

  const chooseColor = (hex: string) => {
    // The palette sets the color for the next passage. pdf.js leaves the last highlight selected.
    if (tool === "highlight") markupRef.current?.clearEditorSelection();
    setColors((c) => ({ ...c, [tool]: hex }));
    const param = COLOR_PARAM[tool];
    if (param) busRef.current!.dispatch("switchannotationeditorparams", { source: null, type: param, value: hex });
  };

  // ── Highlights and notes ────────────────────────────────────────────
  const liveSelection = useTextSelection(container, !!doc && !contentEditing && tool === "select" && !signing && !noteBox, doc?.pdf);
  // Typing in the popover's note field clears the page selection, so the popover keeps its own copy.
  const selection = pinned ?? liveSelection;
  const highlightColor = colors.highlight ?? HIGHLIGHTS[0].hex;
  const toolRef = useRef(tool);
  toolRef.current = tool;

  const quickHighlight = async (hex: string, note = "") => {
    const info = selection;
    if (!info) return;
    setColors((c) => ({ ...c, highlight: hex }));
    if (pinned) {
      const current = document.getSelection();
      current?.removeAllRanges();
      current?.addRange(info.range);
    }
    setPinned(null);
    const owner = docRef.current;
    const id = await markupRef.current!.highlightSelection(hex);
    if (owner !== docRef.current) return;
    if (!id) {
      flash("Couldn't highlight that selection");
      return;
    }
    if (note) markupRef.current!.setNote(id, note);
  };

  const dismissSelection = () => {
    setPinned(null);
    document.getSelection()?.removeAllRanges();
  };

  // Clicking an existing highlight offers recolor, note, Ask AI, and one-click removal.
  useEffect(() => {
    if (!container) return;
    const onDown = () => setPinned(null);
    const onClick = (e: MouseEvent) => {
      if (toolRef.current !== "select" || e.button !== 0) return;
      const current = document.getSelection();
      if (current && !current.isCollapsed) return;
      if ((e.target as HTMLElement).closest("a, .annotationCommentButton, .folio-content-layer")) return;
      const hit = markupRef.current?.markAt(e.clientX, e.clientY);
      if (!hit) return setHighlightMenu(null);
      const stage = container.parentElement!.getBoundingClientRect();
      setHighlightMenu({ id: hit.id, rect: { left: hit.rect.left - stage.left, right: hit.rect.right - stage.left,
        top: hit.rect.top - stage.top, bottom: hit.rect.bottom - stage.top } });
    };
    const onScroll = () => setHighlightMenu(null);
    container.addEventListener("pointerdown", onDown);
    container.addEventListener("click", onClick);
    container.addEventListener("scroll", onScroll, { passive: true });
    return () => {
      container.removeEventListener("pointerdown", onDown);
      container.removeEventListener("click", onClick);
      container.removeEventListener("scroll", onScroll);
    };
  }, [container]);
  useEffect(() => setHighlightMenu(null), [doc, scale, tool]);

  // Clicking an image offers "View larger" and, when it has a web address, "Open in browser".
  useEffect(() => {
    if (!container) return;
    const onClick = (e: MouseEvent) => {
      setImageMenu(null);
      if ((e.target as Element).closest(".folio-content-layer")) return;
      if (toolRef.current !== "select" || e.button !== 0) return;
      const current = document.getSelection();
      if (current && !current.isCollapsed) return;
      if (markupRef.current?.markAt(e.clientX, e.clientY)) return; // highlights take precedence
      const pageEl = (e.target as HTMLElement).closest<HTMLElement>(".page");
      const pageNumber = Number(pageEl?.dataset.pageNumber);
      const view = pageNumber ? (viewerRef.current?.getPageView(pageNumber - 1) as unknown as
        { pdfPage?: { imageCoordinates?: ArrayLike<number> | null }; viewport?: { rotation: number } } | undefined) : undefined;
      const surface = pageEl?.querySelector(".canvasWrapper");
      if (!view?.pdfPage || !view.viewport || !surface) return;
      const box = surface.getBoundingClientRect();
      const x = (e.clientX - box.left) / box.width;
      const y = (e.clientY - box.top) / box.height;
      const hits = imageRegions(view.pdfPage, view.viewport.rotation).filter((r) => {
        const w = (r.right - r.left) * box.width;
        const h = (r.bottom - r.top) * box.height;
        // Skip tiny icons and full-page images (scanned pages), where a menu would just get in the way.
        return x >= r.left && x <= r.right && y >= r.top && y <= r.bottom && w >= 40 && h >= 40 && (r.right - r.left) * (r.bottom - r.top) < 0.85;
      });
      if (!hits.length) return;
      const region = hits.reduce((a, b) => ((a.right - a.left) * (a.bottom - a.top) <= (b.right - b.left) * (b.bottom - b.top) ? a : b));
      const link = document.elementsFromPoint(e.clientX, e.clientY)
        .find((el): el is HTMLAnchorElement => el instanceof HTMLAnchorElement && /^https?:/i.test(el.href) && !!el.closest(".annotationLayer"));
      e.preventDefault();
      e.stopPropagation();
      setHighlightMenu(null);
      const stage = container.parentElement!.getBoundingClientRect();
      setImageMenu({ target: { pageNumber, region, rotation: view.viewport.rotation, href: link?.href ?? null }, x: e.clientX - stage.left, y: e.clientY - stage.top });
    };
    const onScroll = () => setImageMenu(null);
    container.addEventListener("click", onClick, true);
    container.addEventListener("scroll", onScroll, { passive: true });
    return () => {
      container.removeEventListener("click", onClick, true);
      container.removeEventListener("scroll", onScroll);
    };
  }, [container]);
  useEffect(() => { setImageMenu(null); setImageView(null); }, [doc]);
  const menuMark = highlightMenu ? markup.marks.find((m) => m.id === highlightMenu.id) : undefined;

  const removeHighlight = (id: string) => {
    markupRef.current!.remove(id);
    setHighlightMenu(null);
    flash("Highlight removed");
  };

  const askAbout = (text: string, pageNumber: number) => {
    setAiContext({ text, page: pageNumber });
    setAiOpen(true);
    setHistoryOpen(false);
    document.getSelection()?.removeAllRanges();
  };

  useEffect(() => setContainer(containerRef.current), []);
  // Panels and focus mode change the space the page has: re-fit "auto"/"fit" zoom once they finish sliding.
  useEffect(() => {
    const refit = () => {
      const v = viewerRef.current;
      if (!v?.pdfDocument || !FIT_SCALES.includes(v.currentScaleValue)) return;
      // PDF.js hides the ink layer during a refit. Its pointer offsets are invalid until it is visible again.
      if (containerRef.current?.querySelector(".annotationEditorLayer.inkEditing.drawing")) {
        timer = window.setTimeout(refit, 280);
        return;
      }
      v.currentScaleValue = v.currentScaleValue;
    };
    let timer = window.setTimeout(refit, 280);
    return () => window.clearTimeout(timer);
  }, [aiOpen, sidebar, sidebarTab, focus]);
  useEffect(() => {
    const stage = container?.parentElement;
    if (!stage) return;
    const observer = new ResizeObserver(() => setStageWidth(stage.clientWidth));
    observer.observe(stage);
    return () => observer.disconnect();
  }, [container]);

  // pdf.js shows a small note button on highlights that have a note; open Folio's note box from it.
  useEffect(() => {
    if (!container) return;
    const onClick = (e: MouseEvent) => {
      const button = (e.target as HTMLElement).closest(".annotationCommentButton");
      const editorDiv = button?.parentElement?.closest<HTMLElement>("[id]");
      const markId = editorDiv && markupRef.current?.markForEditorId(editorDiv.id);
      if (!button || !markId) return;
      e.preventDefault();
      e.stopPropagation();
      const mark = markupRef.current!.state.marks.find((m) => m.id === markId);
      const stage = container.parentElement!.getBoundingClientRect();
      const r = button.getBoundingClientRect();
      setNoteBox({ markId, anchor: { x: r.left + r.width / 2 - stage.left, top: r.top - stage.top, bottom: r.bottom - stage.top },
        quote: mark?.text ?? "", initial: mark?.note ?? "" });
    };
    container.addEventListener("click", onClick, true);
    return () => container.removeEventListener("click", onClick, true);
  }, [container]);

  // Save notes and highlights before the window goes away.
  useEffect(() => {
    const flush = () => { markupRef.current?.sync(); markupRef.current?.flush(); };
    window.addEventListener("pagehide", flush);
    return () => window.removeEventListener("pagehide", flush);
  }, []);

  const currentBytes = async (doc = docRef.current, pages = pageOrder) => {
    if (!doc || exporting.current === doc.pdf) return null;
    exporting.current = doc.pdf;
    const signal = opening.current!.signal;
    // Marks on pages that haven't been drawn yet aren't in pdf.js's annotation storage; add them for this copy only.
    // (pdf.js saves only entries keyed with its editor prefix.)
    const storage = doc.pdf.annotationStorage;
    const keys: string[] = [];
    try {
      const unrendered = await markupRef.current!.unrenderedMarks(doc.pdf);
      for (const data of unrendered) {
        const key = `pdfjs_internal_editor_folio-unrendered-${keys.length}`;
        storage.setValue(key, data);
        keys.push(key);
      }
      const annotated = storage.size > 0;
      const fonted = annotated ? markupRef.current!.fontedTextBoxes() : [];
      const bytes = pages?.length ? await copyPdfPages(doc.pdf, pages, signal, doc.password)
        : annotated ? await doc.pdf.saveDocument() : await doc.pdf.getData();
      // pdf.js writes text boxes in Helvetica. Redrawing them rewrites the file, which would break a digital
      // signature, so signed PDFs keep pdf.js's drawing.
      if (!fonted.length || (await doc.pdf.getSignatures())?.length) return { bytes, annotated };
      return { bytes: await drawTextFonts(bytes, fonted, pages?.length ? pages.map(n => n - 1) : undefined), annotated };
    } finally {
      for (const key of keys) storage.remove(key);
      if (exporting.current === doc.pdf) exporting.current = null;
    }
  };

  const copyName = pageOrder && doc ? doc.name.replace(/\.pdf$/i, "") + " (reordered).pdf" : doc?.name ?? "document.pdf";
  const saveBytes = (bytes: Uint8Array) => {
    saveBlob(pdfBlob(bytes), copyName.toLowerCase().endsWith(".pdf") ? copyName : `${copyName}.pdf`);
  };

  const exportPages = async (pages: number[]) => {
    if (contentRef.current && !(await contentRef.current.finish())) return;
    const source = docRef.current;
    if (!source || (source.sourceId ?? source.id) !== (doc?.sourceId ?? doc?.id)) return;
    try {
      const copy = await currentBytes(source, pages);
      if (docRef.current !== source) return;
      if (!copy) { flash("Another export is in progress. Try again in a moment."); return; }
      saveBlob(pdfBlob(copy.bytes), source.name.replace(/\.pdf$/i, "") + " (selected pages).pdf");
    } catch {
      if (docRef.current === source) flash("Couldn't export the selected pages.");
    }
  };

  const exportNotes = async (format: "markdown" | "json") => {
    if (!doc) return;
    const controller = markupRef.current!;
    controller.sync();
    try {
      if (format === "markdown") {
        const ok = await copyText(notesMarkdown(doc.name, controller.state));
        if (docRef.current === doc) flash(ok ? "Markdown copied" : "Copy failed");
      } else {
        const json = await notesJson(doc, controller.state);
        if (docRef.current === doc) saveBlob(new Blob([json], { type: "application/json" }), doc.name.replace(/\.pdf$/i, "") + ".notes.json");
      }
    } catch {
      if (docRef.current === doc) flash("Couldn't export notes");
    }
  };

  const download = async () => {
    if (contentRef.current && !(await contentRef.current.finish())) return;
    const doc = docRef.current;
    if (!doc || savingCopy.current) return;
    savingCopy.current = true;
    const revision = editRevision.current;
    try {
      const copy = await currentBytes(doc, pageOrder);
      if (!copy || docRef.current !== doc) return;
      const name = doc.sourceId ? copyName.replace(/\.pdf$/i, "") + " (edited).pdf" : copyName;
      const blob = pdfBlob(copy.bytes);
      saveBlob(blob, /\.pdf$/i.test(name) ? name : name + ".pdf");
      if (editRevision.current === revision) setDirty(false);
      const changes = changesForCopy(doc, copy.annotated, pageOrder);
      let historyMessage = "";
      if (changes.length && doc.historyAllowed && blob.size <= MAX_RECENT_FILE_BYTES) {
        try {
          const original = doc.edits?.original ?? pdfBlob(await doc.pdf.getData());
          await saveVersion({ id: doc.sourceId ?? doc.id, name: doc.name, data: original },
            { id: await contentKey(copy.bytes), name, data: blob }, changes);
        } catch { historyMessage = "Copy exported. Couldn’t save its version history on this device."; }
      } else if (changes.length) {
        historyMessage = "Copy exported. History is available for unencrypted PDFs up to 50 MB.";
      }
      if (docRef.current !== doc) return;
      setError(null);
      if (doc.sourceId) {
        await openBytes(name, copy.bytes, undefined, beginOpen(), pageOrder ? pageOrder.indexOf(pageRef.current) + 1 : pageRef.current, undefined, doc, pageOrder);
      } else if (pageOrder) {
        setPageDraft(draft => draft?.order === pageOrder ? { order: pageOrder, copy: copy.bytes } : draft);
        flash("Downloading reordered copy");
      } else {
        flash(copy.annotated ? "Saved a copy with your highlights and notes" : "Download started");
      }
      if (historyMessage) flash(historyMessage);
    } catch {
      if (docRef.current === doc) {
        if (dirty) setDirty(true);
        setError("Couldn't save a copy of this document.");
      }
    } finally { savingCopy.current = false; }
  };

  const print = async () => {
    if (printCopy || preparingPrint.current) return;
    preparingPrint.current = true;
    try {
      if (contentRef.current && !(await contentRef.current.finish())) return;
      const source = docRef.current;
      if (!source) return;
      const copy = await currentBytes(source);
      if (!copy || docRef.current !== source) return;
      if (!window.folioDesktop) { printPdfBytes(copy.bytes); return; }
      const snapshot = await loadPdf(copy.bytes, opening.current!.signal, source.password);
      if (docRef.current !== source) { void snapshot.loadingTask.destroy(); return; }
      setPrintCopy({ pdf: snapshot, bytes: copy.bytes, page: pageOrder ? pageOrder.indexOf(page) + 1 : page, name: source.name });
    } catch {
      if (docRef.current === doc) setError("Couldn't print this document.");
    } finally {
      preparingPrint.current = false;
    }
  };

  const printPdfBytes = (bytes: Uint8Array) => {
    const url = URL.createObjectURL(pdfBlob(bytes));
    const frame = document.createElement("iframe");
    frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
    frame.src = url;
    frame.onload = () => {
      setTimeout(() => {
        try {
          frame.contentWindow?.focus();
          frame.contentWindow?.print();
        } catch {
          // The desktop shell denies blob: pop-ups, so a saved copy is the working fallback there.
          if (window.folioDesktop) { saveBytes(bytes); flash("Couldn't print — saved a copy instead"); }
          else window.open(url, "_blank");
        }
      }, 250);
    };
    document.body.appendChild(frame);
    setTimeout(() => {
      frame.remove();
      URL.revokeObjectURL(url);
    }, 60_000);
  };

  useEffect(() => {
    const open = async (source: OpenSource) => {
      switch (source.kind) {
        case "bytes":
          await openBytes(source.name, new Uint8Array(source.data.buffer, source.data.byteOffset, source.data.byteLength), source.password);
          if (source.note && docRef.current?.name === source.name) flash(source.note);
          break;
        case "blob":
          if (!await openBlob(source.name, source.blob) && source.recentId) void removeRecent(source.recentId);
          break;
        case "word": {
          const signal = beginOpen();
          setLoading(true); setError(null); setToast(`Converting ${source.name} to PDF…`);
          try {
            const file = await window.folioDesktop!.openDocument(source.file);
            if (signal.aborted) return;
            setToast(null);
            await openBytes(file.name, new Uint8Array(file.data), undefined, signal);
            if (file.note && docRef.current?.name === file.name) flash(file.note);
          } catch (error) {
            if (signal.aborted) return;
            setToast(null); setLoading(false); tabs.failed(tabKey);
            setError(`Couldn't convert “${source.name}” to PDF. ${(error as Error).message}`);
          }
          break;
        }
        case "sample": await openSample(); break;
        case "url": await openUrl(source.url); break;
        case "converting": setLoading(true); setToast(`Converting ${source.name} to PDF…`); break;
        case "error": setLoading(false); setToast(null); setError(source.message); tabs.failed(tabKey); break;
      }
    };
    tabs.bind(tabKey, { open: source => { void open(source); }, flush: async () => {
      if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
      markupRef.current?.sync();
      await markupRef.current?.flush();
      opening.current?.abort();
    } });
  }, [tabs, tabKey, openBytes, openBlob, openSample, openUrl, beginOpen, flash]);
  useEffect(() => {
    const sync = (event: StorageEvent) => {
      if (event.key === "folio:theme" && (event.newValue === "light" || event.newValue === "dark")) {
        document.documentElement.dataset.theme = event.newValue; setTheme(event.newValue);
      }
      if (event.key === "folio:night") setNight(event.newValue === "1");
      if (event.key === "folio:wheel-paging") setWheelPaging(event.newValue === "1");
      if (event.key === "folio:scan-language" || event.key === null) setScanLanguages(scanLanguagesFrom(event.newValue));
    };
    window.addEventListener("storage", sync);
    return () => window.removeEventListener("storage", sync);
  }, []);
  useEffect(() => { if (active) window.folioDesktop?.setTheme(theme); }, [active, theme]);
  // The update banner appears once a new version has been downloaded (the status reports that right away too).
  useEffect(() => window.folioDesktop?.onUpdateStatus((status) => {
    if (status.state === "ready" && status.version) setUpdateVersion(status.version);
  }), []);

  const setThemeTo = (next: "light" | "dark") => {
    document.documentElement.dataset.theme = next;
    writeSetting("folio:theme", next);
    setTheme(next);
  };

  const toggleFocus = async () => {
    const next = !focus;
    if (next && contentEditing) leaveEditMode();
    setFocus(next);
    try {
      if (next && !document.fullscreenElement) await document.documentElement.requestFullscreen();
      if (!next && document.fullscreenElement) await document.exitFullscreen();
    } catch {
      /* fullscreen can be refused; focus mode still hides the chrome */
    }
  };

  useEffect(() => {
    const onFs = () => !document.fullscreenElement && setFocus(false);
    document.addEventListener("fullscreenchange", onFs);
    return () => document.removeEventListener("fullscreenchange", onFs);
  }, []);

  useEffect(() => { writeSetting("folio:sidebar", sidebar ? "1" : "0"); }, [sidebar]);
  const setNightTo = (next: boolean) => {
    if (!writeSetting("folio:night", next ? "1" : "0")) flash("PDF dark mode changed, but couldn't save the preference.");
    setNight(next);
  };

  // ── Keyboard shortcuts ──────────────────────────────────────────────
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (tabs.active !== tabKey) return;
      const mod = e.ctrlKey || e.metaKey;
      const target = e.target;
      const typing = target instanceof Element && (target.closest("input, textarea, select, .freeTextEditor") ||
        (target instanceof HTMLElement && target.isContentEditable));
      if (e.altKey && !mod && !e.shiftKey && (e.key === "ArrowLeft" || e.key === "ArrowRight")) {
        if (e.defaultPrevented) return;
        if (typing || signing || pending || imageView || printCopy || document.querySelector('[aria-modal="true"]')) {
          // Chromium can leave the web reader here; macOS uses Option+arrow for native text editing.
          if (!IS_MAC && !window.folioDesktop) e.preventDefault();
        } else if (docRef.current) {
          e.preventDefault();
          if (e.key === "ArrowLeft") navigationRef.current?.back();
          else navigationRef.current?.forward();
        }
        return;
      }
      if (printCopy) {
        if ((e.ctrlKey || e.metaKey) && ["p", "o", "w", "t", "s", ",", "tab"].includes(e.key.toLowerCase())) e.preventDefault();
        return;
      }
      if (mod && !e.altKey && !e.shiftKey && e.key.toLowerCase() === "t") {
        e.preventDefault(); tabs.library(); return;
      }
      if (mod && !e.altKey && !e.shiftKey && e.key.toLowerCase() === "w") {
        e.preventDefault(); void tabs.close(tabKey); return;
      }
      if (e.ctrlKey && e.key === "Tab") {
        e.preventDefault(); tabs.cycle(e.shiftKey ? -1 : 1); return;
      }
      if (signing) return;

      if (mod && e.key === ",") {
        e.preventDefault();
        setDialog("settings");
        return;
      }
      if (mod && e.key.toLowerCase() === "o") {
        e.preventDefault();
        if (e.shiftKey) openUrlDialog();
        else pickFile();
        return;
      }
      if (!docRef.current) return;
      // A dialog owns the keyboard, and closes itself on Escape. The reader's shortcuts are swallowed
      // rather than left to act on the document behind it (or to fall through to the browser's own).
      if (dialog || pending || imageView) {
        const key = e.key.toLowerCase();
        if (mod && (["f", "s", "p", "=", "+", "-", "0", "\\"].includes(key) || (!typing && (key === "z" || key === "y")))) e.preventDefault();
        return;
      }

      if (mod && !typing && Number(viewerRef.current?.annotationEditorMode) === AnnotationEditorType.NONE
        && (e.key.toLowerCase() === "z" || e.key.toLowerCase() === "y")) {
        e.preventDefault();
        busRef.current?.dispatch("editingaction", { source: null, name: e.shiftKey || e.key.toLowerCase() === "y" ? "redo" : "undo" });
      } else if (mod && e.key.toLowerCase() === "f") {
        e.preventDefault();
        setSearchOpen(true);
        requestAnimationFrame(() => document.getElementById("folio-search")?.focus());
      } else if (mod && e.key.toLowerCase() === "s") {
        e.preventDefault();
        download();
      } else if (mod && e.key.toLowerCase() === "p") {
        e.preventDefault();
        print();
      } else if (mod && (e.key === "=" || e.key === "+")) {
        e.preventDefault();
        zoom(1);
      } else if (mod && e.key === "-") {
        e.preventDefault();
        zoom(-1);
      } else if (mod && e.key === "0") {
        e.preventDefault();
        setScaleValue("auto");
      } else if (mod && e.key === "\\") {
        e.preventDefault();
        setSidebar((s) => !s);
      } else if (typing || mod || e.altKey) {
        return;
      } else if (selection && tool === "select" && /^h$/i.test(e.key)) {
        e.preventDefault();
        void quickHighlight(highlightColor);
      } else if (selection && tool === "select" && /^n$/i.test(e.key)) {
        e.preventDefault();
        document.querySelector<HTMLInputElement>(".pop-note")?.focus();
      } else if (highlightMenu && (e.key === "Delete" || e.key === "Backspace")) {
        e.preventDefault();
        removeHighlight(highlightMenu.id);
      } else if (e.key === "Escape") {
        if (highlightMenu) setHighlightMenu(null);
        else if (noteBox) setNoteBox(null);
        else if (searchOpen) setSearchOpen(false);
        else if (contentRef.current?.deselect()) return;
        else if (contentEditing || tool !== "select") void leaveEditMode();
        else if (focus) toggleFocus();
      } else if (e.key === "f" || e.key === "F") {
        toggleFocus();
      } else if (e.key === "?") {
        setDialog("shortcuts");
      } else if (e.key === "2") {
        toggleSpread();
      } else if (e.key === "r" || e.key === "R") {
        rotate();
      } else {
        const t = TOOLS.find((t) => t.key.toLowerCase() === e.key.toLowerCase());
        if (t) {
          if (focus && t.id !== "select") void toggleFocus();
          chooseTool(t.id);
        }
        else if (e.key === "Home") goTo(1);
        else if (e.key === "End") goTo(numPages);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  // ── Drag & drop anywhere ────────────────────────────────────────────
  useEffect(() => {
    let depth = 0;
    let internal = false; // dragging text or images from inside Folio isn't an "open" gesture
    const hasFiles = (e: DragEvent) => {
      const types = e.dataTransfer?.types ?? [];
      return !internal && (types.includes("Files") || types.includes("text/uri-list") || types.includes("text/plain"));
    };
    const droppedUrl = (data: DataTransfer) => {
      const fromList = data.getData("text/uri-list").split(/\r?\n/).find((line) => line.trim() && !line.startsWith("#"));
      return parseWebUrl(fromList) ?? parseWebUrl(data.getData("text/plain"));
    };
    const start = () => (internal = true);
    const end = () => (internal = false);
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth++;
      setDragging(true);
    };
    const leave = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      if (--depth <= 0) {
        depth = 0;
        setDragging(false);
      }
    };
    const over = (e: DragEvent) => hasFiles(e) && e.preventDefault();
    const drop = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth = 0;
      setDragging(false);
      const files = e.dataTransfer?.files;
      const url = !files?.length && e.dataTransfer ? droppedUrl(e.dataTransfer) : null;
      if (files?.length) openFiles(files);
      else if (url) openUrlDialog(url);
      else setError("Drop a PDF file or a link to a PDF.");
    };
    // Pasting a link (outside text fields) fills in Open from URL too.
    const paste = (e: ClipboardEvent) => {
      if (signing || (e.target as HTMLElement | null)?.closest?.("input, textarea, [contenteditable=true], .freeTextEditor")) return;
      const url = parseWebUrl(e.clipboardData?.getData("text/plain"));
      if (!url) return;
      e.preventDefault();
      openUrlDialog(url);
    };
    window.addEventListener("dragstart", start);
    window.addEventListener("dragend", end);
    window.addEventListener("paste", paste);
    window.addEventListener("dragenter", enter);
    window.addEventListener("dragleave", leave);
    window.addEventListener("dragover", over);
    window.addEventListener("drop", drop);
    return () => {
      window.removeEventListener("dragstart", start);
      window.removeEventListener("dragend", end);
      window.removeEventListener("paste", paste);
      window.removeEventListener("dragenter", enter);
      window.removeEventListener("dragleave", leave);
      window.removeEventListener("dragover", over);
      window.removeEventListener("drop", drop);
    };
  }, [openFiles, signing]);

  const runFind = (query: string, opts: { again?: boolean; previous?: boolean; caseSensitive?: boolean; entireWord?: boolean }) =>
    busRef.current!.dispatch("find", {
      source: null,
      type: opts.again ? "again" : "",
      query,
      caseSensitive: !!opts.caseSensitive,
      entireWord: !!opts.entireWord,
      highlightAll: true,
      findPrevious: !!opts.previous,
      matchDiacritics: false,
    });

  const closeSearch = () => {
    setSearchOpen(false);
    busRef.current!.dispatch("findbarclose", { source: null });
  };

  return (
    <div
      className="app"
      data-focus={focus || undefined}
      data-sidebar={(sidebar && !!doc && !focus) || undefined}
      data-night={night || undefined}
      data-content-edit={contentEditing || undefined}
      data-tool={tool}
    >
      <header className="topbar" ref={headerRef}>
        <div className="topbar-left">
          <button className="brand" onClick={doc ? () => tabs.library() : undefined} title={doc ? "Back to library" : "Folio"}>
            <span className="brand-mark" />
            <span className="brand-name">Folio</span>
          </button>
          {doc && (
            <>
              <IconButton icon="sidebar" label={keys("Toggle sidebar (Ctrl+\\)")} active={sidebar} onClick={() => setSidebar(!sidebar)} />
              {!multipleTabs && <div className="doc-title" title={doc.name}>
                <span>{doc.name}</span>
                {unsaved && <span className="dirty-dot" role="img" aria-label="Unsaved changes" title="Unsaved changes" />}
              </div>}
            </>
          )}
        </div>

        <div className="topbar-center">
          {doc && (
            <div className="segmented" role="toolbar" aria-label="Tools">
              {TOOLS.map((t) => (
                <button
                  key={t.id}
                  className={!contentEditing && tool === t.id ? "on" : ""}
                  aria-label={t.label}
                  aria-pressed={!contentEditing && tool === t.id}
                  onClick={() => chooseTool(t.id)}
                  title={`${t.id === "highlight" ? "Automatically highlight selected text" : t.label} (${t.key})`}
                >
                  <Icon name={t.icon} size={16} />
                  <span>{t.label}</span>
                </button>
              ))}
              <button onClick={startSigning} aria-label="Sign" title="Add a signature"><Icon name="signature" size={16} /><span>Sign</span></button>
              <button className={contentEditing ? "on edit-mode-toggle" : "edit-mode-toggle"} aria-pressed={contentEditing}
                aria-label={contentEditing ? "Exit edit mode" : "Edit mode"} title={contentEditing ? "Exit edit mode" : "Edit existing text and images"}
                onClick={() => contentEditing ? leaveEditMode() : chooseTool("select", true)}>
                <Icon name={contentEditing ? "check" : "edit"} size={16} /><span>{contentEditing ? "Exit edit mode" : "Edit mode"}</span>
              </button>
              {(doc.sourceId || editorDirty) && <button aria-label="Save copy" title="Save edited copy" onClick={download}><Icon name="download" size={16} /><span>Save copy</span></button>}
            </div>
          )}
        </div>

        <div className="topbar-right">
          {doc && (
            <>
              <IconButton icon="sparkle" label="Ask AI" active={aiOpen} onClick={() => { setHistoryOpen(false); setAiOpen(!aiOpen); }} />
              <IconButton icon="search" label={keys("Search (Ctrl+F)")} active={searchOpen} onClick={() => (searchOpen ? closeSearch() : setSearchOpen(true))} />
            </>
          )}
          <MoreMenu
            items={[
              ...(doc
                ? ([
                    ...(navigation.back ? [{ kind: "action", icon: "chevronLeft", label: "Back", shortcut: keys("Alt+←"), onSelect: () => navigationRef.current?.back() }] : []),
                    ...(navigation.forward ? [{ kind: "action", icon: "chevronRight", label: "Forward", shortcut: keys("Alt+→"), onSelect: () => navigationRef.current?.forward() }] : []),
                    ...(navigation.back || navigation.forward ? [{ kind: "divider" }] : []),
                    { kind: "action", icon: "plus", label: "New tab", shortcut: keys("Ctrl T"), onSelect: () => tabs.library() },
                    { kind: "action", icon: "close", label: "Close document", shortcut: keys("Ctrl W"), onSelect: () => void tabs.close(tabKey) },
                    { kind: "divider" },
                    { kind: "action", icon: "download", label: "Save a copy", shortcut: keys("Ctrl S"), onSelect: download },
                    { kind: "action", icon: "history", label: "Version history", onSelect: () => { setAiOpen(false); setHistoryOpen(true); } },
                    { kind: "action", icon: "print", label: "Print", shortcut: keys("Ctrl P"), onSelect: print },
                    { kind: "action", icon: "info", label: "Document info", onSelect: () => setDialog("info") },
                    { kind: "divider" },
                    { kind: "toggle", icon: "moon", label: "PDF dark mode", checked: night, onSelect: () => setNightTo(!night) },
                    { kind: "action", icon: "expand", label: "Focus mode", shortcut: "F", onSelect: toggleFocus },
                  ] as MoreItem[])
                : []),
              { kind: "action", icon: "settings", label: "Settings", shortcut: keys("Ctrl ,"), onSelect: () => setDialog("settings") },
              { kind: "action", icon: "keyboard", label: "Keyboard shortcuts", shortcut: "?", onSelect: () => setDialog("shortcuts") },
              { kind: "action", icon: "info", label: "About Folio", onSelect: () => setDialog("about") },
            ]}
          />
        </div>
      </header>
      {!focus && multipleTabs && <DocumentTabBar tabs={tabs} tabKey={tabKey} />}

      <div className="body" id="reader-panel" role={!focus && multipleTabs ? "tabpanel" : undefined}
        aria-labelledby={!focus && multipleTabs ? "selected-document-tab" : undefined}>
        {doc && (
          <Sidebar
            key={doc.pdf.loadingTask.docId}
            pdf={doc.pdf}
            page={page}
            rotation={rotation}
            onSelectPage={goTo}
            bookmarkedPages={markup.bookmarkedPages}
            onToggleBookmark={n => markupRef.current!.toggleBookmark(n)}
            order={pageDraft?.order ?? null}
            onOrder={order => setPageDraft(order ? { order } : null)}
            onSaveOrder={download}
              onExportPages={exportPages}
            onOpenCopy={pageDraft?.copy ? () => {
              tabs.open([{ kind: "bytes", name: copyName, data: pageDraft.copy!, password: doc.password }]);
            } : undefined}
            onSelectDest={(dest) => linkRef.current!.goToDestination(dest as string)}
            tab={sidebarTab}
            onTab={setSidebarTab}
            noteCount={markup.marks.filter((m) => m.type === "highlight").length}
            notes={
              <NotesPanel
                key={notesGeneration}
                state={markup}
                onExport={exportNotes}
                persisted={persistRef.current}
                onReveal={(id) => markupRef.current!.reveal(id)}
                onNote={(id, note) => markupRef.current!.setNote(id, note)}
                onRemove={(id) => markupRef.current!.remove(id)}
                onDocNote={(note) => markupRef.current!.setDocNote(note)}
                onAsk={(mark: Mark) => askAbout(mark.text, mark.page + 1)}
              />
            }
          />
        )}

        <main className="stage">
          <div ref={containerRef} className="viewer-container" tabIndex={-1}
            onInput={event => {
              // Form values live in the PDF until a copy is saved. Folio only auto-saves its own markup.
              if ((event.target as HTMLElement).closest(".annotationLayer :is(input, textarea, select)")) markDirty();
            }}>
            <div ref={innerRef} className="pdfViewer" />
          </div>
          {contentEditing && doc && <ContentEditor key={doc.pdf.loadingTask.docId} ref={contentRef} nextSelection={nextSelection} pdf={doc.pdf} viewer={viewerRef.current!} bus={busRef.current!}
            onDirty={setEditorDirty} snapshot={async () => {
              if (doc.password || await doc.pdf.getPermissions() || (await doc.pdf.getSignatures())?.length) throw new Error("Editing requires an unsigned, unencrypted PDF.");
              const copy = await currentBytes(doc, []);
              if (!copy || docRef.current !== doc) throw new Error("The document changed. Try again.");
              return copy.bytes;
            }} apply={async (bytes, change) => docRef.current === doc && !!await openBytes(doc.name, bytes, undefined, opening.current!.signal, undefined, { doc, change })} />}
          {doc && <div className="page-text-status" role="status">{textStatus}</div>}

          {doc && searchOpen && (
            <>
              <SearchBar result={find} onFind={runFind} onClose={closeSearch} />
              <SearchResults pdf={doc.pdf} controller={findRef.current!} bus={busRef.current!} onSelect={goTo} />
            </>
          )}

          {doc && !contentEditing && PALETTE[tool] && (
            <div className="color-pill" role="toolbar" aria-label="Annotation options">
              {tool === "ink" && (
                <select className="draw-tool" aria-label="Drawing tool" value={drawTool}
                  onChange={e => setDrawTool(e.target.value as typeof drawTool)}>
                  {["Pen", "Rectangle", "Oval", "Arrow"].map(name => <option key={name}>{name}</option>)}
                </select>
              )}
              {tool === "freetext" && <>
                <select className="draw-tool" aria-label="Font" title="Font" value={textStyle.font}
                  onChange={e => chooseTextStyle({ font: e.target.value as TextFont })}>
                  {TEXT_FONTS.map(f => <option key={f.id} value={f.id} style={f.css ? { fontFamily: f.css } : undefined}>{f.name}</option>)}
                </select>
                <select className="draw-tool" aria-label="Font size" title="Font size" value={textStyle.size}
                  onChange={e => chooseTextStyle({ size: Number(e.target.value) })}>
                  {(TEXT_SIZES.includes(textStyle.size) ? TEXT_SIZES : [...TEXT_SIZES, textStyle.size].sort((a, b) => a - b))
                    .map(n => <option key={n} value={n}>{n} pt</option>)}
                </select>
              </>}
              {tool === "ink" && drawTool === "Pen" && (
                <select className="draw-tool" aria-label="Pen width" title="Pen width" defaultValue={markupRef.current?.penWidth ?? 1}
                  onChange={e => markupRef.current?.setPenWidth(Number(e.target.value))}>
                  <option value={1}>Thin</option>
                  <option value={3}>Medium</option>
                  <option value={6}>Thick</option>
                </select>
              )}
              <div className="color-options" role="radiogroup" aria-label="Color">
                {PALETTE[tool]!.map((c) => (
                  <button
                    key={c.hex}
                    className={colors[tool] === c.hex ? "on" : ""}
                    style={{ "--swatch": c.hex } as React.CSSProperties}
                    onClick={() => chooseColor(c.hex)}
                    title={c.name}
                    role="radio"
                    aria-checked={colors[tool] === c.hex}
                  />
                ))}
              </div>
              <span className="divider" />
              <button className="pill-btn" onClick={() => busRef.current!.dispatch("editingaction", { source: null, name: "undo" })} title={keys("Undo (Ctrl+Z)")}>
                <Icon name="undo" size={16} />
              </button>
              <button className="pill-done" onClick={() => chooseTool("select")}>Done</button>
            </div>
          )}

          {doc && (
            <div className="pill" role="toolbar" aria-label="Navigation">
              <button className="pill-btn" onClick={() => goTo(page - 1)} disabled={page <= 1} title="Previous page">
                <Icon name="chevronUp" size={17} />
              </button>
              <form
                className="page-field"
                onSubmit={(e) => {
                  e.preventDefault();
                  const n = parseInt(pageInput, 10);
                  if (Number.isFinite(n)) goTo(n);
                  else setPageInput(String(page));
                  (document.activeElement as HTMLElement)?.blur();
                }}
              >
                <input
                  value={pageInput}
                  onChange={(e) => setPageInput(e.target.value.replace(/\D/g, ""))}
                  onFocus={(e) => e.target.select()}
                  onBlur={() => setPageInput(String(pageRef.current))}
                  style={{ width: `calc(${Math.max(2, String(numPages).length)}ch + 22px)` }}
                  aria-label="Page number"
                />
                <span className="page-total">of {numPages}</span>
              </form>
              <button className="pill-btn" onClick={() => goTo(page + 1)} disabled={page >= numPages} title="Next page">
                <Icon name="chevronDown" size={17} />
              </button>
              <span className="divider" />
              <button className="pill-btn" onClick={() => zoom(-1)} title={keys("Zoom out (Ctrl −)")}>
                <Icon name="minus" size={17} />
              </button>
              <ZoomMenu scale={scale} onPick={setScaleValue} spread={spread} onToggleSpread={toggleSpread} />
              <button className="pill-btn" onClick={() => zoom(1)} title={keys("Zoom in (Ctrl +)")}>
                <Icon name="plus" size={17} />
              </button>
              <span className="divider wide-only" />
              <button className="pill-btn wide-only" onClick={rotate} title="Rotate (R)">
                <Icon name="rotate" size={17} />
              </button>
            </div>
          )}

          {selection && !noteBox && !highlightMenu && (
            <SelectionPopover
              key={`${selection.page}:${selection.top}:${selection.text.slice(0, 40)}`}
              info={selection}
              stageWidth={stageWidth}
              colors={HIGHLIGHTS}
              current={highlightColor}
              onHighlight={(hex, note) => void quickHighlight(hex, note)}
              onAsk={() => { setPinned(null); askAbout(selection.text, selection.page); }}
              onPin={() => setPinned(selection)}
              onDismiss={dismissSelection}
            />
          )}
          {imageMenu && (
            <ImageMenu
              x={imageMenu.x}
              y={imageMenu.y}
              stageWidth={stageWidth}
              href={imageMenu.target.href}
              onView={() => { setImageView(imageMenu.target); setImageMenu(null); }}
              onOpen={() => { window.open(imageMenu.target.href!, "_blank", "noopener"); setImageMenu(null); }}
            />
          )}
          {highlightMenu && menuMark && !noteBox && (
            <HighlightMenu
              rect={highlightMenu.rect}
              stageWidth={stageWidth}
              colors={HIGHLIGHTS}
              current={menuMark.color}
              hasNote={!!menuMark.note}
              onRecolor={(hex) => markupRef.current!.recolor(menuMark.id, hex)}
              onNote={() => {
                const r = highlightMenu.rect;
                setNoteBox({ markId: menuMark.id, anchor: { x: (r.left + r.right) / 2, top: r.top, bottom: r.bottom }, quote: menuMark.text, initial: menuMark.note });
                setHighlightMenu(null);
              }}
              onAsk={() => { askAbout(menuMark.text, menuMark.page + 1); setHighlightMenu(null); }}
              onRemove={() => removeHighlight(menuMark.id)}
            />
          )}
          {noteBox && (
            <NoteBox
              key={noteBox.markId}
              anchor={noteBox.anchor}
              stageWidth={stageWidth}
              quote={noteBox.quote}
              initial={noteBox.initial}
              onSave={(note) => markupRef.current!.setNote(noteBox.markId, note)}
              onClose={() => setNoteBox(null)}
            />
          )}

          {focus && (
            <button className="exit-focus" onClick={toggleFocus} title="Exit focus mode">
              <kbd>Esc</kbd> to exit focus mode
            </button>
          )}
        </main>

        {doc && historyOpen && !focus && <VersionHistoryPanel key={doc.sourceId ?? doc.id} id={doc.sourceId ?? doc.id}
          allowed={doc.historyAllowed} active={active} pending={changesForCopy(doc, dirty || markup.marks.length > 0, pageOrder)}
          canOpen={id => !tabs.all.some(t => t.contentId === id && t.dirty)}
          onOpen={(v, blob) => tabs.open([{ kind: "blob", name: v.name, blob }])}
          onDownload={(v, blob) => saveBlob(blob, v.name.replace(/\.pdf$/i, "") + ` (v${v.number}).pdf`)}
          onClose={() => setHistoryOpen(false)} />}
        {doc && aiOpen && !focus && (
          <AiPanel
            key={`${doc.sourceId ?? doc.id}:${doc.password ? 0 : chatGeneration}`}
            pdf={doc.pdf}
            docId={doc.sourceId ?? doc.id}
            persist={!doc.password}
            docName={doc.name}
            context={aiContext}
            onClearContext={() => setAiContext(null)}
            onGoToPage={goTo}
            onSettings={() => setDialog("ai-settings")}
            onClose={() => setAiOpen(false)}
          />
        )}
      </div>

      {!doc && <Welcome active={active} onPick={pickFile} onOpenUrl={() => openUrlDialog()} onOpenRecent={file => tabs.open([{ kind: "blob", name: file.name, blob: file.data, recentId: file.id }])}
        onSample={() => tabs.open([{ kind: "sample", name: "Folio sample.pdf" }])} loading={loading} />}

      {loading && doc && <div className="loading-bar" />}

      {dragging && (
        <div className="drop-overlay">
          <div className="drop-card">
            <Icon name="upload" size={30} />
            <p>Drop a PDF or link to open</p>
          </div>
        </div>
      )}

      {pending && (
        <PasswordDialog
          name={pending.name}
          error={pending.error}
          onSubmit={(pw) => openBytes(pending.name, pending.data, pw)}
          onCancel={() => { setPending(null); tabs.failed(tabKey); }}
        />
      )}
      {dialog === "info" && doc && <InfoDialog key={doc.pdf.loadingTask.docId} pdf={doc.pdf} name={doc.name} size={doc.size} onClose={() => setDialog(null)} />}
      {printCopy && <PrintDialog pdf={printCopy.pdf} currentPage={printCopy.page} name={printCopy.name}
        onPrint={() => { printPdfBytes(printCopy.bytes); setPrintCopy(null); }} onClose={() => setPrintCopy(null)} />}
      {dialog === "shortcuts" && <ShortcutsDialog onClose={() => setDialog(null)} />}
      {dialog === "about" && <AboutDialog onClose={() => setDialog(null)} />}
      {(dialog === "settings" || dialog === "ai-settings") && <SettingsDialog initialSection={dialog === "ai-settings" ? "Ask AI" : "General"}
        theme={theme} onTheme={setThemeTo} night={night} onNight={setNightTo}
        scanLanguages={scanLanguages} onScanLanguages={next => {
          setScanLanguages(next);
          try { localStorage.setItem("folio:scan-language", next.join("+")); }
          catch { flash("Scan languages changed, but couldn't save the preference."); }
        }}
        wheelPaging={wheelPaging} onWheelPaging={(next) => {
          setWheelPaging(next);
          try { localStorage.setItem("folio:wheel-paging", next ? "1" : "0"); }
          catch { flash("Page scrolling changed, but couldn't save the preference."); }
        }}
        onOpenChangelog={() => setDialog("changelog")} onClose={() => setDialog(null)} />}
      {/* Closing goes back to Settings, where the changelog is opened from. */}
      {dialog === "changelog" && <ChangelogDialog onClose={() => setDialog("settings")} />}
      {dialog === "upgrade" && <ChangelogDialog upgradeNotes={upgradeNotes} onClose={() => setDialog(null)} />}
      {doc && imageView && <ImageViewer pdf={doc.pdf} target={imageView} onClose={() => setImageView(null)} />}
      {dialog === "url" && <OpenUrlDialog key={linkDraft} initial={linkDraft} onClose={() => setDialog(null)} onSubmit={(url) => { setDialog(null); tabs.open([{ kind: "url", name: new URL(url).hostname, url }]); }} />}
      {doc && signing && <SignatureDialog key={doc.id + ":" + page} pdf={doc.pdf} pageNumber={page} targets={signatureTargets}
        onClose={() => setSigning(false)} onApprove={async (signature, rect) => {
          await placeSignature(viewerRef.current!, doc.pdf, page, signature, rect);
          markDirty(); setSignatureDismissed(true); flash(keys("Signature added. Ctrl+Z to undo; Save copy to export."));
        }} />}
      {doc && !contentEditing && !signing && signatureTargets.length > 0 && !signatureDismissed && <div className="signature-suggestion" role="status">
        <span>{signatureTargets[0].kind === "field" ? "Signature field detected" : "Possible signature area"} on page {page}</span>
        <button className="btn primary" onClick={startSigning}>Review signature</button>
        <button className="pill-btn" aria-label="Dismiss signature suggestion" onClick={() => setSignatureDismissed(true)}><Icon name="close" size={14} /></button>
      </div>}

      {updateVersion && (
        <div className="update-card" role="status">
          <div>
            <strong>Folio {updateVersion} is ready</strong>
            <span>Restart to update. Installation on close follows your auto-update setting.</span>
          </div>
          <button
            className="btn primary"
            onClick={() => {
              window.folioDesktop?.installUpdate();
            }}
          >
            Restart
          </button>
          <button className="pill-btn" onClick={() => setUpdateVersion(null)} aria-label="Later" title="Later">
            <Icon name="close" size={15} />
          </button>
        </div>
      )}

      {active && doc && backHint && !toast && !error && !dialog && !pending && !signing && !imageView && !printCopy &&
        <div role="status"><button type="button" ref={backHintButton} className="toast link-return" onBlur={() => setBackHint(null)}
          onClick={() => navigationRef.current?.back()}><Icon name="chevronLeft" size={15} />{backHint}</button></div>}
      {(toast || error) && (
        <div className={`toast ${error ? "error" : ""}`} onClick={() => setError(null)}>
          {error ?? toast}
        </div>
      )}
    </div>
  );
}

function IconButton({ icon, label, onClick, active }: { icon: IconName; label: string; onClick: () => void; active?: boolean }) {
  return (
    <button className={`icon-btn ${active ? "on" : ""}`} onClick={onClick} title={label} aria-label={label} aria-pressed={active}>
      <Icon name={icon} />
    </button>
  );
}
