import { useEffect, useImperativeHandle, useLayoutEffect, useRef, useState, type CSSProperties, type Ref, type RefObject } from "react";
import { createPortal } from "react-dom";
import type { PageViewport } from "pdfjs-dist";
import type { PDFDocumentProxy } from "./pdf";
import type { EventBus, PDFViewer } from "./viewer";
import type { Change, EditorReply, EditorRequest, EditTarget } from "./contentEditor.worker";
import type { PdfChange } from "./versionHistory";

export interface ContentEditorHandle { finish: () => Promise<boolean>; deselect: () => boolean }
interface Surface { page: number; host: HTMLDivElement; viewport: PageViewport }
interface Selected { page: number; target: EditTarget; paper: string }
export interface NextSelection { from: PDFDocumentProxy; page: number; target: EditTarget }

function rectangle(t: EditTarget, v: PageViewport) {
  const [x1, y1] = v.convertToViewportPoint(t.box[0], t.box[1]);
  const [x2, y2] = v.convertToViewportPoint(t.box[2], t.box[3]);
  return { left: Math.min(x1,x2), top: Math.min(y1,y2), width: Math.abs(x2-x1), height: Math.abs(y2-y1) };
}

function paperColor(surface: Surface, target: EditTarget) {
  const canvas = surface.host.parentElement?.querySelector<HTMLCanvasElement>(".canvasWrapper canvas");
  const context = canvas?.getContext("2d");
  if (!canvas || !context) return "white";
  const { left, top } = rectangle(target, surface.viewport);
  const x = Math.max(0, Math.min(canvas.width - 1, Math.floor(left / surface.viewport.width * canvas.width) - 1));
  const y = Math.max(0, Math.min(canvas.height - 1, Math.floor(top / surface.viewport.height * canvas.height) - 1));
  const [r,g,b] = context.getImageData(x,y,1,1).data;
  return 'rgb(' + [r,g,b].join(" ") + ')';
}

function inputStyle(t: EditTarget, v: PageViewport, paper: string): CSSProperties {
  const s = t.style!, [x,y] = s.origin;
  const [ox,oy] = v.convertToViewportPoint(x,y);
  const [ex,ey] = v.convertToViewportPoint(x+s.axis[0], y+s.axis[1]);
  const [ux,uy] = v.convertToViewportPoint(x,y+s.size);
  const size = Math.max(1, Math.hypot(ux-ox,uy-oy));
  const angle = Math.atan2(ey-oy,ex-ox), box = rectangle(t,v);
  const width = Math.abs(Math.cos(angle))*box.width + Math.abs(Math.sin(angle))*box.height;
  return { left: ox, top: oy, minWidth: Math.max(width+4, size), fontSize: size,
    fontFamily: s.flags & 1 ? "monospace" : s.flags & 2 ? "serif" : "sans-serif",
    fontWeight: s.bold ? 700 : 400, fontStyle: s.flags & 64 ? "italic" : "normal", color: s.color, background: paper,
    transform: 'rotate(' + angle + 'rad) translateY(-0.88em)' };
}

export function ContentEditor({ ref, nextSelection, pdf, viewer, bus, snapshot, apply, onDirty }: {
  ref: Ref<ContentEditorHandle>; pdf: PDFDocumentProxy; viewer: PDFViewer; bus: EventBus;
  nextSelection: RefObject<NextSelection | null>;
  snapshot: () => Promise<Uint8Array>; apply: (bytes: Uint8Array, change: PdfChange) => Promise<boolean>; onDirty: (dirty: boolean) => void;
}) {
  const [surfaces, setSurfaces] = useState<Surface[]>([]);
  const [targets, setTargets] = useState<Map<number, EditTarget[]>>(() => new Map());
  const [selected, setSelected] = useState<Selected | null>(null);
  const [face, setFace] = useState<{ target: EditTarget; family: string } | null>(null);
  const [value, setValue] = useState("");
  const [customFont, setCustomFont] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const input = useRef<HTMLInputElement>(null);
  const file = useRef<HTMLInputElement>(null);
  const worker = useRef<Worker | null>(null);
  const waiter = useRef<{ resolve: (data: Uint8Array) => void; reject: (reason: Error) => void } | null>(null);
  const pending = useRef<Promise<boolean> | null>(null);
  const selectedRef = useRef(selected);
  selectedRef.current = selected;
  const modified = selected?.target.kind === "text" && (value !== selected.target.label || customFont);
  useEffect(() => { onDirty(!!modified || busy); return () => onDirty(false); }, [modified, busy, onDirty]);

  useEffect(() => {
    const w = new Worker(new URL("./contentEditor.worker.ts", import.meta.url), { type: "module" });
    worker.current = w;
    const hosts = new Map<number, HTMLDivElement>(), requested = new Set<number>();
    let initialized = false, frame = 0;
    const refresh = () => {
      if (viewer.pdfDocument !== pdf || !viewer.pagesCount) return;
      const visible = viewer._getVisiblePages() as { views?: { id: number }[] };
      const pages = new Set(visible.views?.map(v => v.id));
      if (selectedRef.current) pages.add(selectedRef.current.page);
      const next: Surface[] = [];
      for (const page of pages) {
        const view = viewer.getPageView(page-1);
        if (!view?.div.isConnected || !view.pdfPage || !view.div.querySelector(".canvasWrapper")) continue;
        let host = hosts.get(page);
        if (!host) { host = document.createElement("div"); host.className = "folio-content-layer"; host.tabIndex = -1; hosts.set(page,host); }
        if (host.parentElement !== view.div) {
          view.div.append(host);
          if (selectedRef.current?.page === page && input.current && host.contains(input.current)) {
            input.current.focus({ preventScroll: true });
          }
        }
        next.push({ page, host, viewport: view.viewport });
        if (initialized && !requested.has(page)) { requested.add(page); w.postMessage({ page } satisfies EditorRequest); }
      }
      for (const [page,host] of hosts) if (!pages.has(page)) { host.remove(); hosts.delete(page); }
      setSurfaces(old => old.length === next.length && next.every((s,i) => s.host === old[i].host && s.viewport === old[i].viewport) ? old : next);
    };
    const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(refresh); };
    const events = ["pagesinit", "updateviewarea", "pagerendered", "scalechanging", "rotationchanging"];
    for (const event of events) bus.on(event, schedule);
    const fail = (reason: Error) => { setError(reason.message); waiter.current?.reject(reason); waiter.current = null; };
    w.onmessage = ({ data }: MessageEvent<EditorReply>) => {
      if ("error" in data) {
        if (data.page === undefined) fail(new Error(data.error));
        else setError(data.error);
      } else if ("bytes" in data) { waiter.current?.resolve(data.bytes); waiter.current = null; }
      else setTargets(old => new Map(old).set(data.page, data.targets));
    };
    w.onerror = () => fail(new Error("The PDF editor stopped. Exit and re-enter Edit mode to try again."));
    void pdf.getData().then(source => {
      if (worker.current !== w) return;
      w.postMessage({ source } satisfies EditorRequest); initialized = true; schedule();
    }).catch(() => { if (worker.current === w) fail(new Error("Couldn't read this PDF for editing.")); });
    schedule();
    return () => {
      for (const event of events) bus.off(event, schedule);
      cancelAnimationFrame(frame);
      w.terminate(); worker.current = null;
      waiter.current?.reject(new Error("The document changed.")); waiter.current = null;
      for (const host of hosts.values()) host.remove();
    };
  }, [pdf, viewer, bus]);

  useLayoutEffect(() => {
    if (selected?.target.kind === "text") { input.current?.focus({ preventScroll: true }); input.current?.select(); }
  }, [selected]);

  const commit = (change: Change): Promise<boolean> => {
    if (pending.current) return pending.current;
    const w = worker.current;
    if (!w || !selected) return Promise.resolve(false);
    setBusy(true); setError("");
    const transaction = (async () => {
      try {
        const data = await snapshot();
        if (worker.current !== w) return false;
        const result = await new Promise<Uint8Array>((resolve,reject) => {
          waiter.current = { resolve, reject };
          w.postMessage({ page: selected.page, change, snapshot: data } satisfies EditorRequest);
        });
        const detail: PdfChange = change.kind === "text"
          ? { kind: "text", page: selected.page, before: selected.target.label, after: change.text }
          : { kind: "image", page: selected.page, fileName: change.file.name };
        return worker.current === w && await apply(result, detail);
      } catch (e) {
        if (worker.current === w) { setError(e instanceof Error ? e.message : "Couldn't apply the edit."); input.current?.focus({ preventScroll: true }); }
        return false;
      } finally { pending.current = null; if (worker.current === w) setBusy(false); }
    })();
    pending.current = transaction; return transaction;
  };
  const finish = () => {
    if (pending.current) return pending.current;
    if (selected && modified) return commit({ kind: "text", index: selected.target.index, text: value,
      ...(customFont ? { font: "easyjet" } : {}) });
    setSelected(null); return Promise.resolve(true);
  };
  useImperativeHandle(ref, () => ({ finish, deselect: () => {
    if (!selected || pending.current) return false;
    selectedRef.current = null; setSelected(null); setError(""); return true;
  } }));

  const pick = (s: Surface, target: EditTarget) => {
    if (pending.current) return;
    setError(""); setCustomFont(false); setValue(target.label); setSelected({ page: s.page, target, paper: target.style ? paperColor(s,target) : "" });
    if (!target.style) return;
    // Use the same embedded or system font as the page renderer while the caret is visible.
    const origin = target.style.origin;
    void pdf.getPage(s.page).then(async page => {
      const content = await page.getTextContent();
      let distance = Infinity, name = "";
      for (const item of content.items) {
        if (!("str" in item) || !item.str.trim()) continue;
        const delta = Math.hypot(item.transform[4] - origin[0], item.transform[5] - origin[1]);
        if (delta < distance) { distance = delta; name = item.fontName; }
      }
      if (distance > target.style!.size || !page.commonObjs.has(name)) return;
      const font = page.commonObjs.get(name);
      const family = font.systemFontInfo?.css || `"${font.loadedName}", ${font.fallbackName || content.styles[name].fontFamily}`;
      if (selectedRef.current?.target === target) setFace({ target, family });
    }).catch(() => {});
  };
  useEffect(() => {
    const queued = nextSelection.current;
    if (!queued || queued.from === pdf) return;
    const surface = surfaces.find(s => s.page === queued.page), items = targets.get(queued.page);
    if (!surface || !items) return;
    nextSelection.current = null;
    const next = items.find(t => t.kind === queued.target.kind && t.label === queued.target.label
      && t.box.every((n,i) => Math.abs(n - queued.target.box[i]) < 0.1));
    if (next) pick(surface, next);
  }, [pdf, surfaces, targets, nextSelection]);

  useEffect(() => {
    if (!selected) return;
    const outside = (e: PointerEvent) => {
      if (!pending.current && e.target instanceof Element && !e.target.closest(".folio-inline-text, .folio-font-actions, .folio-image-actions, .folio-content-target")) void finish();
    };
    document.addEventListener("pointerdown", outside);
    return () => document.removeEventListener("pointerdown", outside);
  }, [selected, value, customFont]);

  return <>
    {surfaces.map(s => createPortal(<>
      {(targets.get(s.page) ?? []).map(t => <button key={t.index} type="button" className="folio-content-target"
        data-kind={t.kind} aria-label={t.label} aria-pressed={selected?.page === s.page && selected.target === t}
        disabled={busy} style={rectangle(t,s.viewport)} onClick={() => pick(s,t)}
        onPointerDown={e => {
          if (!modified || e.button !== 0) return;
          e.preventDefault();
          nextSelection.current = { from: pdf, page: s.page, target: t };
          void finish().then(ok => { if (!ok) nextSelection.current = null; });
        }} />)}
      {selected?.page === s.page && (selected.target.kind === "text" ? <>
        <div className="folio-old-text" style={{ ...rectangle(selected.target,s.viewport), background: selected.paper }} />
        <input ref={input} className="folio-inline-text" aria-label="Edit text on page" maxLength={32767}
          style={{ ...inputStyle(selected.target,s.viewport,selected.paper),
            ...(face?.target === selected.target ? { fontFamily: face.family } : {}),
            ...(customFont ? { fontFamily: '"easyJet Rounded Book"', fontWeight: 400, fontStyle: "normal" } : {}) }} value={value} readOnly={busy}
          onChange={e => { setValue(e.target.value); setError(""); }}
          onKeyDown={e => {
            if (e.nativeEvent.isComposing) return;
            if (e.key === "Tab" && !pending.current) void finish();
            if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); void finish(); }
            if (e.key === "Escape" && !pending.current) { e.stopPropagation(); selectedRef.current = null; setSelected(null); setError(""); }
          }} />
      </> : <div className="folio-image-actions" style={{ left: rectangle(selected.target,s.viewport).left, top: rectangle(selected.target,s.viewport).top }}>
        <button className="btn" disabled={busy} onClick={() => file.current?.click()}>Replace image…</button>
        <input hidden ref={file} type="file" aria-label="Replacement image" accept="image/png,image/jpeg" onChange={e => {
          const picked = e.target.files?.[0]; e.target.value = "";
          if (picked) void commit({ kind: "image", index: selected.target.index, file: picked });
        }} />
      </div>)}
    </>,s.host,s.page))}
    {selected?.target.kind === "text" && <div className="folio-font-actions">
      <label>Font <select aria-label="Text font" disabled={busy} value={customFont ? "easyjet" : "original"}
        onChange={e => setCustomFont(e.target.value === "easyjet")}>
        <option value="original">Original font</option>
        <option value="easyjet">easyJet Rounded Book</option>
      </select></label>
      <button className="btn" disabled={busy} onClick={() => void finish()}>Apply</button>
    </div>}
    <div className="folio-edit-status" role={error ? "alert" : "status"}>{error || (busy ? "Applying edit…" :
      surfaces.length > 0 && surfaces.every(s => targets.get(s.page)?.length === 0) ? "No editable text or images on this page." : "")}</div>
  </>;
}
