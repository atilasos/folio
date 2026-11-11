// End-to-end smoke test, run in CI on Windows, macOS, and Linux (and locally with `npm run smoke`):
// launches the built app in Electron, opens the sample document, and checks that pages render, text is
// selectable and highlightable, text boxes survive reopening, and the window chrome matches the platform.
// Run `npm run build` first.
import { _electron as electron } from "playwright-core";
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import http from "node:http";
import { existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import assert from "node:assert/strict";
import { electronPath } from "./electron-path.mjs";
import { readerPage } from "./reader-page.mjs";
import { checkReaderTools } from "./smoke-reader-tools.mjs";
import { checkOpenRouter } from "./smoke-openrouter.mjs";

const root = path.resolve(import.meta.dirname, "..");
const require = createRequire(import.meta.url);
const shots = path.join(root, "release", "smoke");
mkdirSync(shots, { recursive: true });

// The window stays hidden so a local run doesn't pop up over other work; FOLIO_SMOKE_SHOW=1 shows it.
const env = { ...process.env, FOLIO_TEST_USER_DATA: mkdtempSync(path.join(tmpdir(), "folio-smoke-")),
  FOLIO_TEST_HIDDEN: process.env.FOLIO_SMOKE_SHOW === "1" ? "0" : "1" };
delete env.ELECTRON_RUN_AS_NODE;
const executablePath = electronPath(env);
// Containers often run as root, where Chromium's sandbox can't start. The made-up names stand in for web addresses
// that lead to this computer (see the local network checks below).
const args = [root, "--mute-audio", "--host-resolver-rules=MAP folio-local.test 127.0.0.1, MAP folio-rebind.test 127.0.0.1, MAP folio-pinned.test 127.0.0.2",
  ...(process.platform === "linux" && process.getuid?.() === 0 ? ["--no-sandbox"] : [])];

// The smallest .docx Word opens: one paragraph of `text`.
async function wordDocument(text) {
  const JSZip = require("jszip");
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `${xml}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file("_rels/.rels", `${xml}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/document.xml", `${xml}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
    + `<w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer" });
}

const app = await electron.launch({ executablePath, args, env, timeout: 60_000 });
const failures = [];
// A hidden run has to stay out of the way on macOS too: no Dock icon, never the app in front, no window on a display.
async function checkStaysHidden() {
  if (process.platform !== "darwin" || env.FOLIO_TEST_HIDDEN !== "1") return;
  const seen = await app.evaluate(({ app, BrowserWindow, screen }) => {
    const displays = screen.getAllDisplays().map((display) => display.bounds);
    const onDisplay = ({ x, y, width, height }) => displays.some((d) => x < d.x + d.width && x + width > d.x && y < d.y + d.height && y + height > d.y);
    return { dockIcon: app.dock.isVisible(), inFront: app.isActive(),
      onScreen: BrowserWindow.getAllWindows().filter((w) => w.isVisible() && onDisplay(w.getBounds())).length };
  });
  assert.deepEqual(seen, { dockIcon: false, inFront: false, onScreen: 0 }, "the hidden test run showed up on macOS");
  // The Dock makes room for an ordinary app as it starts, before Folio can say otherwise, so the run has to have
  // started from an app that macOS treats as an agent (see electron-path.mjs).
  const plist = path.join(await app.evaluate(() => process.execPath), "..", "..", "Info.plist");
  assert.equal(spawnSync("/usr/bin/plutil", ["-extract", "LSUIElement", "raw", plist], { encoding: "utf8" }).stdout.trim(), "true",
    "the hidden test run started as an ordinary app, which the Dock makes room for");
}
try {
  const page = readerPage(await app.firstWindow());
  assert.ok(await app.evaluate(({ app }) => app.commandLine.hasSwitch("mute-audio")));
  page.on("pageerror", (error) => {
    console.error(error.stack ?? String(error));
    failures.push(String(error));
  });
  await page.getByRole("button", { name: /Try a sample/ }).waitFor({ timeout: 30_000 });

  const platformClass = await page.evaluate(() => [...document.documentElement.classList].find((c) => c.startsWith("platform-")));
  assert.equal(platformClass, `platform-${process.platform}`);
  console.log("ok  platform:", platformClass);
  await checkStaysHidden();

  const whatsNew = page.getByRole("dialog", { name: "What's new" });
  const installed = await app.evaluate(({ app }) => app.getVersion());
  assert.equal(await whatsNew.count(), 0, "a fresh install opens straight to the library");
  assert.equal(await page.evaluate(() => localStorage.getItem("folio:last-opened-version")), installed);

  // An older profile has a sidebar preference but has never recorded its Folio version.
  await page.evaluate(() => localStorage.removeItem("folio:last-opened-version"));
  await page.reload();
  await whatsNew.getByRole("heading", { name: "Folio " + installed, exact: true }).waitFor();
  assert.equal(await whatsNew.locator("section").count(), 1);
  assert.ok(await whatsNew.getByRole("listitem").count() > 0);
  await whatsNew.getByRole("button", { name: "Done" }).click();
  assert.equal(await page.getByRole("dialog").count(), 0);

  await page.evaluate(() => localStorage.setItem("folio:last-opened-version", "1.13.1"));
  await page.reload();
  await whatsNew.getByRole("heading", { name: "Folio 1.14.0", exact: true }).waitFor();
  assert.equal(await whatsNew.locator("h3").first().textContent(), "Folio " + installed);
  assert.equal(await whatsNew.getByRole("heading", { name: "Folio 1.13.1", exact: true }).count(), 0);
  await page.keyboard.press("Escape");
  assert.equal(await page.getByRole("dialog").count(), 0);
  await page.reload();
  await page.getByRole("button", { name: /Try a sample/ }).waitFor();
  assert.equal(await whatsNew.count(), 0, "reopening the same version doesn't repeat the popup");
  console.log("ok  upgrades show bundled changes once, including skipped releases");

  await page.getByRole("button", { name: /Try a sample/ }).click();
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor({ timeout: 30_000 });
  assert.match(await page.locator(".textLayer").first().innerText(), /A short tour of Folio/);
  console.log("ok  sample rendered with a text layer");

  // The list of matching pages quotes each page's first match. The words it marks are the ones searched for,
  // also when the match sits many lines down the page. (Text with ligatures is checked at the end.)
  await page.getByRole("button", { name: /^Search/ }).click();
  await page.getByLabel("Find in document").fill("highlight");
  const quoted = page.locator(".search-result-list mark");
  await quoted.nth(2).waitFor();
  assert.deepEqual((await quoted.allTextContents()).map((text) => text.toLowerCase()), ["highlight", "highlight", "highlight"]);
  await page.getByLabel("Find in document").press("Escape");
  console.log("ok  search results mark the words that were searched for");

  // A dialog keeps the keyboard: the reader's one-key shortcuts don't act on the document behind it.
  const layout = () => page.evaluate(() => {
    const box = document.querySelector(".pdfViewer .page").getBoundingClientRect();
    return { upright: box.height > box.width, spreads: document.querySelectorAll(".pdfViewer .spread").length };
  });
  const before = await layout();
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: "About Folio" }).click();
  await page.getByRole("dialog", { name: "About Folio" }).waitFor();
  await page.keyboard.press("r");
  await page.keyboard.press("2");
  await page.waitForTimeout(200);
  assert.deepEqual(await layout(), before, "a shortcut reached the document behind the About dialog");
  await page.keyboard.press("Escape");
  assert.equal(await page.getByRole("dialog").count(), 0);
  console.log("ok  shortcuts wait while a dialog is open");
  await checkOpenRouter(app, page);

  const toolbar = page.getByRole("toolbar", { name: "Tools" });
  for (const name of ["Select", "Highlight", "Draw", "Text", "Sign", "Edit mode"]) await toolbar.getByRole("button", { name, exact: true }).waitFor();
  await toolbar.getByRole("button", { name: "Edit mode", exact: true }).click();
  await toolbar.getByRole("button", { name: "Highlight", exact: true }).waitFor();
  await toolbar.getByRole("button", { name: "Exit edit mode", exact: true }).click();

  const docs = page.getByRole("tablist", { name: "Open documents" });
  assert.equal(await page.locator(".document-tabs").count(), 0);
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: /New tab/ }).click();
  await docs.getByRole("tab", { name: "Library", selected: true }).waitFor();
  assert.equal(await docs.getByRole("tab").count(), 2);
  await docs.getByRole("button", { name: "Close Library" }).click();
  await page.locator(".doc-title", { hasText: "Folio sample.pdf" }).waitFor();
  assert.equal(await page.locator(".document-tabs").count(), 0);
  console.log("ok  the tab strip appears for a second tab and disappears when one remains");

  const sidebarToggle = page.getByRole("button", { name: /Toggle sidebar/ });
  const sidebarWasOpen = await sidebarToggle.getAttribute("aria-pressed") === "true";
  if (!sidebarWasOpen) await sidebarToggle.click();
  await page.getByRole("tab", { name: "Pages" }).click();
  const bookmark = page.getByRole("button", { name: "Bookmark page 2", exact: true });
  await bookmark.click();
  assert.equal(await bookmark.getAttribute("aria-pressed"), "true");
  assert.equal(await page.getByRole("textbox", { name: "Page number" }).inputValue(), "1");
  await page.keyboard.press("ControlOrMeta+w");
  await page.getByRole("button", { name: /Try a sample/ }).click();
  await page.getByRole("button", { name: "Bookmark page 2", pressed: true }).waitFor();
  await bookmark.press("Enter");
  assert.equal(await bookmark.getAttribute("aria-pressed"), "false");
  await bookmark.click();
  const filter = page.getByRole("button", { name: "Bookmarked", exact: true });
  const bookmark3 = page.getByRole("button", { name: "Bookmark page 3", exact: true });
  const pageNumber = page.getByRole("textbox", { name: "Page number" });
  await bookmark3.click();
  await filter.press("Enter");
  assert.equal(await filter.getAttribute("aria-pressed"), "true");
  assert.deepEqual(await page.locator(".thumbs .thumb").evaluateAll(rows => rows.map(row => row.dataset.page)), ["2", "3"]);
  assert.equal(await pageNumber.inputValue(), "1");
  await page.getByRole("button", { name: "Page 3", exact: true }).click();
  await page.waitForFunction(() => document.querySelector('input[aria-label="Page number"]').value === "3");
  await bookmark3.click();
  assert.equal(await page.locator(".thumbs .thumb").count(), 1);
  await bookmark.click();
  await page.getByText("No bookmarked pages yet.", { exact: true }).waitFor();
  assert.equal(await page.locator(".thumbs .thumb").count(), 0);
  await filter.click();
  assert.equal(await pageNumber.inputValue(), "3");
  await page.getByRole("button", { name: "Page 1", exact: true }).click();
  await bookmark.click();
  if (!sidebarWasOpen) await sidebarToggle.click();
  console.log("ok  a sidebar bookmark persists on its own and toggling it leaves the current page in place");
  console.log("ok  Bookmarked filters the sidebar, jumps to the original page, and updates as bookmarks are removed");

  // A signature is an unsaved edit that must survive a tab switch.
  const photo = await page.evaluate(() => {
    const canvas = Object.assign(document.createElement("canvas"), { width: 600, height: 200 });
    const g = canvas.getContext("2d");
    g.fillStyle = "#fff";
    g.fillRect(0, 0, 600, 200);
    g.lineWidth = 10;
    g.beginPath();
    g.moveTo(40, 150);
    g.bezierCurveTo(150, 10, 250, 190, 350, 60);
    g.lineTo(560, 120);
    g.stroke();
    return canvas.toDataURL("image/png").split(",")[1];
  });
  await page.getByTitle("Add a signature").click();
  await page.locator('input[aria-label="Upload signature photo"]').setInputFiles({ name: "signature.png", mimeType: "image/png", buffer: Buffer.from(photo, "base64") });
  await page.getByRole("button", { name: /Use this signature/ }).click();
  await page.locator('select[aria-label="Signature placement"]').selectOption("manual");
  await page.locator('[aria-label^="Place signature on page"]').click();
  await page.getByRole("button", { name: /Approve and add signature/ }).click();
  await page.locator(".dirty-dot").waitFor();

  // Stub the PDF viewer handoff in this renderer so printing cannot open a system dialog.
  await page.evaluate(interceptPdfPrint);
  await page.keyboard.press("ControlOrMeta+p");
  const printDialog = page.getByRole("dialog", { name: "Print", exact: true });
  await printDialog.getByRole("img", { name: "Preview of page 1" }).waitFor({ timeout: 15_000 });
  await printDialog.screenshot({ path: path.join(shots, "print-basic.png") });
  await printDialog.getByRole("tab", { name: "Advanced" }).click();
  await printDialog.getByRole("button", { name: "Booklet" }).click();
  await printDialog.getByRole("img", { name: "Preview of sheet 1" }).waitFor();
  await printDialog.screenshot({ path: path.join(shots, "print-advanced.png") });
  await printDialog.getByRole("tab", { name: "Basic" }).click();
  await printDialog.getByRole("button", { name: "Print…", exact: true }).click();
  await page.waitForFunction(() => window.printedPdf);
  assert.ok(await page.evaluate(() => window.printedPdf.type === "application/pdf" && /\/Subtype\s*\/Stamp\b/.test(window.printedPdf.bytes)));
  await page.evaluate(() => window.restorePrintFrame());
  assert.ok(await page.locator(".dirty-dot").isVisible(), "printing cleared the unsaved marker");
  console.log("ok  printing keeps unsaved work marked as unsaved");

  const sample = await page.evaluate(async () => Array.from(new Uint8Array(await (await fetch("sample.pdf")).arrayBuffer())));
  const other = Buffer.concat([Buffer.from(sample), Buffer.from("\\n% Another PDF")]);
  const select = name => docs.getByRole("tab", { name, exact: true }).click();
  await page.getByTitle("Next page").click();
  // Stub the native confirmation only in this test window.
  await page.window.evaluate(() => {
    const original = window.confirm;
    window.tabPrompts = [];
    window.confirm = message => { window.tabPrompts.push(message); return window.tabPrompts.length === 2; };
    window.restoreConfirm = () => { window.confirm = original; };
  });
  await app.evaluate(({ BrowserWindow }, bytes) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Another.pdf", data: Buffer.from(bytes) }), Array.from(other));
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor();
  await docs.getByRole("tab", { name: "Another.pdf", selected: true }).waitFor();
  assert.equal(await page.locator(".stampEditor").count(), 0);
  await select("Folio sample.pdf");
  assert.equal(await page.getByLabel("Page number", { exact: true }).inputValue(), "2");
  await page.getByTitle("Previous page").click();
  await page.locator(".stampEditor").waitFor();
  assert.deepEqual(await page.window.evaluate(() => window.tabPrompts), []);
  await select("Another.pdf");
  const closeSigned = docs.getByRole("button", { name: "Close Folio sample.pdf, unsaved changes" });
  await closeSigned.click();
  assert.equal(await docs.getByRole("tab").count(), 2);
  await closeSigned.click();
  await closeSigned.waitFor({ state: "detached" });
  assert.equal(await page.window.evaluate(() => window.tabPrompts.length), 2);
  await page.window.evaluate(() => window.restoreConfirm());

  // Batch opening keeps picker order, never replaces the current reader, and gives each file its own view.
  const choosing = page.waitForEvent("filechooser");
  await page.keyboard.press("ControlOrMeta+o");
  const chooser = await choosing;
  assert.equal(chooser.isMultiple(), true);
  await chooser.setFiles(["One.pdf", "Two.pdf"].map(name => ({ name, mimeType: "application/pdf",
    buffer: Buffer.concat([other, Buffer.from(name)]) })));
  await docs.getByRole("tab", { name: "One.pdf", selected: true }).waitFor();
  assert.deepEqual(await docs.getByRole("tab").allTextContents(), ["Another.pdf", "One.pdf", "Two.pdf"]);
  await select("Two.pdf");
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor();
  await page.locator(".document-tabs").screenshot({ path: path.join(shots, `document-tabs-${process.platform}.png`) });
  await page.keyboard.press("Control+Shift+Tab");
  await docs.getByRole("tab", { name: "One.pdf", selected: true }).waitFor();
  await page.evaluate(data => {
    const transfer = new DataTransfer();
    for (const name of ["Drop A.pdf", "Drop B.pdf"]) {
      transfer.items.add(new File([Uint8Array.from(data), name], name, { type: "application/pdf" }));
    }
    window.dispatchEvent(new DragEvent("drop", { dataTransfer: transfer }));
  }, sample);
  await docs.getByRole("tab", { name: "Drop A.pdf", selected: true }).waitFor();
  assert.equal(await docs.getByRole("tab").count(), 5);
  for (const name of ["Drop A.pdf", "Drop B.pdf", "One.pdf", "Two.pdf"]) {
    const close = docs.getByRole("button", { name: "Close " + name, exact: true });
    await close.click(); await close.waitFor({ state: "detached" });
  }
  await page.locator(".doc-title", { hasText: "Another.pdf" }).waitFor();
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: /Close document/ }).click();
  await page.getByRole("button", { name: /Try a sample/ }).click();
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor();
  console.log("ok  document tabs keep signatures and pages, confirm discard, and open multiple files");

  const viewport = await page.evaluate(() => ({ width: innerWidth, height: innerHeight }));
  await page.window.setViewportSize({ width: 420, height: viewport.height });
  assert.ok(await page.locator(".topbar button").evaluateAll(buttons => buttons.every(button => {
    const rect = button.getBoundingClientRect();
    return button.contains(document.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2));
  })), "a toolbar button is covered at the minimum width");
  await page.getByRole("button", { name: "Highlight", exact: true }).click();
  await page.window.setViewportSize(viewport);
  const textLayer = page.locator('.page[data-page-number="1"] .textLayer.highlighting');
  await textLayer.waitFor();
  await page.getByRole("radio", { name: "Green" }).click();
  const dragAcross = async line => {
    await line.scrollIntoViewIfNeeded();
    const box = await line.boundingBox();
    const y = box.y + box.height / 2;
    await page.mouse.move(box.x + 1, y);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width - 1, y, { steps: 8 });
  };
  await dragAcross(textLayer.getByText(/Drag a PDF file anywhere/));
  await page.waitForFunction(() => getSelection().toString().includes("Drag a PDF file anywhere"));
  const passage = await page.evaluate(() => getSelection().toString().trim());
  assert.equal(await page.locator(".highlightEditor").count(), 0);
  await page.mouse.up();
  await page.locator(".highlightEditor").waitFor();
  await page.getByRole("radio", { name: "Pink" }).click();
  await page.keyboard.down("Shift");
  await textLayer.getByText(/page\. The sidebar on the left/).evaluate(span => {
    const text = span.firstChild;
    const start = text.textContent.indexOf("The sidebar");
    getSelection().setBaseAndExtent(text, start, text, start + "The sidebar".length);
  });
  await page.waitForFunction(() => getSelection().toString() === "The sidebar");
  assert.equal(await page.locator(".highlightEditor").count(), 1);
  await page.keyboard.up("Shift");
  await page.waitForFunction(() => document.querySelectorAll(".highlightEditor").length === 2);
  assert.equal(await page.locator(".selection-popover").count(), 0);
  await page.getByRole("button", { name: "Done", exact: true }).click();
  await page.getByRole("button", { name: "Select", pressed: true }).waitFor();
  await textLayer.waitFor({ state: "detached" });
  console.log("ok  Highlight mode marks consecutive selections on release in the chosen color");

  await page.getByTitle("Draw (D)").click();
  const drawingTool = page.getByLabel("Drawing tool");
  const inkPage = page.locator('.page[data-page-number="1"]');
  await inkPage.locator(".annotationEditorLayer.inkEditing").waitFor();
  await page.getByRole("radio", { name: "Blue" }).click();
  for (const [index, shape] of ["Rectangle", "Oval", "Arrow"].entries()) {
    if (index === 1) {
      await page.getByTitle("Rotate (R)").click();
      await inkPage.locator('.annotationEditorLayer.inkEditing[data-main-rotation="90"]').waitFor();
    }
    await drawingTool.selectOption(shape);
    // Zoom and rotation can leave only part of the page visible. Wait for a clear drag on the
    // drawing layer, avoiding existing annotations and the floating reader controls.
    const start = await page.waitForFunction(() => {
      const el = document.querySelector('.page[data-page-number="1"]');
      const layer = el?.querySelector(".annotationEditorLayer.inkEditing");
      if (!layer) return null;
      const box = el.querySelector(".canvasWrapper").getBoundingClientRect();
      const view = el.closest(".viewer-container").getBoundingClientRect();
      const left = Math.ceil(Math.max(0, box.left, view.left)) + 8;
      const top = Math.ceil(Math.max(0, box.top, view.top)) + 8;
      const right = Math.min(innerWidth, box.right, view.right) - 68;
      const bottom = Math.min(innerHeight, box.bottom, view.bottom) - 44;
      for (let y = top; y <= bottom; y += 24) {
        for (let x = left; x <= right; x += 32) {
          if (document.elementFromPoint(x, y) === layer && document.elementFromPoint(x + 60, y + 36) === layer) return { x, y };
        }
      }
      return null;
    });
    const { x, y } = await start.jsonValue();
    await start.dispose();
    // pdf.js orders editors by page position, so a new shape need not be the last DOM node.
    const excludeExisting = await inkPage.locator(".inkEditor").evaluateAll(editors =>
      editors.map(editor => `:not(#${CSS.escape(editor.id)})`).join(""));
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 60, y + 36);
    await inkPage.locator(".shape-preview path").waitFor({ state: "attached" });
    await page.mouse.up();
    const editor = inkPage.locator(`.inkEditor${excludeExisting}`);
    await editor.waitFor({ state: "visible" });
    const placed = await editor.boundingBox();
    assert.ok(Math.abs(placed.x + placed.width / 2 - (x + 30)) < 8 && Math.abs(placed.y + placed.height / 2 - (y + 18)) < 8,
      `${shape} should appear where it was drawn`);
  }
  await page.getByTitle(/Undo \(/).click();
  await page.waitForFunction(() => document.querySelectorAll(".inkEditor").length === 2);
  await page.getByRole("button", { name: "Done", exact: true }).click();
  await page.keyboard.press("ControlOrMeta+Shift+z");
  await inkPage.locator(".inkEditor").nth(2).waitFor({ state: "attached" });
  for (let i = 0; i < 3; i++) await page.getByTitle("Rotate (R)").click();
  console.log("ok  Draw creates the three shapes and supports undo and redo");

  // Select a sentence and highlight it in one click.
  await dragAcross(page.locator('.page[data-page-number="1"] .textLayer').getByText(/A short tour of Folio/));
  await page.mouse.up();
  await page.locator(".selection-popover").waitFor();

  // Copy puts the selection on the system clipboard (whose contents are put back afterwards). Clipboard writes
  // need a focused window, which a CI desktop doesn't always give the app. Only CI runs it: on someone's own
  // computer it would briefly replace what they've copied (and anything but plain text wouldn't come back).
  if (!process.env.CI) {
    console.log("--  clipboard check skipped: it only runs in CI, to leave your clipboard alone");
  } else if (await page.evaluate(() => document.hasFocus())) {
    const clipboardBefore = await app.evaluate(({ clipboard }) => clipboard.readText());
    await page.locator(".selection-popover").getByRole("button", { name: "Copy" }).click();
    await page.locator('.selection-popover button[title="Copied"]').waitFor({ timeout: 5_000 });
    const copied = await app.evaluate(({ clipboard }) => clipboard.readText());
    await app.evaluate(({ clipboard }, text) => clipboard.writeText(text), clipboardBefore);
    assert.ok(/A short tour of Folio/.test(copied), "Copy didn't reach the clipboard");
    console.log("ok  selection copied to the clipboard");
  } else {
    console.log("--  clipboard check skipped: the window isn't focused");
  }

  // A refused clipboard write must be reported, not dressed up as a copy. The write is blocked here so the button's
  // failure path runs on every machine, not just where the OS happens to refuse it.
  await page.evaluate(() => {
    globalThis.__clipboardWriteText = navigator.clipboard.writeText;
    Object.defineProperty(navigator.clipboard, "writeText", {
      configurable: true,
      value: () => Promise.reject(new DOMException("Write permission denied.", "NotAllowedError")),
    });
  });
  await page.locator(".selection-popover").getByRole("button", { name: "Copy" }).click();
  await page.locator('.selection-popover button[title="Copy failed"]').waitFor({ timeout: 5_000 });
  await page.evaluate(() => {
    Object.defineProperty(navigator.clipboard, "writeText", {
      configurable: true,
      value: globalThis.__clipboardWriteText,
    });
  });
  console.log("ok  a refused clipboard write is reported as Copy failed");

  await page.getByRole("button", { name: "Highlight yellow" }).click();
  await page.waitForFunction(() => document.querySelectorAll(".highlightEditor").length === 3, null, { timeout: 15_000 });
  console.log("ok  highlight created");

  // A saved copy includes every saved mark, also on pages that haven't been drawn since the document was reopened.
  const lastPage = await page.locator(".page").count();
  // A page's text can be selected before pdf.js has built its highlight layer: it first waits for the page's links
  // and form fields from its worker, which can take a while on a busy computer. Hold that reply back so this
  // highlight always lands in the gap.
  await page.evaluate(() => {
    const post = Worker.prototype.postMessage;
    window.folioPostMessage = post;
    Worker.prototype.postMessage = function (message, ...rest) {
      if (message?.action !== "GetAnnotations") return post.call(this, message, ...rest);
      setTimeout(() => post.call(this, message, ...rest), 2500);
    };
  });
  await page.evaluate((n) => document.querySelector(`.page[data-page-number="${n}"]`).scrollIntoView(), lastPage);
  await page.locator(`.page[data-page-number="${lastPage}"] .textLayer span:not(.markedContent)`).first().evaluate((span) => {
    const range = document.createRange();
    range.selectNodeContents(span);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await page.locator(".selection-popover").waitFor();
  await page.getByRole("button", { name: "Highlight yellow" }).click();
  await page.waitForFunction(() => document.querySelectorAll(".highlightEditor").length === 4, null, { timeout: 15_000 });
  await page.evaluate(() => { Worker.prototype.postMessage = window.folioPostMessage; });
  await page.evaluate(() => document.querySelector('.page[data-page-number="1"]').scrollIntoView());
  await page.waitForTimeout(800);
  await page.keyboard.press("ControlOrMeta+w");
  await page.getByRole("button", { name: /Try a sample/ }).click();
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor({ timeout: 30_000 });
  const copy = path.join(env.FOLIO_TEST_USER_DATA, "copy.pdf");
  await app.evaluate(({ session }, file) => {
    globalThis.folioCopySaved = new Promise((resolve) => session.defaultSession.once("will-download", (_event, item) => {
      item.setSavePath(file);
      item.once("done", (_done, state) => resolve(state));
    }));
  }, copy);
  await page.keyboard.press("ControlOrMeta+s");
  assert.equal(await app.evaluate(() => globalThis.folioCopySaved), "completed");
  assert.equal(readFileSync(copy, "latin1").match(/\/Subtype\s*\/Highlight\b/g)?.length, 4);
  const savedInk = await page.evaluate(async bytes => {
    const task = pdfjsLib.getDocument({ data: Uint8Array.from(bytes) });
    try {
      const pdf = await task.promise;
      const annotations = await (await pdf.getPage(1)).getAnnotations();
      return annotations.filter(a => a.subtype === "Ink").map(a => ({ color: Array.from(a.color), paths: a.inkLists.length }));
    } finally { await task.destroy(); }
  }, Array.from(readFileSync(copy)));
  assert.deepEqual(savedInk.map(a => a.paths).sort(), [1, 3, 4]);
  assert.ok(savedInk.every(a => String(a.color) === "47,111,237"));
  console.log("ok  all three shapes survive reopening and are embedded in the saved PDF");
  console.log("ok  saved copy includes marks on pages not drawn yet");

  if (await sidebarToggle.getAttribute("aria-pressed") !== "true") await sidebarToggle.click();
  await page.getByRole("tab", { name: /^Notes/ }).click();
  for (const [quote, color] of [[passage, "#5edc9a"], ["The sidebar", "#ff8fc7"]]) {
    const card = page.locator(".mark-card").filter({ hasText: quote });
    assert.equal(await card.evaluate(el => el.style.getPropertyValue("--mark")), color);
  }
  await page.getByLabel("Notes about this document").fill("Compare these passages.");
  const jsonFile = path.join(env.FOLIO_TEST_USER_DATA, "reading-notes.json");
  await app.evaluate(({ session }, file) => {
    globalThis.folioNotesSaved = new Promise(resolve => session.defaultSession.once("will-download", (_, item) => {
      item.setSavePath(file);
      item.once("done", (_, state) => resolve([state, item.getFilename()]));
    }));
  }, jsonFile);
  await page.getByRole("button", { name: "Export" }).click();
  await page.getByRole("menuitem", { name: "Save JSON…" }).click();
  const [state, filename] = await app.evaluate(() => globalThis.folioNotesSaved);
  assert.equal(state, "completed");
  assert.match(filename, /\.notes\.json$/);
  const exported = JSON.parse(readFileSync(jsonFile, "utf8"));
  assert.equal(exported.schemaVersion, 1);
  assert.equal(exported.documentNote, "Compare these passages.");
  assert.equal(exported.document.pageCount, lastPage);
  assert.ok(exported.bookmarks.some(b => b.page >= 1));
  assert.deepEqual(exported.bookmarkedPages, [2]);
  assert.equal(new Set(exported.annotations.map(a => a.id)).size, exported.annotations.length);
  assert.ok(exported.annotations.some(a => a.type === "highlight" && a.page === lastPage && a.quadPoints.length >= 8));
  const shapes = exported.annotations.filter(a => a.type === "shape");
  assert.deepEqual(shapes.map(a => a.shape).sort(), ["arrow", "oval", "rectangle"]);
  assert.ok(shapes.every(a => a.page === 1 && a.color.toLowerCase() === "#2f6fed" && a.strokes[0].every(Number.isFinite)));
  await page.evaluate(() => {
    window.folioOriginalClipboard = Object.getOwnPropertyDescriptor(navigator.clipboard, "writeText");
    Object.defineProperty(navigator.clipboard, "writeText", { configurable: true, value: async text => { window.folioMarkdown = text; } });
  });
  await page.getByRole("button", { name: "Export" }).click();
  await page.getByRole("menuitem", { name: "Copy Markdown" }).click();
  assert.match(await page.evaluate(() => window.folioMarkdown), /Compare these passages\./);
  assert.match(await page.evaluate(() => window.folioMarkdown), /## Bookmarked pages\n\n- p\. 2/);
  await page.evaluate(() => {
    if (window.folioOriginalClipboard) Object.defineProperty(navigator.clipboard, "writeText", window.folioOriginalClipboard);
    else delete navigator.clipboard.writeText;
  });
  console.log("ok  Notes exports Markdown and JSON with drawings, offscreen highlights and outline destinations");
  await page.getByRole("tab", { name: /Pages/ }).click();

  await page.getByRole("button", { name: "More" }).click();
  await page.getByRole("menuitem", { name: "About Folio" }).click();
  const about = page.getByRole("dialog", { name: "About Folio" });
  assert.doesNotMatch(await about.innerText(), /Ask AI accounts/);
  assert.equal(await about.getByRole("button", { name: "Settings" }).count(), 0);
  await page.keyboard.press("Escape");

  // The reader shortcut and Settings control the same saved PDF appearance.
  await page.getByRole("button", { name: "More" }).click();
  const darkMode = page.getByRole("menuitemcheckbox", { name: "PDF dark mode" });
  assert.equal(await darkMode.getAttribute("aria-checked"), "false");
  await darkMode.click();
  await page.locator(".app[data-night]").waitFor();
  await page.getByRole("button", { name: "More" }).click();
  assert.equal(await darkMode.getAttribute("aria-checked"), "true");
  await page.getByRole("menuitem", { name: /Settings/ }).click();
  await page.getByRole("navigation", { name: "Settings sections" }).getByRole("button", { name: "General" }).waitFor();
  const darkModeSwitch = page.getByRole("switch", { name: "PDF dark mode" });
  assert.equal(await darkModeSwitch.getAttribute("aria-checked"), "true");
  await darkModeSwitch.click();
  await page.locator(".app:not([data-night])").waitFor();
  assert.equal(await page.evaluate(() => localStorage.getItem("folio:night")), "0");
  await page.keyboard.press("Escape");

  // Settings links to the bundled changelog, and closing it returns to Settings.
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Comma" : "Control+Comma");
  await page.getByRole("button", { name: "View changelog" }).click();
  const summaries = whatsNew.locator("summary");
  await summaries.first().waitFor();
  const firstPage = await summaries.allTextContents();
  assert.equal(firstPage.length, 5);
  assert.equal(firstPage[0], "Folio " + installed);
  assert.equal(await whatsNew.locator("details[open]").count(), 1);
  await summaries.nth(1).click();
  assert.ok(await whatsNew.locator("details").nth(1).locator("li").first().isVisible());
  const older = whatsNew.getByRole("button", { name: "Older releases" });
  const newer = whatsNew.getByRole("button", { name: "Newer releases" });
  assert.ok(await newer.isDisabled());
  await older.click();
  assert.ok(!(await summaries.allTextContents()).some(version => firstPage.includes(version)));
  await newer.click();
  assert.deepEqual(await summaries.allTextContents(), firstPage);
  assert.equal(await whatsNew.getByRole("link", { name: "View releases on GitHub" }).getAttribute("href"),
    "https://github.com/Forward-Future/folio-releases/releases");
  await page.keyboard.press("Escape");
  await page.getByRole("dialog", { name: "Settings" }).waitFor();
  await page.keyboard.press("Escape");
  console.log("ok  changelog opens from Settings");

  // A text box is saved automatically and comes back, text and all, when the document is opened again.
  // The sample can reopen at the previous reading position. Bring its first page into view so pdf.js builds its editor layer.
  const pageField = page.getByLabel("Page number", { exact: true });
  await pageField.fill("1");
  await pageField.press("Enter");
  await page.locator('button[title^="Text ("]').click();
  // The font and size chosen before clicking apply to the new box, and are saved with it.
  await page.getByRole("combobox", { name: "Font", exact: true }).selectOption("serif");
  await page.getByRole("combobox", { name: "Font size", exact: true }).selectOption("24");
  // Clicked in the middle of the page's visible part: a fixed spot near the top lands under the toolbar's color
  // picker in a small window, as on CI runners.
  await page.locator('.page[data-page-number="1"] .annotationEditorLayer').click();
  // Filled in one go: pdf.js moves the caret while it focuses a new box, which can scramble key-by-key typing.
  await page.locator('.freeTextEditor .internal[contenteditable="true"]').fill("Saved text box");
  await page.getByRole("button", { name: "Done", exact: true }).click();
  await page.waitForTimeout(1000);
  await page.keyboard.press("ControlOrMeta+w");
  await page.getByRole("button", { name: /Try a sample/ }).click();
  const restored = page.locator(".freeTextEditor", { hasText: "Saved text box" });
  await restored.waitFor({ timeout: 15_000 });
  const style = await restored.locator(".internal").evaluate(el => ({ family: el.style.fontFamily, size: el.style.fontSize }));
  assert.match(style.family, /Times New Roman/);
  assert.match(style.size, /^calc\(24px/);
  // A selected box shows its own font, and a new choice changes that box instead of the next one.
  await page.locator('button[title^="Text ("]').click();
  await restored.click();
  const fontPicker = page.getByRole("combobox", { name: "Font", exact: true });
  await page.waitForFunction(() => document.querySelector('select[aria-label="Font"]')?.value === "serif");
  const boxFont = () => restored.locator(".internal").evaluate(el => el.style.fontFamily);
  await fontPicker.selectOption("easyjet");
  assert.match(await boxFont(), /easyJet Rounded Book/);
  assert.equal(await page.evaluate(() => localStorage.getItem("folio:text-font")), "serif");
  // Undo puts the box's previous font back, and the picker follows it.
  await page.locator('button.pill-btn[title^="Undo"]').click();
  assert.match(await boxFont(), /Times New Roman/);
  await page.waitForFunction(() => document.querySelector('select[aria-label="Font"]')?.value === "serif");
  await page.getByRole("button", { name: "Done", exact: true }).click();
  console.log("ok  text box restored after reopening, in its font and size");

  await page.waitForTimeout(500);
  await page.screenshot({ path: path.join(shots, `smoke-${process.platform}.png`) });
  await checkReaderTools(app, page, env.FOLIO_TEST_USER_DATA, shots);

  // Thumbnails keep each page's shape and relative size: letter portrait, letter landscape, and an 11x17 sheet.
  const mixed = pdfWithPageSizes([[612, 792], [792, 612], [1224, 792]]);
  await app.evaluate(({ BrowserWindow }, data) =>
    BrowserWindow.getAllWindows()[0].webContents.send("folio:open-file", { name: "Mixed sizes.pdf", data: Buffer.from(data) }), [...mixed]);
  await page.window.waitForFunction(() => document.title.startsWith("Mixed sizes.pdf"), null, { timeout: 15_000 });
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: /^Print\b/ }).click();
  const recordPrint = process.env.FOLIO_PRINT_VIDEO === "1";
  if (recordPrint) {
    await printDialog.getByRole("img", { name: "Preview of page 1" }).waitFor();
    const size = await page.window.evaluate(() => ({ width: innerWidth, height: innerHeight }));
    await page.window.screencast.start({ path: path.join(shots, "print-preview.webm"), size, fps: 12 });
  }
  for (const [i, ratio] of [612 / 792, 792 / 612, 1224 / 792].entries()) {
    if (i) await printDialog.getByRole("button", { name: "Next page" }).click();
    await printDialog.getByRole("img", { name: `Preview of page ${i + 1}` }).waitFor();
    const bounds = await printDialog.locator(".print-paper").boundingBox();
    assert.ok(Math.abs(bounds.width / bounds.height - ratio) < 0.01, "print preview lost the page dimensions");
    if (recordPrint) await page.waitForTimeout(1050);
  }
  if (recordPrint) {
    await printDialog.getByRole("tab", { name: "Advanced" }).click();
    await printDialog.getByRole("button", { name: "Booklet" }).click();
    await printDialog.getByRole("img", { name: "Preview of sheet 1" }).waitFor();
    await page.waitForTimeout(1300);
    await page.window.screencast.stop();
  }
  await printDialog.getByRole("button", { name: "Close", exact: true }).click();
  if (!(await page.locator(".sidebar").count())) await page.keyboard.press("ControlOrMeta+Backslash");
  await page.getByRole("tab", { name: "Pages" }).click();
  // Waits until the page sizes have loaded and the third thumbnail is wider than the first.
  await page.waitForFunction(() => {
    const f = [...document.querySelectorAll(".thumb-frame")];
    return f.length === 3 && f[2].getBoundingClientRect().width > f[0].getBoundingClientRect().width;
  }, null, { timeout: 15_000 });
  const frames = await page.locator(".thumb-frame").evaluateAll((els) => els.map((el) => {
    const r = el.getBoundingClientRect();
    return { width: r.width, height: r.height };
  }));
  const near = (a, b) => Math.abs(a - b) <= 1;
  assert.ok(frames[0].height > frames[0].width, `letter portrait thumbnail isn't portrait: ${JSON.stringify(frames[0])}`);
  assert.ok(frames[1].width > frames[1].height, `letter landscape thumbnail isn't landscape: ${JSON.stringify(frames[1])}`);
  assert.ok(near(frames[0].width, frames[1].height) && near(frames[0].height, frames[1].width),
    `letter pages aren't the same size turned: ${JSON.stringify(frames)}`);
  // 11x17 landscape is 11in tall like portrait letter, and 17in wide against letter's 8.5in.
  assert.ok(near(frames[2].height, frames[0].height) && Math.abs(frames[2].width / frames[0].width - 2) < 0.05,
    `11x17 thumbnail isn't to scale with letter: ${JSON.stringify(frames)}`);
  console.log("ok  thumbnails keep each page's shape and size:", frames.map((f) => `${Math.round(f.width)}x${Math.round(f.height)}`).join(", "));

  // The bookmark filter can show page 2 in row 1 while the document's largest sheet is hidden.
  await page.getByRole("button", { name: "Bookmark page 2", exact: true }).click();
  await filter.click();
  await page.waitForFunction(() => document.querySelectorAll(".thumb").length === 1);
  const filtered = await page.locator('.thumb[data-page="2"] .thumb-frame').boundingBox();
  assert.ok(filtered && near(filtered.width, frames[1].width) && near(filtered.height, frames[1].height));
  await filter.click();

  // In the page organizer, a moved page takes its own shape to its new row: drag the 11x17 sheet to the top.
  await page.getByRole("button", { name: "Reorder pages" }).click();
  const organized = (n) => page.locator(`.thumb[data-page="${n}"]`);
  await organized(3).dragTo(organized(1));
  assert.equal(await organized(3).getAttribute("aria-label"), "Page 3, position 1");
  const rows = await page.locator(".thumb").evaluateAll((els) => els.map((el) => {
    const r = el.querySelector(".thumb-frame").getBoundingClientRect();
    return { page: Number(el.dataset.page), top: el.getBoundingClientRect().top, width: r.width, height: r.height };
  }));
  assert.deepEqual(rows.map((r) => r.page), [3, 1, 2]);
  assert.ok(rows[0].top < rows[1].top && rows[1].top < rows[2].top, `rows are out of order: ${JSON.stringify(rows)}`);
  const sameFrame = (row, frame) => near(row.width, frame.width) && near(row.height, frame.height);
  assert.ok(sameFrame(rows[0], frames[2]) && sameFrame(rows[1], frames[0]) && sameFrame(rows[2], frames[1]),
    `reordered thumbnails lost their shape: ${JSON.stringify(rows)}`);
  await page.getByRole("button", { name: "Cancel" }).click();
  console.log("ok  reordered thumbnails keep each page's shape:", rows.map((r) => `page ${r.page} ${Math.round(r.width)}x${Math.round(r.height)}`).join(", "));

  // A Word document opens converted to PDF (by Word, LibreOffice, or Folio itself, whichever the computer has),
  // and the PDF is saved next to it. Opening the document again reuses that PDF.
  const folder = mkdtempSync(path.join(tmpdir(), "folio-word-"));
  const docx = path.join(folder, "Quarterly report.docx");
  writeFileSync(docx, await wordDocument("Quarterly report from Word"));
  const neighbor = path.join(folder, "Quarterly report.pdf");
  writeFileSync(neighbor, Buffer.from(sample));
  const openWordDocument = async (expectedName) => {
    await page.getByTitle("Back to library").click();
    const chooser = page.waitForEvent("filechooser");
    await page.getByRole("button", { name: /Open a PDF/ }).click();
    await (await chooser).setFiles(docx);
    // Stops at a conversion error too, and shows what each converter said.
    const settled = await page.window.waitForFunction((name) => {
      const reader = document.querySelector("iframe.document-reader.active")?.contentDocument;
      return reader?.title.startsWith(name) || reader?.querySelector(".toast.error")?.textContent;
    }, expectedName, { timeout: 150_000 }).catch((error) => error);
    if (settled instanceof Error || !(await page.title()).startsWith(expectedName)) {
      const log = path.join(env.FOLIO_TEST_USER_DATA, "convert.log");
      console.log(await page.locator(".toast").textContent().catch(() => "(no message)"));
      console.log(existsSync(log) ? readFileSync(log, "utf8") : "(no convert.log)");
      throw settled instanceof Error ? settled : new Error("the Word document didn't convert");
    }
    await page.locator(".textLayer span:not(.markedContent)").first().waitFor({ timeout: 30_000 });
    assert.match(await page.locator(".textLayer").first().innerText(), /Quarterly report from Word/);
  };
  const converted = "Quarterly report (2).pdf";
  await openWordDocument(converted);
  assert.deepEqual(readFileSync(neighbor), Buffer.from(sample));
  await openWordDocument(converted);
  assert.deepEqual(readdirSync(folder).sort(), [converted, "Quarterly report.docx", "Quarterly report.pdf"]);
  writeFileSync(path.join(folder, converted), Buffer.from(sample));
  await openWordDocument("Quarterly report (3).pdf");
  console.log("ok  Word import preserves unrelated neighbors, reuses a verified PDF, and reconverts a replaced copy");

  // A pasted link fills in Open from URL; nothing is contacted until the user presses Open.
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData("text/plain", "https://example.com/paper.pdf");
    document.querySelector(".viewer-container").dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true }));
  });
  const urlDialog = page.getByRole("dialog", { name: "Open from URL" });
  assert.equal(await urlDialog.getByLabel("PDF URL").inputValue(), "https://example.com/paper.pdf");
  await page.keyboard.press("Escape");
  console.log("ok  a pasted link waits in Open from URL");

  // Questions Folio asks in the main process are answered here, and links are recorded instead of reaching the browser.
  await app.evaluate(({ dialog, shell }) => {
    const check = globalThis.folioCheck = { asked: [], opened: [], answer: 1 };
    dialog.showMessageBox = async (_window, options) => { check.asked.push(options.message); return { response: check.answer }; };
    shell.openExternal = async (url) => { check.opened.push(url); };
  });
  const answer = (response) => app.evaluate((_electron, response) => { globalThis.folioCheck.answer = response; }, response);
  const settled = async (count) => {
    for (let i = 0; i < 50 && (await app.evaluate(() => globalThis.folioCheck.asked.length)) < count; i++) await page.waitForTimeout(100);
    return app.evaluate(() => globalThis.folioCheck);
  };

  // A link from a document opens in the browser only after the user has seen where it goes; Folio's own links just open.
  await page.evaluate(() => window.open("https://example.com/invoice", "_blank"));
  assert.deepEqual(await settled(1), { asked: ["Open example.com in your browser?"], opened: [], answer: 1 });
  await answer(0);
  await page.evaluate(() => window.open("https://example.com/invoice", "_blank"));
  assert.deepEqual((await settled(2)).opened, ["https://example.com/invoice"]);
  await page.evaluate(() => window.open("https://claude.com/claude-code", "_blank"));
  await page.waitForTimeout(300);
  const linkCheck = await settled(2);
  assert.deepEqual([linkCheck.asked.length, linkCheck.opened.at(-1)], [2, "https://claude.com/claude-code"]);
  console.log("ok  links from documents ask before opening");

  // Addresses on this computer or the local network are reached only once the user agrees, also when the page asks
  // directly (as a document that took over the page might) and when a name only resolves there.
  const served = [];
  const server = http.createServer((req, res) => {
    served.push(req.url);
    const trap = `http://127.0.0.1:${server.address().port}/private-trap`;
    if (req.url === "/redirect") { res.writeHead(302, { location: trap }); return res.end(); }
    if (req.url === "/page") {
      res.writeHead(200, { "content-type": "text/html" });
      return res.end(`<title>Guarded page</title><script>
        alert("dismissed");
        const origScrollTo = window.scrollTo.bind(window);
        window.scrollTo = (...args) => {
          if (navigator.userActivation?.isActive) document.title = "User gesture leaked";
          return origScrollTo(...args);
        };
      </script><img src="/same-origin"><img src="${trap}">`);
    }
    if (req.url === "/non-pdf-stream") {
      res.writeHead(200, { "content-type": "application/octet-stream" });
      res.write(Buffer.alloc(1024, 0x41));
      return;
    }
    if (req.url === "/stall") return; // never answers
    res.writeHead(200, { "content-type": "application/pdf" }); res.end(Buffer.from(sample));
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const fetchPdf = (url) => page.evaluate((url) => window.folioDesktop.fetchPdf(url).then((file) => file?.name ?? null, (error) => error.message), url);
  try {
    await answer(1);
    assert.equal(await fetchPdf(`http://127.0.0.1:${port}/local.pdf`), null);
    assert.equal(await fetchPdf(`http://folio-local.test:${port}/named.pdf`), null);
    assert.deepEqual(served, []);
    assert.deepEqual((await settled(4)).asked.slice(2), [`Open 127.0.0.1:${port}?`, `Open folio-local.test:${port}?`]);
    await answer(0);
    assert.equal(await fetchPdf(`http://folio-local.test:${port}/named.pdf`), "named.pdf");
    assert.deepEqual(served, ["/named.pdf"]);
    // DNS changes after the consent check. The site must receive no request, even one whose response would be discarded.
    await app.evaluate(({ session }) => {
      const sessions = Object.getPrototypeOf(session.defaultSession);
      const resolveHost = sessions.resolveHost;
      let first = true;
      sessions.resolveHost = function (host, options) {
        if (host === "folio-rebind.test" && first) { first = false; return Promise.resolve({ endpoints: [{ address: "93.184.215.14" }] }); }
        // Chromium's own lookup points to .2. Only a connection pinned to the vetted .1 endpoint can succeed.
        if (host === "folio-pinned.test") return Promise.resolve({ endpoints: [{ address: "127.0.0.1" }] });
        return resolveHost.call(this, host, options);
      };
    });
    assert.match(await fetchPdf(`http://folio-rebind.test:${port}/rebound.pdf`), /local network address/);
    assert.equal((await settled(5)).asked.length, 5);
    assert.ok(!served.includes("/rebound.pdf"));
    assert.equal(await fetchPdf(`http://folio-pinned.test:${port}/pinned.pdf`), "pinned.pdf");
    assert.match(await fetchPdf(`http://folio-pinned.test:${port}/non-pdf-stream`), /isn't a PDF or a webpage/);
    assert.match(await fetchPdf(`http://folio-pinned.test:${port}/redirect`), /local network address/);
    assert.equal(await fetchPdf(`http://folio-pinned.test:${port}/page`), "Guarded page.pdf");
    assert.ok(served.includes("/same-origin"), "allowed page subresource did not load");
    assert.ok(!served.includes("/private-trap"), "redirect or hidden page contacted a different local origin");
    // One link opens at a time. A link that is given up on (its tab closed, another document opened instead) is
    // stopped by the page, which lets the next one open straight away instead of after its timeout.
    const stalled = page.evaluate((url) => window.folioDesktop.fetchPdf(url, "stalled").then((file) => file?.name ?? null, (error) => error.message),
      `http://folio-pinned.test:${port}/stall`);
    for (let i = 0; i < 100 && !served.includes("/stall"); i++) await page.waitForTimeout(50);
    assert.match(await fetchPdf(`http://folio-pinned.test:${port}/pinned.pdf`), /Another link is still opening/);
    await page.evaluate(() => window.folioDesktop.cancelFetchPdf("stalled"));
    assert.equal(await stalled, null);
    assert.equal(await fetchPdf(`http://folio-pinned.test:${port}/pinned.pdf`), "pinned.pdf");
  } finally {
    server.close();
  }
  console.log("ok  local network addresses need the user's OK");
  console.log("ok  a link that is given up on stops, and the next one opens");

  // The question asked before closing over unsaved work names this platform's Save shortcut.
  const closePrompt = await app.evaluate(({ BrowserWindow, dialog }) => {
    const ask = dialog.showMessageBoxSync;
    let detail = null;
    dialog.showMessageBoxSync = (_window, options) => { detail = options.detail; return 1; };
    try { BrowserWindow.getAllWindows()[0].webContents.emit("will-prevent-unload", { preventDefault() {} }); }
    finally { dialog.showMessageBoxSync = ask; }
    return detail;
  });
  assert.match(closePrompt, process.platform === "darwin" ? /⌘S/ : /Ctrl\+S/);
  console.log("ok  the unsaved-work prompt shows this platform's Save shortcut");

  // Closing Ask AI mid-answer stops the answer, like Stop: it doesn't run on unseen or come back unfinished. A stand-in
  // provider starts each answer and never finishes it.
  await app.evaluate(({ ipcMain, BrowserWindow }) => {
    const ai = globalThis.folioAi = { asked: [], cancelled: [], reply: "" };
    for (const channel of ["folio:ai-status", "folio:ai-ask", "folio:ai-cancel"]) ipcMain.removeHandler(channel);
    ipcMain.handle("folio:ai-status", () => ({ chatgpt: { connected: false, email: "", name: "", welcomeSeen: true, encrypted: true },
      claude: { installed: true, loggedIn: true, subscription: true, version: "test", authMethod: "claude.ai" } }));
    ipcMain.handle("folio:ai-ask", (_event, request) => {
      ai.asked.push(request.id);
      if (ai.reply) BrowserWindow.getAllWindows()[0].webContents.send("folio:ai-event", { id: request.id, type: "delta", text: ai.reply });
      return true;
    });
    ipcMain.handle("folio:ai-cancel", (_event, id) => { ai.cancelled.push(id); return true; });
  });
  await page.evaluate(() => localStorage.setItem("folio:ai-provider", "claude"));
  const aiPanel = page.getByRole("complementary", { name: "Ask AI" });
  const askAndClose = async (question, reply) => {
    const asked = await app.evaluate((_electron, reply) => { globalThis.folioAi.reply = reply; return globalThis.folioAi.asked.length; }, reply);
    await aiPanel.getByRole("textbox").fill(question);
    await aiPanel.getByRole("textbox").press("Enter");
    // Wait until the question has reached the provider, so the answer is under way rather than still being prepared.
    for (let i = 0; i < 50 && (await app.evaluate(() => globalThis.folioAi.asked.length)) === asked; i++) await page.waitForTimeout(100);
    if (reply) await aiPanel.getByText(reply).waitFor();
    await aiPanel.getByRole("button", { name: "Close Ask AI" }).click();
    await page.getByRole("button", { name: "Ask AI", exact: true }).click();
    await aiPanel.getByText(question).waitFor();
  };
  await page.getByRole("button", { name: "Ask AI", exact: true }).click();
  await askAndClose("First question", "Partial answer");
  await askAndClose("Second question", "");
  const aiCheck = await app.evaluate(() => globalThis.folioAi);
  assert.equal(aiCheck.asked.length, 2);
  assert.deepEqual(aiCheck.cancelled, aiCheck.asked, "closing Ask AI left an answer running");
  assert.equal(await aiPanel.locator(".typing").count(), 0, "a stopped answer still shows as in progress");
  assert.ok(await aiPanel.getByText("Partial answer").isVisible(), "the part already answered was dropped");
  assert.equal(await aiPanel.getByRole("button", { name: "Stop" }).count(), 0, "the panel still thinks an answer is running");
  await aiPanel.getByRole("button", { name: "Close Ask AI" }).click();
  console.log("ok  closing Ask AI stops the answer in progress");

  await page.getByRole("button", { name: "Ask AI", exact: true }).click();
  await aiPanel.getByText("First question").waitFor();
  if (await sidebarToggle.getAttribute("aria-pressed") !== "true") await sidebarToggle.click();
  await page.getByRole("tab", { name: /^Notes/ }).click();
  const documentNote = page.getByLabel("Notes about this document");
  await documentNote.fill("Note in the other tab");
  const readerName = await page.evaluate(() => document.title.replace(/ — Folio$/, ""));
  await page.keyboard.press("ControlOrMeta+t");
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: /^Settings/ }).click();
  const settings = page.getByRole("dialog", { name: "Settings" });
  await settings.getByRole("button", { name: "Storage", exact: true }).click();
  const stored = settings.locator(".stored-data");
  const clearNotes = stored.getByRole("button", { name: "Clear annotations…", exact: true });
  const clearChats = stored.getByRole("button", { name: "Clear chats…", exact: true });
  assert.equal(await stored.evaluate(el => el.open), false);
  await settings.screenshot({ path: path.join(shots, "stored-data-collapsed.png") });
  await stored.locator("summary").click();
  await stored.getByText(/^\d+ documents?$/).first().waitFor();
  assert.ok(await clearNotes.isEnabled() && await clearChats.isEnabled());
  await stored.screenshot({ path: path.join(shots, "stored-data-expanded.png") });
  const confirmation = stored.getByRole("alertdialog");
  for (const [trigger, question] of [[clearNotes, "Clear all saved annotations?"], [clearChats, "Clear all saved chats?"]]) {
    await trigger.click();
    await confirmation.getByText(question, { exact: true }).waitFor();
    assert.ok(await confirmation.getByRole("button", { name: "Cancel" }).evaluate(el => el === document.activeElement));
    await confirmation.getByRole("button", { name: "Cancel" }).click();
    assert.ok(await clearNotes.isEnabled() && await clearChats.isEnabled());
  }
  await clearNotes.click();
  await confirmation.getByRole("button", { name: "Clear annotations", exact: true }).click();
  await stored.getByText("0 documents", { exact: true }).waitFor();
  assert.ok(await clearNotes.isDisabled() && await clearChats.isEnabled());
  await clearChats.click();
  await confirmation.getByRole("button", { name: "Clear chats", exact: true }).click();
  await confirmation.waitFor({ state: "detached" });
  assert.ok(await clearChats.isDisabled());
  await settings.getByRole("button", { name: "Close" }).click();
  await docs.getByRole("tab", { name: readerName, exact: true }).click();
  assert.equal(await documentNote.inputValue(), "");
  assert.equal(await aiPanel.getByText("First question").count(), 0);
  await aiPanel.getByRole("button", { name: "Close Ask AI" }).click();
  console.log("ok  Stored data counts, cancellation, separate confirmations and open readers");

  const longPdf = pdfWithPageSizes(Array.from({ length: 101 }, () => [612, 792]));
  await app.evaluate(({ BrowserWindow }, data) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Print 101 pages.pdf", data: Buffer.from(data) }), [...longPdf]);
  await page.window.waitForFunction(() => document.title.startsWith("Print 101 pages.pdf"));
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: /^Print\b/ }).click();
  await printDialog.getByRole("img", { name: "Preview of page 1" }).waitFor();
  await printDialog.getByRole("tab", { name: "Advanced" }).click();
  await printDialog.getByRole("alert").filter({ hasText: /100 sheets/ }).waitFor();
  await printDialog.getByRole("tab", { name: "Basic" }).click();
  await page.evaluate(interceptPdfPrint);
  await printDialog.getByRole("button", { name: "Print…", exact: true }).click();
  await page.waitForFunction(() => window.printedPdf);
  assert.equal(await page.evaluate(() => window.printedPdf.bytes), longPdf.toString("utf8"));
  await page.evaluate(() => window.restorePrintFrame());
  console.log("ok  ordinary PDF printing remains available beyond the advanced sheet limit");

  // Find positions count a ligature ("ﬁ") as the one character it is in the PDF. The quoted text has to count
  // the same way, or the words it marks drift to the left of the match.
  await app.evaluate(({ BrowserWindow }, data) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Ligatures.pdf", data: Buffer.from(data) }), [...pdfWithLigatures()]);
  await page.window.waitForFunction(() => document.title.startsWith("Ligatures.pdf"));
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor();
  await page.getByRole("button", { name: /^Search/ }).click();
  await page.getByLabel("Find in document").fill("zebra");
  await quoted.first().waitFor();
  assert.deepEqual(await quoted.allTextContents(), ["zebra"]);
  console.log("ok  search results mark the right words in text with ligatures");

  // Asked again at the end: the windows that convert a webpage or a Word document didn't bring it forward either.
  await checkStaysHidden();
  assert.deepEqual(failures, []);
  console.log(`PASS (${process.platform})`);
} finally {
  await app.close();
}

function interceptPdfPrint() {
  const proto = HTMLIFrameElement.prototype;
  const src = Object.getOwnPropertyDescriptor(proto, "src"), content = Object.getOwnPropertyDescriptor(proto, "contentWindow");
  window.printedPdf = null;
  Object.defineProperty(proto, "src", { configurable: true, set(url) { this.dataset.printUrl = url; } });
  Object.defineProperty(proto, "contentWindow", { configurable: true, get() { return {
    focus() {}, print: async () => {
      const response = await fetch(this.dataset.printUrl);
      window.printedPdf = { type: response.headers.get("Content-Type"), bytes: await response.text() };
    },
  }; } });
  window.restorePrintFrame = () => { Object.defineProperty(proto, "src", src); Object.defineProperty(proto, "contentWindow", content); };
}

// Vector artwork makes orientation and clipping visible in the preview recording.
function pdfWithPageSizes(sizes) {
  const fontId = sizes.length + 3;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    `<< /Type /Pages /Kids [${sizes.map((_, i) => `${i + 3} 0 R`).join(" ")}] /Count ${sizes.length} >>`,
    ...sizes.map(([w, h], i) => `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${fontId + 1 + i} 0 R >>`),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ...sizes.map(([w, h], i) => {
      const lines = [
        `0.98 0.97 0.94 rg 0 0 ${w} ${h} re f`,
        `0.13 0.24 0.22 rg BT /F1 16 Tf 40 ${h - 55} Td (NORTH COVE FIELD SURVEY) Tj ET`,
        `BT /F1 34 Tf 40 ${h - 116} Td (Shoreline profile) Tj ET`,
        `0.85 0.90 0.84 rg 40 ${h * .32} ${w - 80} ${h * .35} re f`,
        `0.44 0.61 0.54 rg 40 ${h * .32} m 40 ${h * .47} l ${w * .28} ${h * .42} ${w * .54} ${h * .60} ${w - 40} ${h * .52} c ${w - 40} ${h * .32} l h f`,
        `0.13 0.24 0.22 rg BT /F1 18 Tf 40 ${h * .26} Td (Field observations / Station ${i + 1}) Tj ET`,
      ].join("\n");
      return `<< /Length ${lines.length} >>\nstream\n${lines}\nendstream`;
    }),
  ];
  return pdfFromObjects(objects);
}

// One page whose first line is written with the "fi" ligature glyph, so its text holds the single character "ﬁ"
// eight times before the words on the second line.
function pdfWithLigatures() {
  const text = "BT /F1 14 Tf 72 720 Td (\\001rst \\001ne \\001ve \\001x \\001sh \\001g \\001t \\001ll) Tj 0 -24 Td (target word zebra here) Tj ET";
  return pdfFromObjects([
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [1 /fi] >> >>",
    `<< /Length ${text.length} >>\nstream\n${text}\nendstream`,
  ]);
}

function pdfFromObjects(objects) {
  let body = "%PDF-1.4\n";
  const offsets = objects.map((object, i) => {
    const offset = body.length;
    body += `${i + 1} 0 obj\n${object}\nendobj\n`;
    return offset;
  });
  const xref = body.length;
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  body += offsets.map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("");
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(body, "latin1");
}
