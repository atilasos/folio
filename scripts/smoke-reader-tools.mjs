import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import { deflateSync } from "node:zlib";
import ocrCatalog from "../electron/ocr-models.json" with { type: "json" };
import { PDFDocument, PDFDict, PDFName } from "pdf-lib";

// A two-page form with a text field on its first page.
function formPdf() {
  return testPdf([
    "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] /DA (/Helv 12 Tf 0 g) /DR << /Font << /Helv 6 0 R >> >> >> >>",
    "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << >> /Annots [5 0 R] >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << >> >>",
    "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /Rect [30 180 260 210] /P 3 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ]);
}

function testPdf(objects) {
  let pdf = "%PDF-1.4\n";
  const entries = objects.map((obj, i) => {
    const offset = pdf.length;
    pdf += `${i + 1} 0 obj\n${obj}\nendobj\n`;
    return String(offset).padStart(10, "0") + " 00000 n \n";
  });
  const size = objects.length + 1;
  return pdf + `xref\n0 ${size}\n0000000000 65535 f \n${entries.join("")}trailer\n<< /Size ${size} /Root 1 0 R >>\nstartxref\n${pdf.length}\n%%EOF`;
}

async function checkScanSelection(app, page) {
  const { pixels, right } = await page.evaluate(() => {
    const canvas = Object.assign(document.createElement("canvas"), { width: 1000, height: 500 });
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, 1000, 500);
    ctx.fillStyle = "black";
    ctx.font = "56px sans-serif";
    ctx.fillText("Select these words", 100, 225);
    ctx.fillText("from a scanned page.", 100, 300);
    ctx.fillText("Grüße aus Köln", 100, 375);
    ctx.fillText("El niño español", 100, 450);
    const rgba = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    const bits = new Uint8Array(canvas.width * canvas.height / 8);
    for (let i = 0; i < canvas.width * canvas.height; i++) {
      if (rgba[4 * i] >= 128) bits[i >> 3] |= 0x80 >> (i % 8);
    }
    return { pixels: Array.from(bits), right: 100 + ctx.measureText("Select these words").width };
  });
  const image = deflateSync(Buffer.from(pixels)).toString("hex") + ">";
  // The scan is upright when displayed, on a rotated, offset page with a non-default unit.
  const draw = "q 0 400 -200 0 220 30 cm /Scan Do Q";
  const native = draw + "\nBT /F 12 Tf 30 50 Td (Embedded text) Tj ET";
  const pageDict = "/Type /Page /Parent 2 0 R /MediaBox [20 30 220 430] /Rotate 90 /UserUnit 1.25";
  const resources = scan => `/Resources << /XObject << /Scan ${scan} 0 R >> /Font << /F 8 0 R >> >>`;
  const imageObject = kind => `<< /Type /XObject /Subtype /Image /Width 1000 /Height 500 ${kind} /BitsPerComponent 1 /Filter [/ASCIIHexDecode /FlateDecode] /Length ${image.length} >>\nstream\n${image}\nendstream`;
  const pdf = testPdf([
    "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 5 /Kids [4 0 R 3 0 R 9 0 R 11 0 R 12 0 R] >>",
    `<< ${pageDict} ${resources(5)} /Contents 6 0 R >>`, `<< ${pageDict} ${resources(5)} /Contents 7 0 R >>`,
    imageObject("/ColorSpace /DeviceGray"),
    ...[draw, native].map(s => `<< /Length ${s.length} >>\nstream\n${s}\nendstream`),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< ${pageDict} ${resources(10)} /Contents 6 0 R >>`, imageObject("/ImageMask true"),
    `<< ${pageDict} ${resources(5)} /Contents 6 0 R >>`, `<< ${pageDict} ${resources(5)} /Contents 6 0 R >>`,
  ]);
  // Hold OCR at the worker boundary so queue priority and loading messages are independent of CPU speed.
  await page.window.addInitScript(() => {
    const post = Worker.prototype.postMessage;
    let jobs = 0;
    Worker.prototype.postMessage = function (message, ...rest) {
      if (message?.action !== "recognize") return post.call(this, message, ...rest);
      window.ocrJob = ++jobs;
      window.resumeOcr = () => post.call(this, message, ...rest);
      if (jobs === 4) {
        Worker.prototype.postMessage = post;
        window.failOcr = () => this.dispatchEvent(new MessageEvent("message", {
          data: { action: "recognize", jobId: message.jobId, status: "reject", data: "Test recognition failure" },
        }));
      }
    };
  });
  await app.evaluate(({ BrowserWindow }, pdf) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Mixed scan.pdf", data: Buffer.from(pdf) }), pdf);
  await page.window.waitForFunction(() => document.title.startsWith("Mixed scan.pdf"));
  const scan = page.locator('.page[data-page-number="2"]');
  const pageNumber = page.getByLabel("Page number");
  const goTo = async n => { await pageNumber.fill(String(n)); await pageNumber.press("Enter"); };
  const status = page.locator(".page-text-status");
  const waitForJob = n => page.waitForFunction(n => window.ocrJob === n, n, { timeout: 45_000 });
  await waitForJob(1);
  assert.equal(await pageNumber.inputValue(), "1");
  const mixed = page.locator('.page[data-page-number="1"] .textLayer');
  await mixed.getByText("Embedded text", { exact: true }).waitFor();
  assert.equal((await mixed.innerText()).trim(), "Embedded text");
  await goTo(2);
  await status.filter({ hasText: "Recognizing text on page 2…" }).waitFor();
  await page.evaluate(() => window.resumeOcr());
  const firstWord = scan.locator(".textLayer span", { hasText: /^Select\s*$/ });
  await firstWord.waitFor({ timeout: 45_000 });
  await waitForJob(2); // Page 3 is already being recognized while the reader is still on page 2.
  assert.equal(await pageNumber.inputValue(), "2");
  assert.equal(await status.innerText(), "");
  await page.getByTitle(/^Zoom in/).click();
  await page.getByRole("button", { name: "More" }).click();
  await page.getByRole("menuitemcheckbox", { name: "PDF dark mode" }).click();
  await scan.locator(".canvasWrapper > .folio-image-overlay").waitFor();
  await firstWord.scrollIntoViewIfNeeded();
  const box = await scan.locator(".canvasWrapper").boundingBox();
  const first = await firstWord.boundingBox();
  const last = await scan.locator(".textLayer span", { hasText: /^words$/ }).boundingBox();
  assert.ok(Math.abs(first.x - box.x - box.width * 0.1) < box.width * 0.01);
  assert.ok(Math.abs(last.x + last.width - box.x - box.width * right / 1000) < box.width * 0.01);
  const y = box.y + box.height * 205 / 500;
  const clip = { x: Math.ceil(first.x + 2), y: Math.floor(y - 2), width: Math.floor(first.width / 2), height: 4 };
  const unselected = await page.screenshot({ clip });
  // Readers often drag from just before the first letter to just past the last.
  await page.mouse.move(first.x - 5, y);
  await page.mouse.down();
  await page.mouse.move(last.x + last.width + 5, y, { steps: 12 });
  await page.waitForFunction(() => getSelection().toString().trim() === "Select these words");
  assert.ok(!unselected.equals(await page.screenshot({ clip })), "the scan hid the live selection in PDF dark mode");
  await page.mouse.up();
  await page.getByRole("dialog", { name: "Selected text" }).waitFor();
  assert.equal((await page.evaluate(() => getSelection().toString())).trim(), "Select these words");
  await page.getByRole("button", { name: "Highlight yellow" }).click();
  await scan.locator(".highlightEditor").waitFor();
  await goTo(5);
  await status.filter({ hasText: "Preparing text on page 5…" }).waitFor();
  await page.evaluate(() => window.resumeOcr());
  await waitForJob(3);
  assert.equal(await status.innerText(), "Recognizing text on page 5…");
  await page.evaluate(() => window.resumeOcr());
  await page.locator('.page[data-page-number="5"] .ocr-word').first().waitFor();
  assert.equal(await status.innerText(), "");
  await goTo(3);
  await page.locator('.page[data-page-number="3"] .textLayer').getByText("Select", { exact: true }).waitFor();
  await waitForJob(4);
  await goTo(4);
  await page.evaluate(() => window.failOcr());
  await status.filter({ hasText: "Text unavailable on page 4." }).waitFor();
  await goTo(3);
  await goTo(4);
  await page.locator('.page[data-page-number="4"] .ocr-word').first().waitFor();
  assert.equal(await status.innerText(), "");
  console.log("ok  background OCR prepares unvisited scans, prioritizes page jumps, and shows selectable text in dark mode");

  // Exercise the real desktop download and checksum path without reaching GitHub from CI.
  await app.evaluate(({ net }, { base, fixtures }) => {
    const original = net.fetch;
    const requests = [];
    net.fetch = async (url, options) => {
      if (!String(url).startsWith(base)) return original.call(net, url, options);
      const file = url.slice(base.length);
      if (!Object.hasOwn(fixtures, file)) throw new Error("Unexpected language download");
      requests.push(file);
      const bytes = process.getBuiltinModule("fs").readFileSync(fixtures[file]);
      if (requests.length === 1) bytes[0] ^= 1;
      return new Response(bytes);
    };
    globalThis.finishModelDownloads = () => { net.fetch = original; return requests; };
  }, { base: ocrCatalog.baseUrl, fixtures: Object.fromEntries(["spa", "deu"].map(code => {
    const file = code + ".traineddata.gz";
    return [file, path.resolve(import.meta.dirname, "../node_modules/@tesseract.js-data", code, "4.0.0_best_int", file)];
  })) });
  const settingsShortcut = process.platform === "darwin" ? "Meta+Comma" : "Control+Comma";
  await page.keyboard.press(settingsShortcut);
  const picker = page.getByRole("button", { name: /^Scan languages:/ });
  await picker.click();
  const languages = page.getByRole("dialog", { name: "Choose scan languages" });
  const search = languages.getByRole("searchbox");
  const apply = languages.getByRole("button", { name: "Apply" });
  const download = languages.getByRole("button", { name: /^Download ·/ });
  assert.equal(await languages.getByRole("checkbox").count(), Object.keys(ocrCatalog.languages).length);
  assert.ok(await languages.getByRole("checkbox", { name: "English", exact: true }).isChecked());
  await languages.getByRole("checkbox", { name: "English", exact: true }).uncheck();
  assert.ok(await apply.isDisabled());
  await search.fill("espanol");
  await languages.getByRole("checkbox", { name: "Spanish", exact: true }).check();
  await download.click();
  await languages.getByRole("alert").waitFor();
  assert.equal(await picker.innerText(), "English", "a damaged pack must not replace the active selection");
  await download.click();
  await languages.waitFor({ state: "hidden" });
  await page.keyboard.press("Escape");
  const currentScan = page.locator('.page[data-page-number="4"] .textLayer');
  await currentScan.getByText("niño", { exact: true }).waitFor({ timeout: 45_000 });
  await page.keyboard.press(settingsShortcut);
  await picker.click();
  await search.fill("Deutsch");
  await languages.getByRole("checkbox", { name: "German", exact: true }).check();
  await download.click();
  await languages.waitFor({ state: "hidden" });
  await page.keyboard.press("Escape");
  await currentScan.getByText("Grüße", { exact: true }).waitFor({ timeout: 45_000 });
  await currentScan.getByText("niño", { exact: true }).waitFor();
  assert.equal(await page.evaluate(() => localStorage.getItem("folio:scan-language")), "spa+deu");
  await page.keyboard.press(settingsShortcut);
  await picker.click();
  await languages.getByRole("checkbox", { name: "German", exact: true }).uncheck();
  // Change languages again while the preceding run's scan image is still being encoded.
  await page.evaluate(() => {
    const toBlob = HTMLCanvasElement.prototype.toBlob;
    HTMLCanvasElement.prototype.toBlob = function (callback, ...options) {
      if (this.isConnected) return toBlob.call(this, callback, ...options);
      HTMLCanvasElement.prototype.toBlob = toBlob;
      return toBlob.call(this, blob => { window.finishScanEncoding = () => callback(blob); }, ...options);
    };
  });
  await apply.click();
  await languages.waitFor({ state: "hidden" });
  await page.waitForFunction(() => typeof window.finishScanEncoding === "function");
  await picker.click();
  await search.fill("German");
  await languages.getByRole("checkbox", { name: "German", exact: true }).check();
  await apply.click();
  await languages.waitFor({ state: "hidden" });
  await page.evaluate(() => window.finishScanEncoding());
  await page.keyboard.press("Escape");
  await currentScan.getByText("Grüße", { exact: true }).waitFor({ timeout: 45_000 });
  assert.deepEqual(await app.evaluate(() => globalThis.finishModelDownloads()), ["spa.traineddata.gz", "spa.traineddata.gz", "deu.traineddata.gz"]);
  await goTo(1);
  assert.equal((await mixed.innerText()).trim(), "Embedded text");
  console.log("ok  scan languages verify downloads, retry a corrupt pack, and reuse saved models for mixed-language scans");
}

async function captureDownload(app, file) {
  await app.evaluate(({ session }, file) => {
    globalThis.folioReordered = new Promise(resolve => session.defaultSession.once("will-download", (_, item) => {
      item.setSavePath(file);
      item.once("done", (_, state) => resolve(state));
    }));
  }, file);
}

async function checkPenDrawing(app, page, folder, shots) {
  const blank = testPdf(["<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 480 320] /Resources << >> >>"]);
  await app.evaluate(({ BrowserWindow }, pdf) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Pen widths.pdf", data: Buffer.from(pdf) }), blank);
  await page.window.waitForFunction(() => document.title.startsWith("Pen widths.pdf"));
  await page.getByTitle("Draw (D)").click();
  const width = page.getByRole("combobox", { name: "Pen width" });
  const layer = page.locator(".page .annotationEditorLayer.inkEditing");
  await layer.waitFor();
  assert.equal(await width.inputValue(), "1");
  const zoom = page.getByTitle("Zoom", { exact: true });
  for (const [index, name] of ["Thin", "Medium", "Thick"].entries()) {
    await width.selectOption({ label: name });
    if (name === "Thick") {
      await zoom.click();
      await page.getByRole("menuitem", { name: "Fit width" }).click();
      await layer.waitFor();
    }
    const box = await layer.boundingBox();
    const x = box.x + box.width * 0.2;
    const y = box.y + box.height * (0.3 + index * 0.15);
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 25, y - 10);
    await page.mouse.move(x + 50, y + 10);
    await page.mouse.move(x + 100, y);
    const initialZoom = await zoom.innerText();
    if (name === "Thick") {
      // Keep the pen down through the sidebar's delayed automatic refit.
      await page.keyboard.press("ControlOrMeta+\\");
      await page.waitForTimeout(450);
    }
    const heldZoom = await zoom.innerText();
    const releaseBox = await layer.boundingBox();
    await page.mouse.up();
    assert.ok(releaseBox, "the ink layer was hidden at pointer release");
    // An older ink fitter showed the tip while drawing, then pulled it halfway back on release.
    const end = await page.locator(".page svg.draw").last().evaluate(svg => {
      const path = svg.querySelector("path");
      const tip = path.getPointAtLength(path.getTotalLength());
      const bounds = svg.viewBox.baseVal;
      return { x: (tip.x - bounds.x) / bounds.width, y: (tip.y - bounds.y) / bounds.height };
    });
    await page.screenshot({ path: path.join(shots, `pen-release-${name.toLowerCase()}.png`) });
    const dx = releaseBox.x + end.x * releaseBox.width - (x + 100);
    const dy = releaseBox.y + end.y * releaseBox.height - y;
    assert.ok(Math.hypot(dx, dy) < 4, `${name} pen ended away from the release point by (${dx}, ${dy})`);
    if (name === "Thick") {
      assert.equal(heldZoom, initialZoom, "automatic fit resized the page during a pen stroke");
      await page.waitForFunction(previous => document.querySelector(".zoom-value").textContent !== previous, initialZoom);
      await page.getByRole("button", { name: /Toggle sidebar/ }).click();
    }
  }
  const tool = page.getByRole("combobox", { name: "Drawing tool" });
  await tool.selectOption("Rectangle");
  assert.equal(await width.count(), 0);
  await tool.selectOption("Pen");
  assert.equal(await width.inputValue(), "6");
  await page.getByRole("button", { name: "Done", exact: true }).click();
  const file = path.join(folder, "pen-widths.pdf");
  await captureDownload(app, file);
  await page.keyboard.press("ControlOrMeta+s");
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  const savedWidths = await page.evaluate(async bytes => {
    const task = pdfjsLib.getDocument({ data: Uint8Array.from(bytes) });
    try {
      const pdf = await task.promise;
      const annotations = await (await pdf.getPage(1)).getAnnotations();
      return annotations.filter(a => a.subtype === "Ink").map(a => a.borderStyle.rawWidth).sort((a, b) => a - b);
    } finally { await task.destroy(); }
  }, Array.from(readFileSync(file)));
  assert.deepEqual(savedWidths, [1, 3, 6]);
  console.log("ok  freehand strokes keep their release points and three independent pen widths in the saved PDF");
}

async function checkContentEditing(app, page, folder, shots) {
  // Two identical text runs and two placements of one image, also reused on the second page.
  const words = "BT /F 16 Tf 12 170 Td (Orig) Tj 29.34 0 Td (inal) Tj -29.34 -25 Td (Original) Tj ET";
  const images = "% /Logo Do\n/Artifact << /Details (literal /Logo Do) >> BDC q 10 20 24 40 re W n "
    + "40 0 0 40 10 20 cm /Logo Do Q EMC\nq 40 0 0 40 80 20 cm /Logo Do Q";
  const pageDict = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources 10 0 R /Contents 9 0 R >>";
  const stream = s => `<< /Length ${s.length} >>\nstream\n${s}\nendstream`;
  const original = testPdf([
    "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
    pageDict, pageDict,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 7 >>\nstream\nff0000>\nendstream",
    stream(words), stream(images), "[7 0 R 8 0 R]", "<< /Font << /F 5 0 R >> /XObject << /Logo 6 0 R >> >>",
  ]);
  await app.evaluate(({ BrowserWindow }, file) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Objects.pdf", data: Buffer.from(file) }), original);
  await page.window.waitForFunction(() => document.title.startsWith("Objects.pdf"));
  const sidebar = page.getByRole("button", { name: /Toggle sidebar/ });
  if (await sidebar.getAttribute("aria-pressed") !== "true") await sidebar.click();
  await page.getByRole("tab", { name: "Pages", exact: true }).click();
  await page.getByRole("button", { name: "Bookmark page 2", exact: true }).click();
  const openEditor = async () => {
    await page.getByRole("button", { name: "Edit mode", exact: true }).click();
    assert.equal(await page.getByRole("dialog").count(), 0);
  };
  const sheet = page.locator('.page[data-page-number="1"]');
  const item = name => sheet.getByRole("button", { name, exact: true });
  const input = sheet.getByRole("textbox", { name: "Edit text on page" });
  const save = page.getByRole("button", { name: "Save copy", exact: true });
  await sheet.locator(".textLayer").getByText("Original", { exact: true }).first().waitFor();
  await page.screenshot({ path: path.join(shots, "before-pdf-edit.png") });
  await openEditor();
  const rotation = page.getByTitle("Rotate (R)");
  await rotation.click();
  await page.keyboard.press("ControlOrMeta+=");
  const checkRotatedTarget = () => page.waitForFunction(() => {
    const p = document.querySelector('.page[data-page-number="1"]');
    const canvas = p?.querySelector(".canvasWrapper")?.getBoundingClientRect();
    const logo = p?.querySelector('button[aria-label="Image 1"]')?.getBoundingClientRect();
    return canvas && logo && Math.abs((logo.left - canvas.left) / canvas.width - 0.1) < 0.012
      && Math.abs((logo.top - canvas.top) / canvas.height - 0.05) < 0.012;
  }, null, { timeout: 15_000 });
  await checkRotatedTarget();
  const pageWidth = await sheet.locator(".canvasWrapper").evaluate(el => el.getBoundingClientRect().width);
  await item("Original").first().click();
  await input.fill("Cancel me"); await input.press("Escape");
  assert.equal(await input.count(), 0);
  await item("Original").first().click();
  await input.fill("𐍈"); await input.press("Enter");
  await page.getByText(/This font can.t display all of those characters/).waitFor();
  await input.fill("Revised");
  await page.getByRole("combobox", { name: "Text font", exact: true }).selectOption("easyjet");
  await item("Image 1").click();
  await item("Revised").waitFor();
  await sheet.locator(".textLayer").getByText("Revised", { exact: true }).waitFor();
  await sheet.getByRole("button", { name: /Replace image/ }).waitFor();
  await checkRotatedTarget();
  assert.ok(Math.abs(await sheet.locator(".canvasWrapper").evaluate(el => el.getBoundingClientRect().width) - pageWidth) < 2);
  const png = await page.evaluate(() => {
    const canvas = Object.assign(document.createElement("canvas"), { width: 2, height: 2 });
    const g = canvas.getContext("2d");
    g.fillStyle = "#0000ff"; g.fillRect(0, 0, 2, 2);
    return canvas.toDataURL().split(",")[1];
  });
  await sheet.getByLabel("Replacement image").setInputFiles({ name: "blue.png", mimeType: "image/png", buffer: Buffer.from(png, "base64") });
  await page.getByRole("button", { name: "Exit edit mode" }).click();
  await page.getByRole("button", { name: "Edit mode", exact: true }).waitFor();
  assert.equal(await sheet.locator(".folio-content-layer").count(), 0);
  await openEditor();
  const file = path.join(folder, "edited.pdf");
  await captureDownload(app, file);
  await save.click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  await page.window.waitForFunction(() => document.title.startsWith("Objects (edited).pdf"));
  const exported = await PDFDocument.load(readFileSync(file));
  const embedded = exported.context.enumerateIndirectObjects().some(([, obj]) => obj instanceof PDFDict
    && obj.get(PDFName.of("FontName"))?.toString().includes("easyJetRoundedBook")
    && obj.has(PDFName.of("FontFile2")));
  assert.ok(embedded, "the edited PDF embeds easyJet Rounded Book for other readers");
  await page.getByRole("button", { name: "Bookmark page 2", pressed: true }).waitFor();
  // Read and rasterize the downloaded bytes with the independent reader, including the untouched page.
  const result = await page.evaluate(async bytes => {
    const task = pdfjsLib.getDocument({ data: Uint8Array.from(bytes) });
    const pdf = await task.promise;
    try {
      const pages = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const p = await pdf.getPage(i), canvas = document.createElement("canvas");
        canvas.width = 200; canvas.height = 200;
        await p.render({ canvas, viewport: p.getViewport({ scale: 1 }) }).promise;
        const g = canvas.getContext("2d");
        pages.push({ text: (await p.getTextContent()).items.map(x => x.str).filter(Boolean),
          images: [30, 45, 100].map(x => [...g.getImageData(x, 160, 1, 1).data]), annotations: (await p.getAnnotations()).length });
      }
      return pages;
    } finally { await task.destroy(); }
  }, Array.from(readFileSync(file)));
  const red = [255, 0, 0, 255], blue = [0, 0, 255, 255], white = [255, 255, 255, 255];
  assert.deepEqual(result, [
    { text: ["Revised", "Original"], images: [blue, white, red], annotations: 0 },
    { text: ["Original", "Original"], images: [red, white, red], annotations: 0 },
  ]);
  console.log("ok  edited copy contains replacement text and image pixels only at the chosen placements");

  await openEditor();
  await item("Original").click();
  await input.fill("");
  await captureDownload(app, path.join(folder, "removed-text.pdf"));
  await page.keyboard.press("ControlOrMeta+s");
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  await page.window.waitForFunction(() => document.title.startsWith("Objects (edited) (edited).pdf"));
  const removed = page.locator('.page[data-page-number="1"] .textLayer');
  await removed.getByText("Revised", { exact: true }).waitFor();
  assert.equal((await removed.innerText()).trim(), "Revised");
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: "Version history" }).click();
  const history = page.getByRole("complementary", { name: "Version history" });
  await history.getByRole("heading", { name: "Version 3 This PDF" }).waitFor();
  assert.equal(await history.locator(".version-entry").count(), 3);
  await history.getByText("blue.png").waitFor();
  await page.screenshot({ path: path.join(shots, "after-pdf-edit.png") });
  assert.equal(await history.locator("ins").innerText(), "Revised");
  const recovered = path.join(folder, "original-version.pdf");
  await captureDownload(app, recovered);
  await history.getByRole("button", { name: "Download version 1" }).click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  assert.deepEqual(readFileSync(recovered), Buffer.from(original));
  await history.getByRole("button", { name: "Open version 1" }).click();
  await page.window.waitForFunction(() => document.title.startsWith("Objects.pdf"));
  await sheet.locator(".textLayer").getByText("Original", { exact: true }).first().waitFor();
  console.log("ok  version history groups edits by save and recovers the exact original");

  await app.evaluate(({ BrowserWindow }, data) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Text runs.pdf", data: Buffer.concat([Buffer.from(data), Buffer.from("\n% Inline editing fixture\n")]) }),
    await page.evaluate(async () => Array.from(new Uint8Array(await (await fetch("sample.pdf")).arrayBuffer()))));
  await page.window.waitForFunction(() => document.title.startsWith("Text runs.pdf"));
  const number = page.getByLabel("Page number", { exact: true });
  await number.fill("1"); await number.press("Enter");
  await openEditor();
  await sheet.getByRole("button", { name: /A short tour of Folio/ }).first().click();
  await input.fill("A tour of Folio"); await input.press("Enter");
  await item("A tour of Folio").waitFor();
  await captureDownload(app, path.join(folder, "edited-sample.pdf"));
  await save.click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  await page.window.waitForFunction(() => document.title.startsWith("Text runs (edited).pdf"));
  await page.locator('.page[data-page-number="1"] .textLayer').getByText("A tour of Folio", { exact: true }).waitFor();
  console.log("ok  fragmented embedded-font text in Folio's sample edits as a line and reopens in the reader");
}

export function linkHistoryPdf() {
  const content = "BT /F 16 Tf 40 1100 Td (Follow a PDF link, then go Back.) Tj ET";
  return testPdf([
    "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Names [(second) [4 0 R /XYZ 0 1000 1.5] (third) [5 0 R /XYZ 0 960 1.25]] >> >> /Outlines 13 0 R >>",
    "<< /Type /Pages /Count 3 /Kids [3 0 R 4 0 R 5 0 R] >>",
    ...["7 0 R 8 0 R 12 0 R 15 0 R", "9 0 R", "10 0 R"].map(annots => `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 1200] /Resources << /Font << /F 6 0 R >> >> /Contents 11 0 R /Annots [${annots}] >>`),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    "<< /Type /Annot /Subtype /Link /Rect [40 1000 160 1040] /Dest (second) >>",
    "<< /Type /Annot /Subtype /Link /Rect [200 1000 320 1040] /Dest (third) >>",
    "<< /Type /Annot /Subtype /Link /Rect [40 1000 160 1040] /Dest (third) >>",
    "<< /Type /Annot /Subtype /Link /Rect [40 1000 160 1040] /A << /S /Named /N /GoBack >> >>",
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    "<< /Type /Annot /Subtype /Link /Rect [360 1000 480 1040] /A << /S /Named /N /GoForward >> >>",
    "<< /Type /Outlines /First 14 0 R /Last 14 0 R /Count 1 >>",
    "<< /Title (Third page) /Parent 13 0 R /Dest (third) >>",
    "<< /Type /Annot /Subtype /Link /Rect [40 900 160 940] /Dest [3 0 R /XYZ 0 500 1.25] >>",
  ]);
}

async function checkPageExport(app, page, folder, shots) {
  await app.evaluate(({ BrowserWindow }, pdf) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Destinations.pdf", data: Buffer.from(pdf) }), linkHistoryPdf() + "\n% page export");
  await page.window.waitForFunction(() => document.title.startsWith("Destinations.pdf"));
  const toggle = page.getByRole("button", { name: /Toggle sidebar/ });
  if (await toggle.getAttribute("aria-pressed") !== "true") await toggle.click();
  await page.getByRole("tab", { name: "Pages", exact: true }).click();
  const sidebar = page.locator("aside.sidebar");
  const thumb = n => sidebar.getByRole("button", { name: "Page " + n, exact: true });
  const picked = () => sidebar.locator('.thumb[aria-pressed="true"]').evaluateAll(els => els.map(el => Number(el.dataset.page)));
  const modifier = process.platform === "darwin" ? "Meta" : "Control";
  await thumb(1).waitFor();
  assert.equal(await sidebar.getByRole("button", { name: /export|extract/i }).count(), 0);
  await sidebar.screenshot({ path: path.join(shots, "export-pages-before.png") });
  await thumb(3).click();
  await thumb(1).click({ modifiers: ["Shift"] });
  assert.deepEqual(await picked(), [1, 2, 3]);
  await thumb(2).click({ modifiers: [modifier] });
  assert.deepEqual(await picked(), [1, 3]);
  const action = page.getByRole("menuitem", { name: "Export 2 selected pages…" });
  await thumb(1).click({ button: "right" });
  await action.waitFor();
  await page.keyboard.press("Escape");
  assert.ok(await thumb(1).evaluate(el => el === document.activeElement));
  assert.deepEqual(await picked(), [1, 3]);

  // Right-clicking a different thumbnail targets just that page. An empty selection stays empty.
  await thumb(2).click({ button: "right" });
  await page.getByRole("menuitem", { name: "Export this page…" }).waitFor();
  await page.keyboard.press("Escape");
  await thumb(2).click({ modifiers: [modifier] });
  assert.deepEqual(await picked(), []);
  await thumb(3).click({ modifiers: [modifier] });
  await thumb(1).click({ modifiers: [modifier] });
  assert.deepEqual(await picked(), [1, 3]);
  await thumb(1).press("Shift+F10");
  await action.waitFor();
  const panel = await sidebar.boundingBox(), popup = await action.boundingBox();
  await page.window.screenshot({ path: path.join(shots, "export-pages-menu.png"), clip: {
    x: panel.x, y: panel.y, width: Math.max(panel.width, popup.x + popup.width + 10 - panel.x), height: panel.height,
  } });
  const file = path.join(folder, "Selected pages.pdf");
  await captureDownload(app, file);
  await action.click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  const result = await page.evaluate(async bytes => {
    const task = pdfjsLib.getDocument({ data: Uint8Array.from(bytes) });
    try {
      const pdf = await task.promise;
      const destination = async dest => {
        if (typeof dest === "string") dest = await pdf.getDestination(dest);
        if (!dest) return null;
        return typeof dest[0] === "number" ? dest[0] : await pdf.getPageIndex(dest[0]);
      };
      const links = await (await pdf.getPage(1)).getAnnotations();
      const linkAt = x => links.find(link => link.rect[0] === x && link.rect[1] === 1000);
      return { count: pdf.numPages, retained: await destination(linkAt(200)?.dest),
        omitted: await destination(linkAt(40)?.dest), outline: await destination((await pdf.getOutline())[0].dest),
        last: (await (await pdf.getPage(2)).getAnnotations()).some(link => link.action === "GoBack") };
    } finally { await task.destroy(); }
  }, Array.from(readFileSync(file)));
  assert.deepEqual(result, { count: 2, retained: 1, omitted: null, outline: 1, last: true });
  for (const n of [1, 3]) await sidebar.getByRole("button", { name: "Bookmark page " + n, exact: true }).click();
  await sidebar.getByRole("button", { name: "Bookmarked", exact: true }).click();
  await thumb(1).click();
  await thumb(3).click({ modifiers: ["Shift"] });
  assert.deepEqual(await picked(), [1, 3], "a filtered range can only export visible bookmarks");
  await page.getByRole("tab", { name: "Outline", exact: true }).click();
  assert.equal(await page.getByRole("menuitem").count(), 0);
  await page.getByRole("button", { name: "Close Destinations.pdf", exact: true }).click();
  await page.getByRole("tab", { name: "Destinations.pdf", exact: true }).waitFor({ state: "detached" });
  console.log("ok  selected pages export in document order with remapped links and outline destinations");
}

async function checkLinkHistory(app, page, shots) {
  const viewport = await page.window.evaluate(() => ({ width: innerWidth, height: innerHeight }));
  await page.window.setViewportSize({ width: 1000, height: 800 });
  const pdf = linkHistoryPdf();
  const open = name => app.evaluate(({ BrowserWindow }, { name, pdf }) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name, data: Buffer.from(pdf) }), { name, pdf });
  await open("Link history.pdf");
  await page.window.waitForFunction(() => document.title.startsWith("Link history.pdf"));
  const more = page.getByRole("button", { name: "More", exact: true });
  const back = page.getByRole("menuitem", { name: /^Back\b/ });
  const forward = page.getByRole("menuitem", { name: /^Forward\b/ });
  const choose = async item => { await more.click(); await item.click(); };
  const checkMenu = async (canBack, canForward) => {
    await more.click();
    assert.equal(await back.count(), Number(canBack));
    assert.equal(await forward.count(), Number(canForward));
    await page.keyboard.press("Escape");
  };
  const pageNumber = page.getByLabel("Page number");
  const at = n => page.waitForFunction(n => document.querySelector('[aria-label="Page number"]').value === String(n), n);
  await checkMenu(false, false);
  const link = (n, name) => page.locator(`.page[data-page-number="${n}"] .annotationLayer a[href$="#${name}"]`);
  const position = () => page.locator(".viewer-container").evaluate(el => ({ top: el.scrollTop, left: el.scrollLeft }));
  const zoom = () => page.getByTitle("Zoom", { exact: true }).innerText();
  await link(1, "second").waitFor();
  await page.locator(".viewer-container").evaluate(el => { el.scrollTop = 50; });
  const start = await position(), startZoom = await zoom();
  const record = process.env.FOLIO_LINK_VIDEO === "1";
  if (record) {
    await page.window.screencast.start({ path: path.join(shots, "link-history.webm"), size: { width: 1000, height: 800 }, fps: 12 });
    await page.waitForTimeout(800);
  }
  await link(1, "second").click(); await at(2);
  const hint = page.getByRole("button", { name: "Back to page 1", exact: true });
  await hint.waitFor();
  if (record) await page.waitForTimeout(900);
  await more.click(); await back.waitFor();
  assert.equal(await forward.count(), 0);
  await page.screenshot({ path: path.join(shots, "link-history-menu-and-hint.png") });
  if (record) await page.waitForTimeout(800);
  await page.keyboard.press("Escape");
  await hint.click(); await at(1);
  if (record) {
    await more.click(); await forward.waitFor();
    await page.waitForTimeout(1000);
    await page.window.screencast.stop();
    await page.keyboard.press("Escape");
  }
  assert.ok(Math.abs((await position()).top - start.top) <= 3);
  await page.locator('.page[data-page-number="1"] .annotationLayer a').nth(3).click();
  assert.equal(await page.locator(".link-return").count(), 0, "the first-use hint repeated");
  await choose(back); await at(1);
  assert.ok(Math.abs((await position()).top - start.top) <= 3);
  await link(1, "second").click(); await at(2);
  await checkMenu(true, false);
  await page.locator(".viewer-container").evaluate(el => { el.scrollTop += 60; el.scrollLeft = 60; });
  const target = await position(), targetZoom = await zoom();
  assert.ok(target.left > 0, "the fixture exercises horizontal scrolling");
  await page.locator(".viewer-container").focus();
  await page.keyboard.press("Alt+ArrowLeft"); await at(1);
  assert.ok(Math.abs((await position()).top - start.top) <= 3);
  assert.equal(await zoom(), startZoom);
  await page.keyboard.press("Alt+ArrowRight"); await at(2);
  assert.ok(Math.abs((await position()).top - target.top) <= 3);
  assert.ok(Math.abs((await position()).left - target.left) <= 3);
  assert.equal(await zoom(), targetZoom);
  // A shortcut in an editable control or dialog must not traverse PDF locations.
  await pageNumber.focus(); await page.keyboard.press("Alt+ArrowLeft");
  assert.equal(await pageNumber.inputValue(), "2");
  await page.locator(".viewer-container").focus();
  await page.keyboard.press("?");
  const shortcuts = page.getByRole("dialog", { name: "Keyboard shortcuts" });
  await shortcuts.getByText("Back / forward through links", { exact: true }).waitFor();
  await page.keyboard.press("Alt+ArrowLeft");
  assert.equal(await pageNumber.inputValue(), "2");
  await shortcuts.waitFor(); await page.keyboard.press("Escape");
  await choose(back); await at(1);
  await link(1, "third").click(); await at(3);
  await checkMenu(true, false);
  await page.locator('.page[data-page-number="3"][data-loaded="true"] .canvasWrapper canvas').waitFor();
  await page.screenshot({ path: path.join(shots, "after-link-history.png") });
  await choose(back); await at(1);
  await choose(forward); await at(3);
  const namedBack = page.locator('.page[data-page-number="3"] .annotationLayer a');
  // Forward changed the zoom, so page 3 is being redrawn and its links are hidden until that finishes. A click
  // that starts before then fails its first attempt, and Playwright's retry scrolls the link to the bottom of
  // the window, where page 2 fills more of the view than page 3. Back then records page 2, not page 3.
  await namedBack.waitFor();
  await namedBack.click(); await at(1);
  await page.locator('.page[data-page-number="1"] .annotationLayer a').nth(2).click(); await at(3);
  await choose(back); await at(1);
  const sidebar = page.getByRole("button", { name: /Toggle sidebar/ });
  if (await sidebar.getAttribute("aria-pressed") !== "true") await sidebar.click();
  await page.getByRole("tab", { name: "Outline" }).click();
  await page.getByRole("button", { name: "Third page", exact: true }).click(); await at(3);
  await choose(back); await at(1);
  const sourceTitle = await page.title();
  await page.getByRole("button", { name: "More", exact: true }).click();
  await page.getByRole("menuitem", { name: /New tab/ }).click();
  await page.getByRole("button", { name: /Try a sample/ }).click();
  await checkMenu(false, false);
  await page.keyboard.press("Control+Shift+Tab");
  assert.equal(await page.title(), sourceTitle);
  await checkMenu(false, true);
  await page.locator(".viewer-container").focus();
  await page.keyboard.press("Alt+ArrowRight"); await at(3);
  await page.window.setViewportSize({ width: 420, height: 800 });
  await page.window.waitForFunction(() => innerWidth <= 420);
  const box = await page.getByRole("toolbar", { name: "Navigation" }).boundingBox();
  const width = await page.window.evaluate(() => innerWidth);
  assert.ok(box && box.x >= 0 && box.x + box.width <= width, `the navigation controls fit a narrow window: ${JSON.stringify({ width, box })}`);
  await page.screenshot({ path: path.join(shots, "link-history-narrow.png") });
  await page.window.setViewportSize(viewport);
  console.log("ok  internal links and outlines support a one-time Back hint, contextual More commands, per-tab position and zoom, shortcuts and forward branching");
}

export async function checkReaderTools(app, page, folder, shots) {
  const settings = () => page.keyboard.press(process.platform === "darwin" ? "Meta+Comma" : "Control+Comma");
  const wheelSwitch = page.getByRole("switch", { name: "Scroll one page at a time" });
  const wheel = options => page.locator(".viewer-container").evaluate((el, options) =>
    el.dispatchEvent(new WheelEvent("wheel", { bubbles: true, cancelable: true, ...options })), options);
  const atPage = n => page.waitForFunction(n => document.querySelector('[aria-label="Page number"]').value === String(n), n);
  await settings();
  assert.equal(await wheelSwitch.getAttribute("aria-checked"), "false");
  await wheelSwitch.click();
  await page.keyboard.press("Escape");
  await page.locator(".viewer-container").hover();
  await page.mouse.wheel(0, 100);
  await atPage(2);
  await wheel({ deltaY: 3, deltaMode: 1 });
  await atPage(3);
  await wheel({ deltaY: -1, deltaMode: 2 });
  await atPage(2);
  assert.equal(await wheel({ deltaX: 100 }), true);
  await page.locator(".viewer-container").evaluate(el => {
    for (const deltaY of [5, 10, 12, 100, 8, 3, 1]) el.dispatchEvent(new WheelEvent("wheel", { deltaY, cancelable: true }));
  });
  await atPage(3);
  await page.reload();
  await page.locator(".textLayer span:not(.markedContent)").first().waitFor();
  await settings();
  assert.equal(await wheelSwitch.getAttribute("aria-checked"), "true");
  await wheelSwitch.click();
  await page.keyboard.press("Escape");
  assert.equal(await wheel({ deltaY: 3, deltaMode: 1 }), true);
  console.log("ok  wheel notches and trackpad swipes turn pages, and the setting survives reopening");

  // A deterministic response exercises the real stream and buttons without signing in or sending document data.
  const answer = "**Two points** (p. 1)\n\n- First\n- Second";
  await app.evaluate(({ ipcMain }, answer) => {
    for (const name of ["folio:ai-status", "folio:ai-chatgpt-models", "folio:ai-ask"]) ipcMain.removeHandler(name);
    ipcMain.handle("folio:ai-status", () => ({ chatgpt: { connected: true, welcomeSeen: true }, claude: { installed: false } }));
    ipcMain.handle("folio:ai-chatgpt-models", () => ({ ok: true, models: [{ id: "local", name: "Local test" }] }));
    ipcMain.handle("folio:ai-ask", (event, { id }) => {
      event.sender.send("folio:ai-event", { id, type: "delta", text: answer });
      event.sender.send("folio:ai-event", { id, type: "done" });
      return true;
    });
  }, answer);
  await page.getByRole("button", { name: "Ask AI", exact: true }).click();
  await page.getByRole("button", { name: "Summarize this document" }).click();
  await page.evaluate(() => {
    navigator.clipboard.writeText = async text => { window.folioCopiedResponse = text; };
  });
  const copyResponse = page.getByTitle("Copy response as Markdown");
  await copyResponse.click();
  await page.getByRole("button", { name: "Copied", exact: true }).waitFor();
  assert.equal(await page.evaluate(() => window.folioCopiedResponse), answer);
  await page.evaluate(() => { navigator.clipboard.writeText = async () => { throw new Error("Denied"); }; });
  await copyResponse.click();
  await page.getByRole("button", { name: "Copy failed" }).waitFor();
  await page.evaluate(() => { delete navigator.clipboard.writeText; });
  const newConversation = page.getByRole("button", { name: "New conversation", exact: true });
  const composer = page.locator(".composer textarea");
  const confirmation = page.getByRole("dialog", { name: "Start a new conversation?" });
  const savedAnswer = () => page.evaluate(answer => new Promise((resolve, reject) => {
    const request = indexedDB.open("folio");
    request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      const db = request.result, rows = db.transaction("ai").objectStore("ai").getAll();
      rows.onsuccess = () => { db.close(); resolve(rows.result.some(row => row.messages.some(m => m.content === answer))); };
      rows.onerror = () => { db.close(); reject(rows.error); };
    };
  }), answer);
  assert.equal(await savedAnswer(), true);
  await composer.fill("Keep this draft");
  await page.screenshot({ path: path.join(shots, "before-new-conversation.png") });
  await newConversation.click();
  await confirmation.getByText("The current conversation for this document will be deleted. This cannot be undone.").waitFor();
  await page.locator(".modal-backdrop").evaluate(el => Promise.all(el.getAnimations({ subtree: true }).map(animation => animation.finished)));
  await page.screenshot({ path: path.join(shots, "after-new-conversation.png") });
  assert.equal(await confirmation.getByRole("button", { name: "Cancel", exact: true }).evaluate(el => el === document.activeElement), true);
  // The modal keeps keyboard focus inside it and never uses Enter to choose deletion by default.
  await page.keyboard.press("Shift+Tab");
  await page.keyboard.press("Shift+Tab");
  assert.equal(await confirmation.getByRole("button", { name: "Delete conversation" }).evaluate(el => el === document.activeElement), true);
  await page.keyboard.press("Tab");
  await page.keyboard.press("Tab");
  await page.keyboard.press("Enter");
  await confirmation.waitFor({ state: "detached" });
  assert.equal(await newConversation.evaluate(el => el === document.activeElement), true);
  assert.equal(await savedAnswer(), true);
  for (const dismiss of ["Escape", "Close", "Backdrop"]) {
    await newConversation.click();
    await confirmation.waitFor();
    if (dismiss === "Escape") await page.keyboard.press("Escape");
    else if (dismiss === "Close") await confirmation.getByRole("button", { name: "Close", exact: true }).click();
    else await page.locator(".modal-backdrop").click({ position: { x: 3, y: 3 } });
    await confirmation.waitFor({ state: "detached" });
    await copyResponse.waitFor();
    assert.equal(await savedAnswer(), true);
    assert.equal(await composer.inputValue(), "Keep this draft");
  }
  await page.getByRole("button", { name: "Close Ask AI" }).click();
  await page.getByRole("button", { name: "Ask AI", exact: true }).click();
  await copyResponse.waitFor();
  const viewport = await page.window.evaluate(() => ({ width: innerWidth, height: innerHeight }));
  await page.window.setViewportSize({ width: 420, height: 800 });
  await page.window.waitForFunction(() => innerWidth <= 420);
  const theme = await page.evaluate(() => {
    const theme = document.documentElement.dataset.theme;
    document.documentElement.dataset.theme = "light";
    return theme;
  });
  await newConversation.click();
  await confirmation.waitFor();
  await page.locator(".modal-backdrop").evaluate(el => Promise.all(el.getAnimations({ subtree: true }).map(animation => animation.finished)));
  const box = await confirmation.boundingBox();
  const width = await page.window.evaluate(() => innerWidth);
  assert.ok(box && box.x >= 0 && box.x + box.width <= width, `the confirmation fits a narrow window: ${JSON.stringify({ width, box })}`);
  for (const name of [/^Search \(/, /^Ask AI$/]) {
    const covered = await page.locator(".topbar").getByRole("button", { name }).evaluate(button => {
      const box = button.getBoundingClientRect();
      return !!document.elementFromPoint(box.x + box.width / 2, box.y + box.height / 2)?.closest(".modal-backdrop");
    });
    assert.equal(covered, true, `the confirmation backdrop covers the narrow-window toolbar: ${name}`);
  }
  await page.screenshot({ path: path.join(shots, "new-conversation-narrow-light.png") });
  const search = page.locator(".topbar").getByRole("button", { name: /^Search \(/ });
  const searchBox = await search.boundingBox();
  await page.mouse.click(searchBox.x + searchBox.width / 2, searchBox.y + searchBox.height / 2);
  await confirmation.waitFor({ state: "detached" });
  assert.equal(await search.getAttribute("aria-pressed"), "false");
  assert.equal(await newConversation.evaluate(el => el === document.activeElement), true);
  assert.equal(await savedAnswer(), true);
  await newConversation.click();
  await confirmation.waitFor();
  await confirmation.getByRole("button", { name: "Delete conversation" }).click();
  await page.evaluate(theme => { document.documentElement.dataset.theme = theme; }, theme);
  await page.window.setViewportSize(viewport);
  await copyResponse.waitFor({ state: "detached" });
  for (let i = 0; i < 100 && await savedAnswer(); i++) await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(await savedAnswer(), false);
  assert.equal(await newConversation.isDisabled(), true);
  await page.getByRole("button", { name: "Close Ask AI" }).click();
  await page.getByRole("button", { name: "Ask AI", exact: true }).click();
  await page.getByRole("button", { name: "Summarize this document" }).waitFor();
  assert.equal(await copyResponse.count(), 0);
  console.log("ok  cancel, Escape, close and backdrop preserve chat; confirmed deletion clears saved history");
  await page.getByRole("button", { name: "Close Ask AI" }).click();
  console.log("ok  Copy response copies the original Markdown and reports a refused write");

  const sidebar = page.getByRole("button", { name: /Toggle sidebar/ });
  if (await sidebar.getAttribute("aria-pressed") !== "true") await sidebar.click();
  await page.getByRole("tab", { name: "Pages", exact: true }).click();
  const thumb = n => page.locator(`.thumb[data-page="${n}"]`);
  await thumb(1).click();
  await thumb(4).click({ modifiers: [process.platform === "darwin" ? "Meta" : "Control"] });
  await thumb(1).click({ button: "right" });
  const selectedSample = path.join(folder, "selected-sample.pdf");
  await captureDownload(app, selectedSample);
  await page.getByRole("menuitem", { name: "Export 2 selected pages…" }).click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  const marks = await page.evaluate(async bytes => {
    const task = pdfjsLib.getDocument({ data: Uint8Array.from(bytes) });
    try {
      const pdf = await task.promise;
      return await Promise.all(Array.from({ length: pdf.numPages }, async (_, n) =>
        (await (await pdf.getPage(n + 1)).getAnnotations()).map(a => a.subtype)));
    } finally { await task.destroy(); }
  }, Array.from(readFileSync(selectedSample)));
  assert.equal(marks.length, 2);
  assert.ok(marks[0].includes("Highlight") && marks[0].includes("FreeText") && marks[1].includes("Highlight"));
  await page.getByRole("tab", { name: "Outline" }).click();
  await page.getByRole("button", { name: "Reorder pages" }).click();
  await thumb(1).dragTo(thumb(2));
  assert.equal(await thumb(1).getAttribute("aria-label"), "Page 1, position 2");
  await page.getByLabel("Move page 1 to").fill("3");
  await page.getByRole("button", { name: "Move", exact: true }).click();
  assert.equal(await thumb(1).getAttribute("aria-label"), "Page 1, position 3");
  const file = path.join(folder, "reordered.pdf");
  await captureDownload(app, file);
  await page.getByRole("button", { name: "Save reordered copy" }).click();
  const openCopy = page.getByRole("button", { name: "Open reordered copy" });
  await openCopy.waitFor();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  const result = await page.evaluate(async bytes => {
    const read = async data => {
      const task = pdfjsLib.getDocument({ data });
      const pdf = await task.promise;
      try {
        const pages = [];
        for (let n = 1; n <= pdf.numPages; n++) {
          const p = await pdf.getPage(n);
          pages.push({ text: (await p.getTextContent()).items.map(i => i.str).join(" "),
            marks: (await p.getAnnotations()).map(a => a.subtype) });
        }
        const [bookmark] = await pdf.getOutline();
        const dest = typeof bookmark.dest === "string" ? await pdf.getDestination(bookmark.dest) : bookmark.dest;
        return { pages, bookmark: await pdf.getPageIndex(dest[0]) };
      } finally { await task.destroy(); }
    };
    return { original: await read(await (await fetch("sample.pdf")).arrayBuffer()), reordered: await read(Uint8Array.from(bytes)) };
  }, Array.from(readFileSync(file)));
  const { original, reordered } = result;
  const order = original.pages.map((_, i) => i);
  order.splice(2, 0, order.shift());
  assert.deepEqual(reordered.pages.map(p => p.text), order.map(i => original.pages[i].text));
  assert.equal(reordered.bookmark, order.indexOf(original.bookmark));
  assert.ok(reordered.pages[2].marks.includes("FreeText") && reordered.pages[2].marks.includes("Highlight"));
  assert.ok(reordered.pages.at(-1).marks.includes("Highlight"));
  await openCopy.click();
  await page.window.waitForFunction(() => document.title.includes("(reordered).pdf"));
  console.log("ok  reordered PDF pages, bookmark destinations and annotations match the sidebar order");

  await app.evaluate(({ BrowserWindow }, form) => BrowserWindow.getAllWindows()[0].webContents.send(
    "folio:open-file", { name: "Form.pdf", data: Buffer.from(form) }), formPdf());
  const field = page.locator(".textWidgetAnnotation input");
  await field.fill("A filled field");
  const documents = page.getByRole("tablist", { name: "Open documents" });
  await documents.getByRole("button", { name: "Close Form.pdf, unsaved changes" }).waitFor();
  await page.keyboard.press("Control+Shift+Tab");
  await documents.getByRole("tab", { name: "Form.pdf", exact: true }).click();
  assert.equal(await field.inputValue(), "A filled field");
  await page.getByRole("tab", { name: "Pages" }).click();
  await page.getByRole("button", { name: "Page 1", exact: true }).click({ button: "right" });
  const selectedForm = path.join(folder, "selected-form.pdf");
  await captureDownload(app, selectedForm);
  await page.getByRole("menuitem", { name: "Export this page…" }).click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  assert.equal(await page.evaluate(async bytes => {
    const task = pdfjsLib.getDocument({ data: Uint8Array.from(bytes) });
    try {
      const pdf = await task.promise;
      return pdf.numPages === 1 && (await (await pdf.getPage(1)).getAnnotations()).find(a => a.fieldName === "Name")?.fieldValue;
    } finally { await task.destroy(); }
  }, Array.from(readFileSync(selectedForm))), "A filled field");
  await documents.getByRole("button", { name: "Close Form.pdf, unsaved changes" }).waitFor();
  await page.getByRole("button", { name: "Reorder pages" }).click();
  await page.getByLabel("Move page 1 to").fill("2");
  await page.getByRole("button", { name: "Move", exact: true }).click();
  await captureDownload(app, path.join(folder, "reordered-form.pdf"));
  await page.keyboard.press("ControlOrMeta+s");
  await openCopy.click();
  assert.equal(await app.evaluate(() => globalThis.folioReordered), "completed");
  await page.window.waitForFunction(() => document.title.startsWith("Form (reordered).pdf"));
  await documents.getByRole("button", { name: "Close Form.pdf", exact: true }).waitFor();
  await page.getByTitle("Next page").click();
  assert.equal(await page.locator('.page[data-page-number="2"] .textWidgetAnnotation input').inputValue(), "A filled field");
  console.log("ok  the normal Save shortcut and Open reordered copy preserve filled form fields");
  await checkScanSelection(app, page);
  await checkContentEditing(app, page, folder, shots);
  await checkLinkHistory(app, page, shots);
  await checkPageExport(app, page, folder, shots);
  await checkPenDrawing(app, page, folder, shots);
}
