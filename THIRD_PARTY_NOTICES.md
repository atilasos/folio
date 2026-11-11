# Third-party notices

Folio is licensed under the [MIT License](LICENSE). It includes the following third-party software, each under its
own license. Full license texts ship with the components: pdf.js asset licenses are in `pdfjs/` alongside the files
they cover, and Electron's and Chromium's licenses are in the app's installation folder (`LICENSE.electron.txt`,
`LICENSES.chromium.html`). Bundled JavaScript and font licenses are in `licenses/THIRD_PARTY_LICENSES.md`.

| Component | License | Source |
| --- | --- | --- |
| [pdf.js](https://github.com/mozilla/pdf.js) (`pdfjs-dist`) | Apache-2.0 | Mozilla Foundation |
| [EmbedPDF PDFium](https://github.com/embedpdf/embed-pdf-viewer/tree/main/packages/pdfium) and [pdf-lib](https://github.com/Hopding/pdf-lib) (content editing) | MIT wrappers, PDFium and bundled dependency notices in `pdfjs/editor-licenses/` | EmbedPDF, PDFium and pdf-lib contributors |
| [@pdf-lib/fontkit](https://github.com/Hopding/fontkit) (text box fonts in saved copies) | MIT, license text in `licenses/THIRD_PARTY_LICENSES.md` | Devon Govett, Andrew Dillon and contributors |
| [Tesseract.js](https://github.com/naptha/tesseract.js), its WebAssembly core and [recognition models](https://github.com/naptha/tessdata) | Apache-2.0, worker dependency notices in `ocr/worker.min.js.LICENSE.txt` | Tesseract.js and Tesseract OCR contributors |
| pdf.js JBIG2 decoder (from PDFium) | BSD-3-Clause | The PDFium Authors |
| pdf.js OpenJPEG decoder | BSD-2-Clause | OpenJPEG contributors |
| pdf.js qcms color management | MIT | Mozilla Corporation, Marti Maria |
| pdf.js standard fonts: Foxit | BSD-3-Clause | The PDFium Authors |
| pdf.js standard fonts: Liberation | Liberation font license (GPLv2 with font exception) | Red Hat, Inc. |
| pdf.js CMaps | BSD-3-Clause | Adobe Systems Incorporated |
| pdf.js ICC profile | CC0-1.0 | |
| [React](https://react.dev) and React DOM | MIT | Meta Platforms, Inc. and affiliates |
| [react-markdown](https://github.com/remarkjs/react-markdown), [remark-math / rehype-katex](https://github.com/remarkjs/remark-math) and their unified parsing dependencies | MIT and ISC | unified contributors |
| [KaTeX](https://github.com/KaTeX/KaTeX), including its bundled fonts | MIT | Khan Academy and other contributors |
| [Electron](https://www.electronjs.org) | MIT | GitHub Inc., OpenJS Foundation and contributors |
| Chromium (bundled with Electron) | BSD-3-Clause and others | The Chromium Authors |
| [electron-updater](https://github.com/electron-userland/electron-builder) and its dependencies | MIT, ISC, BlueOak-1.0.0, Python-2.0 | electron-builder contributors and others |
| [mammoth](https://github.com/mwilliamson/mammoth.js) (Word document conversion) and its dependencies | BSD-2-Clause, MIT, ISC, Zlib, Python-2.0; JSZip under MIT (of MIT OR GPL-3.0-or-later) | Michael Williamson and others |

The sample document (`public/sample.pdf`) was written for Folio and is covered by Folio's MIT License.

## remark-math and rehype-katex

(The MIT License)

Copyright (c) Junyoung Choi <fluke8259@gmail.com>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
