/**
 * AppendicesEngine.ts
 * Formats the appendices section of a thesis manuscript.
 *
 * Sections handled (in document order, after REFERENCES):
 *   • APPENDICES            — title page (alone on page)
 *   • Appendix X            — appendix letter header per appendix
 *   • Appendix title        — first non-empty paragraph after the letter
 *   • Continuation label    — "Continuation of Appendix X..."
 *   • User Manual           — special section: 0.3 cm margins, no header/
 *                             footer template, right-aligned images 18.20×27 cm
 *   • CURRICULUM VITAE      — title page (alone on page), centred images
 *
 * Unit reference
 *   half-points (sz)  : pt × 2    (14 pt → 28, 13 pt → 26, 12 pt → 24)
 *   twips (spacing)   : pt × 20   (1 pt → 20)
 *   line (auto)       : 240 = 1.0×, 480 = 2.0×, 720 = 3.0×
 *   EMU               : 1 cm = 360 000 EMU
 *   twips (margin)    : 1 cm ≈ 567 twips  → 0.3 cm ≈ 170 twips
 *   DrawingML line w  : pt × 12 700 EMU   (3 pt → 38 100)
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

// DrawingML main namespace — a:ln, a:solidFill, a:srgbClr, a:noFill, a:prstDash
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
import type { FormattingConfig } from "../constants";

// ─── EMU constants ─────────────────────────────────────────────────────────────
const UM_IMAGE_W_EMU = 6_552_000; // 18.20 cm × 360 000
const UM_IMAGE_H_EMU = 9_720_000; //  27.00 cm × 360 000

// ─── Appendix large-figure thresholds and targets ─────────────────────────────
const APPENDIX_MIN_H_EMU = 5_760_000; // 16 cm — height threshold
const APPENDIX_MIN_W_EMU = 3_600_000; // 10 cm — width  threshold
const APPENDIX_TARGET_W_EMU = 5_400_000; // 15 cm  — output width
const APPENDIX_TARGET_H_EMU = 6_768_000; // 18.8 cm — output height (after appendix title)
const APPENDIX_TARGET_H_CONT_EMU = 6_984_000; // 19.4 cm — output height (after continuation label)

// 3 pt border in DrawingML line-width units (EMU): 3 × 12 700
const APPENDIX_BORDER_W_EMU = 38_100;

// Full border width in EMU — used as effectExtent on all four sides so
// Word does not clip any part of the border stroke, including top and bottom.
const APPENDIX_BORDER_EXTENT_EMU = APPENDIX_BORDER_W_EMU; // 38 100

function ptsToHalfPts(pts: number): number {
  return Math.round(pts * 2);
}
function linesToTwips(lines: number): number {
  return Math.round(lines * 240);
}
function inchesToTwips(inches: number): number {
  return Math.round(inches * 1440);
}
function ptsToEMU(pts: number): number {
  return Math.round(pts * 12700);
}

// ─── low-level helpers ─────────────────────────────────────────────────────────

function wAttr(el: Element, local: string): string {
  return el.getAttributeNS(W_NS, local) ?? el.getAttribute("w:" + local) ?? "";
}
function setWAttr(el: Element, local: string, val: string) {
  el.setAttributeNS(W_NS, "w:" + local, val);
}
function wElem(doc: Document, local: string): Element {
  return doc.createElementNS(W_NS, "w:" + local);
}
function getChild(parent: Element, local: string): Element | null {
  for (const c of Array.from(parent.childNodes)) {
    if (
      c instanceof Element &&
      c.namespaceURI === W_NS &&
      c.localName === local
    )
      return c;
  }
  return null;
}
function ensureChild(parent: Element, local: string, prepend = false): Element {
  const ex = getChild(parent, local);
  if (ex) return ex;
  const child = wElem(parent.ownerDocument!, local);
  if (prepend && parent.firstChild)
    parent.insertBefore(child, parent.firstChild);
  else parent.appendChild(child);
  return child;
}
function removeChildren(parent: Element, local: string) {
  Array.from(parent.childNodes)
    .filter(
      (c) =>
        c instanceof Element &&
        c.namespaceURI === W_NS &&
        c.localName === local,
    )
    .forEach((c) => parent.removeChild(c));
}
function ensurePPr(p: Element): Element {
  return ensureChild(p, "pPr", true);
}

// ─── text helpers ──────────────────────────────────────────────────────────────

function getParagraphText(p: Element): string {
  const parts: string[] = [];
  p.querySelectorAll("*").forEach((el) => {
    if (el.localName === "t" && el.namespaceURI === W_NS) {
      let ancestor = el.parentElement;
      while (ancestor && ancestor !== p) {
        if (ancestor.localName === "txbxContent") return;
        ancestor = ancestor.parentElement;
      }
      const run = el.parentElement;
      if (run && run.localName === "r") {
        const hasSpecial = Array.from(run.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        );
        if (!hasSpecial) parts.push(el.textContent ?? "");
      }
    }
  });
  return parts.join("");
}

function normalizeText(t: string): string {
  return t.replace(/\s+/gu, " ").trim();
}

function uppercaseParagraph(p: Element) {
  p.querySelectorAll("r > t").forEach((t) => {
    if (t.namespaceURI === W_NS)
      t.textContent = (t.textContent ?? "").toUpperCase();
  });
}

function trimParagraphText(p: Element) {
  p.querySelectorAll("r > t").forEach((t) => {
    if (t.namespaceURI === W_NS) t.textContent = (t.textContent ?? "").trim();
  });
}

// ─── content-control unwrapper ─────────────────────────────────────────────────

function unwrapContentControls(body: Element) {
  let changed = true;
  while (changed) {
    changed = false;
    const sdts = Array.from(body.getElementsByTagNameNS(W_NS, "sdt"));
    for (const sdt of sdts) {
      const parent = sdt.parentElement;
      if (!parent) continue;
      const sdtContent = Array.from(sdt.childNodes).find(
        (c): c is Element =>
          c instanceof Element && c.localName === "sdtContent",
      );
      if (sdtContent) {
        Array.from(sdtContent.childNodes).forEach((child) =>
          parent.insertBefore(child.cloneNode(true), sdt),
        );
      }
      parent.removeChild(sdt);
      changed = true;
      break;
    }
  }
}

// ─── property writers ──────────────────────────────────────────────────────────

function writePAlignment(p: Element, value: string) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "jc");
  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", value);
  pPr.appendChild(jc);
}

function writePIndentZero(p: Element) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "ind");
  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);
}

function writePSpacing(
  p: Element,
  before: number,
  after: number,
  line: number,
  lineRule = "auto",
) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "spacing");
  removeChildren(pPr, "contextualSpacing");
  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", String(before));
  setWAttr(sp, "after", String(after));
  setWAttr(sp, "line", String(line));
  setWAttr(sp, "lineRule", lineRule);
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);
  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);
}

function writePPrRPr(
  p: Element,
  font: string,
  halfPt: number,
  bold: boolean,
  italic: boolean,
) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "rPr");
  const rPr = wElem(p.ownerDocument!, "rPr");
  pPr.appendChild(rPr);

  const rFonts = wElem(p.ownerDocument!, "rFonts");
  for (const a of ["ascii", "hAnsi", "eastAsia", "cs"])
    setWAttr(rFonts, a, font);
  rPr.appendChild(rFonts);

  if (bold) {
    rPr.appendChild(wElem(p.ownerDocument!, "b"));
    rPr.appendChild(wElem(p.ownerDocument!, "bCs"));
  }
  if (italic) {
    rPr.appendChild(wElem(p.ownerDocument!, "i"));
    rPr.appendChild(wElem(p.ownerDocument!, "iCs"));
  }
  const sz = wElem(p.ownerDocument!, "sz");
  setWAttr(sz, "val", String(halfPt));
  rPr.appendChild(sz);
  const szCs = wElem(p.ownerDocument!, "szCs");
  setWAttr(szCs, "val", String(halfPt));
  rPr.appendChild(szCs);
}

function applyRunFormatting(
  p: Element,
  font: string,
  halfPt: number,
  bold: boolean | null,
  italic: boolean | null,
) {
  const sizeStr = String(halfPt);
  const textRuns = Array.from(p.querySelectorAll("r")).filter(
    (r) =>
      r.namespaceURI === W_NS &&
      !Array.from(r.children).some((c) =>
        ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
      ),
  );

  for (const run of textRuns) {
    const existingRPr = getChild(run, "rPr");
    const wasBold =
      existingRPr != null &&
      existingRPr.getElementsByTagNameNS(W_NS, "b").length > 0;
    const wasItalic =
      existingRPr != null &&
      existingRPr.getElementsByTagNameNS(W_NS, "i").length > 0;
    const applyBold = bold ?? wasBold;
    const applyItalic = italic ?? wasItalic;

    removeChildren(run, "rPr");
    const rPr = wElem(run.ownerDocument!, "rPr");
    run.insertBefore(rPr, run.firstChild ?? null);

    const rFonts = ensureChild(rPr, "rFonts");
    for (const attr of ["asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"])
      rFonts.removeAttributeNS(W_NS, attr);
    setWAttr(rFonts, "ascii", font);
    setWAttr(rFonts, "hAnsi", font);
    setWAttr(rFonts, "eastAsia", font);
    setWAttr(rFonts, "cs", font);

    removeChildren(rPr, "sz");
    removeChildren(rPr, "szCs");
    const szEl = ensureChild(rPr, "sz");
    setWAttr(szEl, "val", sizeStr);
    const szCs = ensureChild(rPr, "szCs");
    setWAttr(szCs, "val", sizeStr);

    removeChildren(rPr, "b");
    removeChildren(rPr, "bCs");
    if (applyBold) {
      ensureChild(rPr, "b");
      ensureChild(rPr, "bCs");
    }
    removeChildren(rPr, "i");
    removeChildren(rPr, "iCs");
    if (applyItalic) {
      ensureChild(rPr, "i");
      ensureChild(rPr, "iCs");
    }
    for (const tag of [
      "caps",
      "color",
      "lang",
      "noProof",
      "rStyle",
      "highlight",
      "vertAlign",
      "effect",
    ])
      removeChildren(rPr, tag);
  }
}

function stripPPr(p: Element) {
  const pPr = getChild(p, "pPr");
  if (!pPr) return;
  for (const tag of [
    "pStyle",
    "rPr",
    "widowControl",
    "ind",
    "spacing",
    "jc",
    "pageBreakBefore",
    "keepNext",
    "keepLines",
    "numPr",
    "outlineLvl",
    "contextualSpacing",
    "snapToGrid",
    "pBdr",
    "tabs",
  ])
    removeChildren(pPr, tag);
}

// ─── text-run rewriter ─────────────────────────────────────────────────────────

function rewriteParagraphText(p: Element, text: string) {
  Array.from(p.querySelectorAll("r"))
    .filter(
      (r) =>
        r.namespaceURI === W_NS &&
        !Array.from(r.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        ),
    )
    .forEach((r) => r.parentElement?.removeChild(r));

  const run = wElem(p.ownerDocument!, "r");
  const tEl = wElem(p.ownerDocument!, "t");
  tEl.textContent = text;
  run.appendChild(tEl);
  p.appendChild(run);
}

// ─── paragraph formatters ───────────────────────────────────────────────────────

function applyAppendicesSectionTitle(p: Element, config: FormattingConfig) {
  stripPPr(p);
  if (config.titles.textTransform === "uppercase") {
    uppercaseParagraph(p);
  }

  trimParagraphText(p);

  const pPr = ensurePPr(p);
  pPr.appendChild(wElem(p.ownerDocument!, "pageBreakBefore"));

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", config.titles.alignment);
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", String(inchesToTwips(config.titles.indentation)));
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", String(linesToTwips(config.titles.lineSpacing)));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  const size = ptsToHalfPts(config.titles.fontSize);
  writePPrRPr(
    p,
    config.titles.fontFamily,
    size,
    config.titles.bold ?? true,
    config.titles.italic ?? false,
  );
  applyRunFormatting(
    p,
    config.titles.fontFamily,
    size,
    config.titles.bold ?? true,
    config.titles.italic ?? false,
  );
}

function applyAppendixLetter(p: Element, letter: string, config: FormattingConfig) {
  const canonical = `Appendix ${letter.toUpperCase()}`;
  rewriteParagraphText(p, canonical);
  stripPPr(p);

  const pPr = ensurePPr(p);
  removeChildren(pPr, "sectPr");
  pPr.appendChild(wElem(p.ownerDocument!, "pageBreakBefore"));

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", config.appendixLetter.alignment);
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", String(inchesToTwips(config.appendixLetter.indentation)));
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", String(linesToTwips(config.appendixLetter.lineSpacing)));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  const size = ptsToHalfPts(config.appendixLetter.fontSize);
  writePPrRPr(
    p,
    config.appendixLetter.fontFamily,
    size,
    config.appendixLetter.bold ?? true,
    config.appendixLetter.italic ?? false,
  );
  applyRunFormatting(
    p,
    config.appendixLetter.fontFamily,
    size,
    config.appendixLetter.bold ?? true,
    config.appendixLetter.italic ?? false,
  );
}

function applyAppendixTitle(p: Element, config: FormattingConfig) {
  stripPPr(p);
  if (config.titles.textTransform === "uppercase") {
    uppercaseParagraph(p);
  }

  trimParagraphText(p);

  const pPr = ensurePPr(p);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", config.titles.alignment);
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", String(inchesToTwips(config.titles.indentation)));
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", String(linesToTwips(config.titles.lineSpacing)));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  const size = ptsToHalfPts(config.titles.fontSize);
  writePPrRPr(
    p,
    config.titles.fontFamily,
    size,
    config.titles.bold ?? true,
    config.titles.italic ?? false,
  );
  applyRunFormatting(
    p,
    config.titles.fontFamily,
    size,
    config.titles.bold ?? true,
    config.titles.italic ?? false,
  );
}

function applyContinuationAppendixLabel(
  p: Element,
  currentLetter: string,
  config: FormattingConfig,
) {
  const canonical = currentLetter
    ? `Continuation of Appendix ${currentLetter.toUpperCase()}...`
    : "Continuation of Appendix...";
  rewriteParagraphText(p, canonical);
  stripPPr(p);

  const pPr = ensurePPr(p);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", config.appendixContinuation.alignment);
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", String(inchesToTwips(config.appendixContinuation.indentation)));
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", String(linesToTwips(config.appendixContinuation.lineSpacing)));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  const size = ptsToHalfPts(config.appendixContinuation.fontSize);
  writePPrRPr(
    p,
    config.appendixContinuation.fontFamily,
    size,
    config.appendixContinuation.bold ?? false,
    config.appendixContinuation.italic ?? true,
  );
  applyRunFormatting(
    p,
    config.appendixContinuation.fontFamily,
    size,
    config.appendixContinuation.bold ?? false,
    config.appendixContinuation.italic ?? true,
  );
}

function applyEmptyParagraph(p: Element, config: FormattingConfig) {
  const pPr = ensurePPr(p);
  for (const tag of [
    "pStyle",
    "ind",
    "spacing",
    "jc",
    "rPr",
    "widowControl",
    "pageBreakBefore",
    "keepNext",
    "keepLines",
    "outlineLvl",
    "contextualSpacing",
    "snapToGrid",
  ])
    removeChildren(pPr, tag);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "20");
  setWAttr(sp, "lineRule", "exact");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "1");
  pPr.appendChild(cs);

  const bold = config.body.bold ?? false;
  const italic = config.body.italic ?? false;
  writePPrRPr(
    p,
    config.body.fontFamily,
    ptsToHalfPts(config.body.fontSize),
    bold,
    italic,
  );
}

// ─── image helpers ─────────────────────────────────────────────────────────────

function collectDrawings(node: Element): Element[] {
  const result: Element[] = [];
  for (const c of Array.from(node.childNodes)) {
    if (!(c instanceof Element)) continue;
    if (c.localName === "drawing" && c.namespaceURI === W_NS) result.push(c);
    else result.push(...collectDrawings(c));
  }
  return result;
}

function getDrawingExtentEmu(
  drawing: Element,
): { cx: number; cy: number } | null {
  for (const name of ["inline", "anchor"]) {
    const container = Array.from(drawing.childNodes).find(
      (c): c is Element =>
        c instanceof Element &&
        c.namespaceURI === WP_NS &&
        c.localName === name,
    );
    if (!container) continue;
    const extent = Array.from(container.childNodes).find(
      (c): c is Element => c instanceof Element && c.localName === "extent",
    );
    if (extent) {
      return {
        cx: parseInt(extent.getAttribute("cx") ?? "0", 10),
        cy: parseInt(extent.getAttribute("cy") ?? "0", 10),
      };
    }
  }
  return null;
}

/**
 * Set a solid black picture border of the given EMU width on every spPr
 * inside the drawing.
 *
 * The a:ln element lives inside pic:spPr (or equivalent spPr) in the
 * DrawingML tree — this is the same property exposed by Word's
 * "Format Picture → Picture Border" panel.  It draws an outline directly
 * on the image edge, not around the paragraph.
 *
 * querySelectorAll("spPr") matches by local name regardless of namespace
 * prefix, so it covers both pic:spPr and any other variant correctly.
 */
function setDrawingBorder(drawing: Element, widthEmu: number) {
  drawing.querySelectorAll("spPr").forEach((spPr) => {
    // Remove any pre-existing a:ln
    Array.from(spPr.childNodes)
      .filter((c): c is Element => c instanceof Element && c.localName === "ln")
      .forEach((ln) => spPr.removeChild(ln));

    const doc = drawing.ownerDocument!;

    // <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
    const ln = doc.createElementNS(A_NS, "a:ln");
    ln.setAttribute("w", String(widthEmu));
    ln.setAttribute("cap", "flat");
    ln.setAttribute("cmpd", "sng");
    ln.setAttribute("algn", "ctr");

    // <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
    const solidFill = doc.createElementNS(A_NS, "a:solidFill");
    const srgbClr = doc.createElementNS(A_NS, "a:srgbClr");
    srgbClr.setAttribute("val", "000000");
    solidFill.appendChild(srgbClr);
    ln.appendChild(solidFill);

    // <a:prstDash val="solid"/>
    const prstDash = doc.createElementNS(A_NS, "a:prstDash");
    prstDash.setAttribute("val", "solid");
    ln.appendChild(prstDash);

    spPr.appendChild(ln);
  });
}

/**
 * Explicitly remove any picture border from every spPr in the drawing by
 * replacing a:ln with <a:ln><a:noFill/></a:ln>.
 * This prevents Word from falling back to a theme or style default outline.
 */
function clearDrawingBorder(drawing: Element) {
  drawing.querySelectorAll("spPr").forEach((spPr) => {
    Array.from(spPr.childNodes)
      .filter((c): c is Element => c instanceof Element && c.localName === "ln")
      .forEach((ln) => spPr.removeChild(ln));

    const doc = drawing.ownerDocument!;
    const ln = doc.createElementNS(A_NS, "a:ln");
    ln.appendChild(doc.createElementNS(A_NS, "a:noFill"));
    spPr.appendChild(ln);
  });
}

function convertAnchorToInline(drawing: Element) {
  const alreadyInline = Array.from(drawing.childNodes).some(
    (c): c is Element =>
      c instanceof Element &&
      c.namespaceURI === WP_NS &&
      c.localName === "inline",
  );
  if (alreadyInline) return;

  const anchor = Array.from(drawing.childNodes).find(
    (c): c is Element =>
      c instanceof Element &&
      c.namespaceURI === WP_NS &&
      c.localName === "anchor",
  );
  if (!anchor) return;

  const doc = drawing.ownerDocument!;
  const inline = doc.createElementNS(WP_NS, "wp:inline");
  inline.setAttribute("distT", "0");
  inline.setAttribute("distB", "0");
  inline.setAttribute("distL", "0");
  inline.setAttribute("distR", "0");

  const anchorExtent = Array.from(anchor.childNodes).find(
    (c): c is Element => c instanceof Element && c.localName === "extent",
  );
  if (anchorExtent) {
    const ne = doc.createElementNS(WP_NS, "wp:extent");
    ne.setAttribute("cx", anchorExtent.getAttribute("cx") ?? "0");
    ne.setAttribute("cy", anchorExtent.getAttribute("cy") ?? "0");
    inline.appendChild(ne);
  }

  const eff = doc.createElementNS(WP_NS, "wp:effectExtent");
  eff.setAttribute("l", "0");
  eff.setAttribute("t", "0");
  eff.setAttribute("r", "0");
  eff.setAttribute("b", "0");
  inline.appendChild(eff);

  Array.from(anchor.childNodes).forEach((c) => {
    if (
      c instanceof Element &&
      ["docPr", "cNvGraphicFramePr", "graphic"].includes(c.localName)
    )
      inline.appendChild(c.cloneNode(true));
  });

  drawing.replaceChild(inline, anchor);
}

function resizeInlineImage(
  drawing: Element,
  widthEmu: number,
  heightEmu: number,
) {
  const inline = Array.from(drawing.childNodes).find(
    (c): c is Element =>
      c instanceof Element &&
      c.namespaceURI === WP_NS &&
      c.localName === "inline",
  );
  if (!inline) return;

  const extent = Array.from(inline.childNodes).find(
    (c): c is Element => c instanceof Element && c.localName === "extent",
  );
  if (extent) {
    extent.setAttribute("cx", String(widthEmu));
    extent.setAttribute("cy", String(heightEmu));
  }

  inline.querySelectorAll("ext").forEach((ext) => {
    const parent = ext.parentElement;
    if (parent && parent.localName === "xfrm") {
      ext.setAttribute("cx", String(widthEmu));
      ext.setAttribute("cy", String(heightEmu));
    }
  });
}

/**
 * Set the wp:effectExtent on the inline container so that the picture border
 * (a:ln, rendered at half its declared width outside the image boundary) is
 * not clipped.  Pass the half-width of the border stroke in EMU.
 *
 * Word clips anything that protrudes beyond the effectExtent, so setting all
 * four sides to at least half the stroke width ensures the top and bottom
 * edges of the border are fully visible.
 */
function setInlineEffectExtent(drawing: Element, marginEmu: number) {
  const inline = Array.from(drawing.childNodes).find(
    (c): c is Element =>
      c instanceof Element &&
      c.namespaceURI === WP_NS &&
      c.localName === "inline",
  );
  if (!inline) return;

  // Remove existing effectExtent and replace with new values.
  Array.from(inline.childNodes)
    .filter(
      (c): c is Element =>
        c instanceof Element && c.localName === "effectExtent",
    )
    .forEach((e) => inline.removeChild(e));

  const doc = drawing.ownerDocument!;
  const eff = doc.createElementNS(WP_NS, "wp:effectExtent");
  const val = String(marginEmu);
  eff.setAttribute("l", val);
  eff.setAttribute("t", val);
  eff.setAttribute("r", val);
  eff.setAttribute("b", val);

  // Insert after wp:extent (first child).
  const extent = Array.from(inline.childNodes).find(
    (c): c is Element => c instanceof Element && c.localName === "extent",
  );
  if (extent && extent.nextSibling) {
    inline.insertBefore(eff, extent.nextSibling);
  } else {
    inline.appendChild(eff);
  }
}

function applyUserManualImageParagraph(p: Element, config: FormattingConfig) {
  const pPr = ensurePPr(p);
  for (const tag of [
    "pStyle",
    "widowControl",
    "jc",
    "ind",
    "spacing",
    "pBdr",
    "rPr",
  ])
    removeChildren(pPr, tag);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "right");
  pPr.appendChild(jc);

  writePIndentZero(p);
  writePSpacing(p, 0, 0, 240);

  collectDrawings(p).forEach((drawing) => {
    convertAnchorToInline(drawing);
    resizeInlineImage(drawing, UM_IMAGE_W_EMU, UM_IMAGE_H_EMU);
  });
}

/**
 * CV image paragraph: centred, inline, no border of any kind.
 * clearDrawingBorder() removes any a:ln picture border from the source doc.
 */
function applyCVImageParagraph(p: Element, config: FormattingConfig) {
  const pPr = ensurePPr(p);
  for (const tag of [
    "pStyle",
    "widowControl",
    "jc",
    "ind",
    "spacing",
    "pBdr",
    "rPr",
  ])
    removeChildren(pPr, tag);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);

  writePIndentZero(p);
  writePSpacing(p, 0, 0, 240);

  collectDrawings(p).forEach((drawing) => {
    convertAnchorToInline(drawing);
    clearDrawingBorder(drawing);
  });
}

/**
 * Appendix figure paragraph: centred, inline.
 *
 * Per-drawing logic:
 *   Meets threshold (>= 16 cm tall AND >= 10 cm wide):
 *     → Resize to 19.4 cm H × 15 cm W
 *     → Apply 3 pt solid black border via a:ln on the drawing's spPr
 *        (this is the real picture border on the image edge, not a paragraph box)
 *     → Set effectExtent to half the border stroke width on all four sides
 *        so the top and bottom border edges are not clipped by Word's renderer
 *   Does not meet threshold:
 *     → Leave size unchanged
 *     → Clear any pre-existing picture border
 *
 * No w:pBdr paragraph border is used.
 */
function applyAppendixFigureParagraph(p: Element, targetHeightEmu: number, config: FormattingConfig) {
  const pPr = ensurePPr(p);
  for (const tag of [
    "pStyle",
    "widowControl",
    "jc",
    "ind",
    "spacing",
    "pBdr",
    "rPr",
  ])
    removeChildren(pPr, tag);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);

  writePIndentZero(p);
  writePSpacing(p, 0, 0, 240);

  collectDrawings(p).forEach((drawing) => {
    // Read extent BEFORE converting anchor → inline (anchor is removed after).
    const extent = getDrawingExtentEmu(drawing);
    convertAnchorToInline(drawing);

    const meetsThreshold =
      extent !== null &&
      extent.cy >= APPENDIX_MIN_H_EMU &&
      extent.cx >= APPENDIX_MIN_W_EMU;

    if (meetsThreshold) {
      resizeInlineImage(drawing, APPENDIX_TARGET_W_EMU, targetHeightEmu);
      // Set a 3 pt solid black outline directly on the image via a:ln in spPr.
      setDrawingBorder(drawing, ptsToEMU(config.figure.borderWeight));
      // Expand effectExtent on all four sides by the full stroke width so
      // Word does not clip any part of the border, including top and bottom.
      setInlineEffectExtent(drawing, ptsToEMU(config.figure.borderWeight));
    } else {
      // Ensure no stray picture border from the source document survives.
      clearDrawingBorder(drawing);
    }
  });
}

// ─── section-property helpers ──────────────────────────────────────────────────

function injectSectPr(p: Element, sectPr: Element) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "sectPr");
  pPr.appendChild(sectPr);
}

async function addEmptyHeaderFooterToZip(
  zip: any,
): Promise<{ hdrRelId: string; ftrRelId: string }> {
  const zeroP =
    `<w:p>` +
    `<w:pPr>` +
    `<w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"` +
    ` w:beforeAutospacing="0" w:afterAutospacing="0"/>` +
    `<w:contextualSpacing w:val="0"/>` +
    `<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>` +
    `</w:pPr>` +
    `</w:p>`;

  const emptyHdr =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    zeroP +
    `</w:hdr>`;

  const emptyFtr =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    zeroP +
    `</w:ftr>`;

  zip.file("word/headerUM.xml", emptyHdr);
  zip.file("word/footerUM.xml", emptyFtr);

  const relsPath = "word/_rels/document.xml.rels";
  const relsRaw: string =
    (await zip.file(relsPath)?.async("string")) ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

  const relsParser = new DOMParser();
  const relsDom = relsParser.parseFromString(relsRaw, "application/xml");
  const relsRoot = relsDom.documentElement;

  let maxId = 0;
  Array.from(relsRoot.childNodes).forEach((n) => {
    if (!(n instanceof Element)) return;
    const id = n.getAttribute("Id") ?? "";
    const m = id.match(/^rId(\d+)$/i);
    if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
  });

  const hdrRelId = `rId${maxId + 1}`;
  const ftrRelId = `rId${maxId + 2}`;

  const addRel = (id: string, type: string, target: string) => {
    const rel = relsDom.createElementNS(RELS_NS, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", type);
    rel.setAttribute("Target", target);
    relsRoot.appendChild(rel);
  };

  addRel(
    hdrRelId,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    "headerUM.xml",
  );
  addRel(
    ftrRelId,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    "footerUM.xml",
  );

  zip.file(relsPath, new XMLSerializer().serializeToString(relsDom));

  const ctPath = "[Content_Types].xml";
  const ctRaw: string | undefined = await zip.file(ctPath)?.async("string");
  if (ctRaw) {
    const ctDom = new DOMParser().parseFromString(ctRaw, "application/xml");
    const ctRoot = ctDom.documentElement;

    const addOverride = (partName: string, contentType: string) => {
      const already = Array.from(ctRoot.childNodes).some(
        (n) =>
          n instanceof Element &&
          n.localName === "Override" &&
          n.getAttribute("PartName") === partName,
      );
      if (already) return;
      const ov = ctDom.createElementNS(CT_NS, "Override");
      ov.setAttribute("PartName", partName);
      ov.setAttribute("ContentType", contentType);
      ctRoot.appendChild(ov);
    };

    addOverride(
      "/word/headerUM.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    );
    addOverride(
      "/word/footerUM.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    );

    zip.file(ctPath, new XMLSerializer().serializeToString(ctDom));
  }

  return { hdrRelId, ftrRelId };
}

function buildNormalClosingSectPr(
  doc: Document,
  mainSectPr: Element | null,
): Element {
  let sectPr: Element;

  if (mainSectPr) {
    sectPr = mainSectPr.cloneNode(true) as Element;
  } else {
    sectPr = wElem(doc, "sectPr");
    const pgSz = wElem(doc, "pgSz");
    setWAttr(pgSz, "w", "11906");
    setWAttr(pgSz, "h", "16838");
    sectPr.appendChild(pgSz);
  }

  removeChildren(sectPr, "type");
  const typeEl = wElem(doc, "type");
  setWAttr(typeEl, "val", "nextPage");
  sectPr.insertBefore(typeEl, sectPr.firstChild);

  return sectPr;
}

function buildUserManualSectPr(
  doc: Document,
  mainSectPr: Element | null,
  hdrRelId: string,
  ftrRelId: string,
): Element {
  const sectPr = wElem(doc, "sectPr");

  const typeEl = wElem(doc, "type");
  setWAttr(typeEl, "val", "nextPage");
  sectPr.appendChild(typeEl);

  if (mainSectPr) {
    const pgSz = getChild(mainSectPr, "pgSz");
    if (pgSz) sectPr.appendChild(pgSz.cloneNode(true));
  } else {
    const pgSz = wElem(doc, "pgSz");
    setWAttr(pgSz, "w", "11906");
    setWAttr(pgSz, "h", "16838");
    sectPr.appendChild(pgSz);
  }

  const pgMar = wElem(doc, "pgMar");
  setWAttr(pgMar, "top", "170");
  setWAttr(pgMar, "right", "170");
  setWAttr(pgMar, "bottom", "170");
  setWAttr(pgMar, "left", "170");
  setWAttr(pgMar, "header", "0");
  setWAttr(pgMar, "footer", "0");
  setWAttr(pgMar, "gutter", "0");
  sectPr.appendChild(pgMar);

  const addRef = (tag: string, type: string, relId: string) => {
    const el = wElem(doc, tag);
    setWAttr(el, "type", type);
    el.setAttributeNS(R_NS, "r:id", relId);
    sectPr.appendChild(el);
  };

  addRef("headerReference", "default", hdrRelId);
  addRef("headerReference", "first", hdrRelId);
  addRef("headerReference", "even", hdrRelId);
  addRef("footerReference", "default", ftrRelId);
  addRef("footerReference", "first", ftrRelId);
  addRef("footerReference", "even", ftrRelId);

  return sectPr;
}

// ─── state ─────────────────────────────────────────────────────────────────────

type AppZone = "before" | "appendices" | "user_manual" | "cv";

interface AppState {
  zone: AppZone;
  expectingTitle: boolean;
  currentLetter: string;
  cvSeen: boolean;
  afterContinuation: boolean;
}

// ─── public entry ──────────────────────────────────────────────────────────────

export async function formatAppendices(
  arrayBuffer: ArrayBuffer,
  options: { rules: string[]; config: FormattingConfig },
): Promise<Blob> {
  const rules: Record<string, boolean> = {};
  for (const r of options?.rules ?? []) rules[r] = true;

  const JSZip = (window as any).JSZip;
  if (!JSZip) throw new Error("JSZip not loaded");

  const zip = await JSZip.loadAsync(arrayBuffer);
  const xmlStr: string = await zip.file("word/document.xml").async("string");

  const parser = new DOMParser();
  const dom = parser.parseFromString(xmlStr, "application/xml");

  const body = dom.querySelector("body");
  if (!body) throw new Error("No <body> element found in document.xml");

  const allChildren = Array.from(body.childNodes).filter(
    (c): c is Element => c instanceof Element,
  );

  const mainSectPr =
    (Array.from(body.childNodes).find(
      (c): c is Element =>
        c instanceof Element &&
        c.localName === "sectPr" &&
        c.namespaceURI === W_NS,
    ) as Element | null) ?? null;

  let appendicesStart = -1;
  let seenReferences = false;

  for (let i = 0; i < allChildren.length; i++) {
    const child = allChildren[i];
    if (child.localName !== "p") continue;

    const raw = normalizeText(
      Array.from(child.querySelectorAll("t"))
        .map((n) => n.textContent ?? "")
        .join(""),
    );
    const txt = normalizeText(getParagraphText(child)) || raw;

    if (!seenReferences && /^references$/iu.test(txt)) {
      seenReferences = true;
      continue;
    }
    if (seenReferences && /^appendices$/iu.test(txt)) {
      appendicesStart = i;
      break;
    }
  }

  if (appendicesStart === -1) {
    const serializer = new XMLSerializer();
    zip.file("word/document.xml", serializer.serializeToString(dom));
    return zip.generateAsync({ type: "blob" }) as Promise<Blob>;
  }

  allChildren.slice(appendicesStart).forEach((el) => unwrapContentControls(el));

  const scopedChildren = Array.from(body.childNodes)
    .filter((c): c is Element => c instanceof Element && c !== mainSectPr)
    .slice(appendicesStart);

  let userManualTitleIdx = -1;
  let userManualLastIdx = -1;
  let cvTitleIdx = -1;

  for (let i = 0; i < scopedChildren.length; i++) {
    const child = scopedChildren[i];
    if (child.localName !== "p") continue;

    const raw = normalizeText(
      Array.from(child.querySelectorAll("t"))
        .map((n) => n.textContent ?? "")
        .join(""),
    );
    const txt = normalizeText(getParagraphText(child)) || raw;

    if (/^user\s+manual$/iu.test(txt) && userManualTitleIdx === -1) {
      userManualTitleIdx = i;
      continue;
    }

    if (/^curriculum\s+vitae$/iu.test(txt)) {
      cvTitleIdx = i;
      if (userManualTitleIdx !== -1 && userManualLastIdx === -1) {
        for (let j = i - 1; j > userManualTitleIdx; j--) {
          if (scopedChildren[j].localName === "p") {
            userManualLastIdx = j;
            break;
          }
        }
      }
      break;
    }

    if (
      userManualTitleIdx !== -1 &&
      userManualLastIdx === -1 &&
      i > userManualTitleIdx
    ) {
      if (/^appendix\s+[A-Z]$/iu.test(txt)) {
        for (let j = i - 1; j > userManualTitleIdx; j--) {
          if (scopedChildren[j].localName === "p") {
            userManualLastIdx = j;
            break;
          }
        }
      }
    }
  }

  if (userManualTitleIdx !== -1 && userManualLastIdx === -1) {
    for (let j = scopedChildren.length - 1; j > userManualTitleIdx; j--) {
      if (scopedChildren[j].localName === "p") {
        userManualLastIdx = j;
        break;
      }
    }
  }

  if (userManualTitleIdx !== -1) {
    const { hdrRelId, ftrRelId } = await addEmptyHeaderFooterToZip(zip);

    const umTitleEl = scopedChildren[userManualTitleIdx] as Element;
    const umTitleAlreadyHasSectPr =
      getChild(ensurePPr(umTitleEl), "sectPr") !== null;
    if (!umTitleAlreadyHasSectPr) {
      injectSectPr(umTitleEl, buildNormalClosingSectPr(dom, mainSectPr));
    }

    for (let k = userManualTitleIdx + 1; k < scopedChildren.length; k++) {
      const next = scopedChildren[k];
      if (!(next instanceof Element) || next.localName !== "p") break;
      const hasText =
        normalizeText(
          Array.from(next.querySelectorAll("t"))
            .map((n) => n.textContent ?? "")
            .join(""),
        ) !== "";
      const hasImg = next.querySelectorAll("drawing, pict").length > 0;
      if (hasText || hasImg) break;
      next.parentElement?.removeChild(next);
    }

    if (userManualLastIdx !== -1) {
      injectSectPr(
        scopedChildren[userManualLastIdx] as Element,
        buildUserManualSectPr(dom, mainSectPr, hdrRelId, ftrRelId),
      );
    }
  }

  const state: AppState = {
    zone: "before",
    expectingTitle: false,
    currentLetter: "",
    cvSeen: false,
    afterContinuation: false,
  };

  for (let i = 0; i < scopedChildren.length; i++) {
    const child = scopedChildren[i];
    if (!(child instanceof Element)) continue;
    if (child.parentElement !== body) continue;

    if (child.localName === "tbl") {
      if (state.zone === "cv") {
        child.parentElement?.removeChild(child);
      } else {
        state.expectingTitle = false;
        // Format table cells using config.table
        Array.from(child.querySelectorAll("tc")).forEach((tc) => {
          if (tc.namespaceURI !== W_NS) return;
          Array.from(tc.querySelectorAll("p")).forEach((p) => {
            if (p.namespaceURI !== W_NS) return;
            writePAlignment(p, options.config.table.alignment);
            writePSpacing(p, 0, 0, linesToTwips(options.config.table.lineSpacing));
            const sz = ptsToHalfPts(options.config.table.fontSize);
            writePPrRPr(p, options.config.table.fontFamily, sz, false, false);
            applyRunFormatting(p, options.config.table.fontFamily, sz, false, false);
          });
        });
      }
      continue;
    }

    if (child.localName !== "p") continue;

    const txtClean = normalizeText(getParagraphText(child));
    const txtRaw = normalizeText(
      Array.from(child.querySelectorAll("t"))
        .map((n) => n.textContent ?? "")
        .join(""),
    );
    const txt = txtClean || txtRaw;
    const hasDrawing = child.querySelectorAll("drawing, pict").length > 0;

    const hasSectPr =
      child.querySelector(":scope > pPr > sectPr") !== null ||
      getChild(ensurePPr(child), "sectPr") !== null;

    const isCurriculumVitae = /^curriculum\s+vitae$/iu.test(txt);

    if (isCurriculumVitae) {
      if (!state.cvSeen) {
        state.zone = "cv";
        state.expectingTitle = false;
        state.cvSeen = true;
        applyAppendicesSectionTitle(child, options.config);
      } else {
        child.parentElement?.removeChild(child);
      }
      continue;
    }

    if (state.zone === "cv") {
      if (hasSectPr) {
        // leave alone — defines section break
      } else if (hasDrawing) {
        applyCVImageParagraph(child, options.config);
      } else {
        child.parentElement?.removeChild(child);
      }
      continue;
    }

    const isAppendicesTitle = /^appendices$/iu.test(txt);
    const appendixLetterM = txt.match(/^appendix\s+([A-Z])$/iu);
    const isAppendixLetter = appendixLetterM !== null;
    const isContinuationLabel = /^continuation\s+of\s+appendix/iu.test(txt);
    const isUserManual = /^user\s+manual$/iu.test(txt);

    if (isAppendicesTitle) {
      state.zone = "appendices";
      state.expectingTitle = false;
      applyAppendicesSectionTitle(child, options.config);
      continue;
    }

    if (isAppendixLetter) {
      state.zone = "appendices";
      state.currentLetter = appendixLetterM![1].toUpperCase();
      state.expectingTitle = true;
      state.afterContinuation = false;
      applyAppendixLetter(child, state.currentLetter, options.config);
      continue;
    }

    if (isContinuationLabel) {
      state.expectingTitle = false;
      state.afterContinuation = true;
      applyContinuationAppendixLabel(child, state.currentLetter, options.config);
      continue;
    }

    if (isUserManual) {
      state.zone = "user_manual";
      state.expectingTitle = false;
      applyAppendixTitle(child, options.config);
      continue;
    }

    if (txt === "" && !hasDrawing) {
      if (!hasSectPr) applyEmptyParagraph(child, options.config);
      continue;
    }

    if (state.zone === "user_manual") {
      if (hasDrawing) {
        applyUserManualImageParagraph(child, options.config);
      }
      continue;
    }

    if (state.expectingTitle && txt !== "") {
      state.expectingTitle = false;
      applyAppendixTitle(child, options.config);
      continue;
    }

    if (hasDrawing && state.zone === "appendices") {
      const targetH = state.afterContinuation
        ? APPENDIX_TARGET_H_CONT_EMU
        : APPENDIX_TARGET_H_EMU;
      applyAppendixFigureParagraph(child, targetH, options.config);
      continue;
    }
  }

  const serializer = new XMLSerializer();
  zip.file("word/document.xml", serializer.serializeToString(dom));
  return zip.generateAsync({ type: "blob" }) as Promise<Blob>;
}
