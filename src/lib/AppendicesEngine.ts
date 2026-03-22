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
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

// ─── EMU constants ─────────────────────────────────────────────────────────────
// User-Manual image: width 18.20 cm, height 27 cm
const UM_IMAGE_W_EMU = 6_552_000; // 18.20 cm × 360 000
const UM_IMAGE_H_EMU = 9_720_000; //  27.00 cm × 360 000

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
      // Skip text inside text boxes
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

/** Write paragraph-mark rPr (affects default run style for new/empty runs). */
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

/** Apply font / size / bold / italic to all non-special text runs. */
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

/** Strip common pPr children (NOT sectPr — that must survive). */
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

/**
 * Replace all text content in a paragraph with a single clean run.
 * Drawings / special runs are left untouched.
 */
function rewriteParagraphText(p: Element, text: string) {
  Array.from(p.querySelectorAll("r"))
    .filter(
      (r) =>
        r.namespaceURI === W_NS &&
        !Array.from(r.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        ),
    )
    .forEach((r) => p.removeChild(r));

  const run = wElem(p.ownerDocument!, "r");
  const tEl = wElem(p.ownerDocument!, "t");
  tEl.textContent = text;
  run.appendChild(tEl);
  p.appendChild(run);
}

// ─── paragraph formatters ───────────────────────────────────────────────────────

/**
 * APPENDICES title page / CURRICULUM VITAE title page.
 *
 * • Page break before (alone on page)
 * • UPPERCASE text
 * • Garamond 14 pt
 * • Centered
 * • 3.0 line spacing
 * • 0 before / after paragraph spacing
 * • 0 indentation
 * • No whitespace before / after title words
 */
function applyAppendicesSectionTitle(p: Element) {
  stripPPr(p);
  uppercaseParagraph(p);
  trimParagraphText(p);

  const pPr = ensurePPr(p);

  // Page break before — title must be alone on its page
  pPr.appendChild(wElem(p.ownerDocument!, "pageBreakBefore"));

  // Center alignment
  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);

  // 0 indentation
  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  // 3.0 line spacing, 0 before / after
  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "720"); // 3 × 240
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  // Garamond 14 pt (28 half-points), bold
  writePPrRPr(p, "Garamond", 28, true, false);
  applyRunFormatting(p, "Garamond", 28, true, false);
}

/**
 * Appendix letter header — e.g. "Appendix A".
 *
 * • Rewrites text to canonical "Appendix X"
 * • Page break before (each appendix begins on a new page)
 * • Garamond 14 pt, bold
 * • Centered
 * • 1.0 line spacing
 * • 0 before / after, 0 indentation
 */
function applyAppendixLetter(p: Element, letter: string) {
  const canonical = `Appendix ${letter.toUpperCase()}`;
  rewriteParagraphText(p, canonical);
  stripPPr(p);

  const pPr = ensurePPr(p);

  removeChildren(pPr, "sectPr");
  pPr.appendChild(wElem(p.ownerDocument!, "pageBreakBefore"));

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  // 1.0 line spacing
  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "240"); // 1.0×
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  writePPrRPr(p, "Garamond", 28, true, false);
  applyRunFormatting(p, "Garamond", 28, true, false);
}

/**
 * Appendix title — the single paragraph immediately after the appendix letter.
 *
 * • UPPERCASE, trim whitespace
 * • Garamond 14 pt, bold
 * • Centered
 * • 3.0 line spacing
 * • 0 before / after, 0 indentation
 */
function applyAppendixTitle(p: Element) {
  stripPPr(p);
  uppercaseParagraph(p);
  trimParagraphText(p);

  const pPr = ensurePPr(p);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "720"); // 3.0×
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  writePPrRPr(p, "Garamond", 28, true, false);
  applyRunFormatting(p, "Garamond", 28, true, false);
}

/**
 * Continuation of Appendix label — "Continuation of Appendix X..."
 *
 * • Rewrites to canonical form (exactly 3 dots)
 * • Garamond 13 pt, italic
 * • Left-aligned
 * • 3.0 line spacing
 * • 0 before / after, 0 indentation
 */
function applyContinuationAppendixLabel(p: Element, currentLetter: string) {
  const canonical = currentLetter
    ? `Continuation of Appendix ${currentLetter.toUpperCase()}...`
    : "Continuation of Appendix...";
  rewriteParagraphText(p, canonical);
  stripPPr(p);

  const pPr = ensurePPr(p);

  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", "left");
  pPr.appendChild(jc);

  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "720"); // 3.0×
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  // Garamond 13 pt (26 half-points), italic, not bold
  writePPrRPr(p, "Garamond", 26, false, true);
  applyRunFormatting(p, "Garamond", 26, false, true);
}

/**
 * Zero-height empty paragraph (used between content elements).
 * Leaves any existing sectPr alone.
 */
function applyEmptyParagraph(p: Element) {
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
}

// ─── image helpers ─────────────────────────────────────────────────────────────

/** Walk a subtree and return all w:drawing elements. */
function collectDrawings(node: Element): Element[] {
  const result: Element[] = [];
  for (const c of Array.from(node.childNodes)) {
    if (!(c instanceof Element)) continue;
    if (c.localName === "drawing" && c.namespaceURI === W_NS) result.push(c);
    else result.push(...collectDrawings(c));
  }
  return result;
}

/**
 * Convert a floating wp:anchor to wp:inline so the image flows with text.
 * If the drawing already has an inline child, do nothing.
 */
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

  // Copy extent
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

  // Copy docPr, cNvGraphicFramePr, graphic
  Array.from(anchor.childNodes).forEach((c) => {
    if (
      c instanceof Element &&
      ["docPr", "cNvGraphicFramePr", "graphic"].includes(c.localName)
    )
      inline.appendChild(c.cloneNode(true));
  });

  drawing.replaceChild(inline, anchor);
}

/**
 * Resize an inline image to the target EMU dimensions.
 * Updates wp:extent and pic:spPr a:xfrm a:ext (if present).
 */
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

  // wp:extent
  const extent = Array.from(inline.childNodes).find(
    (c): c is Element => c instanceof Element && c.localName === "extent",
  );
  if (extent) {
    extent.setAttribute("cx", String(widthEmu));
    extent.setAttribute("cy", String(heightEmu));
  }

  // a:xfrm > a:ext inside any pic:spPr
  inline.querySelectorAll("ext").forEach((ext) => {
    const parent = ext.parentElement;
    if (parent && parent.localName === "xfrm") {
      ext.setAttribute("cx", String(widthEmu));
      ext.setAttribute("cy", String(heightEmu));
    }
  });
}

/**
 * Apply User Manual image paragraph formatting:
 *   Right-aligned paragraph, inline image, 18.20 cm × 27 cm.
 */
function applyUserManualImageParagraph(p: Element) {
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

  // Right-aligned
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
 * Apply CV image paragraph formatting:
 *   Centred paragraph, inline image (no resize).
 */
function applyCVImageParagraph(p: Element) {
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

  collectDrawings(p).forEach((drawing) => convertAnchorToInline(drawing));
}

// ─── section-property helpers ──────────────────────────────────────────────────

/**
 * Inject (or replace) a w:sectPr inside a paragraph's w:pPr.
 * The paragraph becomes the last paragraph of its section.
 * NB: sectPr must be the LAST child of pPr per the OOXML schema.
 */
function injectSectPr(p: Element, sectPr: Element) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "sectPr");
  pPr.appendChild(sectPr);
}

/**
 * Create blank header and footer XML files inside the zip, register their
 * relationships and content-type overrides, then return the two rel IDs.
 *
 * Adding explicit (but empty) headerReference / footerReference to the User
 * Manual sectPr is what disables "Link to Previous" in Word — without them
 * the section just inherits whatever header/footer came before it.
 */
async function addEmptyHeaderFooterToZip(
  zip: any,
): Promise<{ hdrRelId: string; ftrRelId: string }> {
  // ── 1. Empty header / footer XML ─────────────────────────────────────────────
  // The paragraph is required by the OOXML schema but we collapse it to
  // 1-twip exact height so it occupies zero visible space.
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

  // ── 2. Update word/_rels/document.xml.rels ────────────────────────────────────
  const relsPath = "word/_rels/document.xml.rels";
  const relsRaw: string =
    (await zip.file(relsPath)?.async("string")) ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

  const relsParser = new DOMParser();
  const relsDom = relsParser.parseFromString(relsRaw, "application/xml");
  const relsRoot = relsDom.documentElement;

  // Find highest existing numeric rId so we don't collide
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

  // ── 3. Update [Content_Types].xml ────────────────────────────────────────────
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

/**
 * Build a sectPr that ends the normal appendices section and starts a new one
 * (type = nextPage). Copies page-size from the main body sectPr; strips
 * headerReference / footerReference so the inline sectPr doesn't carry stale
 * relationship IDs.
 */
function buildNormalClosingSectPr(
  doc: Document,
  mainSectPr: Element | null,
): Element {
  let sectPr: Element;

  if (mainSectPr) {
    sectPr = mainSectPr.cloneNode(true) as Element;
  } else {
    sectPr = wElem(doc, "sectPr");
    // Default A4 page size if the doc has no sectPr
    const pgSz = wElem(doc, "pgSz");
    setWAttr(pgSz, "w", "11906");
    setWAttr(pgSz, "h", "16838");
    sectPr.appendChild(pgSz);
  }

  // Force nextPage type
  removeChildren(sectPr, "type");
  const typeEl = wElem(doc, "type");
  setWAttr(typeEl, "val", "nextPage");
  sectPr.insertBefore(typeEl, sectPr.firstChild);

  return sectPr;
}

/**
 * Build the User Manual section's sectPr:
 *   • Explicit empty headerReference / footerReference — disables "Link to Previous"
 *     so Word shows truly blank headers/footers (no text, shapes, or images)
 *   • All margins 0.3 cm (≈ 170 twips)
 *   • Header distance from top: 0
 *   • Footer distance from bottom: 0
 *   • Page size copied from main sectPr (A4 fallback)
 *   • type = nextPage
 */
function buildUserManualSectPr(
  doc: Document,
  mainSectPr: Element | null,
  hdrRelId: string,
  ftrRelId: string,
): Element {
  const sectPr = wElem(doc, "sectPr");

  // type = nextPage
  const typeEl = wElem(doc, "type");
  setWAttr(typeEl, "val", "nextPage");
  sectPr.appendChild(typeEl);

  // Page size — reuse from main sectPr so paper size stays correct
  if (mainSectPr) {
    const pgSz = getChild(mainSectPr, "pgSz");
    if (pgSz) sectPr.appendChild(pgSz.cloneNode(true));
  } else {
    const pgSz = wElem(doc, "pgSz");
    setWAttr(pgSz, "w", "11906");
    setWAttr(pgSz, "h", "16838");
    sectPr.appendChild(pgSz);
  }

  // 0.3 cm margins on all sides; header / footer distance = 0
  const pgMar = wElem(doc, "pgMar");
  setWAttr(pgMar, "top", "170");
  setWAttr(pgMar, "right", "170");
  setWAttr(pgMar, "bottom", "170");
  setWAttr(pgMar, "left", "170");
  setWAttr(pgMar, "header", "0");
  setWAttr(pgMar, "footer", "0");
  setWAttr(pgMar, "gutter", "0");
  sectPr.appendChild(pgMar);

  // ── Explicit empty header/footer references ───────────────────────────────────
  // Adding these for all three types (default, first, even) fully disables
  // "Link to Previous" — Word will show the blank headerUM/footerUM files
  // instead of inheriting content from any prior section.
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

type AppZone =
  | "before" // haven't reached APPENDICES yet
  | "appendices" // inside the appendices section (general)
  | "user_manual" // inside the User Manual special section
  | "cv"; // inside the Curriculum Vitae section

interface AppState {
  zone: AppZone;
  /** true immediately after an appendix letter — next non-empty para is the title */
  expectingTitle: boolean;
  currentLetter: string;
  /** true once the first CURRICULUM VITAE title has been processed */
  cvSeen: boolean;
}

// ─── public entry ──────────────────────────────────────────────────────────────

export async function formatAppendices(
  arrayBuffer: ArrayBuffer,
): Promise<Blob> {
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

  // The main document sectPr is the sole w:sectPr that is a DIRECT child of w:body.
  const mainSectPr =
    (Array.from(body.childNodes).find(
      (c): c is Element =>
        c instanceof Element &&
        c.localName === "sectPr" &&
        c.namespaceURI === W_NS,
    ) as Element | null) ?? null;

  // ── locate APPENDICES section start ──────────────────────────────────────────
  // We look for the first paragraph whose text is exactly "APPENDICES"
  // (not "LIST OF APPENDICES") that appears after the REFERENCES heading.

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

  // Nothing to do if no APPENDICES section is found
  if (appendicesStart === -1) {
    const serializer = new XMLSerializer();
    zip.file("word/document.xml", serializer.serializeToString(dom));
    return zip.generateAsync({ type: "blob" }) as Promise<Blob>;
  }

  allChildren.slice(appendicesStart).forEach((el) => unwrapContentControls(el));

  const scopedChildren = Array.from(body.childNodes)
    .filter((c): c is Element => c instanceof Element && c !== mainSectPr)
    .slice(appendicesStart);

  // ── PRE-PASS: locate User Manual and CV boundaries ────────────────────────────
  //
  // User Manual is an appendix title (appears right after an appendix letter).
  // The paragraph that IS the "User Manual" title gets a sectPr injected into its
  // pPr — this closes the normal-appendices section and the User Manual CONTENT
  // begins in the next section.
  //
  // The LAST paragraph of User Manual content gets a sectPr with the special
  // (0.3 cm, no header/footer) properties — this closes the User Manual section.
  // Everything after it falls into the final section (uses the body's sectPr).

  let userManualTitleIdx = -1; // index within scopedChildren
  let userManualLastIdx = -1; // index of last UM content paragraph
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
        // Walk backwards to last paragraph before CV
        for (let j = i - 1; j > userManualTitleIdx; j--) {
          if (scopedChildren[j].localName === "p") {
            userManualLastIdx = j;
            break;
          }
        }
      }
      break;
    }

    // Next appendix letter after the UM title also ends the UM section
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

  // If User Manual was found but no explicit end was detected, end at document tail
  if (userManualTitleIdx !== -1 && userManualLastIdx === -1) {
    for (let j = scopedChildren.length - 1; j > userManualTitleIdx; j--) {
      if (scopedChildren[j].localName === "p") {
        userManualLastIdx = j;
        break;
      }
    }
  }

  // ── Inject sectPr elements (must happen BEFORE the formatting pass) ───────────

  if (userManualTitleIdx !== -1) {
    // Create empty header/footer files, register their rels and content types.
    // Must be awaited before we build the sectPr so we have the rel IDs ready.
    const { hdrRelId, ftrRelId } = await addEmptyHeaderFooterToZip(zip);

    // Inject sectPr on UM title only if it doesn't already have one.
    const umTitleEl = scopedChildren[userManualTitleIdx] as Element;
    const umTitleAlreadyHasSectPr =
      getChild(ensurePPr(umTitleEl), "sectPr") !== null;
    if (!umTitleAlreadyHasSectPr) {
      injectSectPr(umTitleEl, buildNormalClosingSectPr(dom, mainSectPr));
    }

    // Delete empty paragraphs immediately after the UM title — these produce
    // blank pages. If an empty paragraph carries a sectPr it is a redundant
    // "Section Break (Next Page)" paragraph left by the source doc; remove it
    // too since the UM title already has the real section break.
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
      // Empty paragraph (with or without sectPr) — remove it.
      next.parentElement?.removeChild(next);
    }

    // Last User Manual content paragraph closes the blank-header section
    if (userManualLastIdx !== -1) {
      injectSectPr(
        scopedChildren[userManualLastIdx] as Element,
        buildUserManualSectPr(dom, mainSectPr, hdrRelId, ftrRelId),
      );
    }
  }

  // ── MAIN FORMATTING PASS ──────────────────────────────────────────────────────

  const state: AppState = {
    zone: "before",
    expectingTitle: false,
    currentLetter: "",
    cvSeen: false,
  };

  for (let i = 0; i < scopedChildren.length; i++) {
    const child = scopedChildren[i];
    if (!(child instanceof Element)) continue;
    if (child.parentElement !== body) continue;

    // Tables: remove entirely in the CV zone; leave as-is everywhere else
    if (child.localName === "tbl") {
      if (state.zone === "cv") {
        child.parentElement?.removeChild(child);
      } else {
        state.expectingTitle = false;
      }
      continue;
    }

    if (child.localName !== "p") continue;

    // ── gather text (both txbxContent-aware and raw) ──────────────────────────
    const txtClean = normalizeText(getParagraphText(child));
    const txtRaw = normalizeText(
      Array.from(child.querySelectorAll("t"))
        .map((n) => n.textContent ?? "")
        .join(""),
    );
    const txt = txtClean || txtRaw;
    const hasDrawing = child.querySelectorAll("drawing, pict").length > 0;

    // ── check whether this paragraph already carries an injected sectPr ───────
    // sectPr-bearing paragraphs define section breaks and must never be deleted.
    const hasSectPr =
      child.querySelector(":scope > pPr > sectPr") !== null ||
      getChild(ensurePPr(child), "sectPr") !== null;

    // ── CURRICULUM VITAE detection (must come before the cv-zone purge below) ──
    // We need to detect the title FIRST so we can format/remove it, then the
    // zone transitions to "cv" and everything after is purged.
    const isCurriculumVitae = /^curriculum\s+vitae$/iu.test(txt);

    if (isCurriculumVitae) {
      if (!state.cvSeen) {
        // First CV title — format it as a standalone title page
        state.zone = "cv";
        state.expectingTitle = false;
        state.cvSeen = true;
        applyAppendicesSectionTitle(child);
      } else {
        // Duplicate CV title — remove it entirely
        child.parentElement?.removeChild(child);
      }
      continue;
    }

    // ── CV zone purge — runs before every other check ─────────────────────────
    // After the CV title page: images are kept (centred, inline), everything
    // else is deleted. sectPr-bearing paragraphs are never removed.
    if (state.zone === "cv") {
      if (hasSectPr) {
        // Leave sectPr paragraphs completely alone — they define section breaks
      } else if (hasDrawing) {
        // Keep images: convert to inline, centre them
        applyCVImageParagraph(child);
      } else {
        // Remove all text, empty lines, continuation labels, etc.
        child.parentElement?.removeChild(child);
      }
      continue;
    }

    // ── pattern detection (non-CV zones only) ─────────────────────────────────
    const isAppendicesTitle = /^appendices$/iu.test(txt);
    const appendixLetterM = txt.match(/^appendix\s+([A-Z])$/iu);
    const isAppendixLetter = appendixLetterM !== null;
    const isContinuationLabel = /^continuation\s+of\s+appendix/iu.test(txt);
    const isUserManual = /^user\s+manual$/iu.test(txt);

    // ── zone transitions ──────────────────────────────────────────────────────

    if (isAppendicesTitle) {
      state.zone = "appendices";
      state.expectingTitle = false;
      applyAppendicesSectionTitle(child);
      continue;
    }

    if (isAppendixLetter) {
      state.zone = "appendices";
      state.currentLetter = appendixLetterM![1].toUpperCase();
      state.expectingTitle = true;
      applyAppendixLetter(child, state.currentLetter);
      continue;
    }

    if (isContinuationLabel) {
      state.expectingTitle = false;
      applyContinuationAppendixLabel(child, state.currentLetter);
      continue;
    }

    if (isUserManual) {
      // Formatted as an appendix title; sectPr was already injected in pre-pass
      state.zone = "user_manual";
      state.expectingTitle = false;
      applyAppendixTitle(child); // UPPERCASE, 14 pt bold, centered, 3.0 line
      continue;
    }

    // ── empty paragraph ───────────────────────────────────────────────────────
    if (txt === "" && !hasDrawing) {
      if (!hasSectPr) applyEmptyParagraph(child);
      continue;
    }

    // ── zone-specific content handling ────────────────────────────────────────

    if (state.zone === "user_manual") {
      // Images in the User Manual: right-aligned, inline, 18.20 × 27 cm
      if (hasDrawing) {
        applyUserManualImageParagraph(child);
      }
      // Non-image text in the User Manual: leave as-is
      continue;
    }

    // ── appendix title (first non-empty paragraph after the letter) ───────────
    if (state.expectingTitle && txt !== "") {
      state.expectingTitle = false;
      // If the title happens to be "User Manual", handled above already.
      // For all other appendix titles:
      applyAppendixTitle(child);
      continue;
    }

    // ── regular appendix body: images → inline + centred ─────────────────────
    if (hasDrawing && state.zone === "appendices") {
      applyCVImageParagraph(child); // centred, inline (same as CV images)
      continue;
    }

    // ── all other appendix body text: leave as-is ─────────────────────────────
    // (body prose, table captions, figures captions inside appendices are not
    //  re-formatted by this engine — they inherit the chapter formatting)
  }

  // ── Serialize and repack ──────────────────────────────────────────────────────
  const serializer = new XMLSerializer();
  zip.file("word/document.xml", serializer.serializeToString(dom));
  return zip.generateAsync({ type: "blob" }) as Promise<Blob>;
}
