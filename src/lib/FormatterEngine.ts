/**
 * FormatterEngine.ts
 * Full client-side port of DocxFormatterService.php
 * Uses JSZip (window.JSZip) + DOMParser for DOCX XML manipulation.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export interface FormatOptions {
  sections: string[];
  rules: string[];
  citationStyle?: string;
}

// ─── helpers ────────────────────────────────────────────────────────────────

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
function isInTable(p: Element): boolean {
  let n: Element | null = p.parentElement;
  while (n) {
    if (n.namespaceURI === W_NS && n.localName === "tc") return true;
    n = n.parentElement;
  }
  return false;
}
function getParagraphText(p: Element): string {
  const parts: string[] = [];
  p.querySelectorAll("*").forEach((el) => {
    if (el.localName === "t" && el.namespaceURI === W_NS) {
      // Skip text inside text boxes — they are floating shapes, not paragraph content
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
  return (t.replace(/\s+/gu, " ") ?? t).trim();
}
function getParagraphStyleId(p: Element): string {
  const pStyle = p.querySelector("pPr > pStyle") as Element | null;
  if (!pStyle) return "";
  return wAttr(pStyle, "val");
}

// ─── text case ──────────────────────────────────────────────────────────────

function uppercaseParagraphText(p: Element) {
  p.querySelectorAll("r > t").forEach((t) => {
    if (t.namespaceURI === W_NS)
      t.textContent = (t.textContent ?? "").toUpperCase();
  });
}
function titleCaseParagraphText(p: Element) {
  p.querySelectorAll("r > t").forEach((t) => {
    if (t.namespaceURI === W_NS)
      t.textContent = (t.textContent ?? "")
        .toLowerCase()
        .replace(/(?:^|\s)\S/gu, (c) => c.toUpperCase());
  });
}

// ─── chapter roman/arabic ────────────────────────────────────────────────────

function chapterToInt(token: string): number {
  if (/^\d+$/.test(token)) return parseInt(token, 10);
  const map: Record<string, number> = {
    M: 1000,
    D: 500,
    C: 100,
    L: 50,
    X: 10,
    V: 5,
    I: 1,
  };
  const s = token.toUpperCase();
  let result = 0,
    prev = 0;
  for (let i = s.length - 1; i >= 0; i--) {
    const val = map[s[i]] ?? 0;
    if (val < prev) result -= val;
    else {
      result += val;
      prev = val;
    }
  }
  return result;
}

// ─── property writers ────────────────────────────────────────────────────────

function writePAlignment(p: Element, value: string) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "jc");
  const jc = ensureChild(pPr, "jc");
  setWAttr(jc, "val", value);
}
function writePIndent(p: Element, firstTwips: number) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "ind");
  const ind = wElem(p.ownerDocument!, "ind");
  pPr.appendChild(ind);
  if (firstTwips > 0) {
    setWAttr(ind, "firstLine", String(firstTwips));
  }
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
}
function writePHangingIndent(
  p: Element,
  leftTwips: number,
  hangingTwips: number,
) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "ind");
  const ind = wElem(p.ownerDocument!, "ind");
  pPr.appendChild(ind);
  setWAttr(ind, "left", String(leftTwips));
  setWAttr(ind, "hanging", String(hangingTwips));
}
function writePSpacing(
  p: Element,
  before: number,
  after: number,
  line: number,
) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "spacing");
  const sp = ensureChild(pPr, "spacing");
  setWAttr(sp, "before", String(before));
  setWAttr(sp, "after", String(after));
  setWAttr(sp, "line", String(line));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
}
function writePageBreakBefore(p: Element, enabled: boolean) {
  const pPr = ensurePPr(p);
  if (enabled) ensureChild(pPr, "pageBreakBefore");
  else removeChildren(pPr, "pageBreakBefore");
}
function removePBdr(p: Element) {
  const pPr = getChild(p, "pPr");
  if (pPr) removeChildren(pPr, "pBdr");
}

function writePPrRPr(
  p: Element,
  font: string,
  size: number,
  bold: boolean,
  italic: boolean,
) {
  const pPr = ensurePPr(p);
  const sizeStr = String(size);
  removeChildren(pPr, "rPr");
  const rPr = wElem(p.ownerDocument!, "rPr");
  pPr.appendChild(rPr);
  const rFonts = wElem(p.ownerDocument!, "rFonts");
  setWAttr(rFonts, "ascii", font);
  setWAttr(rFonts, "hAnsi", font);
  setWAttr(rFonts, "eastAsia", font);
  setWAttr(rFonts, "cs", font);
  rPr.appendChild(rFonts);
  if (bold) {
    rPr.appendChild(wElem(p.ownerDocument!, "b"));
    rPr.appendChild(wElem(p.ownerDocument!, "bCs"));
  }
  if (italic) {
    rPr.appendChild(wElem(p.ownerDocument!, "i"));
    rPr.appendChild(wElem(p.ownerDocument!, "iCs"));
  }
  const szEl = wElem(p.ownerDocument!, "sz");
  setWAttr(szEl, "val", sizeStr);
  const szCs = wElem(p.ownerDocument!, "szCs");
  setWAttr(szCs, "val", sizeStr);
  rPr.appendChild(szEl);
  rPr.appendChild(szCs);
}

function writeRuns(
  scope: Element,
  font: string,
  size: number,
  bold: boolean | null,
  italic: boolean | null,
) {
  const sizeStr = String(size);
  const allRuns = Array.from(scope.querySelectorAll("r")).filter(
    (r) => r.namespaceURI === W_NS,
  );
  const textRuns = allRuns.filter((r) => {
    const special = Array.from(r.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar", "AlternateContent"].includes(
        c.localName,
      ),
    );
    return !special;
  });

  let runsToProcess = textRuns;
  if (runsToProcess.length === 0) {
    const textContent = Array.from(scope.querySelectorAll("t"))
      .map((t) => t.textContent ?? "")
      .join("");
    const run = wElem(scope.ownerDocument!, "r");
    const tEl = wElem(scope.ownerDocument!, "t");
    tEl.textContent = textContent;
    if (textContent !== textContent.trim()) {
      tEl.setAttribute("xml:space", "preserve");
    }
    run.appendChild(tEl);
    scope.appendChild(run);
    runsToProcess = [run];
  }

  for (const run of runsToProcess) {
    const existingRPr = getChild(run, "rPr");
    const existingBold =
      existingRPr != null &&
      existingRPr.getElementsByTagNameNS(W_NS, "b").length > 0;
    const existingItalic =
      existingRPr != null &&
      existingRPr.getElementsByTagNameNS(W_NS, "i").length > 0;
    const applyBold = bold ?? existingBold;
    const applyItalic = italic ?? existingItalic;

    removeChildren(run, "rPr");
    const rPr = wElem(run.ownerDocument!, "rPr");
    if (run.firstChild) run.insertBefore(rPr, run.firstChild);
    else run.appendChild(rPr);

    removeChildren(rPr, "rFonts");
    const rFonts = ensureChild(rPr, "rFonts");
    for (const attr of [
      "asciiTheme",
      "hAnsiTheme",
      "eastAsiaTheme",
      "cstheme",
    ]) {
      rFonts.removeAttributeNS(W_NS, attr);
    }
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
    ]) {
      removeChildren(rPr, tag);
    }
  }
}

// ─── strip helpers ───────────────────────────────────────────────────────────

function stripAll(p: Element) {
  const pPr = getChild(p, "pPr");
  if (pPr) {
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
    ]) {
      removeChildren(pPr, tag);
    }
  }
  Array.from(p.querySelectorAll("r")).forEach((run) => {
    if (run.namespaceURI !== W_NS) return;
    const special = Array.from(run.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
    );
    if (!special) {
      const rPr = getChild(run, "rPr");
      if (rPr) removeChildren(rPr, "rStyle");
    }
  });
}

function stripLeadingTabRuns(p: Element) {
  const runs = Array.from(p.childNodes).filter(
    (c): c is Element => c instanceof Element && c.localName === "r",
  );
  for (const run of runs) {
    const hasTab = run.getElementsByTagNameNS(W_NS, "tab").length > 0;
    const hasTxt = run.getElementsByTagNameNS(W_NS, "t").length > 0;
    const hasOther = Array.from(run.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar", "br"].includes(c.localName),
    );
    if (hasOther) break;
    if (hasTab && !hasTxt) {
      p.removeChild(run);
    } else if (hasTab && hasTxt) {
      Array.from(run.getElementsByTagNameNS(W_NS, "tab")).forEach((tab) => {
        tab.parentElement?.removeChild(tab);
      });
      break;
    } else {
      break;
    }
  }
}

// Strip leading runs that contain only arrow/bullet/symbol characters (e.g. →, •, ►)
function stripLeadingArrowRuns(p: Element) {
  const ARROW_ONLY =
    /^[\u2192\u2190\u2191\u2193\u21D2\u25BA\u25B6\u25CF\u2022\u2023\u2043\u2013\u2014\u2015\s\t]+$/u;
  const LEADING_ARROWS =
    /^[\u2192\u2190\u2191\u2193\u21D2\u25BA\u25B6\u25CF\u2022\u2023\u2043\s\t]+/u;

  const runs = Array.from(p.childNodes).filter(
    (c): c is Element => c instanceof Element && c.localName === "r",
  );
  for (const run of runs) {
    const hasOther = Array.from(run.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar", "br"].includes(c.localName),
    );
    if (hasOther) break;

    const hasTab = run.getElementsByTagNameNS(W_NS, "tab").length > 0;
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    const text = tEl?.textContent ?? "";

    if (ARROW_ONLY.test(text) || (hasTab && text.trim() === "")) {
      p.removeChild(run);
      continue;
    }

    if (tEl) {
      const stripped = text.replace(LEADING_ARROWS, "");
      if (stripped !== text) {
        if (stripped === "") {
          p.removeChild(run);
        } else {
          tEl.textContent = stripped;
          tEl.setAttribute("xml:space", "preserve");
        }
      }
      break;
    }

    break;
  }
}

function stripTrailingTextWhitespace(p: Element) {
  const runs = Array.from(p.childNodes).filter(
    (c): c is Element =>
      c instanceof Element && c.localName === "r" && c.namespaceURI === W_NS,
  );
  for (let i = runs.length - 1; i >= 0; i--) {
    const run = runs[i];
    const hasOther = Array.from(run.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar", "br", "tab"].includes(
        c.localName,
      ),
    );
    if (hasOther) continue;
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    if (!tEl) continue;
    const original = tEl.textContent ?? "";
    const trimmed = original.replace(/[\s\t]+$/, "");
    if (trimmed === "") {
      p.removeChild(run);
      continue;
    }
    if (trimmed !== original) {
      tEl.textContent = trimmed;
      tEl.setAttribute("xml:space", "preserve");
    }
    break;
  }
}

function stripLeadingTextWhitespace(p: Element) {
  const runs = Array.from(p.childNodes).filter(
    (c): c is Element =>
      c instanceof Element && c.localName === "r" && c.namespaceURI === W_NS,
  );
  for (const run of runs) {
    const hasOther = Array.from(run.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar", "br", "tab"].includes(
        c.localName,
      ),
    );
    if (hasOther) continue;
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    if (!tEl) continue;
    const original = tEl.textContent ?? "";
    const trimmed = original.replace(/^[\s\t]+/, "");
    if (trimmed === "") {
      p.removeChild(run);
      continue;
    }
    if (trimmed !== original) {
      tEl.textContent = trimmed;
      tEl.setAttribute("xml:space", "preserve");
    }
    break;
  }
}

function stripOrphanedBookmarks(p: Element) {
  const toRemove: Element[] = [];
  Array.from(p.childNodes).forEach((c) => {
    if (!(c instanceof Element)) return;
    if (!["bookmarkStart", "bookmarkEnd"].includes(c.localName)) return;
    const name = wAttr(c, "name");
    if (name.startsWith("_") || name === "") toRemove.push(c);
  });
  toRemove.forEach((n) => p.removeChild(n));
}

function getPFirstLineIndent(p: Element): number {
  const ind = p.getElementsByTagNameNS(W_NS, "ind").item(0) as Element | null;
  if (!ind) return 0;
  return parseInt(wAttr(ind, "firstLine") || "0", 10);
}

// ─── content control unwrapper ───────────────────────────────────────────────

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
        Array.from(sdtContent.childNodes).forEach((child) => {
          parent.insertBefore(child.cloneNode(true), sdt);
        });
      }
      parent.removeChild(sdt);
      changed = true;
      break;
    }
  }
}

// ─── paragraph formatters ────────────────────────────────────────────────────

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

  const rPr = wElem(p.ownerDocument!, "rPr");
  const rFonts = wElem(p.ownerDocument!, "rFonts");
  for (const a of ["ascii", "hAnsi", "eastAsia", "cs"])
    setWAttr(rFonts, a, "Garamond");
  rPr.appendChild(rFonts);
  const szEl = wElem(p.ownerDocument!, "sz");
  setWAttr(szEl, "val", "2");
  const szCs = wElem(p.ownerDocument!, "szCs");
  setWAttr(szCs, "val", "2");
  rPr.appendChild(szEl);
  rPr.appendChild(szCs);
  pPr.appendChild(rPr);

  Array.from(p.querySelectorAll("r")).forEach((run) => {
    if (run.namespaceURI !== W_NS) return;
    removeChildren(run, "rPr");
    const rr = wElem(p.ownerDocument!, "rPr");
    const s1 = wElem(p.ownerDocument!, "sz");
    setWAttr(s1, "val", "2");
    const s2 = wElem(p.ownerDocument!, "szCs");
    setWAttr(s2, "val", "2");
    rr.appendChild(s1);
    rr.appendChild(s2);
    run.insertBefore(rr, run.firstChild);
  });
}

function applyReferenceEmptyLine(p: Element) {
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
  setWAttr(sp, "line", "240");
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const rPr = wElem(p.ownerDocument!, "rPr");
  const rFonts = wElem(p.ownerDocument!, "rFonts");
  for (const a of ["ascii", "hAnsi", "eastAsia", "cs"])
    setWAttr(rFonts, a, "Garamond");
  rPr.appendChild(rFonts);
  const szEl = wElem(p.ownerDocument!, "sz");
  setWAttr(szEl, "val", "22");
  const szCs = wElem(p.ownerDocument!, "szCs");
  setWAttr(szCs, "val", "22");
  rPr.appendChild(szEl);
  rPr.appendChild(szCs);
  pPr.appendChild(rPr);
}

function buildZeroHeightPageBreakP(doc: Document): Element {
  const p = wElem(doc, "p");
  const pPr = wElem(doc, "pPr");
  p.appendChild(pPr);
  const sp = wElem(doc, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "20");
  setWAttr(sp, "lineRule", "exact");
  pPr.appendChild(sp);
  const rPr = wElem(doc, "rPr");
  const sz = wElem(doc, "sz");
  setWAttr(sz, "val", "2");
  const szC = wElem(doc, "szCs");
  setWAttr(szC, "val", "2");
  rPr.appendChild(sz);
  rPr.appendChild(szC);
  pPr.appendChild(rPr);
  const run = wElem(doc, "r");
  const br = wElem(doc, "br");
  setWAttr(br, "type", "page");
  run.appendChild(br);
  p.appendChild(run);
  return p;
}

function buildContinuationP(doc: Document, label: string): Element {
  const contP = wElem(doc, "p");
  const pPr = wElem(doc, "pPr");
  contP.appendChild(pPr);

  const jc = wElem(doc, "jc");
  setWAttr(jc, "val", "left");
  pPr.appendChild(jc);
  const ind = wElem(doc, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);
  const sp = wElem(doc, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "240");
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const makeRPr = () => {
    const rPr = wElem(doc, "rPr");
    const rFonts = wElem(doc, "rFonts");
    for (const a of ["ascii", "hAnsi", "eastAsia", "cs"])
      setWAttr(rFonts, a, "Garamond");
    rPr.appendChild(rFonts);
    rPr.appendChild(wElem(doc, "i"));
    rPr.appendChild(wElem(doc, "iCs"));
    const sz = wElem(doc, "sz");
    setWAttr(sz, "val", "26");
    rPr.appendChild(sz);
    const sc = wElem(doc, "szCs");
    setWAttr(sc, "val", "26");
    rPr.appendChild(sc);
    return rPr;
  };
  pPr.appendChild(makeRPr());

  const run = wElem(doc, "r");
  run.appendChild(makeRPr());
  const tEl = wElem(doc, "t");
  tEl.textContent = label;
  run.appendChild(tEl);
  contP.appendChild(run);
  return contP;
}

function ensureTrailingPeriod(p: Element) {
  const tNodes = Array.from(p.querySelectorAll("r > t")).filter(
    (t) => t.namespaceURI === W_NS,
  );
  if (tNodes.length === 0) return;
  const last = tNodes[tNodes.length - 1];
  const text = last.textContent ?? "";
  if (!text.trimEnd().endsWith(".")) last.textContent = text.trimEnd() + ".";
}

// ─── apply paragraph types ───────────────────────────────────────────────────

interface Rules {
  spacing?: boolean;
  indentation?: boolean;
  alignment?: boolean;
  captions?: boolean;
  continuation?: boolean;
  borders?: boolean;
  pagination?: boolean;
}

// ─── white cover shape namespace constants ────────────────────────────────────

const WP_NS_DRAW =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS_DRAW =
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

function applyChapterLabel(p: Element, rules: Rules) {
  // First collect the chapter label text before stripping anything
  const rawText = getParagraphText(p);
  const m = normalizeText(rawText).match(/^chapter\s+([ivxlcdm]+|\d+)$/iu);
  const token = m ? (/^\d+$/.test(m[1]) ? m[1] : m[1].toUpperCase()) : null;
  const canonical = token ? "Chapter " + token : normalizeText(rawText);

  // Remove only text runs — any run containing a drawing, shape, or
  // AlternateContent is left completely untouched.
  Array.from(p.childNodes)
    .filter(
      (c): c is Element =>
        c instanceof Element &&
        c.localName === "r" &&
        c.namespaceURI === W_NS &&
        !Array.from(c.children).some((ch) =>
          [
            "drawing",
            "pict",
            "AlternateContent",
            "instrText",
            "fldChar",
          ].includes(ch.localName),
        ),
    )
    .forEach((r) => p.removeChild(r));

  // Insert a single clean run with the canonical text
  const cleanRun = wElem(p.ownerDocument!, "r");
  const tEl = wElem(p.ownerDocument!, "t");
  tEl.textContent = canonical;
  cleanRun.appendChild(tEl);
  p.appendChild(cleanRun);

  stripAll(p);

  // Remove empty paragraphs before this chapter label — but NEVER remove
  // a paragraph that carries a sectPr, as that controls page numbering.
  if (p.parentElement) {
    let prev = p.previousSibling as Element | null;
    while (
      prev instanceof Element &&
      prev.localName === "p" &&
      normalizeText(getParagraphText(prev)) === ""
    ) {
      const pPrEl = prev.getElementsByTagNameNS(W_NS, "pPr").item(0);
      const hasSectPr =
        pPrEl != null &&
        pPrEl.getElementsByTagNameNS(W_NS, "sectPr").length > 0;
      if (hasSectPr) break;
      const toRemove = prev;
      prev = prev.previousSibling as Element | null;
      p.parentElement.removeChild(toRemove);
    }
  }

  writePAlignment(p, "center");
  writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 480);
  if (rules.pagination) writePageBreakBefore(p, true);
  writeRuns(p, "Garamond", 28, true, false);
  writePPrRPr(p, "Garamond", 28, true, false);
}

function applyChapterTitle(p: Element, rules: Rules) {
  stripAll(p);
  stripOrphanedBookmarks(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  uppercaseParagraphText(p);
  writePAlignment(p, "center");
  writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 720);
  writeRuns(p, "Garamond", 28, true, false);
  writePPrRPr(p, "Garamond", 28, true, false);
}

function applyReferencesTitle(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  uppercaseParagraphText(p);
  writePAlignment(p, "center");
  writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 720);
  writeRuns(p, "Garamond", 28, true, false);
  writePPrRPr(p, "Garamond", 28, true, false);
}

function applyHeading(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "left");
  if (rules.indentation) writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 480);
  writeRuns(p, "Garamond", 26, true, false);
  writePPrRPr(p, "Garamond", 26, true, false);
}

function applyItalicHeading(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "left");
  if (rules.spacing) writePSpacing(p, 0, 0, 480);
  const pPr = ensurePPr(p);
  removeChildren(pPr, "ind");
  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "hanging", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);
  writeRuns(p, "Garamond", 24, true, true);
  writePPrRPr(p, "Garamond", 24, true, true);
}

function applyBodyParagraph(p: Element, rules: Rules, beforeSpacing = 0) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "both");
  if (rules.indentation) writePIndent(p, 720);
  else writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, beforeSpacing, 0, 480);
  removePBdr(p);
  writeRuns(p, "Garamond", 24, null, null);
  writePPrRPr(p, "Garamond", 24, false, false);
}

function applyReferenceEntry(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "both");
  if (rules.indentation) writePHangingIndent(p, 720, 720);
  else writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 240);
  writeRuns(p, "Garamond", 22, false, false);
  writePPrRPr(p, "Garamond", 22, false, false);
}

// ─── List paragraph formatter ────────────────────────────────────────────────
//
// Indentation spec (from user requirement):
//   ilvl=0  top-level numbered items  → left=363 twips (0.64cm), hanging=363
//   ilvl=1  sub-items (2.1, 2.2 …)   → left=720 twips (1.27cm), hanging=0
//
// IMPORTANT: We keep w:numPr so Word still renders the actual numbers (1., 2., 3.…).
// We strip only ind and tabs, then write explicit w:ind to override Word's default
// numbering indentation with our desired values.

function getNumIlvl(p: Element): number {
  const numPr = p.querySelector("pPr > numPr");
  if (!numPr) return 0;
  const ilvlEl = getChild(numPr, "ilvl");
  if (!ilvlEl) return 0;
  return parseInt(wAttr(ilvlEl, "val") || "0", 10);
}

function applyListParagraph(p: Element, rules: Rules) {
  // Read indent level BEFORE any stripping
  const ilvl = getNumIlvl(p);

  const pPr = getChild(p, "pPr");
  if (pPr) {
    for (const tag of [
      "pStyle",
      "rPr",
      "widowControl",
      "ind", // remove old ind — we write a new one below
      "spacing",
      "jc",
      "pageBreakBefore",
      "keepNext",
      "keepLines",
      // NOTE: numPr is intentionally NOT stripped — it provides the actual numbers
      "tabs", // remove tabs — ind takes over indentation control
    ])
      removeChildren(pPr, tag);
  }

  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);

  if (rules.alignment) writePAlignment(p, "both");
  if (rules.spacing) writePSpacing(p, 0, 0, 480);

  // Write explicit ind to override whatever the numbering definition says.
  // CRITICAL: ind must be inserted AFTER numPr in the pPr child order.
  // Word resolves indentation by taking the last ind it sees; if ind appears
  // before numPr the numbering definition wins and our values are ignored.
  //   ilvl=0 → left=363 (0.64cm), hanging=363
  //   ilvl=1 → left=720 (1.27cm), firstLine=0
  const pPrEl = ensurePPr(p);
  removeChildren(pPrEl, "ind");
  const ind = wElem(p.ownerDocument!, "ind");
  if (ilvl === 0) {
    setWAttr(ind, "left", "363");
    setWAttr(ind, "hanging", "363");
  } else {
    setWAttr(ind, "left", "720");
    setWAttr(ind, "firstLine", "0");
  }
  setWAttr(ind, "right", "0");
  // Insert ind immediately after numPr so it takes precedence
  const numPrInPPr = getChild(pPrEl, "numPr");
  if (numPrInPPr?.nextSibling) {
    pPrEl.insertBefore(ind, numPrInPPr.nextSibling);
  } else {
    pPrEl.appendChild(ind);
  }

  writeRuns(p, "Garamond", 24, false, false);
  writePPrRPr(p, "Garamond", 24, false, false);
}

function applyFigureCaption(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "center");
  if (rules.indentation) writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 480);
  removePBdr(p);
  ensureTrailingPeriod(p);
  writeRuns(p, "Garamond", 24, false, false);
  writePPrRPr(p, "Garamond", 24, false, false);
}

function applyTableCaption(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "left");
  if (rules.indentation) writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 240);
  removePBdr(p);
  ensureTrailingPeriod(p);
  writeRuns(p, "Garamond", 24, false, false);
  writePPrRPr(p, "Garamond", 24, false, false);
}

function applyLegend(p: Element, rules: Rules) {
  stripAll(p);
  stripLeadingTabRuns(p);
  stripLeadingArrowRuns(p);
  stripLeadingTextWhitespace(p);
  stripTrailingTextWhitespace(p);
  if (rules.alignment) writePAlignment(p, "left");
  if (rules.indentation) writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 240);
  writeRuns(p, "Garamond", 22, false, false);
  writePPrRPr(p, "Garamond", 22, false, false);
}

function applyContinuationLabel(p: Element, rules: Rules, tableNumber: string) {
  stripAll(p);
  if (rules.alignment) writePAlignment(p, "left");
  if (rules.indentation) writePIndent(p, 0);
  if (rules.spacing) writePSpacing(p, 0, 0, 240);

  Array.from(p.querySelectorAll("br")).forEach((br) => {
    if (br.namespaceURI === W_NS && wAttr(br, "type") === "page") {
      const run = br.parentElement;
      if (run) run.parentElement?.removeChild(run);
    }
  });

  if (tableNumber !== "") {
    const canonical = `Continuation of Table ${tableNumber}...`;
    Array.from(p.querySelectorAll("r")).forEach((r) => {
      if (r.namespaceURI === W_NS) r.parentElement?.removeChild(r);
    });
    const run = wElem(p.ownerDocument!, "r");
    const tEl = wElem(p.ownerDocument!, "t");
    tEl.textContent = canonical;
    run.appendChild(tEl);
    p.appendChild(run);
  }

  writeRuns(p, "Garamond", 26, false, true);
  writePPrRPr(p, "Garamond", 26, false, true);
}

function applyFigureParagraph(p: Element, rules?: Rules) {
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
  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", "0");
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);
  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", "240");
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  const rPr = wElem(p.ownerDocument!, "rPr");
  const rFonts = wElem(p.ownerDocument!, "rFonts");
  for (const a of ["ascii", "hAnsi", "eastAsia", "cs"])
    setWAttr(rFonts, a, "Garamond");
  rPr.appendChild(rFonts);
  const szEl = wElem(p.ownerDocument!, "sz");
  setWAttr(szEl, "val", "24");
  rPr.appendChild(szEl);
  const szCs = wElem(p.ownerDocument!, "szCs");
  setWAttr(szCs, "val", "24");
  rPr.appendChild(szCs);
  pPr.appendChild(rPr);

  const allDrawings: Element[] = [];
  const walkForDrawings = (node: Element) => {
    for (const c of Array.from(node.childNodes)) {
      if (!(c instanceof Element)) continue;
      if (c.localName === "drawing" && c.namespaceURI === W_NS)
        allDrawings.push(c);
      else walkForDrawings(c);
    }
  };
  walkForDrawings(p);

  allDrawings.forEach((drawing) => {
    const existingInline = Array.from(drawing.childNodes).find(
      (c): c is Element =>
        c instanceof Element &&
        c.namespaceURI === WP_NS &&
        c.localName === "inline",
    );

    if (existingInline) {
      const existingEff = Array.from(existingInline.childNodes).find(
        (c): c is Element =>
          c instanceof Element && c.localName === "effectExtent",
      );
      if (
        existingEff &&
        existingEff.getAttribute("t") === "38100" &&
        existingEff.getAttribute("b") === "38100"
      )
        return;

      if (existingEff) existingInline.removeChild(existingEff);
      const eff = drawing.ownerDocument!.createElementNS(
        WP_NS,
        "wp:effectExtent",
      );
      eff.setAttribute("l", "38100");
      eff.setAttribute("t", "38100");
      eff.setAttribute("r", "38100");
      eff.setAttribute("b", "38100");
      const extent = Array.from(existingInline.childNodes).find(
        (c): c is Element => c instanceof Element && c.localName === "extent",
      );
      if (extent && extent.nextSibling)
        existingInline.insertBefore(eff, extent.nextSibling);
      else existingInline.appendChild(eff);
      return;
    }

    const anchor = Array.from(drawing.childNodes).find(
      (c): c is Element =>
        c instanceof Element &&
        c.namespaceURI === WP_NS &&
        c.localName === "anchor",
    );
    if (!anchor) return;

    // Strip any text-wrap element so the figure never floats beside text
    Array.from(anchor.childNodes)
      .filter(
        (c): c is Element =>
          c instanceof Element &&
          [
            "wrapSquare",
            "wrapTight",
            "wrapThrough",
            "wrapTopAndBottom",
            "wrapBehindText",
            "wrapInFrontOfText",
          ].includes(c.localName),
      )
      .forEach((c) => anchor.removeChild(c));

    const inline = drawing.ownerDocument!.createElementNS(WP_NS, "wp:inline");
    inline.setAttribute("distT", "0");
    inline.setAttribute("distB", "0");
    inline.setAttribute("distL", "0");
    inline.setAttribute("distR", "0");

    const extent = Array.from(anchor.childNodes).find(
      (c): c is Element => c instanceof Element && c.localName === "extent",
    );
    if (extent) {
      const ne = drawing.ownerDocument!.createElementNS(WP_NS, "wp:extent");
      ne.setAttribute("cx", extent.getAttribute("cx") ?? "0");
      ne.setAttribute("cy", extent.getAttribute("cy") ?? "0");
      inline.appendChild(ne);
    }
    const eff = drawing.ownerDocument!.createElementNS(
      WP_NS,
      "wp:effectExtent",
    );
    eff.setAttribute("l", "38100");
    eff.setAttribute("t", "38100");
    eff.setAttribute("r", "38100");
    eff.setAttribute("b", "38100");
    inline.appendChild(eff);

    Array.from(anchor.childNodes).forEach((c) => {
      if (
        c instanceof Element &&
        ["docPr", "cNvGraphicFramePr", "graphic"].includes(c.localName)
      )
        inline.appendChild(c.cloneNode(true));
    });
    drawing.replaceChild(inline, anchor);
  });

  if (!rules || rules.borders) {
    const makeLn = (doc: Document): Element => {
      const ln = doc.createElementNS(A_NS, "a:ln");
      ln.setAttribute("w", "38100");
      const solidFill = doc.createElementNS(A_NS, "a:solidFill");
      const srgbClr = doc.createElementNS(A_NS, "a:srgbClr");
      srgbClr.setAttribute("val", "000000");
      solidFill.appendChild(srgbClr);
      ln.appendChild(solidFill);
      return ln;
    };

    const applyLnToSpPr = (spPr: Element) => {
      Array.from(spPr.childNodes)
        .filter(
          (c): c is Element => c instanceof Element && c.localName === "ln",
        )
        .forEach((c) => spPr.removeChild(c));
      const ln = makeLn(spPr.ownerDocument!);
      const insertBefore =
        Array.from(spPr.childNodes).find(
          (c): c is Element =>
            c instanceof Element &&
            [
              "solidFill",
              "gradFill",
              "noFill",
              "pattFill",
              "grpFill",
              "blipFill",
              "effectLst",
              "scene3d",
              "sp3d",
              "extLst",
            ].includes(c.localName),
        ) ?? null;
      if (insertBefore) spPr.insertBefore(ln, insertBefore);
      else spPr.appendChild(ln);
    };

    const spPrList: Element[] = [];
    const walkForSpPr = (node: Element) => {
      for (const c of Array.from(node.childNodes)) {
        if (!(c instanceof Element)) continue;
        if (c.localName === "spPr") spPrList.push(c);
        walkForSpPr(c);
      }
    };
    walkForSpPr(p);

    if (spPrList.length > 0) {
      spPrList.forEach(applyLnToSpPr);
    } else {
      const pPrEl = ensurePPr(p);
      removeChildren(pPrEl, "pBdr");
      const pBdr = wElem(p.ownerDocument!, "pBdr");
      for (const side of ["top", "left", "bottom", "right"]) {
        const edge = wElem(p.ownerDocument!, side);
        setWAttr(edge, "val", "single");
        setWAttr(edge, "sz", "24");
        setWAttr(edge, "space", "4");
        setWAttr(edge, "color", "000000");
        pBdr.appendChild(edge);
      }
      pPrEl.appendChild(pBdr);
    }
  }
}

// ─── table border spec ───────────────────────────────────────────────────────

function deFloatTable(tbl: Element) {
  const tblPr = getChild(tbl, "tblPr");
  if (!tblPr) return;
  removeChildren(tblPr, "tblpPr");
  removeChildren(tblPr, "positionH");
  removeChildren(tblPr, "positionV");
  const tblInd = getChild(tblPr, "tblInd");
  if (tblInd) {
    setWAttr(tblInd, "w", "0");
    setWAttr(tblInd, "type", "dxa");
  }
}

function applyTableBordersSpec(tbl: Element) {
  deFloatTable(tbl);

  Array.from(tbl.querySelectorAll("tc")).forEach((tc) => {
    if (tc.namespaceURI !== W_NS) return;
    const cellParas = Array.from(tc.querySelectorAll("p")).filter(
      (p) => p.namespaceURI === W_NS,
    );
    const nonEmpty = cellParas.filter((p) => {
      const t = Array.from(p.querySelectorAll("t"))
        .map((e) => (e.textContent ?? "").trim())
        .join("");
      const hasDrawing = p.querySelectorAll("drawing, pict").length > 0;
      return t !== "" || hasDrawing;
    });
    // Always remove empty paragraphs from cells — including trailing blank
    // lines at the bottom — regardless of whether the cell has content.
    // Keep at least one paragraph if the cell would otherwise be left empty
    // (Word requires at least one paragraph per cell).
    cellParas.forEach((cp) => {
      // Never remove the very last paragraph in a cell — Word requires it
      if (cp === cellParas[cellParas.length - 1] && nonEmpty.length === 0)
        return;
      const t = Array.from(cp.querySelectorAll("t"))
        .map((e) => (e.textContent ?? "").trim())
        .join("");
      const hasDrawing = cp.querySelectorAll("drawing, pict").length > 0;
      if (t === "" && !hasDrawing) cp.parentElement?.removeChild(cp);
    });
  });

  const allRows = Array.from(tbl.querySelectorAll("tr")).filter(
    (r) => r.namespaceURI === W_NS,
  );
  const emptyRows = allRows.filter((r) => {
    const t = Array.from(r.querySelectorAll("t"))
      .map((e) => (e.textContent ?? "").trim())
      .join("");
    const hasDrawing = r.querySelectorAll("drawing, pict").length > 0;
    return t === "" && !hasDrawing;
  });
  if (emptyRows.length < allRows.length)
    emptyRows.forEach((r) => r.parentElement?.removeChild(r));

  Array.from(tbl.querySelectorAll("trHeight")).forEach((trH) => {
    if (trH.namespaceURI === W_NS) trH.parentElement?.removeChild(trH);
  });

  const tblPrW = ensureChild(tbl, "tblPr");
  removeChildren(tblPrW, "tblW");
  const tblW = wElem(tbl.ownerDocument!, "tblW");
  setWAttr(tblW, "type", "pct");
  setWAttr(tblW, "w", "5000");
  tblPrW.appendChild(tblW);
  Array.from(tbl.querySelectorAll("tcW")).forEach((tcW) => {
    if (tcW.namespaceURI === W_NS) tcW.parentElement?.removeChild(tcW);
  });

  const tblPr = ensureChild(tbl, "tblPr");
  removeChildren(tblPr, "tblBorders");
  const tblBorders = wElem(tbl.ownerDocument!, "tblBorders");
  tblPr.appendChild(tblBorders);

  const makeBorder = (local: string, val: string, sz: string) => {
    const el = wElem(tbl.ownerDocument!, local);
    setWAttr(el, "val", val);
    setWAttr(el, "sz", sz);
    setWAttr(el, "space", "0");
    setWAttr(el, "color", "000000");
    tblBorders.appendChild(el);
  };
  makeBorder("top", "double", "4");
  makeBorder("left", "none", "0");
  makeBorder("bottom", "double", "4");
  makeBorder("right", "none", "0");
  makeBorder("insideH", "none", "0");
  makeBorder("insideV", "none", "0");

  const rows = Array.from(tbl.querySelectorAll("tr")).filter(
    (r) => r.namespaceURI === W_NS,
  );
  if (rows.length === 0) return;
  const firstRow = rows[0];
  const lastRow = rows[rows.length - 1];

  rows.forEach((row) => {
    if (row === lastRow) return;
    Array.from(row.querySelectorAll("tc")).forEach((tc) => {
      if (tc.namespaceURI !== W_NS) return;
      const tcPr = getChild(tc, "tcPr");
      if (tcPr) removeChildren(tcPr, "tcBorders");
    });
  });

  Array.from(firstRow.querySelectorAll("tc")).forEach((tc) => {
    if (tc.namespaceURI !== W_NS) return;
    let tcPr = getChild(tc, "tcPr");
    if (!tcPr) {
      tcPr = wElem(tc.ownerDocument!, "tcPr");
      tc.insertBefore(tcPr, tc.firstChild);
    }
    const tcBdr = wElem(tc.ownerDocument!, "tcBorders");
    tcPr.appendChild(tcBdr);
    const el = wElem(tc.ownerDocument!, "bottom");
    setWAttr(el, "val", "single");
    setWAttr(el, "sz", "4");
    setWAttr(el, "space", "0");
    setWAttr(el, "color", "000000");
    tcBdr.appendChild(el);
  });

  if (rows.length > 1) {
    Array.from(lastRow.querySelectorAll("tc")).forEach((tc) => {
      if (tc.namespaceURI !== W_NS) return;
      let tcPr = getChild(tc, "tcPr");
      if (!tcPr) {
        tcPr = wElem(tc.ownerDocument!, "tcPr");
        tc.insertBefore(tcPr, tc.firstChild);
      }

      let existingTopVal = "",
        existingTopSz = "",
        existingTopClr = "",
        existingTopSpc = "";
      const oldTcBdr = getChild(tcPr, "tcBorders");
      if (oldTcBdr) {
        Array.from(oldTcBdr.childNodes).forEach((b) => {
          if (
            b instanceof Element &&
            b.namespaceURI === W_NS &&
            b.localName === "top"
          ) {
            existingTopVal = wAttr(b, "val");
            existingTopSz = wAttr(b, "sz");
            existingTopClr = wAttr(b, "color");
            existingTopSpc = wAttr(b, "space");
          }
        });
      }
      removeChildren(tcPr, "tcBorders");

      const tcBdr = wElem(tc.ownerDocument!, "tcBorders");
      tcPr.appendChild(tcBdr);

      const elTop = wElem(tc.ownerDocument!, "top");
      if (existingTopVal && existingTopVal !== "none") {
        setWAttr(elTop, "val", existingTopVal);
        setWAttr(elTop, "sz", existingTopSz || "4");
        setWAttr(elTop, "space", existingTopSpc || "0");
        setWAttr(elTop, "color", existingTopClr || "000000");
      } else {
        setWAttr(elTop, "val", "none");
        setWAttr(elTop, "sz", "0");
        setWAttr(elTop, "space", "0");
        setWAttr(elTop, "color", "auto");
      }
      tcBdr.appendChild(elTop);

      const elBot = wElem(tc.ownerDocument!, "bottom");
      setWAttr(elBot, "val", "double");
      setWAttr(elBot, "sz", "4");
      setWAttr(elBot, "space", "0");
      setWAttr(elBot, "color", "000000");
      tcBdr.appendChild(elBot);
    });
  }
}

// ─── classification helpers ──────────────────────────────────────────────────

function isHeadingStyleId(styleId: string): boolean {
  if (!styleId) return false;
  if (/^Heading[2-9]$/i.test(styleId)) return true;
  if (/^[2-9]$/.test(styleId)) return true;
  if (/^heading[2-9]$/i.test(styleId)) return true;
  return false;
}

function isHeading(p: Element, text: string): boolean {
  if (!text) return false;
  if (/[.!?]$/u.test(text)) return false;
  if (
    /^(figure|fig\.?|table|chapter|appendix|continuation|legend)\b/iu.test(text)
  )
    return false;
  if (/^\[\d+\]/u.test(text)) return false;
  if (text.length > 150) return false;

  if (/^\d+\.\d+(\.\d+)*(\s+\S.*)?$/u.test(text)) return true;

  if (getPFirstLineIndent(p) > 0) return false;

  const allRuns = Array.from(p.querySelectorAll("r")).filter(
    (r) => r.namespaceURI === W_NS,
  );
  const textRuns = allRuns.filter((r) => {
    return !Array.from(r.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
    );
  });
  if (textRuns.length === 0) return false;

  let textRunsTotal = 0,
    textRunsBold = 0,
    boldItalicRuns = 0,
    sz13Runs = 0;
  for (const run of textRuns) {
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    if (!tEl || !(tEl.textContent ?? "").trim()) continue;
    textRunsTotal++;
    const rPr = getChild(run, "rPr");
    const hasBold = rPr
      ? rPr.getElementsByTagNameNS(W_NS, "b").length > 0
      : false;
    const hasItalic = rPr
      ? rPr.getElementsByTagNameNS(W_NS, "i").length > 0
      : false;
    const szEl = rPr
      ? (rPr.getElementsByTagNameNS(W_NS, "sz").item(0) as Element | null)
      : null;
    if (hasBold && hasItalic) {
      boldItalicRuns++;
      textRunsBold++;
    } else if (hasBold) textRunsBold++;
    if (szEl && wAttr(szEl, "val") === "26") sz13Runs++;
  }
  if (textRunsTotal === 0) return false;
  if (textRunsBold / textRunsTotal >= 0.6 && boldItalicRuns === 0) return true;
  if (textRunsTotal > 0 && sz13Runs / textRunsTotal >= 0.6) return true;
  return false;
}

function isItalicHeading(p: Element, text: string): boolean {
  if (!text) return false;

  if (
    /^\d+\.\s+\S/u.test(text) &&
    text.length <= 300 &&
    !/^\d+\.\s+\S+\s*[—–-]/u.test(text) &&
    !/^(figure|fig\.?|table|chapter|appendix|continuation|legend)\b/iu.test(
      text,
    ) &&
    !/^\[\d+\]/u.test(text)
  ) {
    const runs = Array.from(p.querySelectorAll("r")).filter(
      (r) => r.namespaceURI === W_NS,
    );
    let boldTotal = 0,
      sz12Total = 0,
      runTotal = 0;
    for (const run of runs) {
      if (
        Array.from(run.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        )
      )
        continue;
      const tEl = run
        .getElementsByTagNameNS(W_NS, "t")
        .item(0) as Element | null;
      if (!tEl || !(tEl.textContent ?? "").trim()) continue;
      runTotal++;
      const rPr = getChild(run, "rPr");
      const bold = rPr
        ? rPr.getElementsByTagNameNS(W_NS, "b").length > 0
        : false;
      const szEl = rPr
        ? (rPr.getElementsByTagNameNS(W_NS, "sz").item(0) as Element | null)
        : null;
      if (bold) boldTotal++;
      if (szEl && wAttr(szEl, "val") === "24") sz12Total++;
    }
    if (runTotal > 0 && boldTotal / runTotal >= 0.6) return true;
    if (runTotal > 0 && sz12Total / runTotal >= 0.6) return true;
    const pPrSz = p.querySelector("pPr > rPr > sz") as Element | null;
    if (pPrSz && wAttr(pPrSz, "val") === "24") return true;
  }

  if (/[.!?]$/u.test(text)) return false;
  if (
    /^(figure|fig\.?|table|chapter|appendix|continuation|legend)\b/iu.test(text)
  )
    return false;
  if (/^\[\d+\]/u.test(text)) return false;
  if (text.length > 150) return false;

  const runs = Array.from(p.querySelectorAll("r")).filter(
    (r) => r.namespaceURI === W_NS,
  );
  if (runs.length === 0) return false;

  let textRunsTotal = 0,
    boldItalicRuns = 0,
    sz12Runs = 0;
  for (const run of runs) {
    if (
      Array.from(run.children).some((c) =>
        ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
      )
    )
      continue;
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    if (!tEl || !(tEl.textContent ?? "").trim()) continue;
    textRunsTotal++;
    const rPr = getChild(run, "rPr");
    const hasBold = rPr
      ? rPr.getElementsByTagNameNS(W_NS, "b").length > 0
      : false;
    const hasItalic = rPr
      ? rPr.getElementsByTagNameNS(W_NS, "i").length > 0
      : false;
    const szEl = rPr
      ? (rPr.getElementsByTagNameNS(W_NS, "sz").item(0) as Element | null)
      : null;
    if (hasBold && hasItalic) boldItalicRuns++;
    if (szEl && wAttr(szEl, "val") === "24") sz12Runs++;
  }
  if (textRunsTotal === 0) return false;
  if (boldItalicRuns / textRunsTotal >= 0.6) return true;
  if (boldItalicRuns > 0 && sz12Runs / textRunsTotal >= 0.6) return true;
  const pPrSz = p.querySelector("pPr > rPr > sz") as Element | null;
  if (pPrSz && wAttr(pPrSz, "val") === "24" && boldItalicRuns > 0) return true;
  return false;
}

// ─── table processor ─────────────────────────────────────────────────────────

function processTable(tbl: Element, rules: Rules, state: State) {
  const zone = state.zone;
  if (zone !== "chapters" && zone !== "references") return;

  Array.from(tbl.querySelectorAll("tc")).forEach((tc) => {
    if (tc.namespaceURI !== W_NS) return;
    const cellParas = Array.from(tc.querySelectorAll("p")).filter(
      (p) => p.namespaceURI === W_NS,
    );
    const nonEmpty = cellParas.filter((p) => {
      const t = Array.from(p.querySelectorAll("t"))
        .map((e) => (e.textContent ?? "").trim())
        .join("");
      return t !== "" || p.querySelectorAll("drawing, pict").length > 0;
    });
    cellParas.forEach((cp) => {
      if (cp === cellParas[cellParas.length - 1] && nonEmpty.length === 0)
        return;
      const t = Array.from(cp.querySelectorAll("t"))
        .map((e) => (e.textContent ?? "").trim())
        .join("");
      const hasDrawing = cp.querySelectorAll("drawing, pict").length > 0;
      if (t === "" && !hasDrawing) cp.parentElement?.removeChild(cp);
    });
  });

  const allRows = Array.from(tbl.querySelectorAll("tr")).filter(
    (r) => r.namespaceURI === W_NS,
  );
  const emptyRows = allRows.filter((r) => {
    const t = Array.from(r.querySelectorAll("t"))
      .map((e) => (e.textContent ?? "").trim())
      .join("");
    return t === "" && r.querySelectorAll("drawing, pict").length === 0;
  });
  if (emptyRows.length < allRows.length)
    emptyRows.forEach((r) => r.parentElement?.removeChild(r));

  deFloatTable(tbl);

  Array.from(tbl.querySelectorAll("trHeight")).forEach((trH) => {
    if (trH.namespaceURI === W_NS) trH.parentElement?.removeChild(trH);
  });

  const tblPrW = ensureChild(tbl, "tblPr");
  removeChildren(tblPrW, "tblW");
  const tblW = wElem(tbl.ownerDocument!, "tblW");
  setWAttr(tblW, "type", "pct");
  setWAttr(tblW, "w", "5000");
  tblPrW.appendChild(tblW);
  Array.from(tbl.querySelectorAll("tcW")).forEach((tcW) => {
    if (tcW.namespaceURI === W_NS) tcW.parentElement?.removeChild(tcW);
  });

  if (rules.borders) applyTableBordersSpec(tbl);

  Array.from(tbl.querySelectorAll("tc")).forEach((tc) => {
    if (tc.namespaceURI !== W_NS) return;
    Array.from(tc.querySelectorAll("p")).forEach((p) => {
      if (p.namespaceURI !== W_NS) return;
      const isNumbered = p.querySelectorAll("pPr > numPr").length > 0;
      if (rules.alignment) writePAlignment(p, isNumbered ? "both" : "center");
      if (rules.spacing) writePSpacing(p, 0, 0, 240);
      if (rules.indentation) writePIndent(p, 0);
      writeRuns(p, "Arial", 20, false, false);
      writePPrRPr(p, "Arial", 20, false, false);
    });
  });
}

// ─── paragraph router ────────────────────────────────────────────────────────

interface State {
  zone: string;
  currentChapter: number;
  expectChapterTitle: boolean;
  isFirstParagraph: boolean;
  afterTable: boolean;
  lastTableNumber: string;
}

function processParagraph(p: Element, rules: Rules, state: State) {
  const text = getParagraphText(p);
  const normalized = normalizeText(text);
  const styleId = getParagraphStyleId(p);

  const chapterMatch = normalized.match(/^chapter\s+([ivxlcdm]+|\d+)$/iu);
  const isChapterLabel = chapterMatch !== null;
  const isReferences = /^references$/iu.test(normalized);
  const isAppendixLabel = /^appendi(?:x|ces|xes)$/iu.test(normalized);

  if (isChapterLabel) {
    state.zone = "chapters";
    state.currentChapter = chapterToInt(chapterMatch![1]);
    state.expectChapterTitle = true;
    state.isFirstParagraph = true;
  } else if (isReferences) {
    state.zone = "references";
    state.isFirstParagraph = true;
  } else if (isAppendixLabel) {
    state.zone = "appendices";
  }

  const zone = state.zone;
  if (zone !== "chapters" && zone !== "references") return;

  const hasDrawing = p.querySelectorAll("drawing, pict").length > 0;

  if (hasDrawing && normalized === "" && !isInTable(p)) {
    state.afterTable = false;
    state.isFirstParagraph = false;
    applyFigureParagraph(p, rules);
    return;
  }

  if (normalized === "") {
    if (state.zone === "references") applyReferenceEmptyLine(p);
    else applyEmptyParagraph(p);
    return;
  }

  const hasNumbering = p.querySelectorAll("pPr > numPr").length > 0;
  const inTable = isInTable(p);

  const isFigureCaption =
    !state.isFirstParagraph &&
    /^figure\s+\d+[\-\.]\d+\b/iu.test(normalized) &&
    normalized.length <= 120 &&
    !/^figure\s+[\d\-\.]+\s+\w+\s+(presents|shows|describes|summarizes|lists|displays|illustrates|contains|provides|compares|indicates|reveals|demonstrates)\b/iu.test(
      normalized,
    );

  const isTableCaption =
    !state.isFirstParagraph &&
    /^table\s+\d+[\-\.]\d+\b/iu.test(normalized) &&
    normalized.length <= 120 &&
    !/^table\s+[\d\-\.]+\s+\w+\s+(presents|shows|describes|summarizes|lists|displays|illustrates|contains|provides|compares|indicates|reveals|demonstrates)\b/iu.test(
      normalized,
    );

  const isContinuation =
    /^continuation\s+of\s+(table|figure)(\s+[\d\-\.]+)?/iu.test(normalized);
  const isLegend = /^legend\s*:/iu.test(normalized);

  const applyCorrectHeading = () => {
    if (state.currentChapter === 2 && !/^synthesis\b/iu.test(normalized))
      applyItalicHeading(p, rules);
    else applyHeading(p, rules);
  };

  if (styleId === "Heading1") {
    if (isChapterLabel) {
      state.afterTable = false;
      applyChapterLabel(p, rules);
      return;
    }
    if (
      state.expectChapterTitle &&
      !hasDrawing &&
      !isFigureCaption &&
      !isTableCaption &&
      !hasNumbering &&
      normalized.length <= 200 &&
      !/[.?!]$/.test(normalized)
    ) {
      state.expectChapterTitle = false;
      state.isFirstParagraph = true;
      state.afterTable = false;
      applyChapterTitle(p, rules);
      return;
    }
    if (isReferences) {
      state.afterTable = false;
      if (p.parentElement) {
        p.parentElement.insertBefore(
          buildZeroHeightPageBreakP(p.ownerDocument!),
          p,
        );
      }
      applyReferencesTitle(p, rules);
      return;
    }
    state.afterTable = false;
    applyCorrectHeading();
    return;
  }

  if (isChapterLabel) {
    state.afterTable = false;
    applyChapterLabel(p, rules);
    return;
  }
  if (isReferences) {
    state.afterTable = false;
    if (p.parentElement) {
      p.parentElement.insertBefore(
        buildZeroHeightPageBreakP(p.ownerDocument!),
        p,
      );
    }
    applyReferencesTitle(p, rules);
    return;
  }
  if (
    state.expectChapterTitle &&
    !hasDrawing &&
    !isFigureCaption &&
    !isTableCaption &&
    !hasNumbering &&
    normalized.length <= 200 &&
    !/[.?!]$/.test(normalized)
  ) {
    state.expectChapterTitle = false;
    state.isFirstParagraph = true;
    state.afterTable = false;
    applyChapterTitle(p, rules);
    return;
  }

  state.isFirstParagraph = false;

  if (hasDrawing) {
    state.afterTable = false;
    // Move any text runs that are sitting beside the drawing into a new paragraph below
    if (p.parentElement) {
      const textRuns = Array.from(p.childNodes).filter(
        (c): c is Element =>
          c instanceof Element &&
          c.localName === "r" &&
          c.namespaceURI === W_NS &&
          !Array.from(c.children).some((ch) =>
            ["drawing", "pict", "instrText", "fldChar"].includes(ch.localName),
          ) &&
          (
            c.getElementsByTagNameNS(W_NS, "t").item(0)?.textContent ?? ""
          ).trim() !== "",
      );
      if (textRuns.length > 0) {
        const newP = p.ownerDocument!.createElementNS(W_NS, "w:p");
        textRuns.forEach((r) => {
          p.removeChild(r);
          newP.appendChild(r);
        });
        p.parentElement.insertBefore(newP, p.nextSibling);
        applyFigureCaption(newP, rules);
      }
    }
    applyFigureParagraph(p, rules);
    return;
  }
  if (isFigureCaption) {
    applyFigureCaption(p, rules);
    return;
  }
  if (isTableCaption) {
    const m = normalized.match(/^table\s+([\d]+[\-\.][\d]+)\b/iu);
    if (m) state.lastTableNumber = m[1];
    applyTableCaption(p, rules);
    return;
  }
  if (isContinuation) {
    applyContinuationLabel(p, rules, state.lastTableNumber);
    return;
  }
  if (isLegend) {
    applyLegend(p, rules);
    return;
  }
  if (inTable) return;

  if (state.zone === "references") {
    state.afterTable = false;
    applyReferenceEntry(p, rules);
    return;
  }

  const ch = state.currentChapter;

  // In chapter 1, \d+.\d+ patterns are sub-items of research objectives (not headings)
  if (ch === 1 && /^\d+\.\d+(\.\d+)*(\s+\S.*)?$/u.test(normalized)) {
    const runs = Array.from(p.childNodes).filter(
      (c): c is Element =>
        c instanceof Element && c.localName === "r" && c.namespaceURI === W_NS,
    );
    for (const run of runs) {
      const tEl = run
        .getElementsByTagNameNS(W_NS, "t")
        .item(0) as Element | null;
      if (!tEl) continue;
      const txt = tEl.textContent ?? "";
      const fixed = txt
        .replace(/^(\d+\.\d+(\.\d+)*)(\s)/, "$1.$3")
        .replace(/^(\d+\.\d+(\.\d+)*)$/, "$1.");
      if (fixed !== txt) tEl.textContent = fixed;
      break;
    }
    stripAll(p);
    stripLeadingTabRuns(p);
    stripLeadingArrowRuns(p);
    stripLeadingTextWhitespace(p);
    stripTrailingTextWhitespace(p);
    if (rules.alignment) writePAlignment(p, "both");
    writePHangingIndent(p, 1440, 720);
    writePSpacing(p, 0, 0, 480);
    writeRuns(p, "Garamond", 24, false, false);
    writePPrRPr(p, "Garamond", 24, false, false);
    state.afterTable = false;
    return;
  }

  // Numbered section headings like 4.1, 4.2 — only in chapters 2–5, not chapter 1
  if (ch >= 2 && /^\d+\.\d+(\.\d+)*(\s+\S.*)?$/u.test(normalized)) {
    state.afterTable = false;
    applyCorrectHeading();
    return;
  }

  if (
    [4, 5].includes(ch) &&
    /^\d+\.\s+\S/u.test(normalized) &&
    !/^\d+\.\s+\S+\s*[—–-]/u.test(normalized)
  ) {
    const runs = Array.from(p.querySelectorAll("r")).filter(
      (r) => r.namespaceURI === W_NS,
    );
    let boldSz26 = false;
    for (const run of runs) {
      if (
        Array.from(run.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        )
      )
        continue;
      const tEl = run
        .getElementsByTagNameNS(W_NS, "t")
        .item(0) as Element | null;
      if (!tEl || !(tEl.textContent ?? "").trim()) continue;
      const rPr = getChild(run, "rPr");
      const bold = rPr
        ? rPr.getElementsByTagNameNS(W_NS, "b").length > 0
        : false;
      const szEl = rPr
        ? (rPr.getElementsByTagNameNS(W_NS, "sz").item(0) as Element | null)
        : null;
      if (bold && szEl && wAttr(szEl, "val") === "26") {
        boldSz26 = true;
        break;
      }
    }
    if (boldSz26) {
      state.afterTable = false;
      applyCorrectHeading();
      return;
    }
  }

  if (hasNumbering || styleId === "ListParagraph") {
    if (ch === 1) return;
    // In chapters 4 and 5, if ANY text run is set to Garamond sz=26 (13 pt),
    // the user manually applied heading formatting — treat it as a heading
    // while keeping the original numbering (numPr) intact.
    if ([4, 5].includes(ch)) {
      const textRuns = Array.from(p.querySelectorAll("r"))
        .filter((r) => r.namespaceURI === W_NS)
        .filter(
          (r) =>
            !Array.from(r.children).some((c) =>
              ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
            ),
        );
      const hasGaramond13 = textRuns.some((run) => {
        const rPr = getChild(run, "rPr");
        if (!rPr) return false;
        const rFonts = getChild(rPr, "rFonts");
        const font = rFonts ? wAttr(rFonts, "ascii") : "";
        const szEl = rPr
          .getElementsByTagNameNS(W_NS, "sz")
          .item(0) as Element | null;
        return (
          font === "Garamond" && szEl != null && wAttr(szEl, "val") === "26"
        );
      });
      if (hasGaramond13) {
        // Save numPr before applyCorrectHeading → stripAll removes it
        const pPrEl = getChild(p, "pPr");
        const savedNumPr = pPrEl
          ? (getChild(pPrEl, "numPr")?.cloneNode(true) as Element | undefined)
          : undefined;
        state.afterTable = false;
        applyCorrectHeading();
        // Restore numPr so the user's numbering is preserved
        if (savedNumPr) {
          const pPrAfter = ensurePPr(p);
          removeChildren(pPrAfter, "numPr");
          pPrAfter.insertBefore(savedNumPr, pPrAfter.firstChild);
        }
        return;
      }
    }
    // In chapters 4 and 5, "N. Text." followed by a body paragraph is a heading.
    if (
      [4, 5].includes(ch) &&
      /^\d+\.\s+\S/u.test(normalized) &&
      /\.$/.test(normalized)
    ) {
      // Peek at the next non-empty sibling — if it's a plain body paragraph
      // (no numbering), this is a heading, not a list item.
      let next = p.nextSibling as Element | null;
      while (
        next instanceof Element &&
        next.localName === "p" &&
        normalizeText(getParagraphText(next)) === ""
      ) {
        next = next.nextSibling as Element | null;
      }
      const nextHasNumbering =
        next instanceof Element &&
        next.localName === "p" &&
        next.querySelectorAll("pPr > numPr").length > 0;
      if (!nextHasNumbering) {
        state.afterTable = false;
        applyCorrectHeading();
        return;
      }
    }
    state.afterTable = false;
    applyListParagraph(p, rules);
    return;
  }
  if (isHeadingStyleId(styleId)) {
    state.afterTable = false;
    applyCorrectHeading();
    return;
  }

  const looksLikeItalic =
    [2, 4, 5].includes(ch) && isItalicHeading(p, normalized);
  const looksLikeHeading = looksLikeItalic || isHeading(p, normalized);

  if (looksLikeHeading) {
    state.afterTable = false;
    applyCorrectHeading();
    return;
  }

  const beforeSpacing = state.afterTable ? 240 : 0;
  state.afterTable = false;
  applyBodyParagraph(p, rules, beforeSpacing);
}

// ─── hoist table captions pre-pass ──────────────────────────────────────────

function hoistTableCaptions(body: Element, children: Element[]) {
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (child.localName !== "tbl") continue;

    let captionNode: Element | null = null;
    const emptysBetween: Element[] = [];

    for (let j = i + 1; j < children.length; j++) {
      const sib = children[j];
      if (sib.localName !== "p") break;
      const text = normalizeText(
        Array.from(sib.querySelectorAll("t"))
          .map((t) => t.textContent ?? "")
          .join(""),
      );
      if (text === "") {
        emptysBetween.push(sib);
        continue;
      }
      if (
        /^table\s+\d+[\-\.]\d+\b/iu.test(text) &&
        text.length <= 120 &&
        !/^table\s+[\d\-\.]+\s+\w+\s+(presents|shows|describes|summarizes|lists|displays|illustrates|contains|provides|compares|indicates|reveals|demonstrates)\b/iu.test(
          text,
        )
      )
        captionNode = sib;
      break;
    }
    if (!captionNode) continue;
    body.insertBefore(captionNode, child);
    emptysBetween.forEach((e) => e.parentElement?.removeChild(e));
  }
}

// ─── table continuation pre-pass ─────────────────────────────────────────────

function handleTableContinuation(body: Element, children: Element[]) {
  children.forEach((child) => {
    if (child.localName !== "tbl") return;
    Array.from(child.querySelectorAll("tblHeader")).forEach((th) => {
      if (th.namespaceURI === W_NS) th.parentElement?.removeChild(th);
    });
  });

  let lastTableNumber = "";

  for (let i = 0; i < children.length; i++) {
    const child = children[i];

    if (child.localName === "p") {
      const text = normalizeText(
        Array.from(child.querySelectorAll("t"))
          .map((t) => t.textContent ?? "")
          .join(""),
      );
      const m = text.match(/^table\s+([\d]+[\-\.][\d]+)\b/iu);
      if (m) lastTableNumber = m[1];
      continue;
    }
    if (child.localName !== "tbl") continue;
    if (!child.parentElement) continue;

    let alreadyHasLabel = false;
    for (let k = i + 1; k < children.length; k++) {
      const sib = children[k];
      if (sib.localName === "tbl") break;
      if (sib.localName !== "p") break;
      const t = normalizeText(
        Array.from(sib.querySelectorAll("t"))
          .map((n) => n.textContent ?? "")
          .join(""),
      );
      if (t === "") continue;
      if (/^continuation\s+of\s+(table|figure)/iu.test(t))
        alreadyHasLabel = true;
      break;
    }
    if (alreadyHasLabel) {
      for (let k = i + 1; k < children.length; k++) {
        const sib = children[k];
        if (sib.localName === "tbl") break;
        if (sib.localName !== "p") break;
        const t = normalizeText(
          Array.from(sib.querySelectorAll("t"))
            .map((n) => n.textContent ?? "")
            .join(""),
        );
        if (t === "") continue;
        if (/^continuation\s+of\s+(table|figure)/iu.test(t)) {
          sib.parentElement?.insertBefore(
            buildZeroHeightPageBreakP(body.ownerDocument!),
            sib,
          );
        }
        break;
      }
      continue;
    }

    const between: Element[] = [];
    let nextTbl: Element | null = null;
    let hasExplicitBreak = false;

    for (let j = i + 1; j < children.length; j++) {
      const sib = children[j];
      if (sib.localName === "tbl") {
        nextTbl = sib;
        i = j - 1;
        break;
      }
      if (sib.localName !== "p") break;
      const t = normalizeText(
        Array.from(sib.querySelectorAll("t"))
          .map((n) => n.textContent ?? "")
          .join(""),
      );
      const hasBreak =
        sib.querySelectorAll("br").length > 0 &&
        Array.from(sib.querySelectorAll("br")).some(
          (b) => b.namespaceURI === W_NS && wAttr(b, "type") === "page",
        );
      if (hasBreak) hasExplicitBreak = true;
      if (t === "" || hasBreak) {
        between.push(sib);
        continue;
      }
      break;
    }

    if (!nextTbl) continue;

    if (!hasExplicitBreak) {
      const rows2 = Array.from(nextTbl.querySelectorAll("tr")).filter(
        (r) => r.namespaceURI === W_NS,
      );
      rows2.forEach((r) => child.appendChild(r.cloneNode(true)));
      between.forEach((g) => g.parentElement?.removeChild(g));
      nextTbl.parentElement?.removeChild(nextTbl);
      applyTableBordersSpec(child);
      continue;
    }

    between.forEach((g) => g.parentElement?.removeChild(g));
    applyTableBordersSpec(child);
    applyTableBordersSpec(nextTbl);

    const label =
      lastTableNumber !== ""
        ? `Continuation of Table ${lastTableNumber}...`
        : "Continuation of Table...";

    body.insertBefore(buildZeroHeightPageBreakP(body.ownerDocument!), nextTbl);
    body.insertBefore(buildContinuationP(body.ownerDocument!, label), nextTbl);
  }
}

// ─── main entry ──────────────────────────────────────────────────────────────

export async function formatDocx(
  arrayBuffer: ArrayBuffer,
  options: FormatOptions,
): Promise<Blob> {
  const JSZip = (window as any).JSZip;
  if (!JSZip) throw new Error("JSZip not loaded");

  const zip = await JSZip.loadAsync(arrayBuffer);
  const xmlStr: string = await zip.file("word/document.xml").async("string");

  const parser = new DOMParser();
  const dom = parser.parseFromString(xmlStr, "application/xml");

  const body = dom.querySelector("body");
  if (!body) throw new Error("No body element found in document.xml");

  const rules: Rules = {};
  for (const r of options.rules) {
    (rules as any)[r] = true;
  }

  // Pre-pass: unwrap all w:sdt content controls
  unwrapContentControls(body);

  const allChildren = Array.from(body.childNodes).filter(
    (c): c is Element => c instanceof Element,
  );

  let chapterStart = -1;
  let appendixEnd = allChildren.length;

  for (let idx = 0; idx < allChildren.length; idx++) {
    const child = allChildren[idx];
    if (child.localName !== "p") continue;
    // Use both txbxContent-aware and raw text to catch all cases
    const t = normalizeText(getParagraphText(child));
    const tRaw = normalizeText(
      Array.from(child.querySelectorAll("t"))
        .map((n) => n.textContent ?? "")
        .join(""),
    );
    const text = t || tRaw;
    if (chapterStart === -1 && /^chapter\s+([ivxlcdm]+|\d+)$/iu.test(text)) {
      chapterStart = idx;
    }
    if (chapterStart !== -1 && /^appendi(?:x|ces|xes)$/iu.test(text)) {
      appendixEnd = idx;
      break;
    }
  }

  if (chapterStart === -1) {
    const serializer = new XMLSerializer();
    const newXml = serializer.serializeToString(dom);
    zip.file("word/document.xml", newXml);
    const blob: Blob = await zip.generateAsync({ type: "blob" });
    return blob;
  }

  const scopedChildren = allChildren.slice(chapterStart, appendixEnd);

  hoistTableCaptions(body, scopedChildren);

  if (rules.continuation) {
    handleTableContinuation(body, scopedChildren);
  }

  const state: State = {
    zone: "preliminary",
    currentChapter: 0,
    expectChapterTitle: false,
    isFirstParagraph: true,
    afterTable: false,
    lastTableNumber: "",
  };

  for (const child of scopedChildren) {
    if (!(child instanceof Element)) continue;
    if (child.parentElement !== body) continue;
    if (child.localName === "p") {
      processParagraph(child, rules, state);
    } else if (child.localName === "tbl") {
      processTable(child, rules, state);
      state.afterTable = true;
    }
  }

  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(dom);
  zip.file("word/document.xml", newXml);
  const blob: Blob = await zip.generateAsync({ type: "blob" });
  return blob;
}
