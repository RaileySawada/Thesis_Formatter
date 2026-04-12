/**
 * PreliminaryEngine.ts
 * Formats preliminary pages of a thesis manuscript.
 * Sections handled:
 *   Title Page, Approval Sheet, Abstract, Acknowledgment,
 *   Dedication, Table of Contents, List of Figures,
 *   List of Tables, List of Appendices
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS =
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
import type { FormattingConfig } from "../constants";

// ─── low-level helpers ───────────────────────────────────────────────────────

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

function ptsToHalfPts(pts: number): number {
  return Math.round(pts * 2);
}
function linesToTwips(lines: number): number {
  return Math.round(lines * 240);
}
function inchesToTwips(inches: number): number {
  return Math.round(inches * 1440);
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

// ─── property writers ────────────────────────────────────────────────────────

function writePAlignment(p: Element, value: string) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "jc");
  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", value);
  pPr.appendChild(jc);
}

function writePIndent(p: Element, firstLine: number, left = 0, right = 0) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "ind");
  const ind = wElem(p.ownerDocument!, "ind");
  setWAttr(ind, "firstLine", String(firstLine));
  setWAttr(ind, "left", String(left));
  setWAttr(ind, "right", String(right));
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

function writePageBreakBefore(p: Element) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "pageBreakBefore");
  pPr.appendChild(wElem(p.ownerDocument!, "pageBreakBefore"));
}

function writePPrRPr(
  p: Element,
  font: string,
  size: number,
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
  setWAttr(sz, "val", String(size));
  rPr.appendChild(sz);
  const szCs = wElem(p.ownerDocument!, "szCs");
  setWAttr(szCs, "val", String(size));
  rPr.appendChild(szCs);
}

function stripTabRuns(p: Element) {
  const runs = Array.from(p.childNodes).filter(
    (c): c is Element =>
      c instanceof Element && c.localName === "r" && c.namespaceURI === W_NS,
  );
  for (const run of runs) {
    const hasTabs = run.getElementsByTagNameNS(W_NS, "tab").length > 0;
    const textContent = Array.from(run.getElementsByTagNameNS(W_NS, "t"))
      .map((t) => t.textContent ?? "")
      .join("")
      .trim();
    const hasDrawing = Array.from(run.children).some((c) =>
      ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
    );
    if (hasDrawing) continue;
    if (hasTabs && textContent === "") {
      p.removeChild(run);
    }
  }
  const firstT = p.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
  if (firstT) {
    const trimmed = (firstT.textContent ?? "").replace(/^[\s\t]+/, "");
    if (trimmed !== firstT.textContent) {
      firstT.textContent = trimmed;
      if (trimmed) firstT.setAttribute("xml:space", "preserve");
    }
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
  ])
    removeChildren(pPr, tag);
}

function applyRunFormatting(
  p: Element,
  font: string,
  size: number,
  bold: boolean | null,
  italic: boolean | null,
) {
  const sizeStr = String(size);
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

    removeChildren(rPr, "rFonts");
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

function uppercaseParagraph(p: Element) {
  p.querySelectorAll("r > t").forEach((t) => {
    if (t.namespaceURI === W_NS)
      t.textContent = (t.textContent ?? "").toUpperCase();
  });
}

// ─── format appliers ─────────────────────────────────────────────────────────

function applyPreliminaryTitle(p: Element, config: FormattingConfig) {
  const pPr = ensurePPr(p);
  removeChildren(pPr, "ind");
  removeChildren(pPr, "spacing");
  removeChildren(pPr, "pageBreakBefore");
  removeChildren(pPr, "rPr");
  removeChildren(pPr, "jc");

  if (config.titles.textTransform === "uppercase") {
    uppercaseParagraph(p);
  }

  // Explicitly force alignment
  const jc = wElem(p.ownerDocument!, "jc");
  setWAttr(jc, "val", config.titles.alignment);
  pPr.appendChild(jc);

  // Indent
  const ind = wElem(p.ownerDocument!, "ind");
  const twips = inchesToTwips(config.titles.indentation);
  setWAttr(ind, "firstLine", "0");
  setWAttr(ind, "left", String(twips));
  setWAttr(ind, "right", "0");
  pPr.appendChild(ind);

  // Spacing: using config line spacing
  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", String(linesToTwips(config.titles.lineSpacing)));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);

  // Suppress contextual spacing
  const cs = wElem(p.ownerDocument!, "contextualSpacing");
  setWAttr(cs, "val", "0");
  pPr.appendChild(cs);

  // Page break before
  pPr.appendChild(wElem(p.ownerDocument!, "pageBreakBefore"));

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

function applyPreliminaryBody(p: Element, config: FormattingConfig) {
  stripTabRuns(p);
  stripPPr(p);
  writePAlignment(p, config.body.alignment);
  writePIndent(p, inchesToTwips(config.body.indentation), 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing));
  const size = ptsToHalfPts(config.body.fontSize);
  writePPrRPr(p, config.body.fontFamily, size, false, false);
  applyRunFormatting(p, config.body.fontFamily, size, null, null);
}

function applySignatory(p: Element, config: FormattingConfig) {
  stripPPr(p);
  writePAlignment(p, "center");
  writePIndent(p, 0, 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing / 2));
  const size = ptsToHalfPts(config.body.fontSize);
  writePPrRPr(p, config.body.fontFamily, size, true, false);
  applyRunFormatting(p, config.body.fontFamily, size, true, false);
}

function applyDesignation(p: Element, config: FormattingConfig) {
  stripTabRuns(p);
  stripPPr(p);
  writePAlignment(p, "center");
  writePIndent(p, 0, 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing));
  const size = ptsToHalfPts(config.body.fontSize);
  writePPrRPr(p, config.body.fontFamily, size, false, false);
  applyRunFormatting(p, config.body.fontFamily, size, false, false);
}

function applySignatoryRight(p: Element, config: FormattingConfig) {
  stripTabRuns(p);
  stripPPr(p);
  writePAlignment(p, "right");
  writePIndent(p, 0, 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing / 2));
  const size = ptsToHalfPts(config.body.fontSize);
  writePPrRPr(p, config.body.fontFamily, size, true, false);
  applyRunFormatting(p, config.body.fontFamily, size, true, false);
}

function applyDesignationRight(p: Element, config: FormattingConfig) {
  stripTabRuns(p);
  stripPPr(p);
  writePAlignment(p, "right");
  writePIndent(p, 0, 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing));
  const size = ptsToHalfPts(config.body.fontSize);
  writePPrRPr(p, config.body.fontFamily, size, false, false);
  applyRunFormatting(p, config.body.fontFamily, size, false, false);
}

function applyInitials(p: Element, config: FormattingConfig) {
  stripPPr(p);
  writePAlignment(p, "right");
  writePIndent(p, 0, 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing));
  const size = ptsToHalfPts(config.body.fontSize);
  writePPrRPr(p, config.body.fontFamily, size, true, false);
  applyRunFormatting(p, config.body.fontFamily, size, true, false);
}

function applyKeywordsLine(p: Element, config: FormattingConfig) {
  stripTabRuns(p);
  stripPPr(p);
  writePAlignment(p, config.body.alignment);
  writePIndent(p, 0, 0, 0);
  writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing));
  writePPrRPr(
    p,
    config.body.fontFamily,
    ptsToHalfPts(config.body.fontSize),
    false,
    true,
  );

  const textRuns = Array.from(p.querySelectorAll("r")).filter(
    (r) =>
      r.namespaceURI === W_NS &&
      !Array.from(r.children).some((c) =>
        ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
      ),
  );

  let colonFound = false;
  const sizeVal = String(ptsToHalfPts(config.body.fontSize));
  for (const run of textRuns) {
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    const text = tEl?.textContent ?? "";

    removeChildren(run, "rPr");
    const rPr = wElem(run.ownerDocument!, "rPr");
    run.insertBefore(rPr, run.firstChild ?? null);

    const rFonts = wElem(run.ownerDocument!, "rFonts");
    for (const a of ["ascii", "hAnsi", "eastAsia", "cs"])
      setWAttr(rFonts, a, config.body.fontFamily);
    rPr.appendChild(rFonts);

    const sz = wElem(run.ownerDocument!, "sz");
    setWAttr(sz, "val", sizeVal);
    rPr.appendChild(sz);
    const szCs = wElem(run.ownerDocument!, "szCs");
    setWAttr(szCs, "val", sizeVal);
    rPr.appendChild(szCs);

    rPr.appendChild(wElem(run.ownerDocument!, "i"));
    rPr.appendChild(wElem(run.ownerDocument!, "iCs"));

    if (!colonFound) {
      rPr.appendChild(wElem(run.ownerDocument!, "b"));
      rPr.appendChild(wElem(run.ownerDocument!, "bCs"));
      if (text.includes(":")) colonFound = true;
    }
  }
}

function applyEmptyParagraph(p: Element, config: FormattingConfig, line = 240) {
  // Remove any runs containing only line breaks (soft returns) — these create phantom blank lines
  Array.from(p.childNodes)
    .filter(
      (c): c is Element =>
        c instanceof Element && c.localName === "r" && c.namespaceURI === W_NS,
    )
    .forEach((run) => {
      const hasText = Array.from(run.getElementsByTagNameNS(W_NS, "t")).some(
        (t) => (t.textContent ?? "").trim() !== "",
      );
      if (!hasText) p.removeChild(run);
    });
  const pPr = ensurePPr(p);
  removeChildren(pPr, "spacing");
  const sp = wElem(p.ownerDocument!, "spacing");
  setWAttr(sp, "before", "0");
  setWAttr(sp, "after", "0");
  setWAttr(sp, "line", String(line));
  setWAttr(sp, "lineRule", "auto");
  setWAttr(sp, "beforeAutospacing", "0");
  setWAttr(sp, "afterAutospacing", "0");
  pPr.appendChild(sp);
  writePPrRPr(
    p,
    config.body.fontFamily,
    ptsToHalfPts(config.body.fontSize),
    false,
    false,
  );
}

// ─── section title detection ─────────────────────────────────────────────────

type PrelimSection =
  | "title_page"
  | "approval_sheet"
  | "abstract"
  | "acknowledgment"
  | "dedication"
  | "toc"
  | "list_of_figures"
  | "list_of_tables"
  | "list_of_appendices";

const PRELIM_TITLE_MAP: Array<[RegExp, PrelimSection]> = [
  [/^APPROVAL\s+SHEET$/i, "approval_sheet"],
  [/^ABSTRACT$/i, "abstract"],
  [/^ACKNOWLEDG(E?)MENT$/i, "acknowledgment"],
  [/^DEDICATION$/i, "dedication"],
  [/^TABLE\s+OF\s+CONTENTS$/i, "toc"],
  [/^LIST\s+OF\s+FIGURES?$/i, "list_of_figures"],
  [/^FIGURES?$/i, "list_of_figures"],
  [/^LIST\s+OF\s+TABLES?$/i, "list_of_tables"],
  [/^TABLES?$/i, "list_of_tables"],
  [/^LIST\s+OF\s+APPENDI(CES|X|XES)$/i, "list_of_appendices"],
  [/^APPENDI(CES|X|XES)$/i, "list_of_appendices"],
];

function detectSectionTitle(normalized: string): PrelimSection | null {
  const trimmed = normalized.trim();
  for (const [pattern, section] of PRELIM_TITLE_MAP) {
    if (pattern.test(trimmed)) return section;
  }
  return null;
}

const CANONICAL_TITLE: Partial<Record<PrelimSection, string>> = {
  acknowledgment: "ACKNOWLEDGMENT",
  list_of_figures: "LIST OF FIGURES",
  list_of_tables: "LIST OF TABLES",
  list_of_appendices: "LIST OF APPENDICES",
};

function rewriteTitleText(p: Element, section: PrelimSection) {
  const target = CANONICAL_TITLE[section];
  if (!target) return;
  const tNodes = Array.from(p.querySelectorAll("r > t")).filter(
    (t) => t.namespaceURI === W_NS,
  );
  if (tNodes.length === 0) return;
  tNodes[0].textContent = target;
  for (let i = 1; i < tNodes.length; i++) tNodes[i].textContent = "";
}

// ─── signatory / designation / initials detection ────────────────────────────

function looksLikeSignatory(p: Element, normalized: string): boolean {
  if (!normalized || normalized.length > 120) return false;
  if (/[.!?]$/.test(normalized)) return false;

  const textRuns = Array.from(p.querySelectorAll("r")).filter(
    (r) =>
      r.namespaceURI === W_NS &&
      !Array.from(r.children).some((c) =>
        ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
      ),
  );
  let boldCount = 0,
    total = 0;
  for (const run of textRuns) {
    const tEl = run.getElementsByTagNameNS(W_NS, "t").item(0) as Element | null;
    if (!tEl || !(tEl.textContent ?? "").trim()) continue;
    total++;
    const rPr = getChild(run, "rPr");
    if (rPr && rPr.getElementsByTagNameNS(W_NS, "b").length > 0) boldCount++;
  }
  if (total === 0) return false;
  return boldCount / total >= 0.5;
}

function looksLikeDesignation(normalized: string): boolean {
  if (!normalized || normalized.length > 150) return false;
  return /^(adviser|chairman|chair\s*person|member|academic\s+dean|dean|oic|vice\s+president|college\s+president|president|department\s+chair|professor|instructor|faculty)/iu.test(
    normalized,
  );
}

function looksLikeInitials(p: Element, normalized: string): boolean {
  if (!normalized) return false;

  const hasBr = p.getElementsByTagNameNS(W_NS, "br").length > 0;
  if (hasBr) {
    const runs = Array.from(p.querySelectorAll("r")).filter(
      (r) =>
        r.namespaceURI === W_NS &&
        !Array.from(r.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        ),
    );
    let nameRunCount = 0;
    for (const run of runs) {
      const tEl = run
        .getElementsByTagNameNS(W_NS, "t")
        .item(0) as Element | null;
      const text = (tEl?.textContent ?? "").trim();
      if (/^[A-Za-z\u00C0-\u024F\s-]+,\s+([A-Z]\.\s*)+$/.test(text)) {
        nameRunCount++;
      }
    }
    if (nameRunCount >= 1) return true;
  }

  if (normalized.length <= 80) {
    if (/^[A-Za-z\u00C0-\u024F\s-]+,\s+([A-Z]\.\s*)+$/.test(normalized))
      return true;
  }

  if (normalized.length <= 80 && !looksLikeDesignation(normalized)) {
    const runs = Array.from(p.querySelectorAll("r")).filter(
      (r) =>
        r.namespaceURI === W_NS &&
        !Array.from(r.children).some((c) =>
          ["drawing", "pict", "instrText", "fldChar"].includes(c.localName),
        ),
    );
    let boldCount = 0,
      total = 0;
    for (const run of runs) {
      const tEl = run
        .getElementsByTagNameNS(W_NS, "t")
        .item(0) as Element | null;
      if (!tEl || !(tEl.textContent ?? "").trim()) continue;
      total++;
      const rPr = getChild(run, "rPr");
      if (rPr && rPr.getElementsByTagNameNS(W_NS, "b").length > 0) boldCount++;
    }
    if (total > 0 && boldCount / total >= 0.8 && !/[!?]/.test(normalized))
      return true;
  }

  return false;
}

// ─── state machine ────────────────────────────────────────────────────────────

interface PrelimState {
  section: PrelimSection;
  inInitialsBlock: boolean;
  prevWasSignatory: boolean;
  approvalRightBlockDone: boolean;
  prevWasRightSignatory: boolean;
}

// ─── per-paragraph processor ─────────────────────────────────────────────────

function processPrelimParagraph(
  p: Element,
  state: PrelimState,
  config: FormattingConfig,
) {
  const text = getParagraphText(p);
  const normalized = normalizeText(text);
  const inTable = isInTable(p);

  if (!inTable) {
    const detected = detectSectionTitle(normalized);
    if (detected !== null) {
      state.section = detected;
      state.inInitialsBlock = false;
      state.prevWasSignatory = false;
      state.approvalRightBlockDone = false;
      state.prevWasRightSignatory = false;
      rewriteTitleText(p, detected);
      applyPreliminaryTitle(p, config);
      return;
    }
  }

  if (inTable) return;

  const sec = state.section;

  if (sec === "title_page") return;

  if (normalized === "") {
    const emptyLine = sec === "approval_sheet" ? 240 : 480;
    applyEmptyParagraph(p, config, emptyLine);
    return;
  }

  if (
    sec === "toc" ||
    sec === "list_of_figures" ||
    sec === "list_of_tables" ||
    sec === "list_of_appendices"
  )
    return;

  if (sec === "abstract") {
    state.prevWasSignatory = false;
    if (/^keywords?\s*:/iu.test(normalized)) {
      applyKeywordsLine(p, config);
    } else {
      applyPreliminaryBody(p, config);
    }
    return;
  }

  if (sec === "acknowledgment") {
    if (!state.inInitialsBlock && looksLikeInitials(p, normalized)) {
      state.inInitialsBlock = true;
    }
    if (state.inInitialsBlock) {
      applyInitials(p, config);
    } else {
      applyPreliminaryBody(p, config);
    }
    state.prevWasSignatory = false;
    return;
  }

  if (sec === "approval_sheet") {
    if (!state.approvalRightBlockDone) {
      if (looksLikeDesignation(normalized) || state.prevWasRightSignatory) {
        // Pre-table adviser designation — no formatting applied, leave as-is
        state.prevWasRightSignatory = false;
        state.prevWasSignatory = false;
        return;
      }
      if (looksLikeSignatory(p, normalized)) {
        applySignatoryRight(p, config);
        state.prevWasRightSignatory = true;
        state.prevWasSignatory = false;
        return;
      }
      state.prevWasRightSignatory = false;
      state.prevWasSignatory = false;
      applyPreliminaryBody(p, config);
      return;
    }

    if (looksLikeDesignation(normalized) || state.prevWasSignatory) {
      applyDesignation(p, config);
      state.prevWasSignatory = false;
      return;
    }
    if (looksLikeSignatory(p, normalized)) {
      applySignatory(p, config);
      state.prevWasSignatory = true;
      return;
    }
    state.prevWasSignatory = false;
    applyPreliminaryBody(p, config);
    return;
  }

  if (sec === "dedication") {
    state.prevWasSignatory = false;
    applyPreliminaryBody(p, config);
    return;
  }

  applyPreliminaryBody(p, config);
}

// ─── table unfloat ────────────────────────────────────────────────────────────

function unfloatTable(tbl: Element) {
  const tblPr = tbl.getElementsByTagNameNS(W_NS, "tblPr").item(0);
  if (!tblPr) return;
  // Remove tblpPr — this is the element that makes a table float / wrap text
  Array.from(tblPr.getElementsByTagNameNS(W_NS, "tblpPr")).forEach((el) =>
    tblPr.removeChild(el),
  );
}

// ─── abstract separator line builder ─────────────────────────────────────────

function buildAbstractSeparatorP(doc: Document, config: FormattingConfig): Element {
  const p = doc.createElementNS(W_NS, "w:p");

  const pPr = doc.createElementNS(W_NS, "w:pPr");
  p.appendChild(pPr);
  const sp = doc.createElementNS(W_NS, "w:spacing");
  sp.setAttributeNS(W_NS, "w:line", "480");
  sp.setAttributeNS(W_NS, "w:lineRule", "auto");
  pPr.appendChild(sp);
  const rPr0 = doc.createElementNS(W_NS, "w:rPr");
  pPr.appendChild(rPr0);
  const rFonts0 = doc.createElementNS(W_NS, "w:rFonts");
  rFonts0.setAttributeNS(W_NS, "w:ascii", "Garamond");
  rFonts0.setAttributeNS(W_NS, "w:hAnsi", "Garamond");
  rPr0.appendChild(rFonts0);
  const sz0 = doc.createElementNS(W_NS, "w:sz");
  sz0.setAttributeNS(W_NS, "w:val", "24");
  rPr0.appendChild(sz0);
  const szCs0 = doc.createElementNS(W_NS, "w:szCs");
  szCs0.setAttributeNS(W_NS, "w:val", "24");
  rPr0.appendChild(szCs0);

  const run = doc.createElementNS(W_NS, "w:r");
  p.appendChild(run);
  const rPr = doc.createElementNS(W_NS, "w:rPr");
  run.appendChild(rPr);
  const rFonts = doc.createElementNS(W_NS, "w:rFonts");
  rFonts.setAttributeNS(W_NS, "w:ascii", "Garamond");
  rFonts.setAttributeNS(W_NS, "w:hAnsi", "Garamond");
  rPr.appendChild(rFonts);
  rPr.appendChild(doc.createElementNS(W_NS, "w:noProof"));
  const sz = doc.createElementNS(W_NS, "w:sz");
  sz.setAttributeNS(W_NS, "w:val", "24");
  rPr.appendChild(sz);
  const szCs = doc.createElementNS(W_NS, "w:szCs");
  szCs.setAttributeNS(W_NS, "w:val", "24");
  rPr.appendChild(szCs);

  const altContent = doc.createElementNS(MC_NS, "mc:AlternateContent");
  run.appendChild(altContent);

  const choice = doc.createElementNS(MC_NS, "mc:Choice");
  choice.setAttribute("Requires", "wps");
  altContent.appendChild(choice);

  const drawing = doc.createElementNS(W_NS, "w:drawing");
  choice.appendChild(drawing);

  const inline = doc.createElementNS(WP_NS, "wp:inline");
  inline.setAttribute("distT", "0");
  inline.setAttribute("distB", "0");
  inline.setAttribute("distL", "0");
  inline.setAttribute("distR", "0");
  drawing.appendChild(inline);

  const extent = doc.createElementNS(WP_NS, "wp:extent");
  extent.setAttribute("cx", "5424238");
  extent.setAttribute("cy", "0");
  inline.appendChild(extent);

  const effectExtent = doc.createElementNS(WP_NS, "wp:effectExtent");
  effectExtent.setAttribute("l", "0");
  effectExtent.setAttribute("t", "19050");
  effectExtent.setAttribute("r", "24130");
  effectExtent.setAttribute("b", "19050");
  inline.appendChild(effectExtent);

  const docPr = doc.createElementNS(WP_NS, "wp:docPr");
  docPr.setAttribute("id", "1892502313");
  docPr.setAttribute("name", "Straight Connector 10");
  inline.appendChild(docPr);

  inline.appendChild(doc.createElementNS(WP_NS, "wp:cNvGraphicFramePr"));

  const graphic = doc.createElementNS(A_NS, "a:graphic");
  inline.appendChild(graphic);

  const graphicData = doc.createElementNS(A_NS, "a:graphicData");
  graphicData.setAttribute(
    "uri",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  );
  graphic.appendChild(graphicData);

  const wsp = doc.createElementNS(WPS_NS, "wps:wsp");
  graphicData.appendChild(wsp);

  wsp.appendChild(doc.createElementNS(WPS_NS, "wps:cNvCnPr"));

  const spPr = doc.createElementNS(WPS_NS, "wps:spPr");
  wsp.appendChild(spPr);

  const xfrm = doc.createElementNS(A_NS, "a:xfrm");
  spPr.appendChild(xfrm);
  const off = doc.createElementNS(A_NS, "a:off");
  off.setAttribute("x", "0");
  off.setAttribute("y", "0");
  xfrm.appendChild(off);
  const ext = doc.createElementNS(A_NS, "a:ext");
  ext.setAttribute("cx", "5424238");
  ext.setAttribute("cy", "0");
  xfrm.appendChild(ext);

  const prstGeom = doc.createElementNS(A_NS, "a:prstGeom");
  prstGeom.setAttribute("prst", "line");
  prstGeom.appendChild(doc.createElementNS(A_NS, "a:avLst"));
  spPr.appendChild(prstGeom);

  const ln = doc.createElementNS(A_NS, "a:ln");
  ln.setAttribute("w", "38100");
  spPr.appendChild(ln);

  const style = doc.createElementNS(WPS_NS, "wps:style");
  wsp.appendChild(style);

  const addStyleRef = (tag: string, idx: string, clr: string) => {
    const ref = doc.createElementNS(A_NS, tag);
    ref.setAttribute("idx", idx);
    const sc = doc.createElementNS(A_NS, "a:schemeClr");
    sc.setAttribute("val", clr);
    ref.appendChild(sc);
    style.appendChild(ref);
  };
  addStyleRef("a:lnRef", "3", "dk1");
  addStyleRef("a:fillRef", "0", "dk1");
  addStyleRef("a:effectRef", "2", "dk1");
  addStyleRef("a:fontRef", "minor", "tx1");

  wsp.appendChild(doc.createElementNS(WPS_NS, "wps:bodyPr"));

  const fallback = doc.createElementNS(MC_NS, "mc:Fallback");
  altContent.appendChild(fallback);
  const pict = doc.createElementNS(W_NS, "w:pict");
  fallback.appendChild(pict);

  const vlineStr =
    "<root>" +
    '<v:line xmlns:v="urn:schemas-microsoft-com:vml"' +
    ' xmlns:w10="urn:schemas-microsoft-com:office:word"' +
    ' id="Straight Connector 10"' +
    ' style="visibility:visible;mso-wrap-style:square"' +
    ' from="0,0" to="427.1pt,0"' +
    ` strokecolor="black" strokeweight="${config.figure.borderWeight}pt">` +
    '<v:stroke joinstyle="miter"/>' +
    "<w10:anchorlock/>" +
    "</v:line>" +
    "</root>";

  const tmpDoc = new DOMParser().parseFromString(vlineStr, "application/xml");
  const vline = tmpDoc.documentElement.firstElementChild;
  if (vline) pict.appendChild(doc.importNode(vline, true));

  return p;
}

// ─── abstract table formatter ────────────────────────────────────────────────

function formatAbstractTable(tbl: Element, config: FormattingConfig) {
  const rows = Array.from(tbl.querySelectorAll("tr")).filter(
    (r) => r.namespaceURI === W_NS,
  );

  let firstDataRowIdx = -1;
  const size = ptsToHalfPts(config.body.fontSize);
  for (let i = 0; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll("tc")).filter(
      (tc) => tc.namespaceURI === W_NS,
    );
    if (cells.length === 0) continue;
    const col1Text = normalizeText(
      Array.from(cells[0]?.querySelectorAll("t") ?? [])
        .map((t) => t.textContent ?? "")
        .join(""),
    );
    if (col1Text !== "") {
      firstDataRowIdx = i;
      break;
    }
  }

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const cells = Array.from(row.querySelectorAll("tc")).filter(
      (tc) => tc.namespaceURI === W_NS,
    );
    if (cells.length === 0) continue;

    const isFirstDataRow = rowIdx === firstDataRowIdx;

    cells.forEach((tc, colIdx) => {
      const isCol1 = colIdx === 0;
      const isCol2 = colIdx === 1;

      const cellParas = Array.from(tc.querySelectorAll("p")).filter(
        (p) => p.namespaceURI === W_NS,
      );

      for (const p of cellParas) {
        const pText = normalizeText(getParagraphText(p));

        if (pText === "") {
          const pPr = ensurePPr(p);
          removeChildren(pPr, "spacing");
          const sp = wElem(p.ownerDocument!, "spacing");
          setWAttr(sp, "before", "0");
          setWAttr(sp, "after", "0");
          setWAttr(sp, "line", String(linesToTwips(config.body.lineSpacing / 2)));
          setWAttr(sp, "lineRule", "auto");
          setWAttr(sp, "beforeAutospacing", "0");
          setWAttr(sp, "afterAutospacing", "0");
          pPr.appendChild(sp);
          removeChildren(pPr, "contextualSpacing");
          const cs = wElem(p.ownerDocument!, "contextualSpacing");
          setWAttr(cs, "val", "0");
          pPr.appendChild(cs);
          continue;
        }

        if (isCol1) {
          stripTabRuns(p);
          stripPPr(p);
          writePAlignment(p, config.table.alignment);
          writePSpacing(p, 0, 0, linesToTwips(config.table.lineSpacing / 2));
          const size = ptsToHalfPts(config.table.fontSize);
          writePPrRPr(p, config.table.fontFamily, size, false, false);
          applyRunFormatting(p, config.table.fontFamily, size, false, false);
        } else if (isCol2) {
          if (isFirstDataRow) uppercaseParagraph(p);
          stripTabRuns(p);
          stripPPr(p);
          writePAlignment(p, config.body.alignment);
          writePIndent(p, 0, 0, 0);
          writePSpacing(p, 0, 0, linesToTwips(config.body.lineSpacing / 2));
          writePPrRPr(p, config.body.fontFamily, size, true, false);
          applyRunFormatting(p, config.body.fontFamily, size, true, false);
        }
      }
    });
  }
}

// ─── public entry ─────────────────────────────────────────────────────────────

export async function formatPreliminary(
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
  if (!body) throw new Error("No body element found in document.xml");

  unwrapContentControls(body);

  const allChildren = Array.from(body.childNodes).filter(
    (c): c is Element => c instanceof Element,
  );

  let chapterStart = allChildren.length;
  for (let i = 0; i < allChildren.length; i++) {
    const child = allChildren[i];
    if (child.localName !== "p") continue;
    const t = normalizeText(
      Array.from(child.querySelectorAll("t"))
        .map((n) => n.textContent ?? "")
        .join(""),
    );
    if (/^chapter\s+([ivxlcdm]+|\d+)$/iu.test(t)) {
      chapterStart = i;
      break;
    }
  }

  const prelimChildren = allChildren.slice(0, chapterStart);

  let preSection: PrelimSection = "title_page";
  let lastWasEmpty = false;
  for (const child of [...prelimChildren]) {
    if (!(child instanceof Element)) continue;
    if (child.parentElement !== body) continue;
    if (child.localName !== "p") {
      lastWasEmpty = false;
      continue;
    }
    const txt = normalizeText(getParagraphText(child));
    const det = detectSectionTitle(txt);
    if (det !== null) {
      preSection = det;
      lastWasEmpty = false;
      continue;
    }
    if (txt !== "") {
      lastWasEmpty = false;
      continue;
    }
    // CRITICAL: never delete an empty paragraph that carries a sectPr —
    // it defines the section's page numbering format AND acts as the
    // section break that separates preliminary pages from chapter pages.
    const pPrEl = child.getElementsByTagNameNS(W_NS, "pPr").item(0);
    const hasSectPr =
      pPrEl != null && pPrEl.getElementsByTagNameNS(W_NS, "sectPr").length > 0;
    if (hasSectPr) {
      lastWasEmpty = false;
      continue;
    }
    if (preSection === "title_page") {
      lastWasEmpty = false;
      continue;
    }
    if (preSection === "approval_sheet") {
      if (lastWasEmpty) {
        body.removeChild(child);
      } else {
        lastWasEmpty = true;
      }
    } else {
      body.removeChild(child);
    }
  }

  const state: PrelimState = {
    section: "title_page",
    inInitialsBlock: false,
    prevWasSignatory: false,
    approvalRightBlockDone: false,
    prevWasRightSignatory: false,
  };

  const finalPrelimChildren = Array.from(body.childNodes).filter(
    (c): c is Element => {
      if (!(c instanceof Element)) return false;
      const idx = allChildren.indexOf(c);
      return idx >= 0 && idx < chapterStart;
    },
  );

  for (const child of finalPrelimChildren) {
    if (!(child instanceof Element)) continue;
    if (child.parentElement !== body) continue;
    if (child.localName === "p") {
      processPrelimParagraph(child, state, options.config);
    } else if (child.localName === "tbl") {
      unfloatTable(child);
      if (state.section === "approval_sheet") {
        state.approvalRightBlockDone = true;
        state.prevWasRightSignatory = false;
      }
      if (state.section === "abstract") {
        formatAbstractTable(child, options.config);
        const nextSib = child.nextSibling as Element | null;
        const nextHasDrawing =
          nextSib instanceof Element &&
          nextSib.localName === "p" &&
          (nextSib.getElementsByTagNameNS(W_NS, "drawing").length > 0 ||
            nextSib.getElementsByTagNameNS(W_NS, "pict").length > 0);
        if (!nextHasDrawing) {
          const sep = buildAbstractSeparatorP(child.ownerDocument!, options.config);
          if (child.nextSibling) {
            body.insertBefore(sep, child.nextSibling);
          } else {
            body.appendChild(sep);
          }
        }
      }
    }
  }

  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(dom);
  zip.file("word/document.xml", newXml);
  return zip.generateAsync({ type: "blob" }) as Promise<Blob>;
}
