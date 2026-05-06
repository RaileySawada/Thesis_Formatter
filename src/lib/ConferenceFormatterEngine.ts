// In the acm formatting, let's ignore for now the first page content which are the title of the paper up to the ACM Reference Format and the footnote. Now let's proceed to the introduction and other parts of the ACM. for the heading 1 like "1 INTRODUCTION" the format of it should be Linux Biolinum O with the size of 9 and all uppercase with space before (12 pt) and after (3pt) paragraph and bold. Heading 2  (e.g. "1.1 Accessibility") should have font family of Linux Biolinum O with size of 9, Capital each first letter of words and bold. All headings should have spacing before (12pt) and after (3pt) paragaph
import {
  DEFAULT_CONFERENCE_FORMATTING_CONFIG,
  type AcmFormattingConfig,
  type ConferenceFormat,
  type ConferenceFormattingConfig,
  type ConferenceTextStyle,
  type PublicationFormattingConfig,
} from "../constants";
import { requestPollinations } from "./pollinationsClient";
import { isAiAssistEnabled } from "./aiAssist";

interface ConferenceFormatOptions {
  format: ConferenceFormat;
  styleConfig?: ConferenceFormattingConfig;
}

interface AuthorEntry {
  name: string;
  department: string;
  organization: string;
  cityCountry: string;
  contact: string;
}

const AI_ASSIST_ENABLED = isAiAssistEnabled(
  import.meta.env.VITE_ENABLE_AI_ASSIST,
);

const PUBFORM_SOURCE_FILE =
  "/conference/publication_formatting_guidelines.docx";
const ACM_SOURCE_FILE = "/conference/acm_formatting_guidelines.docx";
const FALLBACK_W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function resolveWNs(doc: Document): string {
  return doc.documentElement.lookupNamespaceURI("w") ?? FALLBACK_W_NS;
}

function wAttr(el: Element, wNs: string, local: string): string {
  return el.getAttributeNS(wNs, local) ?? el.getAttribute(`w:${local}`) ?? "";
}

function setWAttr(el: Element, wNs: string, local: string, val: string) {
  el.setAttributeNS(wNs, `w:${local}`, val);
}

function wElem(doc: Document, wNs: string, local: string): Element {
  return doc.createElementNS(wNs, `w:${local}`);
}

function getChild(parent: Element, wNs: string, local: string): Element | null {
  for (const c of Array.from(parent.childNodes)) {
    if (
      c instanceof Element &&
      c.namespaceURI === wNs &&
      c.localName === local
    ) {
      return c;
    }
  }
  return null;
}

function ensureChild(
  parent: Element,
  wNs: string,
  local: string,
  prepend = false,
): Element {
  const existing = getChild(parent, wNs, local);
  if (existing) return existing;
  const child = wElem(parent.ownerDocument!, wNs, local);
  if (prepend && parent.firstChild)
    parent.insertBefore(child, parent.firstChild);
  else parent.appendChild(child);
  return child;
}

function removeChildren(parent: Element, wNs: string, local: string) {
  Array.from(parent.childNodes)
    .filter(
      (c) =>
        c instanceof Element && c.namespaceURI === wNs && c.localName === local,
    )
    .forEach((c) => parent.removeChild(c));
}

function ensurePPr(p: Element, wNs: string): Element {
  return ensureChild(p, wNs, "pPr", true);
}

function ensureRPr(run: Element, wNs: string): Element {
  return ensureChild(run, wNs, "rPr", true);
}

function getBody(doc: Document, wNs: string): Element | null {
  const bodies = doc.getElementsByTagNameNS(wNs, "body");
  return bodies.length > 0 ? (bodies.item(0) as Element) : null;
}

function getParagraphText(p: Element, wNs: string): string {
  const parts: string[] = [];
  const textNodes = p.getElementsByTagNameNS(wNs, "t");
  for (const t of Array.from(textNodes)) {
    parts.push(t.textContent ?? "");
  }
  return parts.join("");
}

function normalizeText(text: string): string {
  return text.replace(/\s+/gu, " ").trim();
}

function clonePreferredRunProps(p: Element, wNs: string): Element | null {
  const runs = Array.from(p.getElementsByTagNameNS(wNs, "r"));
  for (const run of runs) {
    const hasText = run.getElementsByTagNameNS(wNs, "t").length > 0;
    if (!hasText) continue;
    const rPr = getChild(run, wNs, "rPr");
    if (rPr) return rPr.cloneNode(true) as Element;
  }
  const pPr = getChild(p, wNs, "pPr");
  const paraRPr = pPr ? getChild(pPr, wNs, "rPr") : null;
  if (paraRPr) return paraRPr.cloneNode(true) as Element;
  return null;
}

function setParagraphText(p: Element, wNs: string, text: string) {
  const preservedRPr = clonePreferredRunProps(p, wNs);
  const existingPPr = getChild(p, wNs, "pPr");
  Array.from(p.childNodes).forEach((n) => {
    if (n !== existingPPr) p.removeChild(n);
  });

  const run = wElem(p.ownerDocument!, wNs, "r");
  if (preservedRPr) {
    run.appendChild(preservedRPr);
  }

  const lines = text.split(/\r?\n/u);
  for (let i = 0; i < lines.length; i += 1) {
    if (i > 0) {
      run.appendChild(wElem(p.ownerDocument!, wNs, "br"));
    }
    const t = wElem(p.ownerDocument!, wNs, "t");
    const part = lines[i];
    if (/^\s/u.test(part) || /\s$/u.test(part) || /\s{2,}/u.test(part)) {
      t.setAttribute("xml:space", "preserve");
    }
    t.textContent = part;
    run.appendChild(t);
  }

  p.appendChild(run);
}

interface StyledLine {
  text: string;
  italic: boolean;
}

function applyRunTypography(
  run: Element,
  wNs: string,
  fontFamily: string,
  fontPt: number,
  italic: boolean,
) {
  const rPr = ensureRPr(run, wNs);
  const halfPt = String(ptsToHalfPts(fontPt));

  removeChildren(rPr, wNs, "rFonts");
  const rFonts = ensureChild(rPr, wNs, "rFonts");
  setWAttr(rFonts, wNs, "ascii", fontFamily);
  setWAttr(rFonts, wNs, "hAnsi", fontFamily);
  setWAttr(rFonts, wNs, "eastAsia", fontFamily);
  setWAttr(rFonts, wNs, "cs", fontFamily);

  removeChildren(rPr, wNs, "sz");
  removeChildren(rPr, wNs, "szCs");
  const sz = ensureChild(rPr, wNs, "sz");
  setWAttr(sz, wNs, "val", halfPt);
  const szCs = ensureChild(rPr, wNs, "szCs");
  setWAttr(szCs, wNs, "val", halfPt);

  removeChildren(rPr, wNs, "b");
  removeChildren(rPr, wNs, "bCs");

  removeChildren(rPr, wNs, "i");
  removeChildren(rPr, wNs, "iCs");
  if (italic) {
    rPr.appendChild(wElem(run.ownerDocument!, wNs, "i"));
    rPr.appendChild(wElem(run.ownerDocument!, wNs, "iCs"));
  }
}

function setParagraphStyledLines(
  p: Element,
  wNs: string,
  lines: StyledLine[],
  fontFamily = "Times New Roman",
  fontPt = 9,
) {
  const existingPPr = getChild(p, wNs, "pPr");
  Array.from(p.childNodes).forEach((n) => {
    if (n !== existingPPr) p.removeChild(n);
  });

  lines.forEach((line, index) => {
    const run = wElem(p.ownerDocument!, wNs, "r");
    applyRunTypography(run, wNs, fontFamily, fontPt, line.italic);

    const t = wElem(p.ownerDocument!, wNs, "t");
    const part = line.text;
    if (/^\s/u.test(part) || /\s$/u.test(part) || /\s{2,}/u.test(part)) {
      t.setAttribute("xml:space", "preserve");
    }
    t.textContent = part;
    run.appendChild(t);
    p.appendChild(run);

    if (index < lines.length - 1) {
      const brRun = wElem(p.ownerDocument!, wNs, "r");
      applyRunTypography(brRun, wNs, fontFamily, fontPt, false);
      brRun.appendChild(wElem(p.ownerDocument!, wNs, "br"));
      p.appendChild(brRun);
    }
  });
}

function clearParagraphContent(p: Element, wNs: string) {
  const existingPPr = getChild(p, wNs, "pPr");
  Array.from(p.childNodes).forEach((n) => {
    if (n !== existingPPr) p.removeChild(n);
  });
}

function writeColumnBreakRun(
  p: Element,
  wNs: string,
  fontFamily: string,
  fontPt: number,
) {
  const run = wElem(p.ownerDocument!, wNs, "r");
  applyRunTypography(run, wNs, fontFamily, fontPt, false);
  const br = wElem(p.ownerDocument!, wNs, "br");
  setWAttr(br, wNs, "type", "column");
  run.appendChild(br);
  p.appendChild(run);
}

function setParagraphTwoAuthorColumns(
  p: Element,
  wNs: string,
  left: AuthorEntry,
  right: AuthorEntry,
  fontFamily = "Times New Roman",
  fontPt = 9,
  leadingColumnBreaks = 0,
) {
  clearParagraphContent(p, wNs);
  const leftLines = toStyledAuthorLines(left);
  const rightLines = toStyledAuthorLines(right);

  const writeLine = (line: StyledLine, addBreak: boolean) => {
    const run = wElem(p.ownerDocument!, wNs, "r");
    applyRunTypography(run, wNs, fontFamily, fontPt, line.italic);
    const t = wElem(p.ownerDocument!, wNs, "t");
    const part = line.text;
    if (/^\s/u.test(part) || /\s$/u.test(part) || /\s{2,}/u.test(part)) {
      t.setAttribute("xml:space", "preserve");
    }
    t.textContent = part;
    run.appendChild(t);
    p.appendChild(run);

    if (addBreak) {
      const brRun = wElem(p.ownerDocument!, wNs, "r");
      applyRunTypography(brRun, wNs, fontFamily, fontPt, false);
      brRun.appendChild(wElem(p.ownerDocument!, wNs, "br"));
      p.appendChild(brRun);
    }
  };

  for (let i = 0; i < Math.max(0, leadingColumnBreaks); i += 1) {
    writeColumnBreakRun(p, wNs, fontFamily, fontPt);
  }

  leftLines.forEach((line, idx) => writeLine(line, idx < leftLines.length - 1));
  writeColumnBreakRun(p, wNs, fontFamily, fontPt);
  rightLines.forEach((line, idx) =>
    writeLine(line, idx < rightLines.length - 1),
  );
  writeParagraphLayout(p, wNs, "center", 1.0, 0);
}

function setParagraphAuthorColumns(
  p: Element,
  wNs: string,
  authors: AuthorEntry[],
  fontFamily = "Times New Roman",
  fontPt = 9,
) {
  clearParagraphContent(p, wNs);
  const writeLine = (line: StyledLine, addBreak: boolean) => {
    const run = wElem(p.ownerDocument!, wNs, "r");
    applyRunTypography(run, wNs, fontFamily, fontPt, line.italic);
    const t = wElem(p.ownerDocument!, wNs, "t");
    const part = line.text;
    if (/^\s/u.test(part) || /\s$/u.test(part) || /\s{2,}/u.test(part)) {
      t.setAttribute("xml:space", "preserve");
    }
    t.textContent = part;
    run.appendChild(t);
    p.appendChild(run);
    if (addBreak) {
      const brRun = wElem(p.ownerDocument!, wNs, "r");
      applyRunTypography(brRun, wNs, fontFamily, fontPt, false);
      brRun.appendChild(wElem(p.ownerDocument!, wNs, "br"));
      p.appendChild(brRun);
    }
  };

  authors.forEach((author, authorIndex) => {
    const lines = toStyledAuthorLines(author);
    lines.forEach((line, lineIndex) =>
      writeLine(line, lineIndex < lines.length - 1),
    );
    if (authorIndex < authors.length - 1) {
      writeColumnBreakRun(p, wNs, fontFamily, fontPt);
    }
  });

  writeParagraphLayout(p, wNs, "center", 1.0, 0);
}

function containsContact(text: string): boolean {
  return /@/u.test(text) || /\borcid\b/iu.test(text);
}

function extractContacts(text: string): string[] {
  const emails = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/giu) ?? [];
  const orcids = text.match(/\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b/giu) ?? [];
  const merged = [...emails, ...orcids].map((v) => v.trim()).filter(Boolean);
  if (merged.length > 0) return merged;

  // Fallback when values are not strict email/ORCID patterns.
  const cleaned = text
    .replace(/^(emails?|line\s*5:|orcid)[:\s-]*/iu, "")
    .trim();
  if (!cleaned) return [];
  return cleaned
    .split(/\s*;\s*/u)
    .map((v) => v.trim())
    .filter(Boolean);
}

function splitAuthorNames(text: string): string[] {
  const cleaned = text
    .replace(/^line\s*1\s*:\s*/iu, "")
    .replace(/\s+and\s+/giu, ", ")
    .trim();
  return cleaned
    .split(/\s*,\s*/u)
    .map((v) => v.trim())
    .filter(Boolean);
}

function normalizeInputLinesUntilAbstract(
  doc: Document,
  wNs: string,
): string[] {
  const body = getBody(doc, wNs);
  if (!body) return [];
  const paras = Array.from(body.getElementsByTagNameNS(wNs, "p"));
  const lines: string[] = [];

  for (const p of paras) {
    const text = normalizeText(getParagraphText(p, wNs));
    if (!text) continue;
    if (/^abstract\b/iu.test(text) || /^abstract[-—]/iu.test(text)) break;
    lines.push(text);
  }
  return lines;
}

function parseAuthorEntriesFromFrontMatter(lines: string[]): AuthorEntry[] {
  if (lines.length <= 1) return [];

  const frontMatter = lines.slice(1); // skip title
  const chunks: string[][] = [];
  let current: string[] = [];

  for (const line of frontMatter) {
    current.push(line);
    if (containsContact(line)) {
      chunks.push(current);
      current = [];
    }
  }
  if (current.length > 0) chunks.push(current);

  const authors: AuthorEntry[] = [];

  for (const chunk of chunks) {
    const nonContact = chunk.filter((l) => !containsContact(l));
    const contactLine = chunk.find((l) => containsContact(l)) ?? "";
    const contacts = extractContacts(contactLine);

    const nameLine = nonContact[0] ?? "";
    const names = splitAuthorNames(nameLine);
    if (names.length === 0) continue;

    const deptCandidate =
      nonContact.find((l) => /\bdept\b|\bdepartment\b/iu.test(l)) ??
      nonContact[1] ??
      "";
    let department = deptCandidate.trim();
    let organization = "";

    if (department.includes(",")) {
      const [deptPart, ...orgParts] = department.split(",");
      department = deptPart.trim();
      organization = orgParts.join(",").trim();
    }

    const cityLine =
      nonContact.find((l) =>
        /\bcity\b|\bcountry\b|\bphilippines\b|\busa\b|\buk\b|\bcanada\b/iu.test(
          l,
        ),
      ) ??
      nonContact[nonContact.length - 1] ??
      "";

    if (!organization) {
      organization =
        nonContact.find(
          (l) => l !== deptCandidate && l !== cityLine && !containsContact(l),
        ) ?? department;
    }

    for (let i = 0; i < names.length; i += 1) {
      authors.push({
        name: names[i],
        department,
        organization: organization || department,
        cityCountry: cityLine || "",
        contact: contacts[i] ?? contacts[0] ?? "",
      });
    }
  }

  return authors.slice(0, 6);
}

function toCleanLine(value: unknown): string {
  if (typeof value !== "string") return "";
  return value.replace(/\s+/gu, " ").trim();
}

function normalizeAuthorEntries(entries: AuthorEntry[]): AuthorEntry[] {
  return entries
    .map((a) => ({
      name: toCleanLine(a.name),
      department: toCleanLine(a.department),
      organization: toCleanLine(a.organization),
      cityCountry: toCleanLine(a.cityCountry),
      contact: toCleanLine(a.contact),
    }))
    .filter((a) => a.name.length > 0)
    .slice(0, 6);
}

function extractJsonObject(text: string): string | null {
  const trimmed = text.trim();
  if (!trimmed) return null;

  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/iu);
  if (fenced && fenced[1]) return fenced[1].trim();

  const start = trimmed.indexOf("{");
  const end = trimmed.lastIndexOf("}");
  if (start >= 0 && end > start) return trimmed.slice(start, end + 1);

  return null;
}

function coerceAuthorEntries(raw: unknown): AuthorEntry[] {
  if (!Array.isArray(raw)) return [];
  const entries: AuthorEntry[] = [];

  for (const row of raw) {
    if (!row || typeof row !== "object") continue;
    const obj = row as Record<string, unknown>;
    entries.push({
      name: toCleanLine(obj.name),
      department: toCleanLine(obj.department),
      organization: toCleanLine(obj.organization),
      cityCountry: toCleanLine(obj.cityCountry ?? obj.city_country ?? obj.city),
      contact: toCleanLine(
        obj.contact ?? obj.emailOrOrcid ?? obj.email ?? obj.orcid,
      ),
    });
  }

  return normalizeAuthorEntries(entries);
}

async function tryExtractAuthorsWithAi(
  lines: string[],
): Promise<AuthorEntry[] | null> {
  if (!AI_ASSIST_ENABLED || lines.length === 0) return null;

  const promptLines = lines
    .map((line, idx) => `${idx + 1}. ${line}`)
    .join("\n");
  const model = String(
    import.meta.env.VITE_POLLINATIONS_AUTHOR_MODEL ||
      import.meta.env.VITE_POLLINATIONS_MODEL ||
      "",
  ).trim();

  try {
    const aiResult = await requestPollinations({
      model: model || undefined,
      temperature: 0,
      maxTokens: 900,
      messages: [
        {
          role: "system",
          content:
            "Extract manuscript author metadata. Reply with strict JSON only and no extra text.",
        },
        {
          role: "user",
          content: [
            "Use the front-matter lines below (title + author region, before abstract).",
            "Return this exact JSON shape:",
            '{"authors":[{"name":"","department":"","organization":"","cityCountry":"","contact":""}]}',
            "Rules:",
            "- Keep author order from the manuscript.",
            "- Maximum 6 authors.",
            "- department and organization must be separate fields.",
            "- contact must be email or ORCID for each author when available.",
            "",
            "Front-matter lines:",
            promptLines,
          ].join("\n"),
        },
      ],
    });

    const jsonText = extractJsonObject(aiResult.reply);
    if (!jsonText) return null;

    const parsed = JSON.parse(jsonText) as { authors?: unknown };
    const aiAuthors = coerceAuthorEntries(parsed.authors);
    return aiAuthors.length > 0 ? aiAuthors : null;
  } catch (error) {
    void error;
    return null;
  }
}

function findPubformAuthorParagraphIndexes(
  templateParas: Element[],
  wNs: string,
): { top: number; fifth: number; sixth: number } {
  let top = -1;
  let fifth = -1;
  let sixth = -1;

  for (let i = 0; i < templateParas.length; i += 1) {
    const text = getParagraphText(templateParas[i], wNs);
    if (top < 0 && /line\s*1:\s*1st\s*given\s*name\s*surname/iu.test(text)) {
      top = i;
      continue;
    }
    if (fifth < 0 && /line\s*1:\s*5th\s*given\s*name\s*surname/iu.test(text)) {
      fifth = i;
      continue;
    }
    if (sixth < 0 && /line\s*1:\s*6th\s*given\s*name\s*surname/iu.test(text)) {
      sixth = i;
      continue;
    }
  }

  return { top, fifth, sixth };
}

function getParagraphSectPr(p: Element, wNs: string): Element | null {
  const pPr = getChild(p, wNs, "pPr");
  if (!pPr) return null;
  return getChild(pPr, wNs, "sectPr");
}

function hasParagraphNumbering(p: Element, wNs: string): boolean {
  const pPr = getChild(p, wNs, "pPr");
  if (!pPr) return false;
  return !!getChild(pPr, wNs, "numPr");
}

function getParagraphStyleId(p: Element, wNs: string): string | null {
  const pPr = getChild(p, wNs, "pPr");
  if (!pPr) return null;
  const pStyle = getChild(pPr, wNs, "pStyle");
  if (!pStyle) return null;
  const styleId = wAttr(pStyle, wNs, "val").trim();
  return styleId || null;
}

function collectNumberedParagraphStyleIds(stylesDoc: Document): Set<string> {
  const stylesWNs = resolveWNs(stylesDoc);
  const set = new Set<string>();
  const styleNodes = Array.from(
    stylesDoc.getElementsByTagNameNS(stylesWNs, "style"),
  );
  for (const styleNode of styleNodes) {
    if (wAttr(styleNode, stylesWNs, "type") !== "paragraph") continue;
    const styleId = wAttr(styleNode, stylesWNs, "styleId").trim();
    if (!styleId) continue;

    const pPr = getChild(styleNode, stylesWNs, "pPr");
    const numPr = pPr ? getChild(pPr, stylesWNs, "numPr") : null;
    if (numPr) set.add(styleId);
  }
  return set;
}

function paragraphUsesNumbering(
  p: Element,
  wNs: string,
  numberedStyleIds: Set<string>,
): boolean {
  if (hasParagraphNumbering(p, wNs)) return true;
  const styleId = getParagraphStyleId(p, wNs);
  return !!styleId && numberedStyleIds.has(styleId);
}

function hasLeadingListMarker(text: string): boolean {
  const firstLine = text.split(/\r?\n/u)[0]?.trimStart() ?? "";
  return (
    /^(?:[IVXLCDM]+)[.)]\s+/iu.test(firstLine) ||
    /^(?:[A-Z])[.)]\s+/u.test(firstLine) ||
    /^(?:\d+)(?:\.\d+)*[.)]\s+/u.test(firstLine)
  );
}

function disableParagraphNumbering(p: Element, wNs: string) {
  const pPr = ensurePPr(p, wNs);
  removeChildren(pPr, wNs, "numPr");
  const numPr = ensureChild(pPr, wNs, "numPr");
  const numId = ensureChild(numPr, wNs, "numId");
  setWAttr(numId, wNs, "val", "0");
  removeChildren(numPr, wNs, "ilvl");
}

function stripLeadingListMarkers(text: string): string {
  let out = text.trimStart();
  const markerPatterns = [
    /^(?:[IVXLCDM]+)[.)]\s+/iu,
    /^(?:[A-Z])[.)]\s+/u,
    /^(?:\d+)(?:\.\d+)*[.)]\s+/u,
  ];

  for (let pass = 0; pass < 4; pass += 1) {
    let changed = false;
    for (const pattern of markerPatterns) {
      const next = out.replace(pattern, "").trimStart();
      if (next !== out) {
        out = next;
        changed = true;
        break;
      }
    }
    if (!changed) break;
  }

  return out;
}

function normalizeTemplateNumberedText(
  templateParagraph: Element,
  wNs: string,
  text: string,
): string {
  if (!hasParagraphNumbering(templateParagraph, wNs)) return text;
  const lines = text.split(/\r?\n/u);
  if (lines.length === 0) return text;
  lines[0] = stripLeadingListMarkers(lines[0]);
  return lines.join("\n");
}

function findFirstSectPrParagraphIndex(
  templateParas: Element[],
  wNs: string,
  startIndex: number,
  endIndex: number,
): number {
  const start = Math.max(0, startIndex);
  const end = Math.min(templateParas.length - 1, endIndex);
  for (let i = start; i <= end; i += 1) {
    if (getParagraphSectPr(templateParas[i], wNs)) return i;
  }
  return -1;
}

function setSectionColumnsOnParagraph(
  p: Element,
  wNs: string,
  numColumns: number,
  columnSpace = "10.80pt",
) {
  const sectPr = getParagraphSectPr(p, wNs);
  if (!sectPr) return;
  const cols = ensureChild(sectPr, wNs, "cols");
  setWAttr(cols, wNs, "num", String(Math.max(1, numColumns)));
  setWAttr(cols, wNs, "space", columnSpace);
}

function findFirstParagraphIndexByTextMatch(
  templateParas: Element[],
  wNs: string,
  matcher: RegExp,
  fromIndex = 0,
): number {
  for (let i = Math.max(0, fromIndex); i < templateParas.length; i += 1) {
    const text = normalizeText(getParagraphText(templateParas[i], wNs));
    if (matcher.test(text)) return i;
  }
  return -1;
}

function removeInterveningBodyParagraphs(
  body: Element,
  templateParas: Element[],
  wNs: string,
  startExclusive: number,
  endExclusive: number,
  keepFirstSectionBreak = true,
) {
  if (endExclusive <= startExclusive + 1) return;

  let keptSectionBreak = false;
  for (let i = startExclusive + 1; i < endExclusive; i += 1) {
    const p = templateParas[i];
    if (!p || p.parentNode !== body) continue;

    const hasSectPr = !!getParagraphSectPr(p, wNs);
    if (hasSectPr && keepFirstSectionBreak && !keptSectionBreak) {
      keptSectionBreak = true;
      clearParagraphContent(p, wNs);
      writeParagraphLayout(p, wNs, "center", 1.0, 0);
      continue;
    }

    body.removeChild(p);
  }
}

function tightenPreAbstractGap(
  body: Element,
  templateParas: Element[],
  wNs: string,
  startExclusive: number,
  abstractIndex: number,
) {
  if (abstractIndex <= startExclusive + 1) return;
  for (let i = startExclusive + 1; i < abstractIndex; i += 1) {
    const p = templateParas[i];
    if (!p || p.parentNode !== body) continue;
    const hasSectPr = !!getParagraphSectPr(p, wNs);
    const text = normalizeText(getParagraphText(p, wNs));
    if (hasSectPr) {
      clearParagraphContent(p, wNs);
      writeParagraphLayout(p, wNs, "center", 1.0, 0);
      continue;
    }
    if (text === "") {
      body.removeChild(p);
    }
  }
}

function toStyledAuthorLines(author: AuthorEntry): StyledLine[] {
  return [
    { text: author.name || "", italic: false },
    { text: author.department || "", italic: true },
    { text: author.organization || "", italic: true },
    { text: author.cityCountry || "", italic: false },
    { text: author.contact || "", italic: false },
  ];
}

function writeAuthorParagraphContent(
  p: Element,
  wNs: string,
  authors: AuthorEntry[],
) {
  const lines: StyledLine[] = [];
  authors.forEach((author, index) => {
    if (index > 0) lines.push({ text: "", italic: false });
    lines.push(...toStyledAuthorLines(author));
  });

  if (lines.length === 0) {
    setParagraphText(p, wNs, "");
  } else {
    setParagraphStyledLines(p, wNs, lines, "Times New Roman", 9);
  }
  writeParagraphLayout(p, wNs, "center", 1.0, 0);
}

function ptsToHalfPts(pts: number): number {
  return Math.round(pts * 2);
}

function linesToTwips(lines: number): number {
  return Math.round(lines * 240);
}

function ptToTwips(pt: number): number {
  return Math.round(pt * 20);
}

function cmToTwips(cm: number): number {
  return Math.round((cm / 2.54) * 1440);
}

interface ParagraphIndentOptions {
  firstLineTwips?: number;
  hangingTwips?: number;
  leftTwips?: number;
  rightTwips?: number;
}

function writeParagraphLayout(
  p: Element,
  wNs: string,
  alignment: "left" | "center" | "right" | "both",
  lineSpacing: number,
  afterTwips: number,
  beforeTwips = 0,
) {
  const pPr = ensurePPr(p, wNs);
  removeChildren(pPr, wNs, "jc");
  removeChildren(pPr, wNs, "spacing");

  const jc = ensureChild(pPr, wNs, "jc");
  setWAttr(jc, wNs, "val", alignment);

  const sp = ensureChild(pPr, wNs, "spacing");
  setWAttr(sp, wNs, "before", String(beforeTwips));
  setWAttr(sp, wNs, "after", String(afterTwips));
  setWAttr(sp, wNs, "line", String(linesToTwips(lineSpacing)));
  setWAttr(sp, wNs, "lineRule", "auto");
  setWAttr(sp, wNs, "beforeAutospacing", "0");
  setWAttr(sp, wNs, "afterAutospacing", "0");
}

function writeParagraphIndent(
  p: Element,
  wNs: string,
  indent?: ParagraphIndentOptions,
) {
  const pPr = ensurePPr(p, wNs);
  removeChildren(pPr, wNs, "ind");
  if (!indent) return;

  const hasAny =
    typeof indent.firstLineTwips === "number" ||
    typeof indent.hangingTwips === "number" ||
    typeof indent.leftTwips === "number" ||
    typeof indent.rightTwips === "number";
  if (!hasAny) return;

  const ind = ensureChild(pPr, wNs, "ind");

  if (typeof indent.leftTwips === "number") {
    setWAttr(
      ind,
      wNs,
      "left",
      String(Math.max(0, Math.round(indent.leftTwips))),
    );
  }
  if (typeof indent.rightTwips === "number") {
    setWAttr(
      ind,
      wNs,
      "right",
      String(Math.max(0, Math.round(indent.rightTwips))),
    );
  }
  if (typeof indent.firstLineTwips === "number") {
    setWAttr(
      ind,
      wNs,
      "firstLine",
      String(Math.max(0, Math.round(indent.firstLineTwips))),
    );
  }
  if (typeof indent.hangingTwips === "number") {
    setWAttr(
      ind,
      wNs,
      "hanging",
      String(Math.max(0, Math.round(indent.hangingTwips))),
    );
  }
}

function writeRunFormatting(
  p: Element,
  wNs: string,
  fontFamily: string,
  fontPt: number,
  bold: boolean,
  italic: boolean,
) {
  const runNodes = Array.from(p.getElementsByTagNameNS(wNs, "r"));
  const halfPt = String(ptsToHalfPts(fontPt));

  for (const run of runNodes) {
    const hasText = run.getElementsByTagNameNS(wNs, "t").length > 0;
    if (!hasText) continue;

    const rPr = ensureRPr(run, wNs);

    removeChildren(rPr, wNs, "rFonts");
    const rFonts = ensureChild(rPr, wNs, "rFonts");
    setWAttr(rFonts, wNs, "ascii", fontFamily);
    setWAttr(rFonts, wNs, "hAnsi", fontFamily);
    setWAttr(rFonts, wNs, "eastAsia", fontFamily);
    setWAttr(rFonts, wNs, "cs", fontFamily);

    removeChildren(rPr, wNs, "sz");
    removeChildren(rPr, wNs, "szCs");
    const sz = ensureChild(rPr, wNs, "sz");
    setWAttr(sz, wNs, "val", halfPt);
    const szCs = ensureChild(rPr, wNs, "szCs");
    setWAttr(szCs, wNs, "val", halfPt);

    removeChildren(rPr, wNs, "b");
    removeChildren(rPr, wNs, "bCs");
    if (bold) {
      rPr.appendChild(wElem(p.ownerDocument!, wNs, "b"));
      rPr.appendChild(wElem(p.ownerDocument!, wNs, "bCs"));
    }

    removeChildren(rPr, wNs, "i");
    removeChildren(rPr, wNs, "iCs");
    if (italic) {
      rPr.appendChild(wElem(p.ownerDocument!, wNs, "i"));
      rPr.appendChild(wElem(p.ownerDocument!, wNs, "iCs"));
    }
  }
}

function applyTextStyle(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  afterTwips = 0,
) {
  writeParagraphLayout(p, wNs, style.alignment, style.lineSpacing, afterTwips);
  writeRunFormatting(
    p,
    wNs,
    style.fontFamily,
    style.fontSize,
    !!style.bold,
    !!style.italic,
  );
}

function cloneDeep<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function mergeDeep<T>(defaults: T, source: unknown): T {
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    return cloneDeep(defaults);
  }

  const out: any = { ...(defaults as any) };
  const src = source as Record<string, unknown>;
  for (const key of Object.keys(defaults as Record<string, unknown>)) {
    const baseValue = (defaults as Record<string, unknown>)[key];
    const sourceValue = src[key];
    if (
      baseValue &&
      typeof baseValue === "object" &&
      !Array.isArray(baseValue)
    ) {
      out[key] = mergeDeep(baseValue, sourceValue);
      continue;
    }
    out[key] = sourceValue === undefined ? baseValue : sourceValue;
  }
  return out as T;
}

function resolveConferenceStyleConfig(
  styleConfig?: ConferenceFormattingConfig,
): ConferenceFormattingConfig {
  return mergeDeep(DEFAULT_CONFERENCE_FORMATTING_CONFIG, styleConfig);
}

function isPublicationHeading1(text: string): boolean {
  const normalized = normalizeText(text);
  if (!normalized) return false;
  // Top-level publication headings use Roman numerals (I, II, III, IV, V...).
  // Restricting to I/V/X avoids colliding with lettered sub-headings like "C. ...".
  if (/^[IVX]+\.\s+[A-Za-z]/iu.test(normalized)) return true;
  return /^REFERENCES$/iu.test(normalized);
}

function isPublicationHeading2(text: string): boolean {
  const normalized = normalizeText(text);
  if (!/^[A-Z]\.\s+\S/u.test(normalized)) return false;
  if (/[,;:]/u.test(normalized)) return false;
  if (/\bhttps?:\/\//iu.test(normalized)) return false;
  if (/@/u.test(normalized)) return false;
  if (/\bdoi\b/iu.test(normalized)) return false;
  if (normalized.length > 100) return false;
  return true;
}

function isPublicationReferencesHeading(text: string): boolean {
  const normalized = normalizeText(text)
    .replace(/^[IVXLCDM]+\.\s*/iu, "")
    .trim();
  return normalized.toUpperCase() === "REFERENCES";
}

function isLikelyPublicationCaption(text: string): boolean {
  return /^figure\s+\d+/iu.test(text) || /^table\s+\d+/iu.test(text);
}

function toHeading2TitleCaseWord(token: string): string {
  if (!/[A-Za-z]/u.test(token)) return token;
  if (/\d/u.test(token)) return token;
  if (/^[A-Z]{2,4}$/u.test(token)) return token;

  const parts = token.split(/([\-\/])/u);
  return parts
    .map((part) => {
      if (part === "-" || part === "/") return part;
      if (!/[A-Za-z]/u.test(part)) return part;
      return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
    })
    .join("");
}

function normalizePublicationHeading2(text: string): string {
  const normalized = normalizeText(text);
  const match = normalized.match(/^([A-Z]\.)\s+(.+)$/u);
  if (!match) return normalized;

  const marker = match[1];
  const content = match[2]
    .split(/\s+/u)
    .map((token) => toHeading2TitleCaseWord(token))
    .join(" ");

  return `${marker} ${content}`;
}

function normalizeIeeeReferenceMarker(text: string): string {
  const normalized = normalizeText(text);
  let m = normalized.match(/^\[(\d{1,3})\]\s*(.*)$/u);
  if (m) return `[${m[1]}] ${m[2]}`.trimEnd();
  m = normalized.match(/^\((\d{1,3})\)\s*(.*)$/u);
  if (m) return `[${m[1]}] ${m[2]}`.trimEnd();
  m = normalized.match(/^(\d{1,3})(?:[.)])?\s+(.*)$/u);
  if (m) return `[${m[1]}] ${m[2]}`.trimEnd();
  return text;
}

function applyPublicationRuleBasedFormatting(
  doc: Document,
  styleConfig: PublicationFormattingConfig,
) {
  const wNs = resolveWNs(doc);
  const body = getBody(doc, wNs);
  if (!body) return;

  const paragraphs = Array.from(body.getElementsByTagNameNS(wNs, "p"));
  const abstractIndex = paragraphs.findIndex((p) =>
    /^abstract\b/iu.test(normalizeText(getParagraphText(p, wNs))),
  );
  const startIndex = abstractIndex >= 0 ? abstractIndex + 1 : 0;

  let inReferences = false;
  let referenceIndex = 0;

  for (let i = startIndex; i < paragraphs.length; i += 1) {
    const p = paragraphs[i];
    const rawText = getParagraphText(p, wNs);
    const text = normalizeText(rawText);
    if (!text) continue;

    if (/^keywords\b/iu.test(text) || /^keywords[-—]/iu.test(text)) {
      continue;
    }

    if (inReferences) {
      let ieeeText = text;
      if (styleConfig.references.ieeeStyle) {
        const normalized = normalizeText(text);
        const normalizedWithMarker = normalizeIeeeReferenceMarker(normalized);
        const markerMatch = normalizedWithMarker.match(/^\[(\d{1,3})\]\s+/u);
        if (markerMatch) {
          const markerNum = Number(markerMatch[1]);
          if (!Number.isNaN(markerNum) && markerNum > 0) {
            referenceIndex = Math.max(referenceIndex, markerNum);
          } else {
            referenceIndex += 1;
          }
          ieeeText = normalizedWithMarker;
        } else {
          referenceIndex += 1;
          ieeeText = `[${referenceIndex}] ${normalized}`;
        }
      }
      if (ieeeText !== text) {
        setParagraphText(p, wNs, ieeeText);
      }
      applyTextStyle(p, wNs, styleConfig.references, 0);
      writeParagraphIndent(p, wNs, {
        hangingTwips: cmToTwips(styleConfig.references.hangingIndentCm),
      });
      continue;
    }

    if (isPublicationHeading1(text)) {
      const heading = styleConfig.heading1.uppercase
        ? text.toUpperCase()
        : text;
      if (heading !== text) {
        setParagraphText(p, wNs, heading);
      }
      applyTextStyle(p, wNs, styleConfig.heading1, 0);
      writeParagraphIndent(p, wNs);

      inReferences = isPublicationReferencesHeading(heading);
      continue;
    }

    if (!inReferences && isPublicationHeading2(text)) {
      const heading2 = styleConfig.heading2.titleCase
        ? normalizePublicationHeading2(text)
        : text;
      if (heading2 !== text) {
        setParagraphText(p, wNs, heading2);
      }
      applyTextStyle(p, wNs, styleConfig.heading2, 0);
      writeParagraphIndent(p, wNs);
      continue;
    }

    if (isLikelyPublicationCaption(text)) {
      continue;
    }

    applyTextStyle(
      p,
      wNs,
      styleConfig.body,
      ptToTwips(styleConfig.body.spacingAfterPt),
    );
    writeParagraphIndent(p, wNs, {
      firstLineTwips: cmToTwips(styleConfig.body.firstLineIndentCm),
    });
  }
}

function enforceSingleColumnNoHeaderFooter(body: Element, wNs: string) {
  const sectPrNodes = Array.from(body.getElementsByTagNameNS(wNs, "sectPr"));
  for (const sectPr of sectPrNodes) {
    removeChildren(sectPr, wNs, "headerReference");
    removeChildren(sectPr, wNs, "footerReference");
    const cols = ensureChild(sectPr, wNs, "cols");
    setWAttr(cols, wNs, "num", "1");
    setWAttr(cols, wNs, "space", "708");
  }
}

function isLikelyAcmHeading(text: string): boolean {
  if (text.length < 3 || text.length > 90) return false;
  if (/[.?!:]$/u.test(text)) return false;
  if (/^table\s+\d+/iu.test(text) || /^figure\s+\d+/iu.test(text)) return false;
  if (
    /^(introduction|related work|methodology|methods|results|discussion|conclusion|references)$/iu.test(
      text,
    )
  ) {
    return true;
  }
  if (/^[A-Z][A-Za-z0-9 ,/&\-()]+$/u.test(text)) return true;
  if (/^\d+(\.\d+)*\s+[A-Za-z]/u.test(text)) return true;
  return false;
}

function isLikelyAuthorLine(text: string): boolean {
  return (
    text.includes("@") ||
    /department|institute|university|college|school|center|laboratory/iu.test(
      text,
    ) ||
    /^[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+)*(?:,\s*[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+)*)*(?:\s+and\s+[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+)*)?$/u.test(
      text,
    )
  );
}

function applyAcmRuleBasedFormatting(
  doc: Document,
  styleConfig: AcmFormattingConfig,
) {
  const wNs = resolveWNs(doc);
  const body = getBody(doc, wNs);
  if (!body) return;

  enforceSingleColumnNoHeaderFooter(body, wNs);

  const paragraphs = Array.from(body.getElementsByTagNameNS(wNs, "p"));
  let nonEmptyCount = 0;
  let inFrontMatter = true;
  let inReferences = false;

  for (const p of paragraphs) {
    const text = normalizeText(getParagraphText(p, wNs));

    if (text === "") {
      applyTextStyle(p, wNs, styleConfig.body, 0);
      continue;
    }

    nonEmptyCount += 1;

    if (nonEmptyCount === 1) {
      applyTextStyle(p, wNs, styleConfig.title, 120);
      continue;
    }
    if (nonEmptyCount === 2) {
      applyTextStyle(p, wNs, styleConfig.subtitle, 80);
      continue;
    }
    if (nonEmptyCount === 3) {
      applyTextStyle(p, wNs, styleConfig.subtitle, 120);
      continue;
    }

    if (/^abstract\b/iu.test(text) || /^abstract[-—]/iu.test(text)) {
      inFrontMatter = false;
      applyTextStyle(p, wNs, styleConfig.body, 120);
      continue;
    }

    if (
      /^ccs concepts/iu.test(text) ||
      /^additional keywords/iu.test(text) ||
      /^acm reference format/iu.test(text)
    ) {
      applyTextStyle(p, wNs, styleConfig.body, 100);
      continue;
    }

    if (inFrontMatter && nonEmptyCount <= 12 && isLikelyAuthorLine(text)) {
      applyTextStyle(p, wNs, styleConfig.author, 60);
      continue;
    }

    if (/^references$/iu.test(text)) {
      inFrontMatter = false;
      inReferences = true;
      applyTextStyle(p, wNs, styleConfig.heading, 90);
      continue;
    }

    if (inReferences) {
      applyTextStyle(p, wNs, styleConfig.references, 90);
      continue;
    }

    if (isLikelyAcmHeading(text)) {
      inFrontMatter = false;
      applyTextStyle(p, wNs, styleConfig.heading, 90);
      continue;
    }

    inFrontMatter = false;
    applyTextStyle(p, wNs, styleConfig.body, 90);
  }
}

async function parseDocFromZip(zip: any, path: string): Promise<Document> {
  const entry = zip.file(path);
  if (!entry) throw new Error(`Missing ${path} in DOCX package.`);
  const xmlStr = await entry.async("string");
  return new DOMParser().parseFromString(xmlStr, "application/xml");
}

function serializeDoc(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

async function applyPublicationTemplate(
  inputZip: any,
  pubformZip: any,
  styleConfig: PublicationFormattingConfig,
): Promise<void> {
  const inputDoc = await parseDocFromZip(inputZip, "word/document.xml");
  const templateDoc = await parseDocFromZip(pubformZip, "word/document.xml");
  const templateStylesDoc = await parseDocFromZip(
    pubformZip,
    "word/styles.xml",
  );

  const inputWNs = resolveWNs(inputDoc);
  const templateWNs = resolveWNs(templateDoc);
  const numberedStyleIds = collectNumberedParagraphStyleIds(templateStylesDoc);

  const inputBody = getBody(inputDoc, inputWNs);
  const templateBody = getBody(templateDoc, templateWNs);
  if (!inputBody || !templateBody) throw new Error("Invalid document body.");

  const inputParas = Array.from(
    inputBody.getElementsByTagNameNS(inputWNs, "p"),
  );
  const templateParas = Array.from(
    templateBody.getElementsByTagNameNS(templateWNs, "p"),
  );
  const authorParaIndexes = findPubformAuthorParagraphIndexes(
    templateParas,
    templateWNs,
  );

  const copyCount = Math.min(inputParas.length, templateParas.length);
  for (let i = 0; i < copyCount; i++) {
    const sourceText = getParagraphText(inputParas[i], inputWNs);
    const templateParagraph = templateParas[i];
    const usesNumbering = paragraphUsesNumbering(
      templateParagraph,
      templateWNs,
      numberedStyleIds,
    );

    let text = sourceText;
    if (usesNumbering) {
      if (hasLeadingListMarker(sourceText)) {
        // Keep source numbering marker (e.g. C.) and suppress template auto-numbering.
        disableParagraphNumbering(templateParagraph, templateWNs);
      } else {
        text = normalizeTemplateNumberedText(
          templateParagraph,
          templateWNs,
          sourceText,
        );
      }
    }
    setParagraphText(templateParas[i], templateWNs, text);
  }

  // Rebuild author front-matter as 5-line blocks per author
  // (name, department, organization, city/country, email/ORCID).
  const frontMatterLines = normalizeInputLinesUntilAbstract(inputDoc, inputWNs);
  const parsedAuthors = parseAuthorEntriesFromFrontMatter(frontMatterLines);
  const aiAuthors = await tryExtractAuthorsWithAi(frontMatterLines);
  const authors = normalizeAuthorEntries(aiAuthors ?? parsedAuthors);
  const titleText = frontMatterLines[0] ?? "";
  if (titleText && templateParas[0]) {
    setParagraphText(templateParas[0], templateWNs, titleText);
  }

  // Remove copied front-matter noise between title and the actual author placeholders.
  if (authorParaIndexes.top > 1) {
    for (let i = 1; i < authorParaIndexes.top; i += 1) {
      clearParagraphContent(templateParas[i], templateWNs);
      writeParagraphLayout(templateParas[i], templateWNs, "center", 1.0, 0);
    }
  }

  if (authorParaIndexes.top >= 0) {
    const topAuthors = authors.slice(0, 4);
    if (topAuthors.length > 1) {
      setParagraphAuthorColumns(
        templateParas[authorParaIndexes.top],
        templateWNs,
        topAuthors,
      );
    } else {
      writeAuthorParagraphContent(
        templateParas[authorParaIndexes.top],
        templateWNs,
        topAuthors,
      );
    }
  }

  // Remove empty publication-author gap paragraphs between top row and lower row.
  if (
    authorParaIndexes.top >= 0 &&
    authorParaIndexes.fifth > authorParaIndexes.top + 1
  ) {
    removeInterveningBodyParagraphs(
      templateBody,
      templateParas,
      templateWNs,
      authorParaIndexes.top,
      authorParaIndexes.fifth,
      true,
    );
  }

  // Lower-author region should be two centered columns (5th and 6th) below the top row.
  if (authorParaIndexes.fifth >= 0) {
    const hasTwoLowerAuthors = !!(authors[4] && authors[5]);
    const lowerSectionBreakIndex = findFirstSectPrParagraphIndex(
      templateParas,
      templateWNs,
      Math.max(0, authorParaIndexes.top + 1),
      authorParaIndexes.fifth,
    );
    if (lowerSectionBreakIndex >= 0) {
      setSectionColumnsOnParagraph(
        templateParas[lowerSectionBreakIndex],
        templateWNs,
        hasTwoLowerAuthors ? 4 : 2,
        "10.80pt",
      );
      clearParagraphContent(templateParas[lowerSectionBreakIndex], templateWNs);
      writeParagraphLayout(
        templateParas[lowerSectionBreakIndex],
        templateWNs,
        "center",
        1.0,
        0,
      );
    }

    if (hasTwoLowerAuthors) {
      setParagraphTwoAuthorColumns(
        templateParas[authorParaIndexes.fifth],
        templateWNs,
        authors[4],
        authors[5],
        "Times New Roman",
        9,
        1,
      );
    } else {
      writeAuthorParagraphContent(
        templateParas[authorParaIndexes.fifth],
        templateWNs,
        authors[4] ? [authors[4]] : [],
      );
    }
  }

  if (authorParaIndexes.sixth >= 0) {
    if (authors[4] && authors[5]) {
      clearParagraphContent(
        templateParas[authorParaIndexes.sixth],
        templateWNs,
      );
      writeParagraphLayout(
        templateParas[authorParaIndexes.sixth],
        templateWNs,
        "center",
        1.0,
        0,
      );
    } else {
      writeAuthorParagraphContent(
        templateParas[authorParaIndexes.sixth],
        templateWNs,
        authors[5] ? [authors[5]] : [],
      );
    }
  }

  const abstractParagraphIndex = findFirstParagraphIndexByTextMatch(
    templateParas,
    templateWNs,
    /^abstract\b/iu,
    Math.max(
      authorParaIndexes.sixth,
      authorParaIndexes.fifth,
      authorParaIndexes.top,
    ) + 1,
  );
  const cleanupStart =
    authorParaIndexes.sixth >= 0
      ? authorParaIndexes.sixth + 1
      : authorParaIndexes.fifth + 1;
  if (abstractParagraphIndex > cleanupStart && cleanupStart >= 0) {
    tightenPreAbstractGap(
      templateBody,
      templateParas,
      templateWNs,
      cleanupStart - 1,
      abstractParagraphIndex,
    );
  }

  applyPublicationRuleBasedFormatting(templateDoc, styleConfig);
  pubformZip.file("word/document.xml", serializeDoc(templateDoc));
}

function requireJsZip() {
  const JSZip = (window as any).JSZip;
  if (!JSZip) throw new Error("JSZip not loaded");
  return JSZip;
}

export async function formatDocxPublication(
  arrayBuffer: ArrayBuffer,
  styleConfig: PublicationFormattingConfig,
): Promise<Blob> {
  const JSZip = requireJsZip();
  const [inputZip, pubformResponse] = await Promise.all([
    JSZip.loadAsync(arrayBuffer),
    fetch(PUBFORM_SOURCE_FILE),
  ]);
  if (!pubformResponse.ok) {
    throw new Error("Unable to load publication format source file.");
  }
  const pubformZip = await JSZip.loadAsync(await pubformResponse.arrayBuffer());
  await applyPublicationTemplate(inputZip, pubformZip, styleConfig);
  return pubformZip.generateAsync({ type: "blob" }) as Promise<Blob>;
}

export async function formatDocxAcm(
  arrayBuffer: ArrayBuffer,
  styleConfig: AcmFormattingConfig,
): Promise<Blob> {
  const JSZip = requireJsZip();
  const [targetZip, acmResponse] = await Promise.all([
    JSZip.loadAsync(arrayBuffer),
    fetch(ACM_SOURCE_FILE),
  ]);
  if (!acmResponse.ok) {
    throw new Error("Unable to load ACM conference source file.");
  }

  // Use the local ACM file as the rules reference (no raw XML copy to output).
  await acmResponse.arrayBuffer();

  const targetDoc = await parseDocFromZip(targetZip, "word/document.xml");
  applyAcmRuleBasedFormatting(targetDoc, styleConfig);
  targetZip.file("word/document.xml", serializeDoc(targetDoc));

  return targetZip.generateAsync({ type: "blob" }) as Promise<Blob>;
}

export async function formatDocxConference(
  arrayBuffer: ArrayBuffer,
  options: ConferenceFormatOptions,
): Promise<Blob> {
  const resolvedConfig = resolveConferenceStyleConfig(options.styleConfig);
  if (options.format === "pubform") {
    return formatDocxPublication(arrayBuffer, resolvedConfig.pubform);
  }
  return formatDocxAcm(arrayBuffer, resolvedConfig.acm);
}
