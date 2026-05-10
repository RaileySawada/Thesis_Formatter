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
  aiAssist?: boolean;
}

interface AuthorEntry {
  name: string;
  department: string;
  organization: string;
  cityCountry: string;
  contact: string;
}

type AcmParagraphKind =
  | "ignore"
  | "heading1"
  | "heading2"
  | "heading3"
  | "body"
  | "tableCaption"
  | "figureCaption"
  | "footnote"
  | "equation"
  | "references";

type AcmParagraphHintMap = Map<number, AcmParagraphKind>;

const AI_ASSIST_ENABLED = isAiAssistEnabled(
  import.meta.env.VITE_ENABLE_AI_ASSIST,
);

const PUBFORM_SOURCE_FILE =
  "/conference/publication_formatting_guidelines.docx";
const FALLBACK_W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";

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

function setParagraphTextWithTabs(p: Element, wNs: string, text: string) {
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
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    if (lineIndex > 0) {
      run.appendChild(wElem(p.ownerDocument!, wNs, "br"));
    }

    const segments = lines[lineIndex].split("\t");
    for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
      if (segmentIndex > 0) {
        run.appendChild(wElem(p.ownerDocument!, wNs, "tab"));
      }
      const part = segments[segmentIndex];
      if (part === "") continue;
      const t = wElem(p.ownerDocument!, wNs, "t");
      if (/^\s/u.test(part) || /\s$/u.test(part) || /\s{2,}/u.test(part)) {
        t.setAttribute("xml:space", "preserve");
      }
      t.textContent = part;
      run.appendChild(t);
    }
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

function extractEmailOrOrcid(text: string): string {
  const email = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/iu)?.[0];
  if (email) return email.trim();
  const orcid = text.match(/\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b/iu)?.[0];
  return orcid ? orcid.trim() : "";
}

function stripContactFromAffiliation(text: string): string {
  return normalizeText(text)
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/giu, "")
    .replace(/\b\d{4}-\d{4}-\d{4}-\d{3}[\dX]\b/giu, "")
    .replace(/\s*,\s*$/u, "")
    .replace(/^\s*,\s*/u, "")
    .trim();
}

function looksLikeAcmSubtitleLine(text: string): boolean {
  const normalized = normalizeText(text);
  if (!normalized) return false;
  if (isLikelyAcmAffiliationLine(normalized)) return false;
  if (isLikelyAcmAuthorNameLine(normalized)) return false;
  if (/author/i.test(normalized) && /name/i.test(normalized)) return false;
  return true;
}

function getAcmAuthorExtractionLines(lines: string[]): string[] {
  if (lines.length <= 2) return lines;
  if (!looksLikeAcmSubtitleLine(lines[1])) return lines;
  return [lines[0], ...lines.slice(2)];
}

function parseAcmAuthorEntriesFromFrontMatter(lines: string[]): AuthorEntry[] {
  if (lines.length <= 1) return [];

  const authorLines = getAcmAuthorExtractionLines(lines).slice(1);
  const authors: AuthorEntry[] = [];
  let currentName = "";
  let affiliationParts: string[] = [];
  let contact = "";

  const flush = () => {
    if (!currentName.trim()) return;
    const affiliation = affiliationParts
      .map((part) => normalizeText(part))
      .filter(Boolean)
      .join(", ");
    authors.push({
      name: normalizeText(currentName),
      department: affiliation,
      organization: "",
      cityCountry: "",
      contact: contact.trim(),
    });
    currentName = "";
    affiliationParts = [];
    contact = "";
  };

  for (const line of authorLines) {
    const text = normalizeText(line);
    if (!text) continue;
    if (isAcmAbstractStart(text)) break;
    if (isAcmConceptsParagraph(text)) break;
    if (isAcmKeywordsParagraph(text)) break;
    if (isAcmReferenceFormatParagraph(text)) break;

    const lineIsAffiliation = isLikelyAcmAffiliationLine(text);
    if (!lineIsAffiliation && currentName && (affiliationParts.length > 0 || contact)) {
      flush();
    }

    if (!currentName) {
      currentName = text;
      continue;
    }

    const lineContact = extractEmailOrOrcid(text);
    if (lineContact && !contact) contact = lineContact;
    const affiliationText = stripContactFromAffiliation(text);
    if (affiliationText) affiliationParts.push(affiliationText);
  }

  flush();
  return normalizeAuthorEntries(authors);
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
  enabled = AI_ASSIST_ENABLED,
): Promise<AuthorEntry[] | null> {
  if (!enabled || lines.length === 0) return null;

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
      timeoutMs: 12000,
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

const ACM_PARAGRAPH_KINDS = new Set<AcmParagraphKind>([
  "ignore",
  "heading1",
  "heading2",
  "heading3",
  "body",
  "tableCaption",
  "figureCaption",
  "footnote",
  "equation",
  "references",
]);

function coerceAcmParagraphKind(value: unknown): AcmParagraphKind | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  return ACM_PARAGRAPH_KINDS.has(normalized as AcmParagraphKind)
    ? (normalized as AcmParagraphKind)
    : null;
}

function coerceAcmParagraphHints(raw: unknown): AcmParagraphHintMap {
  const hints: AcmParagraphHintMap = new Map();
  const rows = Array.isArray(raw)
    ? raw
    : raw && typeof raw === "object"
      ? (raw as { paragraphs?: unknown }).paragraphs
      : null;
  if (!Array.isArray(rows)) return hints;

  for (const row of rows) {
    if (!row || typeof row !== "object") continue;
    const obj = row as Record<string, unknown>;
    const index = Number(obj.i ?? obj.index ?? obj.paragraphIndex);
    const kind = coerceAcmParagraphKind(obj.type ?? obj.kind ?? obj.role);
    if (!Number.isInteger(index) || index < 0 || !kind) continue;
    hints.set(index, kind);
  }

  return hints;
}

async function tryClassifyAcmParagraphsWithAi(
  doc: Document,
  enabled = AI_ASSIST_ENABLED,
): Promise<AcmParagraphHintMap> {
  const hints: AcmParagraphHintMap = new Map();
  if (!enabled) return hints;

  const wNs = resolveWNs(doc);
  const body = getBody(doc, wNs);
  if (!body) return hints;

  const children = getTopLevelBodyChildren(body, wNs);
  const rows = children
    .map((child, index) => {
      if (child.localName !== "p") return null;
      const text = normalizeText(getParagraphText(child, wNs));
      if (!text) return null;
      return {
        index,
        text: text.length > 180 ? `${text.slice(0, 180)}...` : text,
      };
    })
    .filter((row): row is { index: number; text: string } => !!row)
    .slice(0, 220);

  if (rows.length === 0) return hints;

  const promptRows = rows.map((row) => `[${row.index}] ${row.text}`).join("\n");
  const model = String(
    import.meta.env.VITE_POLLINATIONS_ACM_MODEL ||
      import.meta.env.VITE_POLLINATIONS_MODEL ||
      "",
  ).trim();

  try {
    const aiResult = await requestPollinations({
      model: model || undefined,
      temperature: 0,
      maxTokens: 1800,
      timeoutMs: 15000,
      messages: [
        {
          role: "system",
          content:
            "Classify ACM paper paragraphs. Reply with strict JSON only and no prose.",
        },
        {
          role: "user",
          content: [
            "Classify these top-level DOCX paragraph indexes for ACM formatting.",
            "Ignore title/authors/abstract/CCS/keywords/ACM Reference Format/front-matter/permission footnote before the main paper body.",
            "Use these exact types only: ignore, heading1, heading2, heading3, body, tableCaption, figureCaption, footnote, equation, references.",
            "heading1 means top-level sections such as Introduction, Related Work, Methodology, Results, Discussion, Conclusion, Acknowledgments, or References.",
            "heading2 means subsection; heading3 means sub-subsection. The heading may be numbered or unnumbered.",
            "references means only the References heading, not each reference entry.",
            "Return only JSON in this shape:",
            '{"paragraphs":[{"i":0,"type":"ignore"}]}',
            "",
            "Paragraphs:",
            promptRows,
          ].join("\n"),
        },
      ],
    });

    const jsonText = extractJsonObject(aiResult.reply);
    if (!jsonText) return hints;
    return coerceAcmParagraphHints(JSON.parse(jsonText));
  } catch (error) {
    void error;
    return hints;
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
  options?: {
    lineRule?: "auto" | "exact" | "atLeast";
    lineTwips?: number;
  },
) {
  const pPr = ensurePPr(p, wNs);
  removeChildren(pPr, wNs, "jc");
  removeChildren(pPr, wNs, "spacing");

  const jc = ensureChild(pPr, wNs, "jc");
  setWAttr(jc, wNs, "val", alignment);

  const sp = ensureChild(pPr, wNs, "spacing");
  setWAttr(sp, wNs, "before", String(beforeTwips));
  setWAttr(sp, wNs, "after", String(afterTwips));
  setWAttr(
    sp,
    wNs,
    "line",
    String(Math.max(0, Math.round(options?.lineTwips ?? linesToTwips(lineSpacing)))),
  );
  setWAttr(sp, wNs, "lineRule", options?.lineRule ?? "auto");
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

function writeRunProperties(
  rPr: Element,
  wNs: string,
  fontFamily: string,
  fontPt: number,
  bold: boolean,
  italic: boolean,
  colorHex?: string,
) {
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
  if (bold) {
    rPr.appendChild(wElem(rPr.ownerDocument!, wNs, "b"));
    rPr.appendChild(wElem(rPr.ownerDocument!, wNs, "bCs"));
  } else {
    const b = wElem(rPr.ownerDocument!, wNs, "b");
    setWAttr(b, wNs, "val", "0");
    rPr.appendChild(b);
    const bCs = wElem(rPr.ownerDocument!, wNs, "bCs");
    setWAttr(bCs, wNs, "val", "0");
    rPr.appendChild(bCs);
  }

  removeChildren(rPr, wNs, "i");
  removeChildren(rPr, wNs, "iCs");
  if (italic) {
    rPr.appendChild(wElem(rPr.ownerDocument!, wNs, "i"));
    rPr.appendChild(wElem(rPr.ownerDocument!, wNs, "iCs"));
  } else {
    const i = wElem(rPr.ownerDocument!, wNs, "i");
    setWAttr(i, wNs, "val", "0");
    rPr.appendChild(i);
    const iCs = wElem(rPr.ownerDocument!, wNs, "iCs");
    setWAttr(iCs, wNs, "val", "0");
    rPr.appendChild(iCs);
  }

  if (typeof colorHex === "string" && colorHex.trim()) {
    removeChildren(rPr, wNs, "color");
    const color = ensureChild(rPr, wNs, "color");
    setWAttr(color, wNs, "val", colorHex.trim());
  }
}

function writeRunFormatting(
  p: Element,
  wNs: string,
  fontFamily: string,
  fontPt: number,
  bold: boolean,
  italic: boolean,
  colorHex?: string,
) {
  const runNodes = Array.from(p.getElementsByTagNameNS(wNs, "r"));

  const pPr = ensurePPr(p, wNs);
  const paragraphRPr = ensureChild(pPr, wNs, "rPr");
  writeRunProperties(
    paragraphRPr,
    wNs,
    fontFamily,
    fontPt,
    bold,
    italic,
    colorHex,
  );

  for (const run of runNodes) {
    const hasText = run.getElementsByTagNameNS(wNs, "t").length > 0;
    if (!hasText) continue;

    const rPr = ensureRPr(run, wNs);
    writeRunProperties(rPr, wNs, fontFamily, fontPt, bold, italic, colorHex);
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

interface ConferenceParagraphStyleOptions {
  alignment?: "left" | "center" | "right" | "both";
  beforePt?: number;
  afterPt?: number;
  firstLineIndentCm?: number;
  hangingIndentCm?: number;
  lineSpacingPt?: number;
  lineSpacingRule?: "auto" | "exact" | "atLeast";
  color?: string;
}

function applyConferenceParagraphLayout(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  options: ConferenceParagraphStyleOptions = {},
) {
  const alignment = options.alignment ?? style.alignment;
  const beforePt = options.beforePt ?? style.spacingBeforePt ?? 0;
  const afterPt = options.afterPt ?? style.spacingAfterPt ?? 0;
  const lineSpacingPt = options.lineSpacingPt ?? style.lineSpacingPt;
  const lineRule = options.lineSpacingRule ?? style.lineSpacingRule ?? "auto";
  const lineTwips =
    typeof lineSpacingPt === "number"
      ? ptToTwips(lineSpacingPt)
      : linesToTwips(style.lineSpacing);

  writeParagraphLayout(
    p,
    wNs,
    alignment,
    style.lineSpacing,
    ptToTwips(afterPt),
    ptToTwips(beforePt),
    {
      lineRule,
      lineTwips,
    },
  );

  const firstLineIndentCm =
    options.firstLineIndentCm ?? style.firstLineIndentCm;
  const hangingIndentCm = options.hangingIndentCm ?? style.hangingIndentCm;
  if (
    typeof firstLineIndentCm === "number" ||
    typeof hangingIndentCm === "number"
  ) {
    writeParagraphIndent(p, wNs, {
      firstLineTwips:
        typeof firstLineIndentCm === "number"
          ? cmToTwips(firstLineIndentCm)
          : undefined,
      hangingTwips:
        typeof hangingIndentCm === "number"
          ? cmToTwips(hangingIndentCm)
          : undefined,
    });
  } else {
    writeParagraphIndent(p, wNs);
  }
}

function applyConferenceTextStyle(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  options: ConferenceParagraphStyleOptions = {},
) {
  applyConferenceParagraphLayout(p, wNs, style, options);
  writeRunFormatting(
    p,
    wNs,
    style.fontFamily,
    style.fontSize,
    !!style.bold,
    !!style.italic,
    options.color ?? style.color ?? "000000",
  );
}

interface StyledRunPart {
  text: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
}

function setParagraphStyledRuns(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  parts: StyledRunPart[],
  options: ConferenceParagraphStyleOptions = {},
) {
  const existingPPr = getChild(p, wNs, "pPr");
  Array.from(p.childNodes).forEach((n) => {
    if (n !== existingPPr) p.removeChild(n);
  });

  applyConferenceParagraphLayout(p, wNs, style, options);

  for (const part of parts) {
    if (part.text === "") continue;
    const run = wElem(p.ownerDocument!, wNs, "r");
    const rPr = ensureRPr(run, wNs);
    writeRunProperties(
      rPr,
      wNs,
      style.fontFamily,
      part.fontSize ?? style.fontSize,
      part.bold ?? !!style.bold,
      part.italic ?? !!style.italic,
      options.color ?? style.color ?? "000000",
    );
    const t = wElem(p.ownerDocument!, wNs, "t");
    if (/^\s/u.test(part.text) || /\s$/u.test(part.text) || /\s{2,}/u.test(part.text)) {
      t.setAttribute("xml:space", "preserve");
    }
    t.textContent = part.text;
    run.appendChild(t);
    p.appendChild(run);
  }
}

function applyParagraphPrefixBold(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  prefixPattern: RegExp,
  options: ConferenceParagraphStyleOptions = {},
) {
  const text = normalizeText(getParagraphText(p, wNs));
  const match = text.match(prefixPattern);
  if (!match || match.index !== 0) {
    applyConferenceTextStyle(p, wNs, style, options);
    return;
  }

  const prefix = match[0];
  const rest = text.slice(prefix.length);
  setParagraphStyledRuns(
    p,
    wNs,
    style,
    [
      { text: prefix, bold: true },
      { text: rest, bold: false },
    ],
    options,
  );
}

function applyAcmKeywordsParagraph(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
) {
  const text = normalizeText(getParagraphText(p, wNs));
  const match = text.match(/^additional\s+keywords\s+and\s+phrases\s*:\s*/iu);
  const keywordStyle: ConferenceTextStyle = {
    ...style,
    fontFamily: "Linux Libertine O",
    fontSize: 9,
    alignment: "both",
    bold: false,
    italic: false,
    lineSpacingPt: 13.5,
    lineSpacingRule: "atLeast",
    color: "000000",
  };

  if (!match || match.index !== 0) {
    applyConferenceTextStyle(p, wNs, keywordStyle, {
      alignment: "both",
      beforePt: style.spacingBeforePt ?? 7,
      afterPt: style.spacingAfterPt ?? 0,
      lineSpacingPt: 13.5,
      lineSpacingRule: "atLeast",
      color: "000000",
    });
    return;
  }

  const prefix = match[0];
  const rest = text.slice(prefix.length);
  setParagraphStyledRuns(
    p,
    wNs,
    keywordStyle,
    [
      { text: prefix, bold: true, fontSize: 8 },
      { text: rest, bold: false, fontSize: 9 },
    ],
    {
      alignment: "both",
      beforePt: style.spacingBeforePt ?? 7,
      afterPt: style.spacingAfterPt ?? 0,
      lineSpacingPt: 13.5,
      lineSpacingRule: "atLeast",
      color: "000000",
    },
  );
}

function setParagraphPageBreakBefore(
  p: Element,
  wNs: string,
  enabled: boolean,
) {
  const pPr = ensurePPr(p, wNs);
  removeChildren(pPr, wNs, "pageBreakBefore");
  if (enabled) {
    pPr.appendChild(wElem(p.ownerDocument!, wNs, "pageBreakBefore"));
  }
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

function getTopLevelBodyChildren(body: Element, wNs: string): Element[] {
  return Array.from(body.childNodes).filter(
    (c): c is Element =>
      c instanceof Element &&
      c.namespaceURI === wNs &&
      (c.localName === "p" || c.localName === "tbl"),
  );
}

function getDirectChildElements(parent: Element, wNs: string, local: string) {
  return Array.from(parent.childNodes).filter(
    (c): c is Element =>
      c instanceof Element && c.namespaceURI === wNs && c.localName === local,
  );
}

function hasMathContent(p: Element, wNs: string): boolean {
  return (
    p.getElementsByTagNameNS(M_NS, "oMath").length > 0 ||
    p.getElementsByTagNameNS(M_NS, "oMathPara").length > 0
  );
}

function hasDrawingOrPicture(p: Element, wNs: string): boolean {
  return (
    p.getElementsByTagNameNS(wNs, "drawing").length > 0 ||
    p.getElementsByTagNameNS(wNs, "pict").length > 0
  );
}

function isAcmFootnoteParagraph(text: string): boolean {
  const normalized = normalizeText(text);
  if (!normalized) return false;
  if (/^[*†‡]\s+/u.test(normalized)) return true;
  return /permission to make digital|personal or classroom use|copyright|conference publication|this work is licensed|for personal or classroom use/iu.test(
    normalized,
  );
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

function isLikelyAcmHeading(text: string): boolean {
  return parseAcmHeading(text) !== null;
}

function isAcmFigureCaption(text: string): boolean {
  const normalized = normalizeText(text);
  return /^figure\s+\d+(?:\.\d+)?[:.]?\s+/iu.test(normalized);
}

function isAcmTableCaption(text: string): boolean {
  const normalized = normalizeText(text);
  return /^table\s+\d+(?:\.\d+)?[:.]?\s+/iu.test(normalized);
}

function isAcmEquationParagraph(p: Element, text: string): boolean {
  if (hasMathContent(p, resolveWNs(p.ownerDocument!))) return true;
  const normalized = normalizeText(text);
  if (!normalized || normalized.length > 80) return false;
  if (!/[=+\-*/^∑∫≤≥≈]/u.test(normalized)) return false;
  return /^[A-Za-z0-9\s()+\-*/=,^_{}\[\].]+$/u.test(normalized);
}

function toAcmTitleCaseWord(token: string): string {
  return toHeading2TitleCaseWord(token);
}

function normalizeAcmHeading2(text: string): string {
  const normalized = normalizeText(text);
  const match = normalized.match(/^(\d+\.\d+)(?:[.)])?\s+(.+)$/u);
  if (!match) {
    return normalized
      .split(/\s+/u)
      .map((token) => toAcmTitleCaseWord(token))
      .join(" ");
  }
  const marker = match[1];
  const content = match[2]
    .split(/\s+/u)
    .map((token) => toAcmTitleCaseWord(token))
    .join(" ");
  return `${marker} ${content}`.trim();
}

interface AcmHeadingMatch {
  level: 1 | 2 | 3;
  marker: string;
  title: string;
}

function isAcmHeadingTitle(title: string): boolean {
  const normalized = normalizeText(title);
  if (!normalized) return false;
  if (normalized.length > 120) return false;
  if (/^(abstract|keywords|ccs concepts|acm reference format)$/iu.test(normalized)) {
    return false;
  }
  return !/[.?!:]$/u.test(normalized);
}

function parseAcmHeading(text: string): AcmHeadingMatch | null {
  const normalized = normalizeText(text);
  if (!normalized) return null;
  if (/^references$/iu.test(normalized)) {
    return { level: 1, marker: "", title: "REFERENCES" };
  }

  if (
    /^(introduction|background|related work|literature review|methodology|methods|materials and methods|implementation|experiment|experiments|evaluation|results|discussion|results and discussion|conclusion|conclusions|future work|acknowledgments?|references)$/iu.test(
      normalized,
    )
  ) {
    return { level: 1, marker: "", title: normalized };
  }

  const h3 = normalized.match(/^(\d+\.\d+\.\d+(?:\.\d+)*)\s+(.+)$/u);
  if (h3 && isAcmHeadingTitle(h3[2])) {
    return { level: 3, marker: h3[1], title: h3[2] };
  }

  const h2 = normalized.match(/^(\d+\.\d+)(?:[.)])?\s+(.+)$/u);
  if (h2 && isAcmHeadingTitle(h2[2])) {
    return { level: 2, marker: h2[1], title: h2[2] };
  }

  const h1 = normalized.match(/^(\d+)(?:[.)])?\s+(.+)$/u);
  if (h1 && isAcmHeadingTitle(h1[2])) {
    return { level: 1, marker: h1[1], title: h1[2] };
  }

  return null;
}

function getTableRows(tbl: Element, wNs: string): Element[] {
  return getDirectChildElements(tbl, wNs, "tr");
}

function getTableCells(row: Element, wNs: string): Element[] {
  return getDirectChildElements(row, wNs, "tc");
}

function getParagraphsInCell(cell: Element, wNs: string): Element[] {
  return getDirectChildElements(cell, wNs, "p");
}

function getRowText(row: Element, wNs: string): string {
  const parts: string[] = [];
  for (const cell of getTableCells(row, wNs)) {
    const cellText = getParagraphsInCell(cell, wNs)
      .map((p) => normalizeText(getParagraphText(p, wNs)))
      .join(" ");
    parts.push(cellText);
  }
  return normalizeText(parts.join(" "));
}

function ensureTableCellBorder(
  cell: Element,
  wNs: string,
  side: "top" | "bottom",
  size = "4",
) {
  let tcPr = getChild(cell, wNs, "tcPr");
  if (!tcPr) {
    tcPr = wElem(cell.ownerDocument!, wNs, "tcPr");
    cell.insertBefore(tcPr, cell.firstChild);
  }

  let tcBorders = getChild(tcPr, wNs, "tcBorders");
  if (!tcBorders) {
    tcBorders = wElem(cell.ownerDocument!, wNs, "tcBorders");
    tcPr.appendChild(tcBorders);
  }

  const border = ensureChild(tcBorders, wNs, side);
  setWAttr(border, wNs, "val", "single");
  setWAttr(border, wNs, "sz", size);
  setWAttr(border, wNs, "space", "0");
  setWAttr(border, wNs, "color", "000000");
}

function clearTableCellBorders(cell: Element, wNs: string) {
  const tcPr = getChild(cell, wNs, "tcPr");
  if (!tcPr) return;
  removeChildren(tcPr, wNs, "tcBorders");
}

function clearTableCellBackground(cell: Element, wNs: string) {
  let tcPr = getChild(cell, wNs, "tcPr");
  if (!tcPr) {
    tcPr = wElem(cell.ownerDocument!, wNs, "tcPr");
    cell.insertBefore(tcPr, cell.firstChild);
  }
  removeChildren(tcPr, wNs, "shd");
  const shd = ensureChild(tcPr, wNs, "shd");
  setWAttr(shd, wNs, "val", "clear");
  setWAttr(shd, wNs, "color", "auto");
  setWAttr(shd, wNs, "fill", "auto");
}

function setTableRowHeaderRepeat(row: Element, wNs: string, repeat: boolean) {
  let trPr = getChild(row, wNs, "trPr");
  if (!trPr) {
    trPr = wElem(row.ownerDocument!, wNs, "trPr");
    row.insertBefore(trPr, row.firstChild);
  }
  removeChildren(trPr, wNs, "tblHeader");
  if (!repeat) return;
  const header = ensureChild(trPr, wNs, "tblHeader");
  setWAttr(header, wNs, "val", "true");
}

function countAcmTableHeaderRows(rows: Element[], wNs: string): number {
  if (rows.length === 0) return 0;
  let headerCount = 1;
  for (let i = 1; i < rows.length; i += 1) {
    const text = getRowText(rows[i], wNs);
    if (!text) break;
    if (text.length > 140) break;
    if (/[.?!]/u.test(text)) break;
    headerCount += 1;
  }
  return headerCount;
}

function deFloatTable(tbl: Element, wNs: string) {
  const tblPr = getChild(tbl, wNs, "tblPr");
  if (!tblPr) return;
  removeChildren(tblPr, wNs, "tblpPr");
  removeChildren(tblPr, wNs, "positionH");
  removeChildren(tblPr, wNs, "positionV");
  const tblInd = getChild(tblPr, wNs, "tblInd");
  if (tblInd) {
    setWAttr(tblInd, wNs, "w", "0");
    setWAttr(tblInd, wNs, "type", "dxa");
  }
}

function applyAcmTable(
  tbl: Element,
  wNs: string,
  styleConfig: AcmFormattingConfig,
) {
  deFloatTable(tbl, wNs);

  const tblPr = ensureChild(tbl, wNs, "tblPr");
  removeChildren(tblPr, wNs, "tblStyle");
  removeChildren(tblPr, wNs, "shd");
  removeChildren(tblPr, wNs, "tblW");
  const tblW = wElem(tbl.ownerDocument!, wNs, "tblW");
  setWAttr(tblW, wNs, "type", "pct");
  setWAttr(tblW, wNs, "w", "5000");
  tblPr.appendChild(tblW);
  Array.from(tbl.getElementsByTagNameNS(wNs, "tcW")).forEach((tcW) => {
    if (tcW.namespaceURI === wNs) tcW.parentElement?.removeChild(tcW);
  });
  Array.from(tbl.getElementsByTagNameNS(wNs, "trHeight")).forEach((trH) => {
    if (trH.namespaceURI === wNs) trH.parentElement?.removeChild(trH);
  });
  removeChildren(tblPr, wNs, "tblBorders");
  const tblBorders = wElem(tbl.ownerDocument!, wNs, "tblBorders");
  tblPr.appendChild(tblBorders);
  const makeBorder = (local: string, val: string, sz: string) => {
    const el = wElem(tbl.ownerDocument!, wNs, local);
    setWAttr(el, wNs, "val", val);
    setWAttr(el, wNs, "sz", sz);
    setWAttr(el, wNs, "space", "0");
    setWAttr(el, wNs, "color", "000000");
    tblBorders.appendChild(el);
  };
  makeBorder("top", "none", "0");
  makeBorder("left", "none", "0");
  makeBorder("bottom", "none", "0");
  makeBorder("right", "none", "0");
  makeBorder("insideH", "none", "0");
  makeBorder("insideV", "none", "0");

  const rows = getTableRows(tbl, wNs);
  if (rows.length === 0) return;

  const headerRowCount = countAcmTableHeaderRows(rows, wNs);
  const lastRow = rows[rows.length - 1];

  rows.forEach((row, index) => {
    setTableRowHeaderRepeat(row, wNs, index < headerRowCount);
    const trPr = getChild(row, wNs, "trPr");
    if (trPr) removeChildren(trPr, wNs, "shd");
    const cells = getTableCells(row, wNs);
    cells.forEach((cell) => {
      clearTableCellBorders(cell, wNs);
      clearTableCellBackground(cell, wNs);
    });
  });

  const firstHeaderRow = rows[0];
  const lastHeaderRow = rows[Math.max(0, headerRowCount - 1)];

  for (const cell of getTableCells(firstHeaderRow, wNs)) {
    ensureTableCellBorder(cell, wNs, "top");
  }
  for (const cell of getTableCells(lastHeaderRow, wNs)) {
    ensureTableCellBorder(cell, wNs, "bottom");
  }
  if (lastRow !== lastHeaderRow) {
    for (const cell of getTableCells(lastRow, wNs)) {
      ensureTableCellBorder(cell, wNs, "bottom");
    }
  }

  for (const row of rows) {
    for (const cell of getTableCells(row, wNs)) {
      for (const p of getParagraphsInCell(cell, wNs)) {
        const text = normalizeText(getParagraphText(p, wNs));
        if (!text && !hasDrawingOrPicture(p, wNs) && !hasMathContent(p, wNs)) {
          applyConferenceTextStyle(p, wNs, styleConfig.table, {
            alignment: "both",
            beforePt: 0,
            afterPt: 0,
            firstLineIndentCm: 0,
            hangingIndentCm: 0,
            lineSpacingPt: styleConfig.table.lineSpacingPt ?? 11,
            lineSpacingRule: styleConfig.table.lineSpacingRule ?? "atLeast",
            color: styleConfig.table.color ?? "000000",
          });
          continue;
        }

        applyConferenceTextStyle(p, wNs, styleConfig.table, {
          alignment: "both",
          beforePt: 0,
          afterPt: 0,
          firstLineIndentCm: 0,
          hangingIndentCm: 0,
          lineSpacingPt: styleConfig.table.lineSpacingPt ?? 11,
          lineSpacingRule: styleConfig.table.lineSpacingRule ?? "atLeast",
          color: styleConfig.table.color ?? "000000",
        });
      }
    }
  }
}

function applyAcmHeadingParagraph(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  level: 1 | 2 | 3,
) {
  const effectiveStyle =
    level === 2 ? { ...style, italic: false } : style;
  const rawText = normalizeText(getParagraphText(p, wNs));
  let nextText = rawText;

  if (level === 1) {
    const match = rawText.match(/^(\d+)(?:[.)])?\s+(.+)$/u);
    if (match) {
      nextText = `${match[1]} ${match[2].toUpperCase()}`.trim();
    } else if (style.uppercase) {
      nextText = rawText.toUpperCase();
    }
    if (/^references$/iu.test(rawText)) {
      nextText = "REFERENCES";
    }
  } else if (level === 2) {
    if (effectiveStyle.titleCase) {
      nextText = normalizeAcmHeading2(rawText);
    }
  }

  if (nextText !== rawText) {
    setParagraphText(p, wNs, nextText);
  }

  applyConferenceTextStyle(p, wNs, effectiveStyle, {
    alignment: effectiveStyle.alignment,
    beforePt: effectiveStyle.spacingBeforePt ?? 12,
    afterPt: effectiveStyle.spacingAfterPt ?? 3,
    color: effectiveStyle.color ?? "000000",
  });
}

function applyAcmReferenceParagraph(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  referenceIndex: number,
): number {
  const rawText = normalizeText(getParagraphText(p, wNs));
  const nextReferenceIndex = referenceIndex + 1;
  const normalizedMarker = normalizeIeeeReferenceMarker(rawText);
  const strippedMarker = normalizedMarker.replace(/^\[\d{1,3}\]\s*/u, "").trim();
  const nextText = `[${nextReferenceIndex}]\t${strippedMarker}`.trimEnd();

  if (hasParagraphNumbering(p, wNs)) {
    disableParagraphNumbering(p, wNs);
  }

  if (nextText !== rawText) {
    setParagraphTextWithTabs(p, wNs, nextText);
  }

  applyConferenceTextStyle(p, wNs, style, {
    alignment: "both",
    beforePt: style.spacingBeforePt ?? 0,
    afterPt: style.spacingAfterPt ?? 3,
    hangingIndentCm: style.hangingIndentCm ?? 0.63,
    lineSpacingPt: style.lineSpacingPt ?? 8.4,
    lineSpacingRule: style.lineSpacingRule ?? "atLeast",
    color: style.color ?? "000000",
  });

  return nextReferenceIndex;
}

function applyAcmSimpleParagraph(
  p: Element,
  wNs: string,
  style: ConferenceTextStyle,
  options: ConferenceParagraphStyleOptions = {},
) {
  applyConferenceTextStyle(p, wNs, style, {
    alignment: options.alignment ?? style.alignment,
    beforePt: options.beforePt ?? style.spacingBeforePt ?? 0,
    afterPt: options.afterPt ?? style.spacingAfterPt ?? 0,
    firstLineIndentCm: options.firstLineIndentCm ?? style.firstLineIndentCm,
    hangingIndentCm: options.hangingIndentCm ?? style.hangingIndentCm,
    lineSpacingPt: options.lineSpacingPt ?? style.lineSpacingPt,
    lineSpacingRule: options.lineSpacingRule ?? style.lineSpacingRule,
    color: options.color ?? style.color ?? "000000",
  });
}

interface AcmFrontMatterState {
  nonEmptyParagraphs: number;
  inAbstract: boolean;
  inReferenceFormat: boolean;
  seenAffiliationLine: boolean;
}

function isAcmAbstractStart(text: string): boolean {
  return /^abstract\b/iu.test(text) || /^abstract[-—]/iu.test(text);
}

function isAcmConceptsParagraph(text: string): boolean {
  return /^ccs\s+concepts\b/iu.test(text);
}

function isAcmKeywordsParagraph(text: string): boolean {
  return /^additional\s+keywords\s+and\s+phrases\s*:/iu.test(text);
}

function isAcmReferenceFormatParagraph(text: string): boolean {
  return /^acm\s+reference\s+format\s*:/iu.test(text);
}

function isAcmSignatoryLineText(text: string): boolean {
  return /^[_\-\u2013\u2014\s]{6,}$/u.test(text);
}

function isAcmPreliminaryFootnoteText(text: string): boolean {
  return (
    /^[*†‡]\s+/u.test(text) ||
    /place the footnote text|permission to make digital|personal or classroom use|copyright|this work is licensed/iu.test(
      text,
    )
  );
}

function isLikelyAcmAffiliationLine(text: string): boolean {
  return (
    /@/u.test(text) ||
    /\borcid\b/iu.test(text) ||
    /\b(affiliation|institution|university|college|school|department|institute|laboratory|lab|center|centre|city|country)\b/iu.test(
      text,
    )
  );
}

function isEmailOnlyLine(text: string): boolean {
  return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}(?:\s*[,;]\s*[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})*$/iu.test(
    normalizeText(text),
  );
}

function isLikelyAcmAuthorNameLine(text: string): boolean {
  const normalized = normalizeText(text);
  if (!normalized || normalized.length > 100) return false;
  if (isLikelyAcmAffiliationLine(normalized)) return false;
  if (/[.?!:]/u.test(normalized)) return false;
  return /^[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+)+(?:\s*,\s*[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+)+)*(?:\s+(?:and|&)\s+[A-Z][A-Za-z.'\-]+(?:\s+[A-Z][A-Za-z.'\-]+)+)?$/u.test(
    normalized,
  );
}

function shouldTreatAsAbstractBodyAfterAuthors(text: string): boolean {
  const normalized = normalizeText(text);
  if (!normalized) return false;
  if (isAcmAbstractStart(normalized)) return true;
  if (isAcmConceptsParagraph(normalized)) return false;
  if (isAcmKeywordsParagraph(normalized)) return false;
  if (isAcmReferenceFormatParagraph(normalized)) return false;
  if (isAcmPreliminaryFootnoteText(normalized)) return false;
  if (isLikelyAcmAffiliationLine(normalized)) return false;
  if (isLikelyAcmAuthorNameLine(normalized)) return false;
  if (/^(this\s+paper|in\s+this\s+paper|this\s+study|the\s+study|we\s+(present|propose|describe|introduce|investigate|examine))\b/iu.test(normalized)) {
    return true;
  }
  if (normalized.length > 120) return true;
  return /[.?!]$/u.test(normalized);
}

function getPreviousTopLevelParagraph(p: Element, wNs: string): Element | null {
  let node = p.previousSibling;
  while (node) {
    if (
      node instanceof Element &&
      node.namespaceURI === wNs &&
      node.localName === "p"
    ) {
      return node;
    }
    if (
      node instanceof Element &&
      node.namespaceURI === wNs &&
      node.localName === "tbl"
    ) {
      return null;
    }
    node = node.previousSibling;
  }
  return null;
}

function getNextTopLevelParagraph(p: Element, wNs: string): Element | null {
  let node = p.nextSibling;
  while (node) {
    if (
      node instanceof Element &&
      node.namespaceURI === wNs &&
      node.localName === "p"
    ) {
      return node;
    }
    if (
      node instanceof Element &&
      node.namespaceURI === wNs &&
      node.localName === "tbl"
    ) {
      return null;
    }
    node = node.nextSibling;
  }
  return null;
}

function mergeFollowingEmailIntoAffiliationLine(
  p: Element,
  wNs: string,
  text: string,
): string {
  const parts = [normalizeText(text)];

  for (let i = 0; i < 4; i += 1) {
    if (/@/u.test(parts.join(" "))) break;

    const next = getNextTopLevelParagraph(p, wNs);
    if (!next) break;

    const nextText = normalizeText(getParagraphText(next, wNs));
    if (!nextText) break;
    if (isAcmAbstractStart(nextText)) break;
    if (isAcmConceptsParagraph(nextText)) break;
    if (isAcmKeywordsParagraph(nextText)) break;
    if (isAcmReferenceFormatParagraph(nextText)) break;
    if (isLikelyAcmAuthorNameLine(nextText)) break;
    if (!isLikelyAcmAffiliationLine(nextText) && !isEmailOnlyLine(nextText)) {
      break;
    }

    parts.push(nextText);
    next.parentNode?.removeChild(next);

    if (isEmailOnlyLine(nextText)) break;
  }

  return parts.join(", ");
}

function acmAuthorAffiliationLine(author: AuthorEntry): string {
  const parts = [
    author.department,
    author.organization,
    author.cityCountry,
    author.contact,
  ]
    .map((part) => normalizeText(part))
    .filter(Boolean);
  const deduped = parts.filter(
    (part, index) =>
      parts.findIndex((candidate) => candidate.toLowerCase() === part.toLowerCase()) === index,
  );
  return deduped.join(", ");
}

function createParagraphBefore(reference: Element, wNs: string): Element {
  const p = wElem(reference.ownerDocument!, wNs, "p");
  reference.parentNode?.insertBefore(p, reference);
  return p;
}

function createTextParagraphBefore(
  reference: Element,
  wNs: string,
  text: string,
): Element {
  const p = createParagraphBefore(reference, wNs);
  setParagraphText(p, wNs, text);
  return p;
}

function rebuildAcmAuthorBlock(
  body: Element,
  wNs: string,
  authors: AuthorEntry[],
) {
  if (authors.length === 0) return;

  const paragraphs = getDirectChildElements(body, wNs, "p");
  const nonEmpty = paragraphs
    .map((p, index) => ({
      p,
      index,
      text: normalizeText(getParagraphText(p, wNs)),
    }))
    .filter((row) => row.text.length > 0);

  if (nonEmpty.length < 4) return;

  const authorStart = nonEmpty[2];
  const abstractStart = nonEmpty.find(
    (row, index) => index > 2 && isAcmAbstractStart(row.text),
  );
  if (!authorStart || !abstractStart) return;

  for (let i = authorStart.index; i < abstractStart.index; i += 1) {
    const p = paragraphs[i];
    if (p?.parentNode === body) {
      body.removeChild(p);
    }
  }

  for (const author of authors) {
    createTextParagraphBefore(abstractStart.p, wNs, author.name);
    const affiliation = acmAuthorAffiliationLine(author);
    if (affiliation) {
      createTextParagraphBefore(abstractStart.p, wNs, affiliation);
    }
  }
}

async function getAcmAuthors(
  doc: Document,
  wNs: string,
  aiAssist: boolean,
): Promise<AuthorEntry[]> {
  const frontMatterLines = normalizeInputLinesUntilAbstract(doc, wNs);
  const extractionLines = getAcmAuthorExtractionLines(frontMatterLines);
  const parsedAuthors = parseAcmAuthorEntriesFromFrontMatter(frontMatterLines);
  const aiAuthors = await tryExtractAuthorsWithAi(extractionLines, aiAssist);
  return normalizeAuthorEntries(aiAuthors ?? parsedAuthors);
}

function applyAcmSignatoryLine(p: Element, wNs: string) {
  setParagraphText(p, wNs, "________________________________");
  applyConferenceTextStyle(
    p,
    wNs,
    {
      fontFamily: "Linux Libertine O",
      fontSize: 7,
      lineSpacing: 1.0,
      alignment: "left",
      bold: false,
      italic: false,
      color: "000000",
    },
    {
      alignment: "left",
      beforePt: 48,
      afterPt: 3,
      lineSpacingRule: "auto",
    },
  );
}

function ensureAcmSignatoryLineBefore(p: Element, wNs: string) {
  const previous = getPreviousTopLevelParagraph(p, wNs);
  if (
    previous &&
    isAcmSignatoryLineText(normalizeText(getParagraphText(previous, wNs)))
  ) {
    applyAcmSignatoryLine(previous, wNs);
    return;
  }

  const line = createParagraphBefore(p, wNs);
  applyAcmSignatoryLine(line, wNs);
}

function applyAcmPreliminaryParagraph(
  p: Element,
  wNs: string,
  styleConfig: AcmFormattingConfig,
  state: AcmFrontMatterState,
) {
  const text = normalizeText(getParagraphText(p, wNs));
  if (!text) return;

  const ordinal = state.nonEmptyParagraphs;
  state.nonEmptyParagraphs += 1;

  if (isAcmSignatoryLineText(text)) {
    applyAcmSignatoryLine(p, wNs);
    return;
  }

  if (isAcmPreliminaryFootnoteText(text)) {
    ensureAcmSignatoryLineBefore(p, wNs);
    applyAcmSimpleParagraph(p, wNs, styleConfig.preliminaryFootnote, {
      alignment: "both",
      beforePt: styleConfig.preliminaryFootnote.spacingBeforePt ?? 0,
      afterPt: styleConfig.preliminaryFootnote.spacingAfterPt ?? 0,
      lineSpacingRule: styleConfig.preliminaryFootnote.lineSpacingRule ?? "auto",
      lineSpacingPt: styleConfig.preliminaryFootnote.lineSpacingPt,
      color: styleConfig.preliminaryFootnote.color ?? "000000",
    });
    state.inReferenceFormat = false;
    return;
  }

  if (isAcmReferenceFormatParagraph(text)) {
    state.inAbstract = false;
    state.inReferenceFormat = true;
    applyParagraphPrefixBold(
      p,
      wNs,
      styleConfig.referenceFormatLabel,
      /^acm\s+reference\s+format\s*:\s*/iu,
      {
        alignment: "both",
        beforePt: styleConfig.referenceFormatLabel.spacingBeforePt ?? 8,
        afterPt: styleConfig.referenceFormatLabel.spacingAfterPt ?? 0,
        lineSpacingPt: styleConfig.referenceFormatLabel.lineSpacingPt ?? 9.6,
        lineSpacingRule:
          styleConfig.referenceFormatLabel.lineSpacingRule ?? "atLeast",
        color: styleConfig.referenceFormatLabel.color ?? "000000",
      },
    );
    return;
  }

  if (state.inReferenceFormat) {
    applyAcmSimpleParagraph(p, wNs, styleConfig.referenceFormatContent, {
      alignment: "both",
      beforePt: styleConfig.referenceFormatContent.spacingBeforePt ?? 1,
      afterPt: styleConfig.referenceFormatContent.spacingAfterPt ?? 0,
      lineSpacingPt: styleConfig.referenceFormatContent.lineSpacingPt ?? 12,
      lineSpacingRule:
        styleConfig.referenceFormatContent.lineSpacingRule ?? "exact",
      color: styleConfig.referenceFormatContent.color ?? "000000",
    });
    return;
  }

  if (isAcmConceptsParagraph(text)) {
    state.inAbstract = false;
    applyAcmSimpleParagraph(p, wNs, styleConfig.concepts, {
      alignment: "both",
      beforePt: styleConfig.concepts.spacingBeforePt ?? 7,
      afterPt: styleConfig.concepts.spacingAfterPt ?? 0,
      lineSpacingPt: styleConfig.concepts.lineSpacingPt ?? 13.5,
      lineSpacingRule: styleConfig.concepts.lineSpacingRule ?? "atLeast",
      color: styleConfig.concepts.color ?? "000000",
    });
    return;
  }

  if (isAcmKeywordsParagraph(text)) {
    state.inAbstract = false;
    applyAcmKeywordsParagraph(p, wNs, styleConfig.keywords);
    return;
  }

  if (
    isAcmAbstractStart(text) ||
    state.inAbstract ||
    (state.seenAffiliationLine && shouldTreatAsAbstractBodyAfterAuthors(text))
  ) {
    state.inAbstract = true;
    const abstractStyle: ConferenceTextStyle = {
      ...styleConfig.abstract,
      fontFamily: "Linux Libertine O",
      fontSize: 8,
      alignment: "both",
      bold: false,
      italic: false,
      lineSpacingPt: 12,
      lineSpacingRule: "atLeast",
      color: "000000",
    };
    applyAcmSimpleParagraph(p, wNs, abstractStyle, {
      alignment: "both",
      beforePt: 10,
      afterPt: abstractStyle.spacingAfterPt ?? 0,
      lineSpacingPt: 12,
      lineSpacingRule: "atLeast",
      color: "000000",
    });
    return;
  }

  if (ordinal === 0) {
    applyAcmSimpleParagraph(p, wNs, styleConfig.title, {
      alignment: "left",
      beforePt: styleConfig.title.spacingBeforePt ?? 0,
      afterPt: styleConfig.title.spacingAfterPt ?? 0,
      lineSpacingPt: styleConfig.title.lineSpacingPt ?? 18,
      lineSpacingRule: styleConfig.title.lineSpacingRule ?? "atLeast",
      color: styleConfig.title.color ?? "000000",
    });
    return;
  }

  if (ordinal === 1) {
    applyAcmSimpleParagraph(p, wNs, styleConfig.subtitle, {
      alignment: "left",
      beforePt: styleConfig.subtitle.spacingBeforePt ?? 0,
      afterPt: styleConfig.subtitle.spacingAfterPt ?? 18,
      lineSpacingPt: styleConfig.subtitle.lineSpacingPt ?? 16.65,
      lineSpacingRule: styleConfig.subtitle.lineSpacingRule ?? "atLeast",
      color: styleConfig.subtitle.color ?? "000000",
    });
    return;
  }

  if (isLikelyAcmAffiliationLine(text)) {
    const affiliationText = mergeFollowingEmailIntoAffiliationLine(p, wNs, text);
    setParagraphText(p, wNs, affiliationText);
    state.seenAffiliationLine = true;
    const affiliationStyle: ConferenceTextStyle = {
      ...styleConfig.authorAffiliation,
      fontFamily: "Linux Libertine O",
      fontSize: 9,
      alignment: "left",
      bold: false,
      italic: false,
      lineSpacingPt: 14.85,
      lineSpacingRule: "atLeast",
      color: "000000",
    };
    applyAcmSimpleParagraph(p, wNs, affiliationStyle, {
      alignment: "left",
      beforePt: affiliationStyle.spacingBeforePt ?? 3,
      afterPt: affiliationStyle.spacingAfterPt ?? 0,
      lineSpacingPt: 14.85,
      lineSpacingRule: "atLeast",
      color: "000000",
    });
    return;
  }

  state.inAbstract = false;
  applyAcmSimpleParagraph(p, wNs, styleConfig.author, {
    alignment: "left",
    beforePt: styleConfig.author.spacingBeforePt ?? 3,
    afterPt: styleConfig.author.spacingAfterPt ?? 0,
    lineSpacingPt: styleConfig.author.lineSpacingPt ?? 16,
    lineSpacingRule: styleConfig.author.lineSpacingRule ?? "atLeast",
    color: styleConfig.author.color ?? "000000",
  });
}

function getAcmHeadingLevelFromKind(
  kind: AcmParagraphKind | undefined,
): 1 | 2 | 3 | null {
  if (kind === "heading1" || kind === "references") return 1;
  if (kind === "heading2") return 2;
  if (kind === "heading3") return 3;
  return null;
}

function getAcmHeadingStyle(
  styleConfig: AcmFormattingConfig,
  level: 1 | 2 | 3,
): ConferenceTextStyle {
  if (level === 1) return styleConfig.heading1;
  if (level === 2) return styleConfig.heading2;
  return styleConfig.heading3;
}

function applyAcmRuleBasedFormatting(
  doc: Document,
  styleConfig: AcmFormattingConfig,
  paragraphHints: AcmParagraphHintMap = new Map(),
  authors: AuthorEntry[] = [],
) {
  const wNs = resolveWNs(doc);
  const body = getBody(doc, wNs);
  if (!body) return;

  enforceSingleColumnNoHeaderFooter(body, wNs);
  rebuildAcmAuthorBlock(body, wNs, authors);

  const children = getTopLevelBodyChildren(body, wNs);
  let mainActive = false;
  let inReferences = false;
  let firstBodyAfterHeading = false;
  let referenceIndex = 0;
  const frontMatterState: AcmFrontMatterState = {
    nonEmptyParagraphs: 0,
    inAbstract: false,
    inReferenceFormat: false,
    seenAffiliationLine: false,
  };

  for (let childIndex = 0; childIndex < children.length; childIndex += 1) {
    const child = children[childIndex];
    if (child.parentNode !== body) continue;
    const hintedKind = paragraphHints.get(childIndex);

    if (child.localName === "tbl") {
      if (mainActive) {
        applyAcmTable(child, wNs, styleConfig);
      }
      continue;
    }

    const p = child;
    const rawText = normalizeText(getParagraphText(p, wNs));
    const localHeading = parseAcmHeading(rawText);
    const hintedHeadingLevel = getAcmHeadingLevelFromKind(hintedKind);
    const headingLevel = hintedHeadingLevel ?? localHeading?.level ?? null;
    const isReferencesHeading =
      hintedKind === "references" ||
      /^references$/iu.test(rawText) ||
      !!(
        headingLevel === 1 &&
        localHeading &&
        /^references$/iu.test(localHeading.title)
      );

    if (!mainActive) {
      if (hintedKind === "ignore" || !headingLevel) {
        applyAcmPreliminaryParagraph(p, wNs, styleConfig, frontMatterState);
        continue;
      }

      mainActive = true;
      setParagraphPageBreakBefore(p, wNs, true);
      if (isReferencesHeading) {
        inReferences = true;
      }

      applyAcmHeadingParagraph(
        p,
        wNs,
        getAcmHeadingStyle(styleConfig, headingLevel),
        headingLevel,
      );

      firstBodyAfterHeading = true;
      continue;
    }

    if (rawText === "") {
      if (inReferences) {
        applyAcmSimpleParagraph(p, wNs, styleConfig.references, {
          alignment: "both",
          beforePt: 0,
          afterPt: 0,
          hangingIndentCm: styleConfig.references.hangingIndentCm ?? 0.63,
          lineSpacingPt: styleConfig.references.lineSpacingPt,
          lineSpacingRule: styleConfig.references.lineSpacingRule,
        });
      } else {
        applyAcmSimpleParagraph(p, wNs, styleConfig.body, {
          beforePt: 0,
          afterPt: 0,
          firstLineIndentCm: 0,
          lineSpacingPt: styleConfig.body.lineSpacingPt,
          lineSpacingRule: styleConfig.body.lineSpacingRule,
        });
      }
      continue;
    }

    if (headingLevel) {
      if (isReferencesHeading) {
        inReferences = true;
      }
      applyAcmHeadingParagraph(
        p,
        wNs,
        getAcmHeadingStyle(styleConfig, headingLevel),
        headingLevel,
      );
      firstBodyAfterHeading = true;
      continue;
    }

    if (hintedKind === "footnote" || isAcmFootnoteParagraph(rawText)) {
      applyAcmSimpleParagraph(p, wNs, styleConfig.footnote, {
        alignment: "center",
        beforePt: styleConfig.footnote.spacingBeforePt ?? 3,
        afterPt: styleConfig.footnote.spacingAfterPt ?? 10,
        lineSpacingRule: styleConfig.footnote.lineSpacingRule ?? "auto",
        lineSpacingPt: styleConfig.footnote.lineSpacingPt,
      });
      continue;
    }

    if (inReferences) {
      referenceIndex = applyAcmReferenceParagraph(
        p,
        wNs,
        styleConfig.references,
        referenceIndex,
      );
      continue;
    }

    if (hintedKind === "tableCaption" || isAcmTableCaption(rawText)) {
      applyAcmSimpleParagraph(p, wNs, styleConfig.tableCaption, {
        alignment: "center",
        beforePt: styleConfig.tableCaption.spacingBeforePt ?? 9,
        afterPt: styleConfig.tableCaption.spacingAfterPt ?? 6,
        lineSpacingPt: styleConfig.tableCaption.lineSpacingPt,
        lineSpacingRule: styleConfig.tableCaption.lineSpacingRule,
      });
      continue;
    }

    if (hintedKind === "figureCaption" || isAcmFigureCaption(rawText)) {
      applyAcmSimpleParagraph(p, wNs, styleConfig.figureCaption, {
        alignment: "center",
        beforePt: styleConfig.figureCaption.spacingBeforePt ?? 3,
        afterPt: styleConfig.figureCaption.spacingAfterPt ?? 9,
        lineSpacingPt: styleConfig.figureCaption.lineSpacingPt,
        lineSpacingRule: styleConfig.figureCaption.lineSpacingRule,
      });
      continue;
    }

    if (hintedKind === "equation" || isAcmEquationParagraph(p, rawText)) {
      applyAcmSimpleParagraph(p, wNs, styleConfig.equation, {
        alignment: "center",
        beforePt: styleConfig.equation.spacingBeforePt ?? 0,
        afterPt: styleConfig.equation.spacingAfterPt ?? 0,
        lineSpacingPt: styleConfig.equation.lineSpacingPt,
        lineSpacingRule: styleConfig.equation.lineSpacingRule,
      });
      continue;
    }

    if (hasDrawingOrPicture(p, wNs)) {
      applyAcmSimpleParagraph(p, wNs, styleConfig.figure, {
        alignment: "center",
        beforePt: styleConfig.figure.spacingBeforePt ?? 6,
        afterPt: styleConfig.figure.spacingAfterPt ?? 10,
        lineSpacingPt: styleConfig.figure.lineSpacingPt,
        lineSpacingRule: styleConfig.figure.lineSpacingRule,
      });
      continue;
    }

    if (firstBodyAfterHeading) {
      applyAcmSimpleParagraph(p, wNs, styleConfig.body, {
        firstLineIndentCm: 0,
        lineSpacingPt: styleConfig.body.lineSpacingPt,
        lineSpacingRule: styleConfig.body.lineSpacingRule,
      });
      firstBodyAfterHeading = false;
      continue;
    }

    applyAcmSimpleParagraph(p, wNs, styleConfig.body, {
      firstLineIndentCm: styleConfig.body.firstLineIndentCm ?? 0.42,
      lineSpacingPt: styleConfig.body.lineSpacingPt,
      lineSpacingRule: styleConfig.body.lineSpacingRule,
    });
  }
}

function applyAcmRuleBasedFormattingLegacy(
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
  aiAssist = AI_ASSIST_ENABLED,
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
  const parsedAuthors = normalizeAuthorEntries(
    parseAuthorEntriesFromFrontMatter(frontMatterLines),
  );
  const localFallbackAuthors = normalizeAuthorEntries(
    parseAcmAuthorEntriesFromFrontMatter(frontMatterLines),
  );
  const aiAuthors = await tryExtractAuthorsWithAi(frontMatterLines, aiAssist);
  const authors = normalizeAuthorEntries(
    aiAuthors && aiAuthors.length > 0
      ? aiAuthors
      : parsedAuthors.length > 0
        ? parsedAuthors
        : localFallbackAuthors,
  );
  const titleText = frontMatterLines[0] ?? "";
  if (titleText && templateParas[0]) {
    setParagraphText(templateParas[0], templateWNs, titleText);
  }

  if (authors.length > 0) {
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
  }

  if (authors.length > 0) {
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
  aiAssist = AI_ASSIST_ENABLED,
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
  await applyPublicationTemplate(inputZip, pubformZip, styleConfig, aiAssist);
  return pubformZip.generateAsync({ type: "blob" }) as Promise<Blob>;
}

export async function formatDocxAcm(
  arrayBuffer: ArrayBuffer,
  styleConfig: AcmFormattingConfig,
  aiAssist = AI_ASSIST_ENABLED,
): Promise<Blob> {
  const JSZip = requireJsZip();
  const targetZip = await JSZip.loadAsync(arrayBuffer);
  const targetDoc = await parseDocFromZip(targetZip, "word/document.xml");
  const wNs = resolveWNs(targetDoc);
  const authors = await getAcmAuthors(targetDoc, wNs, aiAssist);
  const paragraphHints = await tryClassifyAcmParagraphsWithAi(
    targetDoc,
    aiAssist,
  );
  applyAcmRuleBasedFormatting(targetDoc, styleConfig, paragraphHints, authors);
  targetZip.file("word/document.xml", serializeDoc(targetDoc));

  return targetZip.generateAsync({ type: "blob" }) as Promise<Blob>;
}

export async function formatDocxConference(
  arrayBuffer: ArrayBuffer,
  options: ConferenceFormatOptions,
): Promise<Blob> {
  const resolvedConfig = resolveConferenceStyleConfig(options.styleConfig);
  if (options.format === "pubform") {
    return formatDocxPublication(
      arrayBuffer,
      resolvedConfig.pubform,
      options.aiAssist ?? AI_ASSIST_ENABLED,
    );
  }
  return formatDocxAcm(
    arrayBuffer,
    resolvedConfig.acm,
    options.aiAssist ?? AI_ASSIST_ENABLED,
  );
}
