import { useMemo, type CSSProperties } from "react";
import type {
  ConferenceFormat,
  ConferenceFormattingConfig,
  ConferenceTextAlignment,
  ConferenceTextStyle,
} from "../constants";

interface Props {
  format: ConferenceFormat;
  config: ConferenceFormattingConfig;
  onChange: (next: ConferenceFormattingConfig) => void;
}

const FONT_FAMILIES = [
  "Times New Roman",
  "Linux Biolinum O",
  "Linux Libertine O",
  "Garamond",
  "Arial",
  "Calibri",
  "Georgia",
  "Verdana",
];

const ALIGNMENTS: { value: ConferenceTextAlignment; icon: string }[] = [
  { value: "left", icon: "fa-align-left" },
  { value: "center", icon: "fa-align-center" },
  { value: "right", icon: "fa-align-right" },
  { value: "both", icon: "fa-align-justify" },
];

function toPreviewTitleCase(text: string): string {
  return text
    .split(/\s+/u)
    .map((token) => {
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
    })
    .join(" ");
}

function Stepper({
  label,
  value,
  step,
  min = 0,
  onChange,
}: {
  label: string;
  value: number;
  step: number;
  min?: number;
  onChange: (val: number) => void;
}) {
  const digits = step >= 1 ? 0 : 2;
  return (
    <div className="space-y-1">
      <label
        className="text-[10px] font-bold uppercase tracking-wider"
        style={{ color: "var(--text-soft)" }}
      >
        {label}
      </label>
      <div
        className="flex items-center overflow-hidden rounded-xl border"
        style={{ borderColor: "var(--border)", background: "var(--surface)" }}
      >
        <button
          type="button"
          onClick={() =>
            onChange(Number(Math.max(min, value - step).toFixed(digits)))
          }
          className="flex h-9 w-9 items-center justify-center transition hover:bg-black/5"
          style={{ borderRight: "1px solid var(--border)", color: "var(--text-secondary)" }}
        >
          <i className="fa-solid fa-minus text-[10px]" />
        </button>
        <div className="flex-1 text-center text-xs font-bold" style={{ color: "var(--text-primary)" }}>
          {value.toFixed(digits)}
        </div>
        <button
          type="button"
          onClick={() => onChange(Number((value + step).toFixed(digits)))}
          className="flex h-9 w-9 items-center justify-center transition hover:bg-black/5"
          style={{ borderLeft: "1px solid var(--border)", color: "var(--text-secondary)" }}
        >
          <i className="fa-solid fa-plus text-[10px]" />
        </button>
      </div>
    </div>
  );
}

function ToggleGroup({
  value,
  onChange,
}: {
  value: ConferenceTextAlignment;
  onChange: (val: ConferenceTextAlignment) => void;
}) {
  return (
    <div className="space-y-1">
      <label
        className="text-[10px] font-bold uppercase tracking-wider"
        style={{ color: "var(--text-soft)" }}
      >
        Alignment
      </label>
      <div
        className="flex rounded-xl border p-1"
        style={{ borderColor: "var(--border)", background: "var(--surface)" }}
      >
        {ALIGNMENTS.map((opt) => (
          <button
            key={opt.value}
            type="button"
            onClick={() => onChange(opt.value)}
            className="flex-1 flex h-8 items-center justify-center rounded-lg transition"
            style={
              value === opt.value
                ? { background: "var(--accent)", color: "#fff" }
                : { color: "var(--text-muted)" }
            }
          >
            <i className={`fa-solid ${opt.icon} text-xs`} />
          </button>
        ))}
      </div>
    </div>
  );
}

function PreviewCard({
  title,
  sample,
  style,
}: {
  title: string;
  sample: string;
  style: ConferenceTextStyle;
}) {
  const displaySample = style.uppercase
    ? sample.toUpperCase()
    : style.titleCase
      ? toPreviewTitleCase(sample)
      : sample;
  const lineHeight = style.lineSpacingPt
    ? style.lineSpacingPt / style.fontSize
    : style.lineSpacing;
  const css = useMemo<CSSProperties>(
    () => ({
      fontFamily: style.fontFamily,
      fontSize: `${style.fontSize}px`,
      lineHeight,
      textAlign:
        (style.alignment === "both" ? "justify" : style.alignment) as CSSProperties["textAlign"],
      fontWeight: style.bold ? "700" : "400",
      fontStyle: style.italic ? "italic" : "normal",
      textTransform: style.uppercase ? "uppercase" : "none",
      color: "#334155",
      whiteSpace: "pre-line" as const,
      marginTop: style.spacingBeforePt ? `${style.spacingBeforePt * 1.333}px` : undefined,
      marginBottom: style.spacingAfterPt ? `${style.spacingAfterPt * 1.333}px` : undefined,
    }),
    [style, lineHeight],
  );

  return (
    <div className="rounded-2xl border p-3" style={{ borderColor: "var(--border)", background: "#fff" }}>
      <p className="mb-2 text-[10px] font-bold uppercase tracking-wider" style={{ color: "var(--accent)" }}>
        Preview: {title}
      </p>
      <div className="rounded-lg border border-dashed p-3" style={{ borderColor: "#cbd5e1" }}>
        <p style={css}>{displaySample}</p>
      </div>
    </div>
  );
}

function Section({
  title,
  icon,
  children,
}: {
  title: string;
  icon: string;
  children: React.ReactNode;
}) {
  return (
    <div className="mb-3 overflow-hidden rounded-2xl border" style={{ borderColor: "var(--border)", background: "var(--surface-raised)" }}>
      <div className="flex items-center gap-3 px-4 py-3" style={{ background: "var(--accent-subtle)" }}>
        <span
          className="flex h-8 w-8 items-center justify-center rounded-xl"
          style={{ background: "var(--accent-subtle-strong)", color: "var(--accent)" }}
        >
          <i className={`fa-solid ${icon} text-sm`} />
        </span>
        <span className="text-sm font-bold" style={{ color: "var(--text-primary)" }}>
          {title}
        </span>
      </div>
      <div className="space-y-4 p-4">{children}</div>
    </div>
  );
}

function StyleEditor({
  style,
  onChange,
  sample,
  title,
  extra,
}: {
  style: ConferenceTextStyle;
  onChange: (updates: Partial<ConferenceTextStyle>) => void;
  sample: string;
  title: string;
  extra?: React.ReactNode;
}) {
  const usesExactLineSpacing = typeof style.lineSpacingPt === "number";
  const showSpacing =
    typeof style.spacingBeforePt === "number" ||
    typeof style.spacingAfterPt === "number";
  const showIndent =
    typeof style.firstLineIndentCm === "number" ||
    typeof style.hangingIndentCm === "number";
  const showCaseToggle =
    typeof style.uppercase === "boolean" || typeof style.titleCase === "boolean";

  return (
    <div className="space-y-4">
      <PreviewCard title={title} sample={sample} style={style} />
      <div className="grid grid-cols-2 gap-4">
        <div className="space-y-1">
          <label className="text-[10px] font-bold uppercase tracking-wider" style={{ color: "var(--text-soft)" }}>
            Font Family
          </label>
          <select
            value={style.fontFamily}
            onChange={(e) => onChange({ fontFamily: e.target.value })}
            className="w-full rounded-xl border px-3 py-2 text-xs font-bold"
            style={{ borderColor: "var(--border)", background: "var(--surface)", color: "var(--text-primary)" }}
          >
            {FONT_FAMILIES.map((font) => (
              <option key={font} value={font}>
                {font}
              </option>
            ))}
          </select>
        </div>
        <Stepper
          label="Font Size (pt)"
          value={style.fontSize}
          step={1}
          min={6}
          onChange={(v) => onChange({ fontSize: v })}
        />
      </div>
      <div className="grid grid-cols-2 gap-4">
        {usesExactLineSpacing ? (
          <Stepper
            label="Line Spacing (pt)"
            value={style.lineSpacingPt ?? 0}
            step={0.05}
            min={0}
            onChange={(v) => onChange({ lineSpacingPt: v })}
          />
        ) : (
          <Stepper
            label="Line Spacing"
            value={style.lineSpacing}
            step={0.05}
            min={0.8}
            onChange={(v) => onChange({ lineSpacing: v })}
          />
        )}
        <ToggleGroup
          value={style.alignment}
          onChange={(v) => onChange({ alignment: v })}
        />
      </div>
      {showSpacing && (
        <div className="grid grid-cols-2 gap-4">
          {typeof style.spacingBeforePt === "number" ? (
            <Stepper
              label="Spacing Before (pt)"
              value={style.spacingBeforePt}
              step={1}
              min={0}
              onChange={(v) => onChange({ spacingBeforePt: v })}
            />
          ) : (
            <div />
          )}
          {typeof style.spacingAfterPt === "number" ? (
            <Stepper
              label="Spacing After (pt)"
              value={style.spacingAfterPt}
              step={1}
              min={0}
              onChange={(v) => onChange({ spacingAfterPt: v })}
            />
          ) : (
            <div />
          )}
        </div>
      )}
      {showIndent && (
        <div className="grid grid-cols-2 gap-4">
          {typeof style.firstLineIndentCm === "number" ? (
            <Stepper
              label="First Line Indent (cm)"
              value={style.firstLineIndentCm}
              step={0.01}
              min={0}
              onChange={(v) => onChange({ firstLineIndentCm: v })}
            />
          ) : (
            <div />
          )}
          {typeof style.hangingIndentCm === "number" ? (
            <Stepper
              label="Hanging Indent (cm)"
              value={style.hangingIndentCm}
              step={0.01}
              min={0}
              onChange={(v) => onChange({ hangingIndentCm: v })}
            />
          ) : (
            <div />
          )}
        </div>
      )}
      <div className="grid grid-cols-2 gap-2">
        <button
          type="button"
          onClick={() => onChange({ bold: !style.bold })}
          className="rounded-xl border py-2 text-xs font-bold"
          style={
            style.bold
              ? { borderColor: "var(--accent)", background: "var(--accent-subtle-strong)", color: "var(--accent)" }
              : { borderColor: "var(--border)", background: "var(--surface)", color: "var(--text-muted)" }
          }
        >
          <i className="fa-solid fa-bold mr-1" />
          Bold
        </button>
        <button
          type="button"
          onClick={() => onChange({ italic: !style.italic })}
          className="rounded-xl border py-2 text-xs font-bold"
          style={
            style.italic
              ? { borderColor: "var(--accent)", background: "var(--accent-subtle-strong)", color: "var(--accent)" }
              : { borderColor: "var(--border)", background: "var(--surface)", color: "var(--text-muted)" }
          }
          >
          <i className="fa-solid fa-italic mr-1" />
          Italic
        </button>
      </div>
      {showCaseToggle && (
        <div className="grid grid-cols-2 gap-2">
          {typeof style.uppercase === "boolean" ? (
            <button
              type="button"
              onClick={() => onChange({ uppercase: !style.uppercase })}
              className="rounded-xl border py-2 text-xs font-bold"
              style={
                style.uppercase
                  ? {
                      borderColor: "var(--accent)",
                      background: "var(--accent-subtle-strong)",
                      color: "var(--accent)",
                    }
                  : {
                      borderColor: "var(--border)",
                      background: "var(--surface)",
                      color: "var(--text-muted)",
                    }
              }
            >
              <i className="fa-solid fa-text-height mr-1" />
              Uppercase
            </button>
          ) : (
            <div />
          )}
          {typeof style.titleCase === "boolean" ? (
            <button
              type="button"
              onClick={() => onChange({ titleCase: !style.titleCase })}
              className="rounded-xl border py-2 text-xs font-bold"
              style={
                style.titleCase
                  ? {
                      borderColor: "var(--accent)",
                      background: "var(--accent-subtle-strong)",
                      color: "var(--accent)",
                    }
                  : {
                      borderColor: "var(--border)",
                      background: "var(--surface)",
                      color: "var(--text-muted)",
                    }
              }
            >
              <i className="fa-solid fa-font mr-1" />
              Title Case
            </button>
          ) : (
            <div />
          )}
        </div>
      )}
      {extra}
    </div>
  );
}

export default function ConferenceFormattingPanel({
  format,
  config,
  onChange,
}: Props) {
  const updatePub = (
    key: keyof ConferenceFormattingConfig["pubform"],
    updates: Partial<ConferenceFormattingConfig["pubform"][typeof key]>,
  ) => {
    onChange({
      ...config,
      pubform: {
        ...config.pubform,
        [key]: {
          ...config.pubform[key],
          ...updates,
        },
      },
    });
  };

  const updateAcm = (
    key: keyof ConferenceFormattingConfig["acm"],
    updates: Partial<ConferenceFormattingConfig["acm"][typeof key]>,
  ) => {
    onChange({
      ...config,
      acm: {
        ...config.acm,
        [key]: {
          ...config.acm[key],
          ...updates,
        },
      },
    });
  };

  if (format === "pubform") {
    return (
      <div className="pb-4">
        <Section title="Heading 1 (Roman)" icon="fa-heading">
          <StyleEditor
            title="Heading 1"
            sample="I. INTRODUCTION"
            style={config.pubform.heading1}
            onChange={(u) => updatePub("heading1", u)}
          />
        </Section>

        <Section title="Heading 2 (Lettered)" icon="fa-list-ol">
          <StyleEditor
            title="Heading 2"
            sample="A. Sample Title"
            style={config.pubform.heading2}
            onChange={(u) => updatePub("heading2", u)}
          />
        </Section>

        <Section title="Body Paragraph" icon="fa-align-justify">
          <StyleEditor
            title="Body"
            sample="This is a sample body paragraph. It is justified and follows publication paragraph spacing and indentation."
            style={config.pubform.body}
            onChange={(u) => updatePub("body", u)}
            extra={
              <div className="grid grid-cols-2 gap-4">
                <Stepper
                  label="Spacing After (pt)"
                  value={config.pubform.body.spacingAfterPt}
                  step={1}
                  min={0}
                  onChange={(v) => updatePub("body", { spacingAfterPt: v })}
                />
                <Stepper
                  label="First Line Indent (cm)"
                  value={config.pubform.body.firstLineIndentCm}
                  step={0.01}
                  min={0}
                  onChange={(v) => updatePub("body", { firstLineIndentCm: v })}
                />
              </div>
            }
          />
        </Section>

        <Section title="References (IEEE)" icon="fa-book">
          <StyleEditor
            title="References"
            sample="[1] Author, A., and Author, B. 2026. Sample reference title."
            style={config.pubform.references}
            onChange={(u) => updatePub("references", u)}
            extra={
              <div className="grid grid-cols-2 gap-4">
                <Stepper
                  label="Hanging Indent (cm)"
                  value={config.pubform.references.hangingIndentCm}
                  step={0.01}
                  min={0}
                  onChange={(v) => updatePub("references", { hangingIndentCm: v })}
                />
                <button
                  type="button"
                  onClick={() =>
                    updatePub("references", {
                      ieeeStyle: !config.pubform.references.ieeeStyle,
                    })
                  }
                  className="mt-5 rounded-xl border py-2 text-xs font-bold"
                  style={
                    config.pubform.references.ieeeStyle
                      ? { borderColor: "var(--accent)", background: "var(--accent-subtle-strong)", color: "var(--accent)" }
                      : { borderColor: "var(--border)", background: "var(--surface)", color: "var(--text-muted)" }
                  }
                >
                  <i className="fa-solid fa-hashtag mr-1" />
                  IEEE Markers
                </button>
              </div>
            }
          />
        </Section>
      </div>
    );
  }

  return (
    <div className="pb-4">
      <Section title="Title" icon="fa-heading">
        <StyleEditor
          title="ACM Title"
          sample="Submission Template for ACM Papers"
          style={config.acm.title}
          onChange={(u) => updateAcm("title", u)}
        />
      </Section>

      <Section title="Subtitle" icon="fa-file-lines">
        <StyleEditor
          title="Subtitle"
          sample="This is the subtitle of the paper"
          style={config.acm.subtitle}
          onChange={(u) => updateAcm("subtitle", u)}
        />
      </Section>

      <Section title="Author Block" icon="fa-users">
        <StyleEditor
          title="Author Name"
          sample="First Author's Name, Initials, and Last Name"
          style={config.acm.author}
          onChange={(u) => updateAcm("author", u)}
        />
        <StyleEditor
          title="Affiliation and Email"
          sample="First author's affiliation, an Institution with a very long name, xxxx@gmail.com"
          style={config.acm.authorAffiliation}
          onChange={(u) => updateAcm("authorAffiliation", u)}
        />
      </Section>

      <Section title="Abstract and Metadata" icon="fa-align-justify">
        <StyleEditor
          title="Abstract"
          sample="Abstract This paper presents a concise summary of the study."
          style={config.acm.abstract}
          onChange={(u) => updateAcm("abstract", u)}
        />
        <StyleEditor
          title="CCS Concepts"
          sample="CCS CONCEPTS • Insert your first CCS term here • Insert your second CCS term here"
          style={config.acm.concepts}
          onChange={(u) => updateAcm("concepts", u)}
        />
        <StyleEditor
          title="Additional Keywords"
          sample="Additional Keywords and Phrases: Keyword number 1, Keyword number 2"
          style={config.acm.keywords}
          onChange={(u) => updateAcm("keywords", u)}
        />
      </Section>

      <Section title="ACM Reference Format" icon="fa-book-open">
        <StyleEditor
          title="Reference Label"
          sample="ACM Reference Format:"
          style={config.acm.referenceFormatLabel}
          onChange={(u) => updateAcm("referenceFormatLabel", u)}
        />
        <StyleEditor
          title="Reference Content"
          sample="First Author's Name. 2018. The Title of the Paper: ACM Conference Proceedings Manuscript Submission Template."
          style={config.acm.referenceFormatContent}
          onChange={(u) => updateAcm("referenceFormatContent", u)}
        />
        <StyleEditor
          title="Preliminary Footnote"
          sample="* Place the footnote text for the author (if applicable) here."
          style={config.acm.preliminaryFootnote}
          onChange={(u) => updateAcm("preliminaryFootnote", u)}
        />
      </Section>

      <Section title="Heading 1" icon="fa-list-ol">
        <StyleEditor
          title="Heading 1"
          sample="1 INTRODUCTION"
          style={config.acm.heading1}
          onChange={(u) => updateAcm("heading1", u)}
        />
      </Section>

      <Section title="Heading 2" icon="fa-list-ol">
        <StyleEditor
          title="Heading 2"
          sample="1.1 Accessibility"
          style={config.acm.heading2}
          onChange={(u) => updateAcm("heading2", u)}
        />
      </Section>

      <Section title="Heading 3" icon="fa-list-ul">
        <StyleEditor
          title="Heading 3"
          sample="1.1.1 Heading 3 Title"
          style={config.acm.heading3}
          onChange={(u) => updateAcm("heading3", u)}
        />
      </Section>

      <Section title="Body Paragraph" icon="fa-align-left">
        <StyleEditor
          title="Body"
          sample="ACM manuscript body text sample to preview spacing, alignment, and font."
          style={config.acm.body}
          onChange={(u) => updateAcm("body", u)}
        />
      </Section>

      <Section title="Table Caption" icon="fa-closed-captioning">
        <StyleEditor
          title="Table Caption"
          sample="Table 1: Sample table caption"
          style={config.acm.tableCaption}
          onChange={(u) => updateAcm("tableCaption", u)}
        />
      </Section>

      <Section title="Table Text" icon="fa-table">
        <StyleEditor
          title="Table"
          sample="Table cell content sample."
          style={config.acm.table}
          onChange={(u) => updateAcm("table", u)}
        />
      </Section>

      <Section title="Footnote" icon="fa-note-sticky">
        <StyleEditor
          title="Footnote"
          sample="* Permission to make digital or hard copies for personal or classroom use."
          style={config.acm.footnote}
          onChange={(u) => updateAcm("footnote", u)}
        />
      </Section>

      <Section title="Figure" icon="fa-image">
        <StyleEditor
          title="Figure"
          sample="Inline figure paragraph sample."
          style={config.acm.figure}
          onChange={(u) => updateAcm("figure", u)}
        />
      </Section>

      <Section title="Figure Caption" icon="fa-closed-captioning">
        <StyleEditor
          title="Figure Caption"
          sample="Figure 1: Sample Figure."
          style={config.acm.figureCaption}
          onChange={(u) => updateAcm("figureCaption", u)}
        />
      </Section>

      <Section title="Equation" icon="fa-square-root-variable">
        <StyleEditor
          title="Equation"
          sample="E = mc^2"
          style={config.acm.equation}
          onChange={(u) => updateAcm("equation", u)}
        />
      </Section>

      <Section title="References" icon="fa-quote-right">
        <StyleEditor
          title="References"
          sample="[1] Example ACM reference item with conference/journal details."
          style={config.acm.references}
          onChange={(u) => updateAcm("references", u)}
        />
      </Section>
    </div>
  );
}
