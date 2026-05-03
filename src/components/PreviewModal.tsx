import { useEffect, useMemo, useState } from "react";
import type {
  CitationStyle,
  ConferenceFormat,
  ConferenceFormattingConfig,
  FormattingStandard,
} from "../constants";

type RuleItem = [string, string, string];

interface PreviewProps {
  open: boolean;
  onClose: () => void;
  formattingStandard: FormattingStandard;
  citationStyle: CitationStyle;
  conferenceFormat: ConferenceFormat;
  conferenceConfig: ConferenceFormattingConfig;
}

const CCC_PREVIEW_RULES = (citationStyle: CitationStyle): RuleItem[] => [
  [
    "fa-font",
    "Font & Size",
    "Garamond - 14pt titles, 13pt headings, 12pt body, 11pt references.",
  ],
  [
    "fa-text-height",
    "Line Spacing",
    "Double-spaced body text; single-spaced captions and references.",
  ],
  [
    "fa-hashtag",
    "Pagination",
    "Chapter labels trigger page breaks automatically.",
  ],
  [
    "fa-indent",
    "Indentation",
    "First-line indent 1.27 cm for all body paragraphs.",
  ],
  [
    "fa-image",
    "Captions",
    "Centered figure captions; left-aligned table captions, both ending with a period.",
  ],
  [
    "fa-table",
    "Tables",
    "Title above; double outer borders; header row separator; auto full-width.",
  ],
  [
    "fa-book-bookmark",
    "References",
    citationStyle === "apa"
      ? "APA author-date style with configured spacing and hanging indent."
      : "IEEE numbered style with configured spacing and hanging indent.",
  ],
];

function alignLabel(alignment: string): string {
  if (alignment === "both") return "justified";
  return alignment;
}

function conferencePreviewRules(
  conferenceFormat: ConferenceFormat,
  conferenceConfig: ConferenceFormattingConfig,
): RuleItem[] {
  if (conferenceFormat === "pubform") {
    const pub = conferenceConfig.pubform;
    return [
      [
        "fa-heading",
        "Heading 1 (Roman)",
        `${pub.heading1.fontFamily} ${pub.heading1.fontSize}pt, ${alignLabel(pub.heading1.alignment)}, ${pub.heading1.uppercase ? "uppercase" : "normal case"}, ${pub.heading1.italic ? "italic" : "regular"}.`,
      ],
      [
        "fa-list-ol",
        "Heading 2 (Lettered)",
        `${pub.heading2.fontFamily} ${pub.heading2.fontSize}pt, ${alignLabel(pub.heading2.alignment)}, ${pub.heading2.titleCase ? "title case" : "source case"}, ${pub.heading2.italic ? "italic" : "regular"}.`,
      ],
      [
        "fa-align-justify",
        "Body Paragraph",
        `${pub.body.fontFamily} ${pub.body.fontSize}pt, ${alignLabel(pub.body.alignment)}, line spacing ${pub.body.lineSpacing.toFixed(2)}, first-line indent ${pub.body.firstLineIndentCm.toFixed(2)} cm, spacing after ${pub.body.spacingAfterPt.toFixed(0)} pt.`,
      ],
      [
        "fa-book",
        "References",
        `${pub.references.fontFamily} ${pub.references.fontSize}pt, ${alignLabel(pub.references.alignment)}, hanging indent ${pub.references.hangingIndentCm.toFixed(2)} cm, ${pub.references.ieeeStyle ? "IEEE markers [n]" : "manual marker style"}.`,
      ],
    ];
  }

  const acm = conferenceConfig.acm;
  return [
    [
      "fa-heading",
      "Title",
      `${acm.title.fontFamily} ${acm.title.fontSize}pt, ${alignLabel(acm.title.alignment)}, ${acm.title.bold ? "bold" : "regular"}, line spacing ${acm.title.lineSpacing.toFixed(2)}.`,
    ],
    [
      "fa-file-lines",
      "Subtitle / Front Matter",
      `${acm.subtitle.fontFamily} ${acm.subtitle.fontSize}pt, ${alignLabel(acm.subtitle.alignment)}, ${acm.subtitle.italic ? "italic" : "regular"}.`,
    ],
    [
      "fa-users",
      "Author Block",
      `${acm.author.fontFamily} ${acm.author.fontSize}pt, ${alignLabel(acm.author.alignment)}, ${acm.author.bold ? "bold" : "regular"}.`,
    ],
    [
      "fa-list-ol",
      "Section Headings",
      `${acm.heading.fontFamily} ${acm.heading.fontSize}pt, ${alignLabel(acm.heading.alignment)}, ${acm.heading.bold ? "bold" : "regular"}.`,
    ],
    [
      "fa-align-left",
      "Body and References",
      `Body: ${acm.body.fontFamily} ${acm.body.fontSize}pt (${alignLabel(acm.body.alignment)}). References: ${acm.references.fontFamily} ${acm.references.fontSize}pt (${alignLabel(acm.references.alignment)}).`,
    ],
  ];
}

export default function PreviewModal({
  open,
  onClose,
  formattingStandard,
  citationStyle,
  conferenceFormat,
  conferenceConfig,
}: PreviewProps) {
  const [visible, setVisible] = useState(false);
  const [closing, setClosing] = useState(false);

  const rules = useMemo(
    () =>
      formattingStandard === "conference"
        ? conferencePreviewRules(conferenceFormat, conferenceConfig)
        : CCC_PREVIEW_RULES(citationStyle),
    [formattingStandard, conferenceFormat, conferenceConfig, citationStyle],
  );

  const title =
    formattingStandard === "conference"
      ? conferenceFormat === "pubform"
        ? "Publication Rule Preview"
        : "ACM Rule Preview"
      : "Formatting Guide";
  const subText =
    formattingStandard === "conference"
      ? "Editable conference formatting rules currently in use."
      : "Rules applied to your manuscript.";

  useEffect(() => {
    if (open) {
      setClosing(false);
      setVisible(true);
      document.body.style.overflow = "hidden";
    }
  }, [open]);

  const handleClose = () => {
    setClosing(true);
    setTimeout(() => {
      setVisible(false);
      setClosing(false);
      document.body.style.overflow = "";
      onClose();
    }, 340);
  };

  useEffect(() => {
    if (!open && visible && !closing) handleClose();
  }, [open]);

  if (!visible) return null;

  return (
    <div
      className={`sheet-backdrop${visible ? " open" : ""}`}
      onClick={(e) => {
        if (e.target === e.currentTarget) handleClose();
      }}
    >
      <div className={`sheet-modal${closing ? " closing" : ""}`}>
        <div className="px-5 pt-4 pb-2 shrink-0">
          <div className="sheet-drag-handle" />
          <div className="flex items-center justify-between">
            <div>
              <p
                className="text-[10px] font-bold uppercase tracking-[0.2em]"
                style={{ color: "var(--accent)" }}
              >
                <i className="fa-solid fa-scroll mr-1" /> Formatting Rules
              </p>
              <h2
                className="text-xl font-bold mt-0.5"
                style={{ color: "var(--text-primary)" }}
              >
                {title}
              </h2>
            </div>
            <button
              onClick={handleClose}
              className="flex h-8 w-8 items-center justify-center rounded-full transition"
              style={{
                background: "var(--surface-raised)",
                color: "var(--text-secondary)",
              }}
              type="button"
            >
              <i className="fa-solid fa-xmark text-sm" />
            </button>
          </div>
          <p className="mt-1 text-sm" style={{ color: "var(--text-muted)" }}>
            {subText}
          </p>
        </div>

        <div className="flex-1 overflow-y-auto px-5 pb-4 space-y-2.5 mt-2">
          {rules.map(([icon, ruleTitle, desc]) => (
            <div
              key={ruleTitle}
              className="rounded-2xl border p-4"
              style={{
                background: "var(--surface-raised)",
                borderColor: "var(--border)",
              }}
            >
              <div className="flex items-center gap-2.5 mb-1">
                <span
                  className="flex h-7 w-7 shrink-0 items-center justify-center rounded-xl"
                  style={{ background: "var(--accent-subtle-strong)" }}
                >
                  <i
                    className={`fa-solid ${icon} text-xs`}
                    style={{ color: "var(--accent)" }}
                  />
                </span>
                <h3
                  className="text-sm font-semibold"
                  style={{ color: "var(--text-primary)" }}
                >
                  {ruleTitle}
                </h3>
              </div>
              <p
                className="text-xs mt-1 pl-[2.375rem]"
                style={{ color: "var(--text-muted)" }}
              >
                {desc}
              </p>
            </div>
          ))}
        </div>

        <div
          className="shrink-0 border-t px-5 py-4"
          style={{ borderColor: "var(--border)", background: "var(--surface)" }}
        >
          <button
            onClick={handleClose}
            className="w-full rounded-2xl py-3.5 text-sm font-bold transition active:scale-[0.98]"
            style={{
              background: "var(--surface-raised)",
              color: "var(--text-secondary)",
            }}
          >
            <i className="fa-solid fa-xmark mr-2" /> Close
          </button>
        </div>
      </div>
    </div>
  );
}

