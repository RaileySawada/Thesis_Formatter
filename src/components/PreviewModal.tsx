import { useEffect, useState } from "react";

// ── PreviewModal ─────────────────────────────────────────────────────────────

const PREVIEW_RULES: [string, string, string][] = [
  [
    "fa-font",
    "Font & Size",
    "Garamond — 14pt titles, 13pt headings, 12pt body, 11pt references.",
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
    "Centered figure captions; left-aligned table captions — both end with a period.",
  ],
  [
    "fa-table",
    "Tables",
    "Title above; double outer borders; header row separator; auto full-width.",
  ],
  [
    "fa-book-bookmark",
    "References",
    "Garamond 11pt, 1.0× spacing, hanging indent 1.27 cm.",
  ],
];

interface PreviewProps {
  open: boolean;
  onClose: () => void;
}

export function PreviewModal({ open, onClose }: PreviewProps) {
  const [visible, setVisible] = useState(false);
  const [closing, setClosing] = useState(false);

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
        {/* Header */}
        <div className="px-5 pt-4 pb-2 shrink-0">
          <div
            className="sheet-drag-handle"
          />
          <div className="flex items-center justify-between">
            <div>
              <p
                className="text-[10px] font-bold uppercase tracking-[0.2em]"
                style={{ color: "var(--accent)" }}
              >
                <i className="fa-solid fa-scroll mr-1" /> Template Rules
              </p>
              <h2
                className="text-xl font-bold mt-0.5"
                style={{ color: "var(--text-primary)" }}
              >
                Formatting Guide
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
            Rules applied to your manuscript.
          </p>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto px-5 pb-4 space-y-2.5 mt-2">
          {PREVIEW_RULES.map(([icon, title, desc]) => (
            <div
              key={title}
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
                  {title}
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

        {/* Footer */}
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

export default PreviewModal;
