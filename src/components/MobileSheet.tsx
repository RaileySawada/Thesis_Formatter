import { useEffect, useState } from "react";
import { RULES_DEF, CITATION_STYLES } from "../constants";
import type { CitationStyle } from "../constants";

const SECTIONS = [
  {
    value: "preliminary",
    label: "Preliminary Pages",
    sub: "Title page, approval sheet, abstract, TOC",
    icon: "fa-file-lines",
    disabled: false,
  },
  {
    value: "chapters",
    label: "Chapter 1 – References",
    sub: "Chapters, headings, body text, figures",
    icon: "fa-book-open",
    disabled: false,
  },
  {
    value: "appendices",
    label: "Appendices",
    sub: "Appendix headings, labels, CV",
    icon: "fa-paperclip",
    disabled: false,
  },
];

interface Props {
  open: boolean;
  onClose: () => void;
  selectedSections: string[];
  toggleSection: (v: string) => void;
  enabledRules: string[];
  toggleRule: (v: string) => void;
  sectionLabels: Record<string, string>;
  sectionIcons: Record<string, string>;
  citationStyle: CitationStyle;
  setCitationStyle: (v: CitationStyle) => void;
}

export default function MobileSheet({
  open,
  onClose,
  selectedSections,
  toggleSection,
  enabledRules,
  toggleRule,
  citationStyle,
  setCitationStyle,
}: Props) {
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
    }, 320);
  };

  if (!visible) return null;

  return (
    <div
      className={`options-backdrop${visible ? " open" : ""}`}
      onClick={(e) => {
        if (e.target === e.currentTarget) handleClose();
      }}
    >
      <div className={`options-modal${closing ? " closing" : ""}`}>
        <div className="px-5 pt-4 pb-2 shrink-0">
          {/* Drag handle — mobile only */}
          <div
            className="options-drag-handle mx-auto mb-3 h-1 w-10 rounded-full"
            style={{ background: "var(--border)" }}
          />
          <div className="flex items-center justify-between">
            <div>
              <p
                className="text-[10px] font-bold uppercase tracking-[0.2em]"
                style={{ color: "var(--accent)" }}
              >
                <i className="fa-solid fa-sliders mr-1" /> Options
              </p>
              <h2
                className="text-xl font-bold mt-0.5"
                style={{ color: "var(--text-primary)" }}
              >
                Formatting Scope
              </h2>
            </div>
            <button
              onClick={handleClose}
              className="flex h-8 w-8 items-center justify-center rounded-full transition"
              style={{
                background: "var(--surface-raised)",
                color: "var(--text-secondary)",
              }}
            >
              <i className="fa-solid fa-xmark text-sm" />
            </button>
          </div>
          <p className="mt-1 text-sm" style={{ color: "var(--text-muted)" }}>
            Select sections and rules to apply.
          </p>
        </div>

        <div className="flex-1 overflow-y-auto px-5 pb-4 space-y-4">
          {/* Citation Style */}
          <div>
            <p
              className="mb-2 text-[10px] font-bold uppercase tracking-[0.18em]"
              style={{ color: "var(--text-muted)" }}
            >
              <i className="fa-solid fa-quote-left mr-1" /> Citation Style
            </p>
            <div className="flex gap-2">
              {CITATION_STYLES.map((cs) => {
                const active = citationStyle === cs.value;
                return (
                  <button
                    key={cs.value}
                    type="button"
                    onClick={() => setCitationStyle(cs.value)}
                    className="flex-1 flex flex-col items-center gap-1.5 rounded-2xl border-2 px-3 py-3 transition-all text-center"
                    style={
                      active
                        ? {
                            borderColor: "var(--accent)",
                            background: "var(--accent-subtle)",
                          }
                        : {
                            borderColor: "var(--border)",
                            background: "var(--surface-raised)",
                          }
                    }
                  >
                    <span
                      className="flex h-7 w-7 items-center justify-center rounded-xl text-xs"
                      style={
                        active
                          ? {
                              background: "var(--accent-subtle-strong)",
                              color: "var(--accent)",
                            }
                          : {
                              background: "var(--border)",
                              color: "var(--text-secondary)",
                            }
                      }
                    >
                      <i className={`fa-solid ${cs.icon}`} />
                    </span>
                    <span
                      className="text-sm font-bold"
                      style={{
                        color: "var(--text-primary)",
                      }}
                    >
                      {cs.label}
                    </span>
                    <span
                      className="text-[10px]"
                      style={{
                        color: active
                          ? "var(--text-secondary)"
                          : "var(--text-soft)",
                      }}
                    >
                      {cs.sub}
                    </span>
                    {active && (
                      <span
                        className="rounded-full px-2 py-0.5 text-[9px] font-bold uppercase"
                        style={{ background: "var(--accent)", color: "#fff" }}
                      >
                        Active
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          <div>
            <p
              className="mb-2 text-[10px] font-bold uppercase tracking-[0.18em]"
              style={{ color: "var(--text-muted)" }}
            >
              <i className="fa-solid fa-sliders mr-1" /> Options
            </p>
            <div className="space-y-2">
              {SECTIONS.map((sec) => {
                const isSelected = selectedSections.includes(sec.value);
                if (sec.disabled)
                  return (
                    <div
                      key={sec.value}
                      className="section-card section-card--disabled relative flex items-start gap-3 rounded-2xl border px-4 py-3 cursor-not-allowed select-none"
                      style={{ borderColor: "var(--border)", opacity: 0.45 }}
                    >
                      <span
                        className="mt-0.5 flex h-7 w-7 shrink-0 items-center justify-center rounded-xl"
                        style={{
                          background: "var(--surface-raised)",
                          color: "var(--text-muted)",
                        }}
                      >
                        <i className={`fa-solid ${sec.icon} text-xs`} />
                      </span>
                      <span className="flex-1 min-w-0">
                        <span
                          className="block text-sm font-semibold"
                          style={{ color: "var(--text-muted)" }}
                        >
                          {sec.label}
                        </span>
                        <span
                          className="block text-xs"
                          style={{ color: "var(--text-muted)" }}
                        >
                          {sec.sub}
                        </span>
                      </span>
                      <span
                        className="rounded-full px-2 py-0.5 text-[10px] font-bold uppercase"
                        style={{
                          background: "var(--surface-raised)",
                          color: "var(--text-muted)",
                        }}
                      >
                        Soon
                      </span>
                    </div>
                  );
                return (
                  <div
                    key={sec.value}
                    className={`section-card relative flex cursor-pointer items-start gap-3 rounded-2xl px-4 py-3 transition-all${isSelected ? " section-card--selected" : ""}`}
                    style={
                      isSelected
                        ? {
                            border: "2px solid var(--accent)",
                            background: "var(--accent-subtle)",
                          }
                        : { border: "1px solid var(--border)" }
                    }
                    onClick={() => toggleSection(sec.value)}
                  >
                    <span
                      className="mt-0.5 flex h-7 w-7 shrink-0 items-center justify-center rounded-xl"
                      style={
                        isSelected
                          ? {
                              background: "var(--accent-subtle-strong)",
                              color: "var(--accent)",
                            }
                          : {
                              background: "var(--surface-raised)",
                              color: "var(--text-muted)",
                            }
                      }
                    >
                      <i className={`fa-solid ${sec.icon} text-xs`} />
                    </span>
                    <span className="flex-1 min-w-0 pr-1">
                      <span
                        className="block text-sm font-semibold leading-snug"
                        style={{
                          color: isSelected
                            ? "var(--text-primary)"
                            : "var(--text-secondary)",
                        }}
                      >
                        {sec.label}
                      </span>
                      <span
                        className="block text-xs mt-0.5"
                        style={{ color: "var(--text-secondary)" }}
                      >
                        {sec.sub}
                      </span>
                    </span>
                    {isSelected && (
                      <span
                        className="section-card__check mt-0.5 shrink-0 flex h-5 w-5 items-center justify-center rounded-full text-white"
                        style={{ background: "var(--accent)" }}
                      >
                        <i className="fa-solid fa-check text-[10px]" />
                      </span>
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          <div>
            <p
              className="mb-2 text-[10px] font-bold uppercase tracking-[0.18em]"
              style={{ color: "var(--text-muted)" }}
            >
              <i className="fa-solid fa-gear mr-1" /> Advanced Rules
            </p>
            <div className="space-y-2">
              {RULES_DEF.map(([val, icon, label]) => {
                const checked = enabledRules.includes(val);
                return (
                  <label
                    key={val}
                    className="flex items-center gap-3 rounded-xl border px-3 py-3 cursor-pointer transition-colors"
                    style={{
                      background: "var(--surface-raised)",
                      borderColor: "var(--border)",
                    }}
                  >
                    <input
                      type="checkbox"
                      className="sr-only"
                      checked={checked}
                      onChange={() => toggleRule(val)}
                    />
                    <span
                      className="rule-toggle-dot flex h-4 w-4 shrink-0 items-center justify-center rounded-full text-white"
                      style={{
                        background: checked ? "var(--accent)" : "var(--border)",
                      }}
                    >
                      <i
                        className="fa-solid fa-check text-[8px]"
                        style={{ opacity: checked ? 1 : 0 }}
                      />
                    </span>
                    <i
                      className={`fa-solid ${icon} text-xs w-3.5 text-center shrink-0`}
                      style={{ color: "var(--text-muted)" }}
                    />
                    <span
                      className="text-sm font-medium"
                      style={{ color: "var(--text-secondary)" }}
                    >
                      {label}
                    </span>
                  </label>
                );
              })}
            </div>
          </div>
        </div>

        <div
          className="shrink-0 border-t px-5 py-4"
          style={{ borderColor: "var(--border)", background: "var(--surface)" }}
        >
          <button
            type="button"
            onClick={handleClose}
            className="w-full rounded-2xl py-3.5 text-sm font-bold text-white transition active:scale-[0.98]"
            style={{ background: "var(--accent)" }}
          >
            <i className="fa-solid fa-check mr-2" /> Apply &amp; Close
          </button>
        </div>
      </div>
    </div>
  );
}
