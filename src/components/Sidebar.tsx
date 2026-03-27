import { RULES_DEF, CITATION_STYLES } from "../constants";
import type { CitationStyle } from "../constants";

interface Props {
  isDark: boolean;
  toggleTheme: () => void;
  selectedSections: string[];
  toggleSection: (v: string) => void;
  enabledRules: string[];
  toggleRule: (v: string) => void;
  rulesOpen: boolean;
  setRulesOpen: (v: boolean) => void;
  citationStyle: CitationStyle;
  setCitationStyle: (v: CitationStyle) => void;
}

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

export default function Sidebar({
  isDark,
  toggleTheme,
  selectedSections,
  toggleSection,
  enabledRules,
  toggleRule,
  rulesOpen,
  setRulesOpen,
  citationStyle,
  setCitationStyle,
}: Props) {
  return (
    <aside
      id="main-sidebar"
      className="sticky top-0 hidden lg:flex lg:flex-col w-full max-w-[280px] shrink-0 self-start"
    >
      {/* Logo */}
      <div className="flex items-center gap-3 pb-1">
        <img
          src="/images/logo.webp"
          alt="Thesis Formatter"
          className="h-10 w-10 object-contain shrink-0 rounded-xl"
        />
        <div>
          <p
            className="text-[10px] font-bold uppercase tracking-[0.22em]"
            style={{ color: "var(--accent)" }}
          >
            Thesis
          </p>
          <p
            className="text-sm font-bold leading-none"
            style={{ color: "var(--text-primary)" }}
          >
            Formatter
          </p>
        </div>
      </div>

      <div className="h-px" style={{ background: "var(--border)" }} />

      {/* Appearance */}
      <div
        className="rounded-2xl border px-4 py-3 transition-colors"
        style={{
          background: "var(--surface-raised)",
          borderColor: "var(--border)",
        }}
      >
        <p
          className="mb-2.5 text-[10px] font-bold uppercase tracking-[0.18em]"
          style={{ color: "var(--text-soft)" }}
        >
          <i className="fa-solid fa-palette mr-1" /> Appearance
        </p>
        <div className="flex items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            {isDark ? (
              <i
                className="fa-solid fa-moon fa-lg"
                style={{ color: "#60a5fa" }}
              />
            ) : (
              <i
                className="fa-solid fa-sun fa-lg"
                style={{ color: "#f59e0b" }}
              />
            )}
            <span
              className="text-sm font-semibold"
              style={{ color: "var(--text-primary)" }}
            >
              Dark Mode
            </span>
          </div>
          <button
            type="button"
            className={`toggle-track${isDark ? " on" : ""}`}
            aria-pressed={isDark}
            aria-label="Toggle dark mode"
            onClick={toggleTheme}
            style={{ border: "none", padding: 0 }}
          >
            <div className="toggle-thumb" />
          </button>
        </div>
      </div>

      {/* Citation Style */}
      <div
        className="rounded-2xl border px-4 py-3 transition-colors"
        style={{
          background: "var(--surface-raised)",
          borderColor: "var(--border)",
        }}
      >
        <p
          className="mb-2.5 text-[10px] font-bold uppercase tracking-[0.18em]"
          style={{ color: "var(--text-soft)" }}
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
                className="flex-1 flex flex-col items-center gap-1 rounded-xl border-2 px-2 py-2.5 transition-all text-center"
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
                  className="flex h-6 w-6 items-center justify-center rounded-lg text-xs"
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
                  className="text-xs font-bold leading-tight"
                  style={{
                    color: active
                      ? "var(--text-primary)"
                      : "var(--text-primary)",
                  }}
                >
                  {cs.label}
                </span>
                <span
                  className="text-[9px] leading-tight"
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
                    className="mt-0.5 rounded-full px-1.5 py-0.5 text-[9px] font-bold uppercase"
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

      {/* Options */}
      <div>
        <p
          className="mb-2.5 text-[10px] font-bold uppercase tracking-[0.18em]"
          style={{ color: "var(--text-soft)" }}
        >
          <i className="fa-solid fa-sliders mr-1" /> Options
        </p>
        <div className="space-y-2">
          {SECTIONS.map((sec) => {
            const isSelected = selectedSections.includes(sec.value);
            if (sec.disabled) {
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
                      className="block text-xs mt-0.5"
                      style={{ color: "var(--text-muted)" }}
                    >
                      {sec.sub}
                    </span>
                  </span>
                  <span
                    className="mt-0.5 shrink-0 rounded-full px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide"
                    style={{
                      background: "var(--surface-raised)",
                      color: "var(--text-muted)",
                    }}
                  >
                    Soon
                  </span>
                </div>
              );
            }
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

      {/* Advanced Rules */}
      <div>
        <button
          type="button"
          className="w-full flex items-center justify-between rounded-xl px-1 py-1.5 transition cursor-pointer"
          onClick={() => setRulesOpen(!rulesOpen)}
        >
          <p
            className="text-[10px] font-bold uppercase tracking-[0.18em]"
            style={{ color: "var(--text-soft)" }}
          >
            <i className="fa-solid fa-gear mr-1" /> Advanced Rules
          </p>
          <i
            className={`fa-solid fa-chevron-down rules-chevron text-xs${rulesOpen ? "" : " rotated"}`}
            style={{ color: "var(--text-muted)" }}
          />
        </button>
        <div className={`rules-body${rulesOpen ? "" : " collapsed"}`}>
          <div className="space-y-1.5 pt-2">
            {RULES_DEF.map(([val, icon, label]) => {
              const checked = enabledRules.includes(val);
              return (
                <label
                  key={val}
                  className="tf-rule-row flex items-center gap-3 rounded-xl px-3 py-2.5 cursor-pointer transition-colors"
                  style={{ background: "var(--surface-raised)" }}
                >
                  <input
                    type="checkbox"
                    className="sr-only"
                    checked={checked}
                    onChange={() => toggleRule(val)}
                  />
                  <span
                    className="rule-toggle-dot flex h-4 w-4 shrink-0 items-center justify-center rounded-full text-white transition"
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
                    className="text-sm"
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

      {/* Footer */}
      <div
        className="mt-auto pt-2 border-t"
        style={{ borderColor: "var(--border)" }}
      >
        <p
          className="text-[10px] text-center font-semibold"
          style={{ color: "var(--text-muted)" }}
        >
          &copy; 2026 City College of Calamba - OVPREPQA. All rights reserved.
        </p>
        <p
          className="text-[10px] text-center"
          style={{ color: "var(--text-muted)" }}
        >
          Developed by Railey Dela Peña
        </p>
      </div>
    </aside>
  );
}
