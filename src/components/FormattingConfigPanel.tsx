import React, { useState, useMemo } from "react";
import type { FormattingConfig, ElementStyle } from "../constants";

interface Props {
  config: FormattingConfig;
  onChange: (newConfig: FormattingConfig) => void;
  citationStyle: "ieee" | "apa";
}

const FONT_FAMILIES = [
  "Garamond",
  "Arial",
  "Times New Roman",
  "Calibri",
  "Courier New",
  "Georgia",
  "Verdana",
];

const FONT_STACKS: Record<string, string> = {
  "Garamond": "'Garamond', 'EB Garamond', serif",
  "Arial": "'Arial', 'Arimo', sans-serif",
  "Times New Roman": "'Times New Roman', 'Tinos', serif",
  "Calibri": "'Calibri', 'Source Sans 3', sans-serif",
  "Courier New": "'Courier New', 'Cousine', monospace",
  "Georgia": "'Georgia', serif",
  "Verdana": "'Verdana', 'Inter', sans-serif",
};

const ALIGNMENTS = [
  { value: "left", icon: "fa-align-left" },
  { value: "center", icon: "fa-align-center" },
  { value: "right", icon: "fa-align-right" },
  { value: "both", icon: "fa-align-justify" },
];

/**
 * Audit & Scale Logic:
 * 1pt = 1.333px on standard screens.
 */
const PT_TO_PX_SCALE = 1.333 * 0.8;
const INCH_TO_PX = 72 * PT_TO_PX_SCALE;

const contextBodyCss = (alignment: any = "both") => ({
  fontFamily: FONT_STACKS["Garamond"],
  fontSize: `${12 * PT_TO_PX_SCALE}px`,
  lineHeight: 1.5,
  textAlign: alignment === "both" ? "justify" : alignment,
  color: "#334155"
});

// --- STABILIZED SUB-COMPONENTS (Defined outside to prevent re-mount on change) ---

const Stepper = ({ value, onChange, step = 1, min = 0, label }: {
  value: number;
  onChange: (v: number) => void;
  step?: number;
  min?: number;
  label: string;
}) => (
  <div className="space-y-1">
    <label className="text-[10px] font-bold uppercase tracking-wider text-soft" style={{ color: "var(--text-soft)" }}>
      {label}
    </label>
    <div className="flex items-center overflow-hidden rounded-xl border" style={{ borderColor: "var(--border)", background: "var(--surface)" }}>
      <button
        type="button"
        onClick={() => onChange(Number(Math.max(min, Number(value) - Number(step)).toFixed(2)))}
        className="flex h-9 w-9 items-center justify-center transition hover:bg-black/5"
        style={{ borderRight: "1px solid var(--border)", color: "var(--text-secondary)" }}
      >
        <i className="fa-solid fa-minus text-[10px]" />
      </button>
      <div className="flex-1 text-center text-xs font-bold" style={{ color: "var(--text-primary)" }}>
        {!isNaN(Number(value)) ? Number(value).toFixed(step % 1 === 0 ? 0 : 2) : value}
      </div>
      <button
        type="button"
        onClick={() => onChange(Number((Number(value) + Number(step)).toFixed(2)))}
        className="flex h-9 w-9 items-center justify-center transition hover:bg-black/5"
        style={{ borderLeft: "1px solid var(--border)", color: "var(--text-secondary)" }}
      >
        <i className="fa-solid fa-plus text-[10px]" />
      </button>
    </div>
  </div>
);

const ToggleGroup = ({ value, options, onChange, label }: {
  value: string;
  options: { value: string; icon: string }[];
  onChange: (v: any) => void;
  label: string;
}) => (
  <div className="space-y-1">
    <label className="text-[10px] font-bold uppercase tracking-wider text-soft" style={{ color: "var(--text-soft)" }}>
      {label}
    </label>
    <div className="flex rounded-xl border p-1" style={{ borderColor: "var(--border)", background: "var(--surface)" }}>
      {options.map((opt) => (
        <button
          key={opt.value}
          type="button"
          onClick={() => onChange(opt.value)}
          className="flex-1 flex h-8 items-center justify-center rounded-lg transition"
          style={value === opt.value
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

const MultiToggle = ({ items }: { items: { label: string; value: boolean; onChange: (v: boolean) => void; icon: string }[] }) => (
  <div className="grid grid-cols-2 gap-2 mt-2">
    {items.map((item) => (
      <button
        key={item.label}
        type="button"
        onClick={() => item.onChange(!item.value)}
        className="flex items-center justify-center gap-2 rounded-xl border h-10 transition font-bold text-xs"
        style={item.value
          ? { background: "var(--accent-subtle-strong)", borderColor: "var(--accent)", color: "var(--accent)" }
          : { background: "var(--surface)", borderColor: "var(--border)", color: "var(--text-muted)" }
        }
      >
        <i className={`fa-solid ${item.icon}`} />
        {item.label}
      </button>
    ))}
  </div>
);

const PreviewTable = ({ style }: { style: any }) => (
  <div className="overflow-hidden rounded-md border" style={{ borderColor: "#cbd5e1" }}>
    <table className="w-full border-collapse bg-white" style={{
      fontFamily: FONT_STACKS[style?.fontFamily] || style?.fontFamily || "inherit",
      fontSize: `${(style?.fontSize || 10) * PT_TO_PX_SCALE}px`,
      lineHeight: style?.lineSpacing || 1.0,
      textAlign: (style?.alignment === "both" ? "justify" : style?.alignment) || "center",
    }}>
      <thead>
        <tr style={{ background: "#f8fafc", borderBottom: "1px solid #cbd5e1" }}>
          <th className="p-2 border-r" style={{ borderColor: "#cbd5e1" }}>ID</th>
          <th className="p-2">Description</th>
        </tr>
      </thead>
      <tbody>
        <tr style={{ borderBottom: "1px solid #e2e8f0" }}>
          <td className="p-2 border-r" style={{ borderColor: "#e2e8f0" }}>1</td>
          <td className="p-2">Sample cell text following {style?.alignment} alignment.</td>
        </tr>
        <tr>
          <td className="p-2 border-r" style={{ borderColor: "#e2e8f0" }}>2</td>
          <td className="p-2">Row showing accurate {style?.fontSize}pt sizing.</td>
        </tr>
      </tbody>
    </table>
  </div>
);

const Preview = ({
  style,
  type = "body",
  title,
  config,
  citationStyle
}: {
  style: ElementStyle | any;
  type?: "body" | "title" | "heading" | "chapter" | "table" | "figure" | "references" | "tableContinuation" | "appendixContinuation" | "appendixLetter" | "caption" | "legend";
  title: string;
  config: FormattingConfig;
  citationStyle: "ieee" | "apa";
}) => {
  const cssStyle = (s: any) => ({
    fontFamily: FONT_STACKS[s?.fontFamily] || s?.fontFamily || "inherit",
    fontSize: `${(s?.fontSize || 12) * PT_TO_PX_SCALE}px`,
    lineHeight: s?.lineSpacing || 1.5,
    textAlign: (s?.alignment === "both" ? "justify" : s?.alignment) || "left",
    fontWeight: s?.bold ? "bold" : "normal",
    fontStyle: s?.italic ? "italic" : "normal",
    textTransform: s?.textTransform || "none",
    textIndent: s?.indentation ? `${s.indentation * INCH_TO_PX}px` : 0,
    marginBottom: s?.lineSpacing ? `${(s.lineSpacing - 1) * 4}px` : "0.125rem",
  });

  const activeStyle = useMemo(() => cssStyle(style), [style]);

  const renderContent = () => {
    switch (type) {
      case "title":
        return (
          <div className="text-center">
            <div style={activeStyle as any}>Title Title Title</div>
            <div className="text-left" style={{ ...contextBodyCss("both") as any, textIndent: `${0.5 * INCH_TO_PX}px` }}>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</div>
          </div>
        );
      case "heading":
        return (
          <div>
            <div style={activeStyle as any}>Heading</div>
            <div style={{ ...contextBodyCss("both") as any, textIndent: `${0.5 * INCH_TO_PX}px` }}>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore.</div>
          </div>
        );
      case "chapter":
        return (
          <div className="text-center">
            <div style={{ ...activeStyle as any, marginBottom: 0 }}>Chapter 1</div>
            <div style={{ ...cssStyle(config.titles) as any }}>TITLE TITLE TITLE</div>
            <div className="text-left" style={{ ...contextBodyCss("both") as any, textIndent: `${0.5 * INCH_TO_PX}px` }}>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore.</div>
          </div>
        );
      case "appendixLetter":
        return (
          <div className="text-center">
            <div style={{ ...activeStyle as any, marginBottom: 0 }}>Appendix X</div>
            <div style={{ ...cssStyle(config.titles) as any }}>TITLE TITLE TITLE</div>
            <div className="text-left" style={{ ...contextBodyCss("both") as any, textIndent: `${0.5 * INCH_TO_PX}px` }}>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore.</div>
          </div>
        );
      case "references":
        const isIEEE = citationStyle === "ieee";
        return (
          <div className="space-y-3" style={activeStyle as any}>
            <div style={{ paddingLeft: `${0.5 * INCH_TO_PX}px`, textIndent: `-${0.5 * INCH_TO_PX}px` }}>
              {isIEEE ? "[1] " : ""}Doe, J. (2024). Accuracy in bibliography layout. Journal of Thesis Engineering, 12(4), 101-105. Both styles now use correct points-to-pixel scale.
            </div>
            <div style={{ paddingLeft: `${0.5 * INCH_TO_PX}px`, textIndent: `-${0.5 * INCH_TO_PX}px` }}>
              {isIEEE ? "[2] " : ""}Smith, A. B. (2023). Professional document architecture for modern scholars. Advanced Formatting Press.
            </div>
          </div>
        );
      case "table":
        return <PreviewTable style={style} />;
      case "tableContinuation":
        return (
          <div className="space-y-2">
            <div style={activeStyle as any}>Continuation of Table X-X...</div>
            <PreviewTable style={config.table} />
          </div>
        );
      case "figure":
        return (
          <div className="flex flex-col items-center">
            <img
              src="/images/sample_figure.webp"
              alt="Sample Figure"
              className="max-w-full h-32 object-cover rounded shadow-sm border transition-all"
              style={{
                borderColor: "#334155",
                borderWidth: `${config.figure.borderWeight * PT_TO_PX_SCALE}px`,
                marginBottom: `${(config.figure.spacing - 1) * 10}px`
              }}
            />
            <div style={cssStyle(config.figureCaption) as any}>Figure X-X. Figure Caption.</div>
          </div>
        );
      case "caption":
        const isTable = title.includes("Table");
        return (
          <div className="space-y-2">
            <div style={activeStyle as any}>{isTable ? "Table X-X. Table Caption." : "Figure X-X. Figure Caption."}</div>
            {isTable && <PreviewTable style={config.table} />}
            {!isTable && <div style={contextBodyCss("both") as any}>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore.</div>}
          </div>
        );
      case "legend":
        return (
          <div className="space-y-2">
            <PreviewTable style={config.table} />
            <div style={activeStyle as any}>Legend: This is a sample legend text appearing immediately below the table as configured.</div>
          </div>
        );
      case "appendixContinuation":
        return (
          <div className="space-y-4">
            <div style={activeStyle as any}>Continuation of Appendix X...</div>
            <div style={{ ...contextBodyCss("both") as any, lineHeight: 2.0, textIndent: `${0.5 * INCH_TO_PX}px` }}>
              Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. This sample body paragraph now correctly uses 2.0 line spacing.
            </div>
          </div>
        );
      default:
        return (
          <div style={activeStyle as any}>
            Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
          </div>
        );
    }
  };

  return (
    <div className="mb-6 rounded-2xl border p-4 bg-white shadow-inner" style={{ borderColor: "var(--border)" }}>
      <div className="flex items-center justify-between mb-2">
        <span className="text-[10px] font-bold uppercase tracking-widest" style={{ color: "var(--accent)" }}>Live Preview: {title}</span>
        <i className="fa-solid fa-eye text-[10px]" style={{ color: "var(--text-muted)" }} />
      </div>
      <div
        className="p-3 rounded-lg border border-dashed transition-all overflow-hidden"
        style={{ borderColor: "#cbd5e1", minHeight: "80px", background: "#fff", color: "#334155" }}
      >
        {renderContent()}
      </div>
    </div>
  );
};

const Section = ({ id, title, icon, isExpanded, onToggle, children }: {
  id: string;
  title: string;
  icon: string;
  isExpanded: boolean;
  onToggle: () => void;
  children: React.ReactNode
}) => {
  return (
    <div className="mb-2 overflow-hidden rounded-2xl border transition-all duration-300"
      style={{ borderColor: isExpanded ? "var(--accent)" : "var(--border)", background: "var(--surface-raised)" }}>
      <button
        onClick={onToggle}
        className="flex w-full items-center justify-between px-4 py-3 text-left transition-colors"
        style={{ background: isExpanded ? "var(--accent-subtle)" : "transparent" }}
        type="button"
      >
        <div className="flex items-center gap-3">
          <span className="flex h-8 w-8 items-center justify-center rounded-xl transition"
            style={{ background: isExpanded ? "var(--accent-subtle-strong)" : "var(--surface)", color: isExpanded ? "var(--accent)" : "var(--text-muted)", boxShadow: isExpanded ? "none" : "0 1px 3px rgba(0,0,0,0.05)" }}>
            <i className={`fa-solid ${icon} text-sm`} />
          </span>
          <span className="text-sm font-bold" style={{ color: "var(--text-primary)" }}>{title}</span>
        </div>
        <i className={`fa-solid fa-chevron-down text-[10px] transition-transform duration-300 ${isExpanded ? "rotate-180" : ""}`}
          style={{ color: "var(--text-muted)" }} />
      </button>
      {isExpanded && <div className="p-4 pt-1 space-y-5">{children}</div>}
    </div>
  );
};

const StyleControls = ({
  elementKey,
  label,
  config,
  onChange,
  citationStyle,
  previewType = "body"
}: {
  elementKey: keyof FormattingConfig;
  label: string;
  config: FormattingConfig;
  onChange: (key: keyof FormattingConfig, updates: any) => void;
  citationStyle: "ieee" | "apa";
  previewType?: any;
}) => {
  const style = (config as any)[elementKey];
  if (!style) return null;

  return (
    <div className="space-y-4">
      <Preview style={style} type={previewType} title={label} config={config} citationStyle={citationStyle} />

      <div className="grid grid-cols-2 gap-4">
        <div className="space-y-1">
          <label className="text-[10px] font-bold uppercase tracking-wider text-soft" style={{ color: "var(--text-soft)" }}>Font Family</label>
          <div className="relative">
            <select
              value={style.fontFamily}
              onChange={(e) => onChange(elementKey, { fontFamily: e.target.value })}
              className="w-full appearance-none rounded-xl border bg-transparent px-3 py-2 text-xs font-bold outline-none transition focus:border-accent"
              style={{ borderColor: "var(--border)", background: "var(--surface)", color: "var(--text-primary)" }}
            >
              {FONT_FAMILIES.map((f) => <option key={f} value={f}>{f}</option>)}
            </select>
            <i className="fa-solid fa-chevron-down absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-[8px]" style={{ color: "var(--text-muted)" }} />
          </div>
        </div>
        <Stepper label="Font Size (pt)" value={style.fontSize} onChange={(v) => onChange(elementKey, { fontSize: v })} />
      </div>

      <div className="grid grid-cols-2 gap-4">
        <Stepper label="Line Spacing" step={0.1} min={1} value={style.lineSpacing} onChange={(v) => onChange(elementKey, { lineSpacing: v })} />
        {elementKey !== "table" && (
          <Stepper label="Indentation (in)" step={0.25} value={style.indentation} onChange={(v) => onChange(elementKey, { indentation: v })} />
        )}
      </div>

      <div className="grid grid-cols-1 gap-1">
        <ToggleGroup label="Alignment" value={style.alignment} options={ALIGNMENTS} onChange={(v) => onChange(elementKey, { alignment: v })} />
      </div>

      <MultiToggle items={[
        { label: "Bold", value: !!style.bold, icon: "fa-bold", onChange: (v) => onChange(elementKey, { bold: v }) },
        { label: "Italic", value: !!style.italic, icon: "fa-italic", onChange: (v) => onChange(elementKey, { italic: v }) },
      ]} />

      {elementKey === "titles" && (
        <MultiToggle items={[
          {
            label: "All Uppercase",
            value: (style as any).textTransform === "uppercase",
            icon: "fa-font-case",
            onChange: (v) => onChange(elementKey, { textTransform: v ? "uppercase" : "none" })
          },
        ]} />
      )}
    </div>
  );
};

// --- MAIN COMPONENT ---

export default function FormattingConfigPanel({ config, onChange, citationStyle }: Props) {
  const [expandedSection, setExpandedSection] = useState<string | null>("body");

  const updateElement = (key: keyof FormattingConfig, updates: Partial<ElementStyle | any>) => {
    onChange({
      ...config,
      [key]: {
        ...config[key],
        ...updates,
      },
    });
  };

  return (
    <div className="pb-4">
      <Section id="body" title="Body Paragraph" icon="fa-align-left" isExpanded={expandedSection === "body"} onToggle={() => setExpandedSection(expandedSection === "body" ? null : "body")}>
        <StyleControls elementKey="body" label="Body Paragraph" config={config} onChange={updateElement} citationStyle={citationStyle} />
      </Section>

      <Section id="references" title="References" icon="fa-quote-right" isExpanded={expandedSection === "references"} onToggle={() => setExpandedSection(expandedSection === "references" ? null : "references")}>
        <StyleControls elementKey="references" label="References" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="references" />
      </Section>

      <Section id="titles" title="Titles" icon="fa-heading" isExpanded={expandedSection === "titles"} onToggle={() => setExpandedSection(expandedSection === "titles" ? null : "titles")}>
        <StyleControls elementKey="titles" label="Titles" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="title" />
      </Section>

      <Section id="headings" title="Headings" icon="fa-list-ol" isExpanded={expandedSection === "headings"} onToggle={() => setExpandedSection(expandedSection === "headings" ? null : "headings")}>
        <StyleControls elementKey="headings" label="Headings" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="heading" />
      </Section>

      <Section id="chapters" title="Chapter Number" icon="fa-hashtag" isExpanded={expandedSection === "chapters"} onToggle={() => setExpandedSection(expandedSection === "chapters" ? null : "chapters")}>
        <StyleControls elementKey="chapterNumber" label="Chapter Number" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="chapter" />
      </Section>

      <Section id="table" title="Table" icon="fa-table" isExpanded={expandedSection === "table"} onToggle={() => setExpandedSection(expandedSection === "table" ? null : "table")}>
        <StyleControls elementKey="table" label="Table Content" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="table" />
      </Section>

      <Section id="legends" title="Table Legend" icon="fa-circle-info" isExpanded={expandedSection === "legends"} onToggle={() => setExpandedSection(expandedSection === "legends" ? null : "legends")}>
        <StyleControls elementKey="legends" label="Table Legend" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="legend" />
      </Section>

      <Section id="figure" title="Figure" icon="fa-image" isExpanded={expandedSection === "figure"} onToggle={() => setExpandedSection(expandedSection === "figure" ? null : "figure")}>
        <Preview style={config.figure as any} type="figure" title="Figure Assets" config={config} citationStyle={citationStyle} />
        <div className="grid grid-cols-2 gap-4 px-1">
          <Stepper label="Border Weight (pt)" step={0.25} value={config.figure.borderWeight} onChange={(v) => updateElement("figure", { borderWeight: v })} />
          <Stepper label="Line Spacing" step={0.5} min={1} value={config.figure.spacing} onChange={(v) => updateElement("figure", { spacing: v })} />
        </div>
      </Section>

      <Section id="tableCaption" title="Table Caption" icon="fa-closed-captioning" isExpanded={expandedSection === "tableCaption"} onToggle={() => setExpandedSection(expandedSection === "tableCaption" ? null : "tableCaption")}>
        <StyleControls elementKey="tableCaption" label="Table Caption" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="caption" />
      </Section>

      <Section id="figureCaption" title="Figure Caption" icon="fa-closed-captioning" isExpanded={expandedSection === "figureCaption"} onToggle={() => setExpandedSection(expandedSection === "figureCaption" ? null : "figureCaption")}>
        <StyleControls elementKey="figureCaption" label="Figure Caption" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="caption" />
      </Section>

      <Section id="tableContinuation" title="Table Continuation" icon="fa-table-columns" isExpanded={expandedSection === "tableContinuation"} onToggle={() => setExpandedSection(expandedSection === "tableContinuation" ? null : "tableContinuation")}>
        <StyleControls elementKey="tableContinuation" label="Table Continuation" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="tableContinuation" />
      </Section>

      <Section id="appendixLetter" title="Appendix Letter" icon="fa-font" isExpanded={expandedSection === "appendixLetter"} onToggle={() => setExpandedSection(expandedSection === "appendixLetter" ? null : "appendixLetter")}>
        <StyleControls elementKey="appendixLetter" label="Appendix Letter" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="appendixLetter" />
      </Section>

      <Section id="appendixContinuation" title="Appendix Continuation" icon="fa-rotate-right" isExpanded={expandedSection === "appendixContinuation"} onToggle={() => setExpandedSection(expandedSection === "appendixContinuation" ? null : "appendixContinuation")}>
        <StyleControls elementKey="appendixContinuation" label="Appendix Continuation" config={config} onChange={updateElement} citationStyle={citationStyle} previewType="appendixContinuation" />
      </Section>

    </div>
  );
}
