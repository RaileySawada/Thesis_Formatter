import { useEffect, useMemo, useState } from "react";
import type {
  ConferenceFormat,
  ConferenceFormattingConfig,
} from "../constants";
import {
  DEFAULT_ACM_FORMATTING_CONFIG,
  DEFAULT_PUBLICATION_FORMATTING_CONFIG,
} from "../constants";
import ConferenceFormattingPanel from "./ConferenceFormattingPanel";

interface Props {
  open: boolean;
  onClose: () => void;
  format: ConferenceFormat;
  config: ConferenceFormattingConfig;
  onChange: (next: ConferenceFormattingConfig) => void;
  onReset: () => void;
}

export default function MobileConferenceStylesSheet({
  open,
  onClose,
  format,
  config,
  onChange,
  onReset,
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

  useEffect(() => {
    if (!open && visible && !closing) handleClose();
  }, [open]);

  const isConfigChanged = useMemo(() => {
    const defaults =
      format === "pubform"
        ? DEFAULT_PUBLICATION_FORMATTING_CONFIG
        : DEFAULT_ACM_FORMATTING_CONFIG;
    return JSON.stringify(config[format]) !== JSON.stringify(defaults);
  }, [config, format]);

  if (!visible) return null;

  return (
    <div
      className={`sheet-backdrop${visible ? " open" : ""}`}
      onClick={(e) => {
        if (e.target === e.currentTarget) handleClose();
      }}
    >
      <div className={`sheet-modal${closing ? " closing" : ""}`}>
        <div className="sheet-drag-handle" />

        <div className="px-5 pt-2 pb-2 shrink-0">
          <div className="flex items-center justify-between">
            <div>
              <p
                className="text-[10px] font-bold uppercase tracking-[0.2em]"
                style={{ color: "var(--accent)" }}
              >
                <i className="fa-solid fa-wand-magic-sparkles mr-1" /> Styles
              </p>
              <h2
                className="text-xl font-bold mt-0.5"
                style={{ color: "var(--text-primary)" }}
              >
                Conference Configuration
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
          <p
            className="mt-1 text-xs font-semibold uppercase tracking-wider"
            style={{ color: "var(--accent)" }}
          >
            Format: {format === "pubform" ? "Publication Form" : "ACM Conference"}
          </p>
        </div>

        <div className="flex-1 overflow-y-auto px-5 pb-6 mt-2 custom-scrollbar">
          <ConferenceFormattingPanel
            format={format}
            config={config}
            onChange={onChange}
          />
        </div>

        <div
          className="shrink-0 border-t px-5 py-4 safe-area-bottom"
          style={{ borderColor: "var(--border)", background: "var(--surface)" }}
        >
          <div className="flex gap-3">
            {isConfigChanged && (
              <button
                onClick={onReset}
                className="flex-1 rounded-2xl border py-3.5 text-sm font-bold transition active:scale-[0.98]"
                style={{
                  borderColor: "var(--border)",
                  color: "var(--text-muted)",
                  background: "var(--surface-raised)",
                }}
                type="button"
              >
                Restore Defaults
              </button>
            )}
            <button
              onClick={handleClose}
              className="flex-1 rounded-2xl py-3.5 text-sm font-bold transition active:scale-[0.98]"
              style={{
                background: "var(--accent)",
                color: "#ffffff",
                boxShadow: "0 4px 12px var(--accent-glow)",
              }}
            >
              Save Changes
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

