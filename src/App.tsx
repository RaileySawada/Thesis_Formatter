import { useState, useEffect, useCallback, useRef } from "react";
import { formatDocx } from "./lib/FormatterEngine";
import { formatDocxApa } from "./lib/ApaFormatterEngine";
import { formatDocxConference } from "./lib/ConferenceFormatterEngine";
import { formatPreliminary } from "./lib/PreliminaryEngine";
import { formatAppendices } from "./lib/AppendicesEngine";
import {
  RULES_DEF,
  DEFAULT_CONFIG_IEEE,
  DEFAULT_CONFIG_APA,
  CONFERENCE_FORMATS,
  DEFAULT_CONFERENCE_FORMATTING_CONFIG,
  DEFAULT_PUBLICATION_FORMATTING_CONFIG,
  DEFAULT_ACM_FORMATTING_CONFIG,
} from "./constants";
import type {
  ToastMsg,
  CitationStyle,
  FormattingConfig,
  FormattingStandard,
  ConferenceFormat,
  ConferenceFormattingConfig,
} from "./constants";
import Sidebar from "./components/Sidebar";
import FormattingConfigPanel from "./components/FormattingConfigPanel";
import ConferenceFormattingPanel from "./components/ConferenceFormattingPanel";
import UploadZone from "./components/UploadZone";
import StatusPanel from "./components/StatusPanel";
import MobileSheet from "./components/MobileSheet";
import PreviewModal from "./components/PreviewModal";
import MobileStylesSheet from "./components/MobileStylesSheet";
import MobileConferenceStylesSheet from "./components/MobileConferenceStylesSheet";
import Toast from "./components/Toast";
import { requestPollinations } from "./lib/pollinationsClient";
import { isAiAssistEnabled } from "./lib/aiAssist";
import SplashScreen from "./components/SplashScreen";

type AiRunStatus = "not-used" | "running" | "success" | "failed";
type ProcessLogLevel = "info" | "success" | "error";

interface ProcessLogEntry {
  id: number;
  at: string;
  level: ProcessLogLevel;
  message: string;
}

const AI_ASSIST_ENABLED = isAiAssistEnabled(
  import.meta.env.VITE_ENABLE_AI_ASSIST,
);

function formatDuration(ms: number): string {
  if (ms >= 1000) return `${(ms / 1000).toFixed(2)}s`;
  return `${Math.max(0, Math.round(ms))}ms`;
}

function formatLogTime(date = new Date()): string {
  return date.toLocaleTimeString([], { hour12: false });
}

function cloneDeep<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function mergeDeep<T>(defaults: T, source: unknown): T {
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    return cloneDeep(defaults);
  }

  const out: any = Array.isArray(defaults) ? [] : { ...(defaults as any) };
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

function toUserSafeErrorMessage(error: unknown): string {
  const raw =
    error instanceof Error ? error.message.trim() : String(error ?? "").trim();
  if (!raw) return "Formatting failed. Please try again.";

  if (/missing\s+pollinations_api_key/iu.test(raw)) {
    return "AI service key is missing in the server configuration.";
  }
  if (/failed to reach pollinations api/iu.test(raw)) {
    return "AI service is currently unreachable. Local formatting fallback was used.";
  }
  if (/pollinations api returned an error/iu.test(raw)) {
    return "AI service returned an error. Local formatting fallback was used.";
  }
  if (/unable to load .*format.*source file/iu.test(raw)) {
    return "The selected conference format file could not be loaded.";
  }

  return raw;
}

// PWA install prompt is handled entirely by InstallPrompt.tsx

// ── PWA: load JSZip from CDN ──────────────────────────────────────────────
function useJSZip() {
  const [ready, setReady] = useState(false);
  useEffect(() => {
    if ((window as any).JSZip) {
      setReady(true);
      return;
    }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = () => setReady(true);
    document.head.appendChild(s);
  }, []);
  return ready;
}

// Service worker registration is handled by InstallPrompt.tsx

import InstallPrompt from "./components/InstallPrompt";

// ── Main App ──────────────────────────────────────────────────────────────
export default function App() {
  const jsZipReady = useJSZip();

  // ── theme ────────────────────────────────────────────────────────────────
  const [isDark, setIsDark] = useState(() => {
    return (localStorage.getItem("thesis_theme") ?? "light") === "dark";
  });

  useEffect(() => {
    const theme = isDark ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", theme);
    document.body.setAttribute("data-theme", theme);
    document.documentElement.style.background = isDark ? "#0d1117" : "#f0f4ff";
    document.body.style.background = isDark ? "#0d1117" : "#f0f4ff";
    localStorage.setItem("thesis_theme", theme);
  }, [isDark]);

  const toggleTheme = useCallback(() => setIsDark((d) => !d), []);

  // ── responsive ─────────────────────────────────────────────────────────────
  const [isMobile, setIsMobile] = useState(() => window.innerWidth < 640);
  useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth < 640);
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);
  const [selectedSections, setSelectedSections] = useState<string[]>([
    "preliminary",
    "chapters",
    "appendices",
  ]);

  const toggleSection = useCallback((val: string) => {
    setSelectedSections((prev) => {
      if (prev.includes(val)) {
        if (prev.length <= 1) return prev;
        return prev.filter((s) => s !== val);
      }
      return [...prev, val];
    });
  }, []);

  // ── rules ────────────────────────────────────────────────────────────────
  const [enabledRules, setEnabledRules] = useState<string[]>(
    RULES_DEF.map(([val]) => val),
  );

  const toggleRule = useCallback((val: string) => {
    setEnabledRules((prev) =>
      prev.includes(val) ? prev.filter((r) => r !== val) : [...prev, val],
    );
  }, []);

  // ── components state ───────────────────────────────────────────────────
  const [formattingStandard, setFormattingStandard] =
    useState<FormattingStandard>(() => {
      const saved = localStorage.getItem("thesis_formatting_standard");
      return saved === "conference" ? "conference" : "ccc";
    });
  const [conferenceFormat, setConferenceFormat] = useState<ConferenceFormat>(
    () => {
      const saved = localStorage.getItem("thesis_conference_format");
      return saved === "pubform" ? "pubform" : "acm";
    },
  );
  const [citationStyle, setCitationStyle] = useState<CitationStyle>("ieee");

  useEffect(() => {
    localStorage.setItem("thesis_formatting_standard", formattingStandard);
  }, [formattingStandard]);

  useEffect(() => {
    localStorage.setItem("thesis_conference_format", conferenceFormat);
  }, [conferenceFormat]);

  // ── formatting config ────────────────────────────────────────────────────
  // Version key: bump this whenever defaults change structurally (e.g. new fields, changed defaults)
  const CONFIG_VERSION = "v3";

  const loadConfig = (
    key: string,
    defaults: FormattingConfig,
  ): FormattingConfig => {
    const saved = localStorage.getItem(key);
    const savedVersion = localStorage.getItem(key + "_version");
    if (saved && savedVersion === CONFIG_VERSION) {
      try {
        return JSON.parse(saved);
      } catch {
        return defaults;
      }
    }
    // Version mismatch or no saved config: merge saved values on top of new defaults
    // so new default fields (like borderWeight: 2.25) take effect
    if (saved && savedVersion !== CONFIG_VERSION) {
      try {
        const parsed = JSON.parse(saved);
        // Deep merge: defaults take precedence for nested objects, then overlay saved values
        const merged: FormattingConfig = { ...defaults };
        for (const k of Object.keys(defaults) as (keyof FormattingConfig)[]) {
          if (
            parsed[k] !== undefined &&
            typeof parsed[k] === "object" &&
            typeof (defaults[k] as any) === "object"
          ) {
            (merged as any)[k] = {
              ...(defaults[k] as any),
              ...(parsed[k] as any),
            };
          } else if (parsed[k] !== undefined) {
            (merged as any)[k] = parsed[k];
          }
        }
        localStorage.setItem(key + "_version", CONFIG_VERSION);
        return merged;
      } catch {
        return defaults;
      }
    }
    localStorage.setItem(key + "_version", CONFIG_VERSION);
    return defaults;
  };

  const [configIeee, setConfigIeee] = useState<FormattingConfig>(() =>
    loadConfig("thesis_formatting_config_ieee", DEFAULT_CONFIG_IEEE),
  );
  const [configApa, setConfigApa] = useState<FormattingConfig>(() =>
    loadConfig("thesis_formatting_config_apa", DEFAULT_CONFIG_APA),
  );
  const CONFERENCE_CONFIG_VERSION = "v1";
  const loadConferenceConfig = (): ConferenceFormattingConfig => {
    const key = "thesis_conference_formatting_config";
    const saved = localStorage.getItem(key);
    const savedVersion = localStorage.getItem(`${key}_version`);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        const merged = mergeDeep(DEFAULT_CONFERENCE_FORMATTING_CONFIG, parsed);
        localStorage.setItem(`${key}_version`, CONFERENCE_CONFIG_VERSION);
        return merged;
      } catch {
        // fall through to defaults
      }
    }
    if (savedVersion !== CONFERENCE_CONFIG_VERSION) {
      localStorage.setItem(`${key}_version`, CONFERENCE_CONFIG_VERSION);
    }
    return cloneDeep(DEFAULT_CONFERENCE_FORMATTING_CONFIG);
  };
  const [conferenceFormattingConfig, setConferenceFormattingConfig] =
    useState<ConferenceFormattingConfig>(() => loadConferenceConfig());

  useEffect(() => {
    localStorage.setItem(
      "thesis_formatting_config_ieee",
      JSON.stringify(configIeee),
    );
    localStorage.setItem(
      "thesis_formatting_config_ieee_version",
      CONFIG_VERSION,
    );
  }, [configIeee]);

  useEffect(() => {
    localStorage.setItem(
      "thesis_formatting_config_apa",
      JSON.stringify(configApa),
    );
    localStorage.setItem(
      "thesis_formatting_config_apa_version",
      CONFIG_VERSION,
    );
  }, [configApa]);

  useEffect(() => {
    localStorage.setItem(
      "thesis_conference_formatting_config",
      JSON.stringify(conferenceFormattingConfig),
    );
    localStorage.setItem(
      "thesis_conference_formatting_config_version",
      CONFERENCE_CONFIG_VERSION,
    );
  }, [conferenceFormattingConfig]);

  const formattingConfig = citationStyle === "apa" ? configApa : configIeee;
  const setFormattingConfig = (newConfig: FormattingConfig) => {
    if (citationStyle === "apa") setConfigApa(newConfig);
    else setConfigIeee(newConfig);
  };

  const activeConferenceDefaults =
    conferenceFormat === "pubform"
      ? DEFAULT_PUBLICATION_FORMATTING_CONFIG
      : DEFAULT_ACM_FORMATTING_CONFIG;
  const activeConferenceConfig = conferenceFormattingConfig[conferenceFormat];

  const isConfigChanged =
    formattingStandard === "conference"
      ? JSON.stringify(activeConferenceConfig) !==
        JSON.stringify(activeConferenceDefaults)
      : JSON.stringify(formattingConfig) !==
        JSON.stringify(
          citationStyle === "apa" ? DEFAULT_CONFIG_APA : DEFAULT_CONFIG_IEEE,
        );

  const handleResetStyles = () => {
    if (formattingStandard === "conference") {
      const prev = cloneDeep(conferenceFormattingConfig);
      const defaults =
        conferenceFormat === "pubform"
          ? cloneDeep(DEFAULT_PUBLICATION_FORMATTING_CONFIG)
          : cloneDeep(DEFAULT_ACM_FORMATTING_CONFIG);
      setConferenceFormattingConfig((current) => ({
        ...current,
        [conferenceFormat]: defaults,
      }));
      showToast("Styles restored to defaults.", "success", "Undo", () => {
        setConferenceFormattingConfig(prev);
        showToast("Restoration undone.");
      });
      return;
    }

    const prev = { ...(citationStyle === "apa" ? configApa : configIeee) };
    const defaults =
      citationStyle === "apa" ? DEFAULT_CONFIG_APA : DEFAULT_CONFIG_IEEE;

    setFormattingConfig(defaults);

    showToast("Styles restored to defaults.", "success", "Undo", () => {
      setFormattingConfig(prev);
      showToast("Restoration undone.");
    });
  };

  // ── file ─────────────────────────────────────────────────────────────────
  const [file, setFile] = useState<File | null>(null);
  const [rulesOpen, setRulesOpen] = useState(false);
  const [stylesModalOpen, setStylesModalOpen] = useState(false);

  // ── modals ───────────────────────────────────────────────────────────────
  const [mobileSheetOpen, setMobileSheetOpen] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [templateDropdownOpen, setTemplateDropdownOpen] = useState(false);
  const templateDropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!templateDropdownOpen) return;
    const handler = (e: MouseEvent) => {
      if (
        templateDropdownRef.current &&
        !templateDropdownRef.current.contains(e.target as Node)
      ) {
        setTemplateDropdownOpen(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [templateDropdownOpen]);

  useEffect(() => {
    if (formattingStandard === "conference") {
      setStylesModalOpen(false);
      setPreviewOpen(false);
      setRulesOpen(false);
    }
  }, [formattingStandard]);

  // ── processing ───────────────────────────────────────────────────────────
  const [processing, setProcessing] = useState(false);
  const [activeElapsedMs, setActiveElapsedMs] = useState(0);
  const [lastRunMs, setLastRunMs] = useState<number | null>(null);
  const [aiStatus, setAiStatus] = useState<AiRunStatus>("not-used");
  const [lastAiMs, setLastAiMs] = useState<number | null>(null);
  const [aiError, setAiError] = useState<string | null>(null);
  const [processLogs, setProcessLogs] = useState<ProcessLogEntry[]>([]);
  const [toast, setToast] = useState<ToastMsg | null>(null);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const processLogRef = useRef<HTMLDivElement | null>(null);

  const showToast = useCallback(
    (
      msg: string,
      type: "success" | "error" = "success",
      actionLabel?: string,
      onAction?: () => void,
    ) => {
      if (toastTimer.current) clearTimeout(toastTimer.current);

      setToast({ id: Date.now(), msg, type, actionLabel, onAction });

      toastTimer.current = setTimeout(() => {
        setToast(null);
        toastTimer.current = null;
      }, 5000);
    },
    [],
  );

  const pushProcessLog = useCallback(
    (level: ProcessLogLevel, message: string) => {
      setProcessLogs((prev) => [
        ...prev.slice(-79),
        {
          id: Date.now() + Math.floor(Math.random() * 1000),
          at: formatLogTime(),
          level,
          message,
        },
      ]);
    },
    [],
  );

  useEffect(() => {
    if (formattingStandard !== "conference") return;
    const node = processLogRef.current;
    if (!node) return;
    node.scrollTop = node.scrollHeight;
  }, [processLogs, formattingStandard]);

  useEffect(() => {
    if (!processing) return;
    const startedAt = performance.now();
    const tick = window.setInterval(() => {
      setActiveElapsedMs(Math.round(performance.now() - startedAt));
    }, 120);
    return () => window.clearInterval(tick);
  }, [processing]);

  const handleApply = useCallback(async () => {
    if (!file || !jsZipReady || processing) return;
    const startedAt = performance.now();
    setProcessing(true);
    setActiveElapsedMs(0);
    setLastAiMs(null);
    setAiError(null);
    setProcessLogs([
      {
        id: Date.now(),
        at: formatLogTime(),
        level: "info",
        message: `Run started for ${file.name}.`,
      },
    ]);

    let runAiStatus: AiRunStatus = AI_ASSIST_ENABLED ? "running" : "not-used";
    let runAiMs: number | null = null;
    setAiStatus(runAiStatus);

    try {
      if (AI_ASSIST_ENABLED) {
        pushProcessLog("info", "Checking AI assistant connectivity...");
        const aiStartedAt = performance.now();
        try {
          const aiResult = await requestPollinations({
            prompt:
              "Reply only with READY. This is a health check for a formatter workflow.",
            temperature: 0,
            maxTokens: 8,
          });
          void aiResult;
          runAiStatus = "success";
        } catch (aiErr) {
          runAiStatus = "failed";
          setAiError(toUserSafeErrorMessage(aiErr));
        } finally {
          runAiMs = Math.round(performance.now() - aiStartedAt);
          setLastAiMs(runAiMs);
          setAiStatus(runAiStatus);
          if (runAiStatus === "success") {
            pushProcessLog(
              "success",
              `AI check passed in ${formatDuration(runAiMs)}.`,
            );
          } else {
            pushProcessLog(
              "error",
              `AI check failed after ${formatDuration(runAiMs)}. Local fallback active.`,
            );
          }
        }
      } else {
        pushProcessLog(
          "info",
          "AI check disabled. Using local rule-based formatting.",
        );
      }

      const runPreliminary = selectedSections.includes("preliminary");
      const runChapters = selectedSections.includes("chapters");
      const runAppendices = selectedSections.includes("appendices");

      let currentBuffer = await file.arrayBuffer();
      let blob: Blob | null = null;

      if (formattingStandard === "conference") {
        const conferenceLabel =
          CONFERENCE_FORMATS.find((fmt) => fmt.value === conferenceFormat)
            ?.label ?? "Conference";
        pushProcessLog(
          "info",
          `Applying ${conferenceLabel} formatting to the full document...`,
        );
        blob = await formatDocxConference(currentBuffer, {
          format: conferenceFormat,
          styleConfig: conferenceFormattingConfig,
        });
        pushProcessLog("success", `${conferenceLabel} formatting completed.`);
      } else {
        if (runPreliminary) {
          pushProcessLog("info", "Formatting preliminary pages...");
          blob = await formatPreliminary(currentBuffer, {
            rules: enabledRules,
            config: formattingConfig,
          });
          currentBuffer = await blob.arrayBuffer();
          pushProcessLog("success", "Preliminary pages formatted.");
        }
        if (runChapters) {
          pushProcessLog("info", "Formatting chapters and references...");
          if (citationStyle === "apa") {
            blob = await formatDocxApa(currentBuffer, {
              sections: selectedSections,
              rules: enabledRules,
              config: formattingConfig,
            });
          } else {
            blob = await formatDocx(currentBuffer, {
              sections: selectedSections,
              rules: enabledRules,
              citationStyle,
              config: formattingConfig,
            });
          }
          currentBuffer = await blob.arrayBuffer();
          pushProcessLog("success", "Chapters and references formatted.");
        }
        if (runAppendices) {
          pushProcessLog("info", "Formatting appendices...");
          blob = await formatAppendices(currentBuffer, {
            rules: enabledRules,
            config: formattingConfig,
          });
          pushProcessLog("success", "Appendices formatted.");
        }
      }

      const docxMime =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      if (!blob) blob = new Blob([currentBuffer], { type: docxMime });
      else blob = new Blob([await blob.arrayBuffer()], { type: docxMime });

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const safeName = file.name.replace(/\.docx$/i, "") + "_formatted.docx";
      a.href = url;
      a.download = safeName;
      a.click();
      URL.revokeObjectURL(url);

      const totalMs = Math.round(performance.now() - startedAt);
      setLastRunMs(totalMs);

      const aiSummary =
        runAiStatus === "success" && runAiMs !== null
          ? ` AI check: success in ${formatDuration(runAiMs)}.`
          : runAiStatus === "failed"
            ? " AI check: failed. Local formatting fallback used."
            : " AI check: off (local rule-based formatting).";

      showToast(
        `Formatted in ${formatDuration(totalMs)}.${aiSummary} Download started.`,
      );
      pushProcessLog(
        "success",
        `Formatting finished in ${formatDuration(totalMs)}. Download started.`,
      );
    } catch (err) {
      const totalMs = Math.round(performance.now() - startedAt);
      setLastRunMs(totalMs);
      const safeError = toUserSafeErrorMessage(err);
      pushProcessLog("error", safeError);
      showToast(safeError, "error");
    } finally {
      setProcessing(false);
    }
  }, [
    file,
    jsZipReady,
    processing,
    selectedSections,
    enabledRules,
    formattingStandard,
    conferenceFormat,
    conferenceFormattingConfig,
    citationStyle,
    showToast,
    formattingConfig,
    pushProcessLog,
  ]);

  const sectionLabels: Record<string, string> = {
    preliminary: "Preliminary",
    chapters: "Chapter 1 – References",
    appendices: "Appendices",
  };
  const sectionIcons: Record<string, string> = {
    preliminary: "fa-file-lines",
    chapters: "fa-book-open",
    appendices: "fa-paperclip",
  };
  const cccFormatFiles = [
    {
      href: "/template/manuscript_template(IEEE).docx",
      download: "ieee_format_file.docx",
      label: "IEEE",
      sub: "Numbered references [1]",
      icon: "fa-hashtag",
    },
    {
      href: "/template/manuscript_template(APA 7th Edition).docx",
      download: "apa_7th_format_file.docx",
      label: "APA 7th Edition",
      sub: "Author-date (Smith, 2024)",
      icon: "fa-user-pen",
    },
  ];
  const availableFormatFiles =
    formattingStandard === "conference" ? CONFERENCE_FORMATS : cccFormatFiles;
  const activeConferenceFormat =
    CONFERENCE_FORMATS.find((fmt) => fmt.value === conferenceFormat) ??
    CONFERENCE_FORMATS[0];
  const activeStyleLabel =
    formattingStandard === "conference"
      ? activeConferenceFormat.label
      : citationStyle === "apa"
        ? "APA 7th Edition"
        : "IEEE";

  const theme = isDark ? "dark" : "light";

  return (
    <div
      id="formatter-app"
      data-theme={theme}
      className="relative min-h-screen overflow-x-hidden transition-colors duration-300"
      style={{ background: "var(--bg-page)" }}
    >
      {/* Blobs */}
      <div
        className="pointer-events-none absolute inset-0 overflow-hidden"
        aria-hidden="true"
      >
        <div
          className="blob-blue absolute -top-24 -left-16 h-80 w-80 rounded-full blur-3xl opacity-30"
          style={{ background: "var(--blob-1)" }}
        />
        <div
          className="blob-indigo absolute top-1/3 -right-20 h-96 w-96 rounded-full blur-3xl opacity-25"
          style={{ background: "var(--blob-2)" }}
        />
        <div
          className="blob-sky absolute bottom-0 left-1/3 h-72 w-72 rounded-full blur-3xl opacity-20"
          style={{ background: "var(--blob-3)" }}
        />
      </div>

      <div className="relative mx-auto max-w-7xl px-4 py-4 sm:px-6">
        {/* Mobile header */}
        <header className="mb-4 flex items-center justify-between lg:hidden">
          <div className="flex items-center gap-2">
            <img
              src="/images/logo.webp"
              alt="e-Formatter"
              className="h-10 w-10 object-contain shrink-0 rounded-xl"
              style={{
                filter: isDark ? "brightness(0) invert(1)" : "",
              }}
            />
            <span
              className="text-[11px] font-bold tracking-[0.2em]"
              style={{ color: "var(--accent)" }}
            >
              e-<span className="uppercase">Formatter</span>
            </span>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={toggleTheme}
              className="flex items-center gap-1.5 rounded-xl border px-3 py-2 text-xs font-semibold transition"
              style={{
                borderColor: "var(--border)",
                background: "var(--surface)",
                color: "var(--text-secondary)",
              }}
            >
              {isDark ? (
                <>
                  <i
                    className="fa-solid fa-moon fa-sm"
                    style={{ color: "#60a5fa" }}
                  />{" "}
                  Dark
                </>
              ) : (
                <>
                  <i
                    className="fa-solid fa-sun fa-sm"
                    style={{ color: "#f59e0b" }}
                  />{" "}
                  Light
                </>
              )}
            </button>
            <button
              onClick={() => setMobileSheetOpen(true)}
              className="flex items-center gap-2 rounded-2xl px-4 py-2 text-xs font-bold transition active:scale-95"
              style={{
                background: "var(--accent-subtle)",
                border: "1.5px solid var(--accent-muted)",
                color: "var(--accent)",
              }}
            >
              <i className="fa-solid fa-sliders text-sm" />
              Options
            </button>
          </div>
        </header>

        {/* Main layout */}
        <div className="flex w-full gap-6 items-start">
          <Sidebar
            isDark={isDark}
            toggleTheme={toggleTheme}
            selectedSections={selectedSections}
            toggleSection={toggleSection}
            enabledRules={enabledRules}
            toggleRule={toggleRule}
            rulesOpen={rulesOpen}
            setRulesOpen={setRulesOpen}
            formattingStandard={formattingStandard}
            setFormattingStandard={setFormattingStandard}
            citationStyle={citationStyle}
            setCitationStyle={setCitationStyle}
            conferenceFormat={conferenceFormat}
            setConferenceFormat={setConferenceFormat}
          />

          <main className="min-w-0 flex-1 space-y-5">
            {/* Upload card */}
            <div
              className="rounded-3xl border p-5 sm:p-7 transition-colors duration-300"
              style={{
                background: "var(--surface)",
                borderColor: "var(--border)",
                boxShadow: "var(--shadow)",
              }}
            >
              <div
                className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between"
                style={{ position: "relative" }}
              >
                <div>
                  <p
                    className="text-sm leading-7"
                    style={{ color: "var(--text-soft)" }}
                  >
                    {formattingStandard === "conference"
                      ? "Upload your manuscript and apply the selected conference formatting rule set to the full document."
                      : "Upload your manuscript and apply formatting rules for chapters, references, figures, tables, and captions."}
                  </p>
                </div>
                <div
                  className="flex flex-wrap shrink-0 gap-2"
                  style={{ position: "relative", zIndex: 10 }}
                >
                  <div className="relative" ref={templateDropdownRef}>
                    <button
                      type="button"
                      onClick={() => setTemplateDropdownOpen((v) => !v)}
                      className="z-0 inline-flex items-center gap-1.5 rounded-2xl px-4 py-2.5 text-xs font-bold text-white transition hover:opacity-90 active:scale-95 shadow-md"
                      style={{
                        background: "var(--accent)",
                        boxShadow: "0 4px 12px var(--accent-glow)",
                      }}
                    >
                      <i className="fa-solid fa-file-arrow-down text-xs" />
                      Download Format Files
                      <i
                        className={`fa-solid fa-chevron-down text-[9px] transition-transform duration-200${templateDropdownOpen ? " rotate-180" : ""}`}
                      />
                    </button>
                    {templateDropdownOpen && (
                      <div
                        className="template-dropdown absolute right-auto left-0 sm:right-0 sm:left-auto top-full mt-2 min-w-55 rounded-2xl border overflow-hidden"
                        style={{
                          background: "var(--surface-raised)",
                          borderColor: "var(--border)",
                          boxShadow: "0 12px 40px rgba(0,0,0,0.28)",
                          zIndex: 9999,
                        }}
                      >
                        {availableFormatFiles.map((formatFile, idx) => (
                          <div key={formatFile.href}>
                            <a
                              href={formatFile.href}
                              download={formatFile.download}
                              onClick={() => setTemplateDropdownOpen(false)}
                              className="template-dropdown-item flex items-center gap-3 px-4 py-3 text-xs font-semibold transition-colors"
                              style={{ color: "var(--text-primary)" }}
                            >
                              <span
                                className="flex h-7 w-7 shrink-0 items-center justify-center rounded-xl"
                                style={{
                                  background: "var(--accent-subtle-strong)",
                                }}
                              >
                                <i
                                  className={`fa-solid ${formatFile.icon} text-[11px]`}
                                  style={{ color: "var(--accent)" }}
                                />
                              </span>
                              <span className="flex-1">
                                <span
                                  className="block font-bold"
                                  style={{ color: "var(--text-primary)" }}
                                >
                                  {formatFile.label}
                                </span>
                                <span
                                  className="block text-[10px]"
                                  style={{ color: "var(--text-soft)" }}
                                >
                                  {formatFile.sub}
                                </span>
                              </span>
                              <i
                                className="fa-solid fa-arrow-down text-[10px]"
                                style={{ color: "var(--text-muted)" }}
                              />
                            </a>
                            {idx < availableFormatFiles.length - 1 && (
                              <div
                                style={{
                                  height: "1px",
                                  background: "var(--border)",
                                }}
                              />
                            )}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  <button
                    onClick={() => setStylesModalOpen(true)}
                    className="inline-flex items-center gap-1.5 rounded-2xl border px-4 py-2.5 text-xs font-semibold transition hover:opacity-80"
                    style={{
                      borderColor: "var(--border)",
                      color: "var(--accent)",
                      background: "var(--accent-subtle)",
                    }}
                  >
                    <i className="fa-solid fa-wand-magic-sparkles text-xs" />{" "}
                    Formatting Styles
                  </button>
                  <button
                    onClick={() => setPreviewOpen(true)}
                    className="inline-flex items-center gap-1.5 rounded-2xl border px-4 py-2.5 text-xs font-semibold transition hover:opacity-80"
                    style={{
                      borderColor: "var(--border)",
                      color: "var(--text-soft)",
                      background: "var(--surface-raised)",
                    }}
                  >
                    <i className="fa-solid fa-eye text-xs" /> Preview Rules
                  </button>
                </div>
              </div>

              <UploadZone file={file} setFile={setFile} />

              {/* Badge */}
              <div className="mt-4 flex items-center gap-2">
                <span
                  className="text-xs font-medium"
                  style={{ color: "var(--text-soft)" }}
                >
                  Formatting:
                </span>
                {formattingStandard === "conference" ? (
                  <span
                    className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-semibold"
                    style={{
                      background: "var(--accent-subtle-strong)",
                      color: "var(--accent)",
                    }}
                  >
                    <i
                      className={`fa-solid ${activeConferenceFormat.icon} text-[10px]`}
                    />
                    {activeConferenceFormat.label}
                  </span>
                ) : selectedSections.length === 0 ? (
                  <span
                    className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-semibold"
                    style={{
                      background: "rgba(239,68,68,.15)",
                      color: "#ef4444",
                    }}
                  >
                    <i className="fa-solid fa-triangle-exclamation text-[10px]" />{" "}
                    None selected
                  </span>
                ) : (
                  <span
                    className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-semibold"
                    style={{
                      background: "var(--accent-subtle-strong)",
                      color: "var(--accent)",
                    }}
                  >
                    <i
                      className={`fa-solid ${sectionIcons[selectedSections[0]] ?? "fa-layer-group"} text-[10px]`}
                    />
                    {selectedSections
                      .map((v) => sectionLabels[v] ?? v)
                      .join(", ")}
                  </span>
                )}
              </div>

              {/* Apply button */}
              <div className="mt-3">
                <button
                  onClick={handleApply}
                  disabled={!file || processing || !jsZipReady}
                  className="w-full rounded-2xl px-5 py-3.5 text-sm font-bold text-white transition active:scale-[0.98] shadow-md disabled:cursor-not-allowed disabled:opacity-40 disabled:shadow-none disabled:active:scale-100"
                  style={{
                    background: "var(--accent)",
                    boxShadow: "0 4px 14px var(--accent-glow)",
                  }}
                >
                  {processing ? (
                    <span className="inline-flex items-center justify-center gap-2">
                      <svg
                        className="animate-spin h-4 w-4 text-white inline-block"
                        xmlns="http://www.w3.org/2000/svg"
                        fill="none"
                        viewBox="0 0 24 24"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          stroke="currentColor"
                          strokeWidth="4"
                        />
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"
                        />
                      </svg>
                      Processing... {formatDuration(activeElapsedMs)}
                    </span>
                  ) : (
                    <span>
                      <i className="fa-solid fa-bolt mr-2" /> Apply Formatting
                    </span>
                  )}
                </button>
              </div>
            </div>

            {/* Coverage + Status row */}
            <div className="grid gap-5 xl:grid-cols-2">
              {formattingStandard === "conference" ? (
                <div
                  className="rounded-3xl border p-6 transition-colors duration-300"
                  style={{
                    background: "var(--surface)",
                    borderColor: "var(--border)",
                    boxShadow: "var(--shadow)",
                  }}
                >
                  <h2
                    className="text-lg font-bold sm:text-xl"
                    style={{ color: "var(--text-primary)" }}
                  >
                    <i
                      className="fa-solid fa-terminal mr-2"
                      style={{ color: "var(--accent)" }}
                    />
                    Process Log
                  </h2>
                  <p
                    className="mt-1.5 text-sm"
                    style={{ color: "var(--text-soft)" }}
                  >
                    Tracks formatting steps and user-safe run issues.
                  </p>
                  <div
                    ref={processLogRef}
                    className="mt-4 max-h-[330px] overflow-y-auto rounded-2xl border p-3"
                    style={{
                      background: "var(--surface-raised)",
                      borderColor: "var(--border)",
                    }}
                  >
                    {processLogs.length === 0 ? (
                      <p
                        className="text-xs"
                        style={{ color: "var(--text-muted)" }}
                      >
                        No process yet. Upload a file and run formatting.
                      </p>
                    ) : (
                      <div className="space-y-2">
                        {processLogs.map((log) => {
                          const levelColor =
                            log.level === "error"
                              ? "#ef4444"
                              : log.level === "success"
                                ? "#22c55e"
                                : "var(--text-muted)";
                          const levelIcon =
                            log.level === "error"
                              ? "fa-circle-exclamation"
                              : log.level === "success"
                                ? "fa-circle-check"
                                : "fa-circle-dot";
                          return (
                            <div
                              key={log.id}
                              className="rounded-xl border px-2.5 py-2"
                              style={{
                                borderColor: "var(--border)",
                                background: "var(--surface)",
                              }}
                            >
                              <p
                                className="mb-1 text-[10px] font-semibold"
                                style={{ color: "var(--text-muted)" }}
                              >
                                {log.at}
                              </p>
                              <p
                                className="text-xs leading-5"
                                style={{ color: "var(--text-secondary)" }}
                              >
                                <i
                                  className={`fa-solid ${levelIcon} mr-2`}
                                  style={{ color: levelColor }}
                                />
                                {log.message}
                              </p>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div
                  className="rounded-3xl border p-6 transition-colors duration-300"
                  style={{
                    background: "var(--surface)",
                    borderColor: "var(--border)",
                    boxShadow: "var(--shadow)",
                  }}
                >
                  <h2
                    className="text-lg font-bold sm:text-xl"
                    style={{ color: "var(--text-primary)" }}
                  >
                    <i
                      className="fa-solid fa-list-check mr-2"
                      style={{ color: "var(--accent)" }}
                    />
                    Formatting Coverage
                  </h2>
                  <p
                    className="mt-1.5 text-sm"
                    style={{ color: "var(--text-soft)" }}
                  >
                    Normalizes document sections based on the selected
                    formatting rule.
                  </p>
                  <div className="mt-4 space-y-2.5">
                    {[
                      {
                        key: "preliminary",
                        icon: "fa-file-circle-check",
                        label: "Preliminary Pages",
                        desc: "Title page, approval sheet, abstract, acknowledgement.",
                      },
                      {
                        key: "chapters",
                        icon: "fa-book-open",
                        label: "Chapters and References",
                        desc: "Chapter titles, headings, body text, figures, tables, captions, legends, references.",
                      },
                      {
                        key: "appendices",
                        icon: "fa-paperclip",
                        label: "Appendices",
                        desc: "Appendix labels, continuation blocks, User Manual, CV.",
                      },
                    ].map(({ key, icon, label, desc }) => {
                      const active = selectedSections.includes(key);
                      return (
                        <div
                          key={key}
                          className="rounded-2xl border-2 p-4 transition-colors"
                          style={
                            active
                              ? {
                                  background: "var(--accent-subtle)",
                                  borderColor: "var(--accent)",
                                }
                              : {
                                  background: "var(--surface-raised)",
                                  borderColor: "var(--border)",
                                }
                          }
                        >
                          <h3
                            className="text-sm font-semibold"
                            style={{
                              color: active
                                ? "var(--text-primary)"
                                : "var(--text-secondary)",
                            }}
                          >
                            <i
                              className={`fa-solid ${icon} mr-1.5`}
                              style={{
                                color: active
                                  ? "var(--accent)"
                                  : "var(--text-muted)",
                              }}
                            />
                            {label}
                            {active && (
                              <span
                                className="ml-2 rounded-full px-2 py-0.5 text-[10px] font-bold uppercase"
                                style={{
                                  background: "var(--accent-subtle-strong)",
                                  color: "var(--accent)",
                                }}
                              >
                                Active
                              </span>
                            )}
                          </h3>
                          <p
                            className="mt-1 text-xs"
                            style={{
                              color: active
                                ? "var(--text-secondary)"
                                : "var(--text-muted)",
                            }}
                          >
                            {desc}
                          </p>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              <StatusPanel
                formattingStandard={formattingStandard}
                selectedSections={selectedSections}
                sectionLabels={sectionLabels}
                file={file}
                enabledRules={enabledRules}
                totalRules={RULES_DEF.length}
                processing={processing}
                activeElapsedMs={activeElapsedMs}
                lastRunMs={lastRunMs}
                aiAssistEnabled={AI_ASSIST_ENABLED}
                aiStatus={aiStatus}
                lastAiMs={lastAiMs}
                aiError={aiError}
              />
            </div>

            {/* Footer mobile */}
            <div
              className="text-center text-xs pb-4 lg:hidden"
              style={{ color: "var(--text-muted)" }}
            >
              <p className="font-semibold">
                &copy; 2026 City College of Calamba - OVPREPQA. All rights
                reserved.
              </p>
              <p className="mt-1 text-[10px]">Developed by Railey Dela Peña</p>
            </div>
          </main>
        </div>
      </div>

      {/* Modals */}
      <MobileSheet
        open={mobileSheetOpen}
        onClose={() => setMobileSheetOpen(false)}
        selectedSections={selectedSections}
        toggleSection={toggleSection}
        enabledRules={enabledRules}
        toggleRule={toggleRule}
        rulesOpen={rulesOpen}
        setRulesOpen={setRulesOpen}
        sectionLabels={sectionLabels}
        sectionIcons={sectionIcons}
        formattingStandard={formattingStandard}
        setFormattingStandard={setFormattingStandard}
        citationStyle={citationStyle}
        setCitationStyle={setCitationStyle}
        conferenceFormat={conferenceFormat}
        setConferenceFormat={setConferenceFormat}
        onOpenStyles={() => setStylesModalOpen(true)}
        onOpenPreview={() => setPreviewOpen(true)}
      />
      <PreviewModal
        open={previewOpen}
        onClose={() => setPreviewOpen(false)}
        formattingStandard={formattingStandard}
        citationStyle={citationStyle}
        conferenceFormat={conferenceFormat}
        conferenceConfig={conferenceFormattingConfig}
      />

      {/* Styles Modal - Desktop */}
      {!isMobile && stylesModalOpen && (
        <div
          className="sheet-backdrop open"
          onClick={(e) => {
            if (e.target === e.currentTarget) setStylesModalOpen(false);
          }}
          style={{ backdropFilter: "blur(4px)" }}
        >
          <div
            className="sheet-modal w-full max-w-2xl max-h-[90vh]"
            style={{
              background: "var(--surface)",
              borderColor: "var(--border)",
            }}
          >
            <div
              className="flex items-center justify-between border-b px-6 py-4"
              style={{ borderColor: "var(--border)" }}
            >
              <div className="flex items-center gap-3">
                <span
                  className="flex h-10 w-10 items-center justify-center rounded-2xl"
                  style={{
                    background: "var(--accent-subtle)",
                    color: "var(--accent)",
                  }}
                >
                  <i className="fa-solid fa-wand-magic-sparkles text-lg" />
                </span>
                <div>
                  <h2
                    className="text-xl font-bold"
                    style={{ color: "var(--text-primary)" }}
                  >
                    Formatting Styles
                  </h2>
                  <p
                    className="text-xs font-semibold uppercase tracking-wider"
                    style={{ color: "var(--accent)" }}
                  >
                    Style: {activeStyleLabel}
                  </p>
                </div>
              </div>
              <button
                onClick={() => setStylesModalOpen(false)}
                className="flex h-10 w-10 items-center justify-center rounded-xl transition hover:bg-black/5"
                style={{ color: "var(--text-soft)" }}
              >
                <i className="fa-solid fa-xmark text-lg" />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-6 custom-scrollbar">
              {formattingStandard === "conference" ? (
                <ConferenceFormattingPanel
                  format={conferenceFormat}
                  config={conferenceFormattingConfig}
                  onChange={setConferenceFormattingConfig}
                />
              ) : (
                <FormattingConfigPanel
                  config={formattingConfig}
                  onChange={setFormattingConfig}
                  citationStyle={citationStyle}
                />
              )}
            </div>

            <div
              className="px-6 py-4 border-t flex items-center justify-end gap-3"
              style={{
                borderColor: "var(--border)",
                background: "var(--surface)",
              }}
            >
              {isConfigChanged && (
                <button
                  onClick={handleResetStyles}
                  className="rounded-2xl border px-5 py-2.5 text-xs font-bold transition hover:bg-black/5"
                  style={{
                    borderColor: "var(--border)",
                    color: "var(--text-muted)",
                  }}
                  type="button"
                >
                  Restore Defaults
                </button>
              )}
              <button
                onClick={() => setStylesModalOpen(false)}
                className="rounded-2xl px-6 py-2.5 text-xs font-bold text-white transition active:scale-95 shadow-md"
                style={{
                  background: "var(--accent)",
                  boxShadow: "0 4px 12px var(--accent-glow)",
                }}
                type="button"
              >
                Save Changes
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Styles Modal - Mobile */}
      {isMobile &&
        (formattingStandard === "conference" ? (
          <MobileConferenceStylesSheet
            open={stylesModalOpen}
            onClose={() => setStylesModalOpen(false)}
            format={conferenceFormat}
            config={conferenceFormattingConfig}
            onChange={setConferenceFormattingConfig}
            onReset={handleResetStyles}
          />
        ) : (
          <MobileStylesSheet
            open={stylesModalOpen}
            onClose={() => setStylesModalOpen(false)}
            config={formattingConfig}
            onChange={setFormattingConfig}
            citationStyle={citationStyle}
            onReset={handleResetStyles}
          />
        ))}

      {toast && (
        <Toast
          key={toast.id}
          msg={toast.msg}
          type={toast.type}
          actionLabel={toast.actionLabel}
          onAction={toast.onAction}
          onDismiss={() => setToast(null)}
        />
      )}

      {/* PWA Install Prompt */}
      <InstallPrompt isDark={isDark} />
    </div>
  );
}
