import { useState, useEffect, useCallback, useRef } from "react";
import { formatDocx } from "./lib/FormatterEngine";
import { formatDocxApa } from "./lib/ApaFormatterEngine";
import { formatPreliminary } from "./lib/PreliminaryEngine";
import { formatAppendices } from "./lib/AppendicesEngine";
import { RULES_DEF } from "./constants";
import type { ToastMsg, CitationStyle } from "./constants";
import Sidebar from "./components/Sidebar";
import UploadZone from "./components/UploadZone";
import StatusPanel from "./components/StatusPanel";
import MobileSheet from "./components/MobileSheet";
import PreviewModal from "./components/PreviewModal";
import Toast from "./components/Toast";

// ── PWA: load JSZip from CDN ───────────────────────────────────────────────
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

// ── PWA: register service worker ───────────────────────────────────────────
function useServiceWorker() {
  useEffect(() => {
    if ("serviceWorker" in navigator) {
      window.addEventListener("load", () => {
        navigator.serviceWorker.register("/sw.js").catch(() => {
          // SW registration failed silently
        });
      });
    }
  }, []);
}

// ── PWA: capture install prompt ────────────────────────────────────────────
interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

function InstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] =
    useState<BeforeInstallPromptEvent | null>(null);
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const handler = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);
      setVisible(true);
    };
    window.addEventListener("beforeinstallprompt", handler);
    return () => window.removeEventListener("beforeinstallprompt", handler);
  }, []);

  const handleInstall = async () => {
    if (!deferredPrompt) return;
    await deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === "accepted") setVisible(false);
    setDeferredPrompt(null);
  };

  if (!visible) return null;

  return (
    <div
      style={{
        position: "fixed",
        bottom: "24px",
        right: "24px",
        background: "var(--surface-raised)",
        border: "1.5px solid var(--border)",
        borderRadius: "20px",
        padding: "20px 22px",
        boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        zIndex: 99999,
        maxWidth: "320px",
        width: "calc(100vw - 48px)",
        fontFamily: "inherit",
      }}
    >
      {/* Close */}
      <button
        onClick={() => setVisible(false)}
        style={{
          position: "absolute",
          top: "12px",
          right: "14px",
          background: "none",
          border: "none",
          color: "var(--text-muted)",
          fontSize: "16px",
          cursor: "pointer",
          lineHeight: 1,
        }}
        aria-label="Dismiss"
      >
        ✕
      </button>

      {/* Header */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "12px",
          marginBottom: "12px",
        }}
      >
        <img
          src="/images/logo.webp"
          alt="ThesisFormatter"
          style={{
            width: 40,
            height: 40,
            borderRadius: 10,
            objectFit: "contain",
          }}
        />
        <div>
          <div
            style={{
              fontWeight: 700,
              fontSize: "14px",
              color: "var(--text-primary)",
            }}
          >
            Install Manuscript Formatter
          </div>
          <div style={{ fontSize: "11px", color: "var(--text-muted)" }}>
            Publisher: ccc-ovprepqa.vercel.app
          </div>
        </div>
      </div>

      {/* Body */}
      <p
        style={{
          fontSize: "12px",
          color: "var(--text-soft)",
          marginBottom: "8px",
        }}
      >
        Use this site often? Install the app which:
      </p>
      <ul
        style={{
          fontSize: "12px",
          color: "var(--text-soft)",
          paddingLeft: "18px",
          marginBottom: "16px",
          lineHeight: 1.8,
        }}
      >
        <li>Opens in a focused window</li>
        <li>Has quick access like pin to taskbar</li>
        <li>Works offline on your device</li>
      </ul>

      {/* Actions */}
      <div style={{ display: "flex", gap: "10px" }}>
        <button
          onClick={handleInstall}
          style={{
            flex: 1,
            padding: "10px",
            borderRadius: "12px",
            background: "var(--accent)",
            color: "#fff",
            border: "none",
            fontWeight: 700,
            fontSize: "13px",
            cursor: "pointer",
            boxShadow: "0 4px 12px var(--accent-glow)",
          }}
        >
          Install
        </button>
        <button
          onClick={() => setVisible(false)}
          style={{
            flex: 1,
            padding: "10px",
            borderRadius: "12px",
            background: "transparent",
            color: "var(--text-secondary)",
            border: "1.5px solid var(--border)",
            fontWeight: 600,
            fontSize: "13px",
            cursor: "pointer",
          }}
        >
          Not now
        </button>
      </div>
    </div>
  );
}

// ── Main App ───────────────────────────────────────────────────────────────
export default function App() {
  const jsZipReady = useJSZip();
  useServiceWorker();

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

  // ── sections ─────────────────────────────────────────────────────────────
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

  // ── file ─────────────────────────────────────────────────────────────────
  const [file, setFile] = useState<File | null>(null);
  const [rulesOpen, setRulesOpen] = useState(true);
  const [citationStyle, setCitationStyle] = useState<CitationStyle>("ieee");

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

  // ── processing ───────────────────────────────────────────────────────────
  const [processing, setProcessing] = useState(false);
  const [toast, setToast] = useState<ToastMsg | null>(null);

  const showToast = useCallback(
    (msg: string, type: "success" | "error" = "success") => {
      setToast({ id: Date.now(), msg, type });
      setTimeout(() => setToast(null), 4700);
    },
    [],
  );

  const handleApply = useCallback(async () => {
    if (!file || !jsZipReady || processing) return;
    setProcessing(true);
    try {
      const runPreliminary = selectedSections.includes("preliminary");
      const runChapters = selectedSections.includes("chapters");
      const runAppendices = selectedSections.includes("appendices");

      let currentBuffer = await file.arrayBuffer();
      let blob: Blob | null = null;

      if (runPreliminary) {
        blob = await formatPreliminary(currentBuffer);
        currentBuffer = await blob.arrayBuffer();
      }

      if (runChapters) {
        if (citationStyle === "apa") {
          blob = await formatDocxApa(currentBuffer, {
            sections: selectedSections,
            rules: enabledRules,
          });
        } else {
          blob = await formatDocx(currentBuffer, {
            sections: selectedSections,
            rules: enabledRules,
            citationStyle,
          });
        }
        currentBuffer = await blob.arrayBuffer();
      }

      if (runAppendices) {
        blob = await formatAppendices(currentBuffer);
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
      showToast("Your manuscript has been formatted. Download started.");
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err), "error");
    } finally {
      setProcessing(false);
    }
  }, [
    file,
    jsZipReady,
    processing,
    selectedSections,
    enabledRules,
    citationStyle,
    showToast,
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
              alt="Manuscript Formatter"
              className="h-8 w-8 object-contain rounded-xl"
            />
            <span
              className="text-[11px] font-bold uppercase tracking-[0.2em]"
              style={{ color: "var(--accent)" }}
            >
              Manuscript Formatter
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
          {/* Sidebar */}
          <Sidebar
            isDark={isDark}
            toggleTheme={toggleTheme}
            selectedSections={selectedSections}
            toggleSection={toggleSection}
            enabledRules={enabledRules}
            toggleRule={toggleRule}
            rulesOpen={rulesOpen}
            setRulesOpen={setRulesOpen}
            citationStyle={citationStyle}
            setCitationStyle={setCitationStyle}
          />

          {/* Main content */}
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
                    Upload your manuscript and apply formatting rules for
                    chapters, references, figures, tables, and captions.
                  </p>
                </div>
                <div
                  className="flex shrink-0 gap-2"
                  style={{ position: "relative", zIndex: 10 }}
                >
                  {/* Download Template dropdown */}
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
                      Download Template
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
                        <a
                          href="/template/manuscript_template(IEEE).docx"
                          download="manuscript_template(IEEE).docx"
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
                              className="fa-solid fa-hashtag text-[11px]"
                              style={{ color: "var(--accent)" }}
                            />
                          </span>
                          <span className="flex-1">
                            <span
                              className="block font-bold"
                              style={{ color: "var(--text-primary)" }}
                            >
                              IEEE
                            </span>
                            <span
                              className="block text-[10px]"
                              style={{ color: "var(--text-soft)" }}
                            >
                              Numbered references [1]
                            </span>
                          </span>
                          <i
                            className="fa-solid fa-arrow-down text-[10px]"
                            style={{ color: "var(--text-muted)" }}
                          />
                        </a>
                        <div
                          style={{ height: "1px", background: "var(--border)" }}
                        />
                        <a
                          href="/template/manuscript_template(APA 7th Edition).docx"
                          download="manuscript_template(APA 7th Edition).docx"
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
                              className="fa-solid fa-user-pen text-[11px]"
                              style={{ color: "var(--accent)" }}
                            />
                          </span>
                          <span className="flex-1">
                            <span
                              className="block font-bold"
                              style={{ color: "var(--text-primary)" }}
                            >
                              APA 7th Edition
                            </span>
                            <span
                              className="block text-[10px]"
                              style={{ color: "var(--text-soft)" }}
                            >
                              Author-date (Smith, 2024)
                            </span>
                          </span>
                          <i
                            className="fa-solid fa-arrow-down text-[10px]"
                            style={{ color: "var(--text-muted)" }}
                          />
                        </a>
                      </div>
                    )}
                  </div>
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

              {/* Upload zone */}
              <UploadZone file={file} setFile={setFile} />

              {/* Badge */}
              <div className="mt-4 flex items-center gap-2">
                <span
                  className="text-xs font-medium"
                  style={{ color: "var(--text-soft)" }}
                >
                  Formatting:
                </span>
                {selectedSections.length === 0 ? (
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
                      Processing…
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
              {/* Coverage card */}
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
                  Normalizes document sections per the master template.
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

              {/* Status panel */}
              <StatusPanel
                selectedSections={selectedSections}
                sectionLabels={sectionLabels}
                file={file}
                enabledRules={enabledRules}
                totalRules={RULES_DEF.length}
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
        sectionLabels={sectionLabels}
        sectionIcons={sectionIcons}
        citationStyle={citationStyle}
        setCitationStyle={setCitationStyle}
      />

      <PreviewModal open={previewOpen} onClose={() => setPreviewOpen(false)} />

      {toast && (
        <Toast
          key={toast.id}
          msg={toast.msg}
          type={toast.type}
          onDismiss={() => setToast(null)}
        />
      )}

      {/* PWA Install Prompt */}
      <InstallPrompt />
    </div>
  );
}
