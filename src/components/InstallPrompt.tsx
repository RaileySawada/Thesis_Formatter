import { useEffect, useRef, useState } from "react";

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

// Inline SVG fallback — prevents a failed network request for logo.webp
const LOGO_FALLBACK =
  "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='48' height='48' viewBox='0 0 48 48'%3E%3Crect width='48' height='48' rx='10' fill='%231e40af'/%3E%3Ctext x='50%25' y='55%25' dominant-baseline='middle' text-anchor='middle' font-size='24' fill='white'%3E%F0%9F%93%84%3C/text%3E%3C/svg%3E";

const DISMISS_KEY = "mf-install-dismissed";

/** true when already running as an installed PWA */
function isStandalone(): boolean {
  return (
    window.matchMedia("(display-mode: standalone)").matches ||
    ("standalone" in window.navigator &&
      (window.navigator as { standalone?: boolean }).standalone === true)
  );
}

/**
 * true only on phone / tablet.
 * Uses 1023px to match the app's Tailwind `lg` breakpoint (1024px)
 * so the sheet never appears when the desktop sidebar is visible.
 */
function isMobile(): boolean {
  return window.matchMedia("(max-width: 1023px)").matches;
}

/** Edge / Firefox Tracking Prevention can block sessionStorage — always wrap */
function ssGet(key: string): string | null {
  try {
    return sessionStorage.getItem(key);
  } catch {
    return null;
  }
}
function ssSet(key: string, val: string): void {
  try {
    sessionStorage.setItem(key, val);
  } catch {
    /* blocked — safe to ignore */
  }
}

export default function InstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] =
    useState<BeforeInstallPromptEvent | null>(null);
  const [visible, setVisible] = useState(false);
  const [closing, setClosing] = useState(false);
  const [installed, setInstalled] = useState(false);
  const [logoError, setLogoError] = useState(false);
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Register service worker — errors swallowed so the console stays clean
  useEffect(() => {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("/sw.js").catch(() => {});
    }
  }, []);

  /**
   * Attach the beforeinstallprompt listener ONLY when we're on mobile.
   *
   * Why this matters for the console warning:
   * The "Banner not shown: beforeinstallpromptEvent.preventDefault() called"
   * message is emitted by the browser the moment we call e.preventDefault().
   * By never attaching the listener on desktop, we never call preventDefault
   * on desktop, so the warning never appears there.
   */
  useEffect(() => {
    // Exit early on desktop, standalone PWA, or already-dismissed session
    if (isStandalone() || !isMobile() || ssGet(DISMISS_KEY)) return;

    const handler = (e: Event) => {
      // preventDefault is intentional on mobile — it suppresses the browser's
      // own mini install bar so our bottom sheet can appear instead.
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);
      // Wait for the page to finish its first render before sliding in
      timerRef.current = setTimeout(() => setVisible(true), 1200);
    };

    window.addEventListener("beforeinstallprompt", handler);
    return () => {
      window.removeEventListener("beforeinstallprompt", handler);
      if (timerRef.current) clearTimeout(timerRef.current);
    };
  }, []);

  const dismiss = (remember: boolean) => {
    setClosing(true);
    if (remember) ssSet(DISMISS_KEY, "1");
    setTimeout(() => {
      setVisible(false);
      setClosing(false);
    }, 340);
  };

  const handleInstall = async () => {
    if (!deferredPrompt) return;
    await deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    setDeferredPrompt(null);
    if (outcome === "accepted") {
      setInstalled(true);
      setTimeout(() => dismiss(true), 1800);
    } else {
      dismiss(false);
    }
  };

  if (!visible) return null;

  return (
    <>
      {/* ── Backdrop ── */}
      <div
        onClick={() => dismiss(false)}
        style={{
          position: "fixed",
          inset: 0,
          background: "rgba(0,0,0,0.45)",
          backdropFilter: "blur(2px)",
          WebkitBackdropFilter: "blur(2px)",
          zIndex: 99990,
          opacity: closing ? 0 : 1,
          transition: "opacity 0.32s ease",
        }}
      />

      {/* ── Bottom sheet ── */}
      <div
        role="dialog"
        aria-modal="true"
        aria-label="Install Manuscript Formatter"
        style={{
          position: "fixed",
          bottom: 0,
          left: 0,
          right: 0,
          zIndex: 99999,
          background: "var(--surface)",
          borderTop: "1.5px solid var(--border)",
          borderRadius: "24px 24px 0 0",
          paddingBottom: "calc(env(safe-area-inset-bottom, 0px) + 20px)",
          boxShadow: "0 -8px 48px rgba(0,0,0,0.35)",
          transform: closing ? "translateY(100%)" : "translateY(0)",
          transition: "transform 0.34s cubic-bezier(0.32,0.72,0,1)",
          animation: !closing
            ? "slideUpSheet 0.36s cubic-bezier(0.32,0.72,0,1)"
            : undefined,
        }}
      >
        <style>{`
          @keyframes slideUpSheet {
            from { transform: translateY(100%); }
            to   { transform: translateY(0); }
          }
        `}</style>

        {/* Drag handle */}
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            paddingTop: 12,
            paddingBottom: 4,
          }}
        >
          <div
            style={{
              width: 36,
              height: 4,
              borderRadius: 99,
              background: "var(--border)",
            }}
          />
        </div>

        {/* Close × */}
        <button
          onClick={() => dismiss(false)}
          aria-label="Dismiss"
          style={{
            position: "absolute",
            top: 18,
            right: 18,
            width: 30,
            height: 30,
            borderRadius: "50%",
            background: "var(--surface-raised)",
            border: "1px solid var(--border)",
            color: "var(--text-muted)",
            fontSize: 13,
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <i className="fa-solid fa-xmark" />
        </button>

        <div style={{ padding: "12px 22px 0" }}>
          {installed ? (
            /* ── Success confirmation ── */
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                textAlign: "center",
                gap: 12,
                padding: "16px 0 8px",
              }}
            >
              <span
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  width: 56,
                  height: 56,
                  borderRadius: "50%",
                  background: "rgba(16,185,129,0.15)",
                  color: "#10b981",
                  fontSize: 24,
                }}
              >
                <i className="fa-solid fa-circle-check" />
              </span>
              <div>
                <p
                  style={{
                    fontWeight: 700,
                    fontSize: 16,
                    color: "var(--text-primary)",
                  }}
                >
                  App installed!
                </p>
                <p
                  style={{
                    fontSize: 13,
                    color: "var(--text-muted)",
                    marginTop: 4,
                  }}
                >
                  Manuscript Formatter is now on your home screen.
                </p>
              </div>
            </div>
          ) : (
            /* ── Prompt ── */
            <>
              {/* Header */}
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 14,
                  marginBottom: 16,
                }}
              >
                <img
                  src={logoError ? LOGO_FALLBACK : "/images/logo.webp"}
                  alt="Manuscript Formatter"
                  onError={() => setLogoError(true)}
                  style={{
                    width: 48,
                    height: 48,
                    borderRadius: 12,
                    objectFit: "contain",
                    border: "1px solid var(--border)",
                    flexShrink: 0,
                  }}
                />
                <div>
                  <p
                    style={{
                      fontSize: 10,
                      fontWeight: 700,
                      textTransform: "uppercase",
                      letterSpacing: "0.2em",
                      color: "var(--accent)",
                      marginBottom: 2,
                    }}
                  >
                    <i
                      className="fa-solid fa-download"
                      style={{ marginRight: 4 }}
                    />
                    Add to Home Screen
                  </p>
                  <p
                    style={{
                      fontWeight: 700,
                      fontSize: 16,
                      color: "var(--text-primary)",
                      lineHeight: 1.2,
                    }}
                  >
                    Manuscript Formatter
                  </p>
                  <p
                    style={{
                      fontSize: 11,
                      color: "var(--text-muted)",
                      marginTop: 2,
                    }}
                  >
                    By Railey Dela Peña · CCC OVPREPQA
                  </p>
                </div>
              </div>

              {/* Feature rows */}
              <div
                style={{
                  display: "flex",
                  flexDirection: "column",
                  gap: 8,
                  marginBottom: 20,
                }}
              >
                {[
                  {
                    icon: "fa-bolt",
                    text: "Instant access from your home screen",
                  },
                  {
                    icon: "fa-wifi-slash",
                    text: "Works offline — no internet needed",
                  },
                  {
                    icon: "fa-expand",
                    text: "Opens full-screen, no browser chrome",
                  },
                ].map(({ icon, text }) => (
                  <div
                    key={text}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 12,
                      padding: "10px 14px",
                      borderRadius: 14,
                      background: "var(--surface-raised)",
                      border: "1px solid var(--border)",
                    }}
                  >
                    <span
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        width: 28,
                        height: 28,
                        borderRadius: 8,
                        background: "var(--accent-subtle)",
                        color: "var(--accent)",
                        fontSize: 12,
                        flexShrink: 0,
                      }}
                    >
                      <i className={`fa-solid ${icon}`} />
                    </span>
                    <span
                      style={{
                        fontSize: 13,
                        color: "var(--text-secondary)",
                        fontWeight: 500,
                      }}
                    >
                      {text}
                    </span>
                  </div>
                ))}
              </div>

              {/* Actions */}
              <div style={{ display: "flex", gap: 10 }}>
                <button
                  onClick={handleInstall}
                  style={{
                    flex: 1,
                    padding: "13px 16px",
                    borderRadius: 14,
                    background: "var(--accent)",
                    color: "#fff",
                    border: "none",
                    fontWeight: 700,
                    fontSize: 14,
                    cursor: "pointer",
                    boxShadow: "0 4px 16px var(--accent-glow)",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    gap: 8,
                  }}
                >
                  <i className="fa-solid fa-download" /> Install App
                </button>
                <button
                  onClick={() => dismiss(true)}
                  style={{
                    flex: 1,
                    padding: "13px 16px",
                    borderRadius: 14,
                    background: "var(--surface-raised)",
                    color: "var(--text-secondary)",
                    border: "1.5px solid var(--border)",
                    fontWeight: 600,
                    fontSize: 14,
                    cursor: "pointer",
                  }}
                >
                  Not now
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </>
  );
}
