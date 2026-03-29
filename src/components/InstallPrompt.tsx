import { useEffect, useRef, useState } from "react";

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

/** Returns true if the device is likely a phone/tablet. */
function isMobileDevice(): boolean {
  return (
    typeof window !== "undefined" &&
    window.matchMedia("(max-width: 768px)").matches
  );
}

/** Returns true if already running as an installed PWA. */
function isStandalone(): boolean {
  return (
    typeof window !== "undefined" &&
    (window.matchMedia("(display-mode: standalone)").matches ||
      ("standalone" in window.navigator &&
        (window.navigator as { standalone?: boolean }).standalone === true))
  );
}

const DISMISS_KEY = "install-prompt-dismissed";

export default function InstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] =
    useState<BeforeInstallPromptEvent | null>(null);
  const [visible, setVisible] = useState(false);
  const [closing, setClosing] = useState(false);
  const [installed, setInstalled] = useState(false);
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Register service worker
  useEffect(() => {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("/sw.js").catch(() => {
        /* SW registration failed silently */
      });
    }
  }, []);

  // Capture the beforeinstallprompt event — mobile only
  useEffect(() => {
    // Skip if: already installed, running standalone, or not mobile
    if (
      isStandalone() ||
      !isMobileDevice() ||
      sessionStorage.getItem(DISMISS_KEY)
    )
      return;

    const handler = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);
      // Small delay so the page finishes loading before showing the sheet
      timerRef.current = setTimeout(() => setVisible(true), 1200);
    };

    window.addEventListener("beforeinstallprompt", handler);
    return () => {
      window.removeEventListener("beforeinstallprompt", handler);
      if (timerRef.current) clearTimeout(timerRef.current);
    };
  }, []);

  const dismiss = (remember = true) => {
    setClosing(true);
    if (remember) sessionStorage.setItem(DISMISS_KEY, "1");
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
      setTimeout(() => dismiss(true), 1600);
    } else {
      dismiss(false);
    }
  };

  if (!visible) return null;

  return (
    <>
      {/* Backdrop */}
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

      {/* Sheet */}
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
          padding: "0 0 calc(env(safe-area-inset-bottom, 0px) + 20px)",
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

        {/* Close button */}
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
            lineHeight: 1,
          }}
        >
          <i className="fa-solid fa-xmark" />
        </button>

        <div style={{ padding: "12px 22px 0" }}>
          {installed ? (
            /* ── Installed confirmation ── */
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
                  width: 52,
                  height: 52,
                  borderRadius: "50%",
                  background: "rgba(16,185,129,0.15)",
                  color: "#10b981",
                  fontSize: 22,
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
            /* ── Install prompt ── */
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
                  src="/images/logo.webp"
                  alt="Manuscript Formatter"
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
                    manuscript-formatter.netlify.app
                  </p>
                </div>
              </div>

              {/* Feature pills */}
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
                  <i className="fa-solid fa-download" />
                  Install App
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
