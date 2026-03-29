import { useEffect, useState } from "react";

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

export default function InstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] =
    useState<BeforeInstallPromptEvent | null>(null);
  const [showPrompt, setShowPrompt] = useState(false);

  // Capture the install prompt before the browser shows its own
  useEffect(() => {
    const handler = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);
      setShowPrompt(true);
    };
    window.addEventListener("beforeinstallprompt", handler);
    return () => window.removeEventListener("beforeinstallprompt", handler);
  }, []);

  // Register service worker
  useEffect(() => {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("/sw.js").catch(() => {
        /* SW registration failed silently */
      });
    }
  }, []);

  const handleInstall = async () => {
    if (!deferredPrompt) return;
    await deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === "accepted") setShowPrompt(false);
    setDeferredPrompt(null);
  };

  if (!showPrompt) return null;

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
        boxShadow: "0 12px 40px rgba(0,0,0,0.25)",
        zIndex: 99999,
        maxWidth: "320px",
        width: "calc(100vw - 48px)",
      }}
    >
      {/* Close button */}
      <button
        onClick={() => setShowPrompt(false)}
        aria-label="Dismiss"
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
          alt="Manuscript Formatter"
          style={{
            width: 40,
            height: 40,
            borderRadius: 10,
            objectFit: "contain",
            border: "1px solid var(--border)",
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
            Publisher: manuscript-formatter.netlify.app
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
          lineHeight: 1.9,
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
          onClick={() => setShowPrompt(false)}
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
