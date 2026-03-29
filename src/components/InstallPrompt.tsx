import { useEffect, useState } from "react";

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

export default function InstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] =
    useState<BeforeInstallPromptEvent | null>(null);
  const [showPrompt, setShowPrompt] = useState(false);

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
      navigator.serviceWorker.register("/sw.js");
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
        background: "#1e1e1e",
        color: "#fff",
        borderRadius: "12px",
        padding: "20px 24px",
        boxShadow: "0 8px 32px rgba(0,0,0,0.4)",
        zIndex: 9999,
        maxWidth: "320px",
        fontFamily: "sans-serif",
      }}
    >
      <button
        onClick={() => setShowPrompt(false)}
        style={{
          position: "absolute",
          top: "10px",
          right: "14px",
          background: "none",
          border: "none",
          color: "#aaa",
          fontSize: "18px",
          cursor: "pointer",
        }}
      >
        ✕
      </button>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "12px",
          marginBottom: "12px",
        }}
      >
        {/* Replace with your actual app icon */}
        <div
          style={{
            width: 40,
            height: 40,
            borderRadius: 8,
            background: "#1e40af",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: 20,
          }}
        >
          📄
        </div>
        <div>
          <div style={{ fontWeight: 700, fontSize: "15px" }}>
            Install ThesisFormatter
          </div>
          <div style={{ fontSize: "12px", color: "#aaa" }}>
            Publisher: your-domain.com
          </div>
        </div>
      </div>

      <p style={{ fontSize: "13px", marginBottom: "10px" }}>
        Use this site often? Install the app which:
      </p>
      <ul
        style={{
          fontSize: "13px",
          paddingLeft: "18px",
          marginBottom: "16px",
          lineHeight: 1.7,
        }}
      >
        <li>Opens in a focused window</li>
        <li>Has quick access options like pin to taskbar</li>
        <li>Works offline and syncs across devices</li>
      </ul>

      <div style={{ display: "flex", gap: "10px" }}>
        <button
          onClick={handleInstall}
          style={{
            flex: 1,
            padding: "10px",
            borderRadius: "8px",
            background: "#1e40af",
            color: "#fff",
            border: "none",
            fontWeight: 600,
            cursor: "pointer",
            fontSize: "14px",
          }}
        >
          Install
        </button>
        <button
          onClick={() => setShowPrompt(false)}
          style={{
            flex: 1,
            padding: "10px",
            borderRadius: "8px",
            background: "transparent",
            color: "#fff",
            border: "1px solid #555",
            cursor: "pointer",
            fontSize: "14px",
          }}
        >
          Not now
        </button>
      </div>
    </div>
  );
}
