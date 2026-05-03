import { useEffect, useState } from "react";

const HOLD_MS = 2400;
const FADE_MS = 550;

interface Props {
  onFinish: () => void;
}

/** Inline SVG recreation of the e-Formatter logo icon */
function EFormatterIcon({ size = 64 }: { size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 100 100"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden="true"
    >
      {/* ── Laptop base (bottom dish) ── */}
      <path
        d="M8 80 Q8 88 16 88 H84 Q92 88 92 80 L88 76 H12 Z"
        fill="white"
        opacity="0.9"
      />

      {/* ── Laptop screen body ── */}
      <rect
        x="14"
        y="14"
        width="72"
        height="60"
        rx="5"
        ry="5"
        stroke="white"
        strokeWidth="4.5"
        fill="none"
      />

      {/* ── Browser/folder tab on top ── */}
      <path
        d="M14 14 L14 10 Q14 6 18 6 L36 6 Q40 6 42 10 L44 14"
        stroke="white"
        strokeWidth="4.5"
        fill="none"
        strokeLinejoin="round"
      />
      {/* Tab dots */}
      <circle cx="21" cy="10" r="1.8" fill="white" opacity="0.7" />
      <circle cx="27" cy="10" r="1.8" fill="white" opacity="0.7" />
      <circle cx="33" cy="10" r="1.8" fill="white" opacity="0.7" />

      {/* ── Gear outer ring ── */}
      <circle
        cx="35"
        cy="45"
        r="13"
        stroke="white"
        strokeWidth="3.5"
        fill="none"
      />

      {/* ── Gear teeth (8) ── */}
      {[0, 45, 90, 135, 180, 225, 270, 315].map((deg) => {
        const rad = (deg * Math.PI) / 180;
        const x1 = 35 + Math.cos(rad) * 13;
        const y1 = 45 + Math.sin(rad) * 13;
        const x2 = 35 + Math.cos(rad) * 17.5;
        const y2 = 45 + Math.sin(rad) * 17.5;
        return (
          <line
            key={deg}
            x1={x1}
            y1={y1}
            x2={x2}
            y2={y2}
            stroke="white"
            strokeWidth="3.8"
            strokeLinecap="round"
          />
        );
      })}

      {/* ── Inner gear hub ── */}
      <circle
        cx="35"
        cy="45"
        r="7"
        stroke="white"
        strokeWidth="2.5"
        fill="none"
      />

      {/* ── "e" letterform inside gear ── */}
      <path
        d="M30.5 45 H39.5 M31 42 Q30 49 36 50 Q40 50 40 47"
        stroke="white"
        strokeWidth="2.6"
        strokeLinecap="round"
        strokeLinejoin="round"
        fill="none"
      />

      {/* ── Circuit connector dot ── */}
      <circle cx="51" cy="33" r="2" fill="white" opacity="0.45" />
      <line
        x1="51"
        y1="35"
        x2="51"
        y2="40"
        stroke="white"
        strokeWidth="1.5"
        opacity="0.35"
      />

      {/* ── Text lines right side ── */}
      <rect
        x="54"
        y="33"
        width="24"
        height="4"
        rx="2"
        fill="white"
        opacity="0.85"
      />
      <rect
        x="54"
        y="41"
        width="20"
        height="4"
        rx="2"
        fill="white"
        opacity="0.65"
      />
      <rect
        x="54"
        y="49"
        width="22"
        height="4"
        rx="2"
        fill="white"
        opacity="0.65"
      />
      <rect
        x="54"
        y="57"
        width="16"
        height="4"
        rx="2"
        fill="white"
        opacity="0.45"
      />

      {/* ── Bottom bar ── */}
      <rect
        x="20"
        y="67"
        width="60"
        height="3.5"
        rx="1.75"
        fill="white"
        opacity="0.25"
      />
    </svg>
  );
}

export default function SplashScreen({ onFinish }: Props) {
  const [phase, setPhase] = useState<"enter" | "visible" | "exit">("enter");

  useEffect(() => {
    const t0 = requestAnimationFrame(() => setPhase("visible"));
    const t1 = setTimeout(() => setPhase("exit"), HOLD_MS);
    const t2 = setTimeout(onFinish, HOLD_MS + FADE_MS);
    return () => {
      cancelAnimationFrame(t0);
      clearTimeout(t1);
      clearTimeout(t2);
    };
  }, [onFinish]);

  return (
    <>
      <style>{`
        @keyframes sp-ring {
          0%   { transform: scale(1);   opacity: 0.6; }
          100% { transform: scale(3.8); opacity: 0;   }
        }
        @keyframes sp-logo-in {
          from { transform: scale(0.38) translateY(22px); opacity: 0; }
          to   { transform: scale(1)    translateY(0);    opacity: 1; }
        }
        @keyframes sp-text-in {
          from { transform: translateY(16px); opacity: 0; }
          to   { transform: translateY(0);    opacity: 1; }
        }
        @keyframes sp-fade-up {
          from { transform: translateY(10px); opacity: 0; }
          to   { transform: translateY(0);    opacity: 1; }
        }
        @keyframes sp-glow-pulse {
          from {
            box-shadow:
              0 0 0 1px rgba(59,130,246,0.30),
              0 0 26px rgba(37,99,235,0.50),
              0 0 60px rgba(37,99,235,0.20),
              0 14px 44px rgba(0,0,0,0.60),
              inset 0 1px 0 rgba(255,255,255,0.10);
          }
          to {
            box-shadow:
              0 0 0 1px rgba(96,165,250,0.48),
              0 0 50px rgba(59,130,246,0.75),
              0 0 100px rgba(37,99,235,0.32),
              0 14px 44px rgba(0,0,0,0.60),
              inset 0 1px 0 rgba(255,255,255,0.18);
          }
        }
        @keyframes sp-bar-fill {
          from { width: 0%; }
          to   { width: 100%; }
        }
        @keyframes sp-orb-drift {
          0%,100% { transform: translate(0,0); }
          50%     { transform: translate(16px, 12px); }
        }
        @keyframes sp-grid-in {
          from { opacity: 0; }
          to   { opacity: 1; }
        }

        /* ── Root shell ── */
        .sp-root {
          position: fixed;
          inset: 0;
          z-index: 99999;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          overflow: hidden;
          font-family: 'Inter', system-ui, -apple-system, sans-serif;
          /* Matches app dark: --bg-page #0d1117 → --status-bg #1e293b */
          background: linear-gradient(158deg, #0f172a 0%, #1e293b 50%, #0d1117 100%);
          opacity: 0;
          transition: opacity ${FADE_MS}ms cubic-bezier(0.4,0,0.2,1);
          user-select: none;
          -webkit-user-select: none;
        }
        .sp-root.sp-visible { opacity: 1; }
        .sp-root.sp-exit    { opacity: 0; pointer-events: none; }

        /* ── Dot grid (mirrors app blob texture idea) ── */
        .sp-grid {
          position: absolute;
          inset: 0;
          background-image:
            linear-gradient(rgba(37,99,235,0.065) 1px, transparent 1px),
            linear-gradient(90deg, rgba(37,99,235,0.065) 1px, transparent 1px);
          background-size: 46px 46px;
          animation: sp-grid-in 0.9s ease both;
        }

        /* ── Floating orbs (matches app --blob-* colors in dark) ── */
        .sp-orb {
          position: absolute;
          border-radius: 50%;
          pointer-events: none;
          filter: blur(70px);
        }
        .sp-orb-1 {
          width: 400px; height: 400px;
          top: -130px; left: -110px;
          background: rgba(29, 78, 216, 0.22);   /* --blob-1 dark */
          animation: sp-orb-drift 10s ease-in-out infinite;
        }
        .sp-orb-2 {
          width: 320px; height: 320px;
          bottom: -90px; right: -90px;
          background: rgba(79, 70, 229, 0.18);   /* --blob-2 dark */
          animation: sp-orb-drift 13s ease-in-out infinite reverse;
        }
        .sp-orb-3 {
          width: 220px; height: 220px;
          top: 50%; left: 50%;
          transform: translate(-50%, -50%);
          background: rgba(3, 105, 161, 0.14);   /* --blob-3 dark */
          filter: blur(55px);
        }

        /* ── Pulse rings ── */
        .sp-ring {
          position: absolute;
          width: 144px; height: 144px;
          border-radius: 50%;
          border: 1px solid rgba(59,130,246,0.25);
          animation: sp-ring 3s ease-out infinite;
        }
        .sp-ring-2 { animation-delay: 0.9s;  border-color: rgba(59,130,246,0.13); }
        .sp-ring-3 { animation-delay: 1.8s;  border-color: rgba(59,130,246,0.07); }

        /* ── Logo group ── */
        .sp-logo-wrap {
          position: relative;
          z-index: 10;
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 24px;
          animation: sp-logo-in 0.8s cubic-bezier(0.34,1.56,0.64,1) 0.1s both;
        }

        /* ── Icon box (glassmorphism + accent blue glow) ── */
        .sp-icon-box {
          width: 116px; height: 116px;
          border-radius: 34px;
          display: flex;
          align-items: center;
          justify-content: center;
          /* Gradient echoes --accent (#2563eb) in dark mode */
          background: linear-gradient(148deg,
            rgba(37,99,235,0.50) 0%,
            rgba(30,64,175,0.65) 100%
          );
          border: 1.5px solid rgba(147,197,253,0.20);
          backdrop-filter: blur(18px);
          -webkit-backdrop-filter: blur(18px);
          animation: sp-glow-pulse 2.4s ease-in-out 1s infinite alternate;
        }

        /* ── Text group ── */
        .sp-text {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 6px;
          animation: sp-text-in 0.6s cubic-bezier(0.22,1,0.36,1) 0.55s both;
        }

        /* App name — matches app font weight 800 usage */
        .sp-name {
          font-size: 33px;
          font-weight: 800;
          letter-spacing: -0.6px;
          color: #e2e8f0;              /* ~--text-primary dark */
          line-height: 1;
          margin: 0;
        }
        /* "e-" in app's --accent-light #93c5fd */
        .sp-name em {
          font-style: normal;
          color: #93c5fd;
        }

        /* Tagline matches app's small uppercase tracking labels */
        .sp-tagline {
          font-size: 10px;
          font-weight: 500;
          letter-spacing: 3.2px;
          text-transform: uppercase;
          color: rgba(148,163,184,0.60);   /* --text-muted dark */
          margin: 0;
        }

        /* ── Institution badge ── */
        .sp-badge {
          margin-top: 8px;
          display: inline-flex;
          align-items: center;
          gap: 7px;
          padding: 5px 13px;
          border-radius: 999px;
          border: 1px solid rgba(59,130,246,0.20);
          background: rgba(37,99,235,0.10);
          animation: sp-fade-up 0.45s ease 1s both;
        }
        .sp-badge-dot {
          width: 5px; height: 5px;
          border-radius: 50%;
          background: #3b82f6;
          box-shadow: 0 0 5px rgba(59,130,246,0.9);
        }
        .sp-badge-label {
          font-size: 9.5px;
          font-weight: 600;
          letter-spacing: 0.3px;
          color: rgba(147,197,253,0.70);  /* --accent-light dark */
        }

        /* ── Progress bar ── */
        .sp-bar-wrap {
          position: absolute;
          bottom: 50px;
          left: 50%;
          transform: translateX(-50%);
          width: 76px;
          opacity: 0;
          animation: sp-fade-up 0.35s ease 1.2s forwards;
        }
        .sp-bar-track {
          width: 100%;
          height: 2px;
          border-radius: 2px;
          background: rgba(255,255,255,0.07);
          overflow: hidden;
        }
        .sp-bar-fill {
          height: 100%;
          border-radius: 2px;
          /* Matches --accent → --accent-light */
          background: linear-gradient(90deg, #2563eb 0%, #93c5fd 100%);
          animation: sp-bar-fill ${HOLD_MS - 1000}ms cubic-bezier(0.25,0,0.35,1) 1.2s both;
        }

        /* ── Dev credit ── */
        .sp-credit {
          position: absolute;
          bottom: 24px;
          font-size: 9px;
          font-weight: 500;
          letter-spacing: 0.3px;
          color: rgba(100,116,139,0.55);  /* --text-muted dark */
          opacity: 0;
          animation: sp-fade-up 0.35s ease 1.4s forwards;
        }
      `}</style>

      <div
        className={[
          "sp-root",
          phase === "visible" ? "sp-visible" : "",
          phase === "exit" ? "sp-exit" : "",
        ].join(" ")}
        aria-label="Loading e-Formatter"
        role="status"
      >
        {/* Background */}
        <div className="sp-grid" />
        <div className="sp-orb sp-orb-1" />
        <div className="sp-orb sp-orb-2" />
        <div className="sp-orb sp-orb-3" />

        {/* Pulse rings */}
        <div className="sp-ring" />
        <div className="sp-ring sp-ring-2" />
        <div className="sp-ring sp-ring-3" />

        {/* Logo + Text */}
        <div className="sp-logo-wrap">
          <div className="sp-icon-box">
            <EFormatterIcon size={70} />
          </div>

          <div className="sp-text">
            <h1 className="sp-name">
              <em>e-</em>Formatter
            </h1>
            <p className="sp-tagline">Digital Manuscript Processing</p>
            <div className="sp-badge">
              <div className="sp-badge-dot" />
              <span className="sp-badge-label">
                City College of Calamba · OVPREPQA
              </span>
            </div>
          </div>
        </div>

        {/* Progress bar */}
        <div className="sp-bar-wrap">
          <div className="sp-bar-track">
            <div className="sp-bar-fill" />
          </div>
        </div>

        <p className="sp-credit">Developed by Railey Dela Peña</p>
      </div>
    </>
  );
}
