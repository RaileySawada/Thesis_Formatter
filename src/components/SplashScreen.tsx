import { useEffect, useState } from "react";

const HOLD_MS = 2000;
const FADE_MS = 480;

interface Props {
  onFinish: () => void;
}

export default function SplashScreen({ onFinish }: Props) {
  // Start visible immediately — no enter phase to avoid flash of app content
  const [phase, setPhase] = useState<"visible" | "exit">("visible");

  useEffect(() => {
    const exitTimer = window.setTimeout(() => setPhase("exit"), HOLD_MS);
    const finishTimer = window.setTimeout(onFinish, HOLD_MS + FADE_MS);

    return () => {
      window.clearTimeout(exitTimer);
      window.clearTimeout(finishTimer);
    };
  }, [onFinish]);

  return (
    <>
      <style>{`
        /* ── Keyframes ─────────────────────────────────────────── */

        @keyframes ef-logo-reveal {
          from { opacity: 0; transform: scale(0.88); }
          to   { opacity: 1; transform: scale(1); }
        }

        @keyframes ef-text-rise {
          from { opacity: 0; transform: translateY(14px); }
          to   { opacity: 1; transform: translateY(0); }
        }

        @keyframes ef-divider-expand {
          from { transform: scaleX(0); opacity: 0; }
          to   { transform: scaleX(1); opacity: 1; }
        }

        @keyframes ef-progress-fill {
          from { transform: scaleX(0); }
          to   { transform: scaleX(1); }
        }

        @keyframes ef-ray-sweep {
          0%   { opacity: 0; transform: translateX(-120%) skewX(-18deg); }
          30%  { opacity: 1; }
          70%  { opacity: 1; }
          100% { opacity: 0; transform: translateX(220%) skewX(-18deg); }
        }

        @keyframes ef-ambient-pulse {
          0%, 100% { opacity: 0.18; }
          50%       { opacity: 0.28; }
        }

        /* ── Root container ────────────────────────────────────── */

        .ef-splash {
          position: fixed;
          inset: 0;
          z-index: 99999;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          overflow: hidden;
          min-height: 100svh;
          padding:
            max(24px, env(safe-area-inset-top))
            max(20px, env(safe-area-inset-right))
            max(32px, env(safe-area-inset-bottom))
            max(20px, env(safe-area-inset-left));
          font-family: Inter, ui-sans-serif, system-ui, -apple-system,
                       BlinkMacSystemFont, "Segoe UI", sans-serif;
          /* Navy blue palette — matches system splash background_color */
          background:
            radial-gradient(ellipse 70% 50% at 50% 0%,   rgba(59, 130, 246, 0.14), transparent),
            radial-gradient(ellipse 55% 40% at 100% 100%, rgba(37,  99, 235, 0.10), transparent),
            linear-gradient(170deg, #0f1e3d 0%, #0a1628 55%, #050d1a 100%);
          opacity: 1;
          user-select: none;
          -webkit-user-select: none;
          transition: opacity ${FADE_MS}ms cubic-bezier(0.4, 0, 0.2, 1);
        }

        .ef-splash--exit { opacity: 0; pointer-events: none; }

        /* ── Ambient glow orb ──────────────────────────────────── */

        .ef-splash__orb {
          position: absolute;
          width: min(72vw, 520px);
          aspect-ratio: 1;
          border-radius: 50%;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.22), transparent 68%);
          filter: blur(40px);
          animation: ef-ambient-pulse 3.6s ease-in-out infinite;
          pointer-events: none;
        }

        /* ── Subtle light ray ──────────────────────────────────── */

        .ef-splash__ray {
          position: absolute;
          top: 0;
          left: 0;
          width: 38%;
          height: 100%;
          background: linear-gradient(
            105deg,
            transparent 30%,
            rgba(147, 197, 253, 0.055) 50%,
            transparent 70%
          );
          animation: ef-ray-sweep 2.8s cubic-bezier(0.4, 0, 0.6, 1) 0.3s both;
          pointer-events: none;
        }

        /* ── Top rule line ─────────────────────────────────────── */

        .ef-splash__rule {
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          height: 2px;
          background: linear-gradient(
            90deg,
            transparent,
            rgba(147, 197, 253, 0.55) 40%,
            rgba(96, 165, 250, 0.80) 50%,
            rgba(147, 197, 253, 0.55) 60%,
            transparent
          );
        }

        /* ── Content card ──────────────────────────────────────── */

        .ef-splash__body {
          position: relative;
          z-index: 2;
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 20px;
          width: min(92vw, 440px);
        }

        /* ── Logo ──────────────────────────────────────────────── */

        .ef-splash__logo-wrap {
          position: relative;
          width: clamp(96px, 28vw, 140px);
          aspect-ratio: 1;
        }

        .ef-splash__logo-glow {
          position: absolute;
          inset: -18px;
          border-radius: 0;
          background: radial-gradient(circle, rgba(96, 165, 250, 0.30), transparent 62%);
          filter: blur(14px);
        }

        .ef-splash__logo {
          position: relative;
          display: block;
          width: 100%;
          height: 100%;
          object-fit: contain;
          border-radius: 0;         /* No border radius — square logo */
          filter: drop-shadow(0 4px 18px rgba(96, 165, 250, 0.35));
          animation: ef-logo-reveal 600ms cubic-bezier(0.22, 1, 0.36, 1) both;
        }

        /* ── Text block ────────────────────────────────────────── */

        .ef-splash__texts {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 6px;
          animation: ef-text-rise 600ms cubic-bezier(0.22, 1, 0.36, 1) 180ms both;
        }

        .ef-splash__eyebrow {
          margin: 0;
          font-size: clamp(9px, 2.2vw, 11px);
          font-weight: 700;
          letter-spacing: 0.22em;
          text-transform: uppercase;
          color: rgba(147, 197, 253, 0.72);
        }

        .ef-splash__title {
          margin: 0;
          font-size: clamp(34px, 9.5vw, 60px);
          font-weight: 900;
          letter-spacing: -0.045em;
          line-height: 1;
          text-align: center;
          color: #e2eaf8;
          text-shadow:
            0 1px 0 rgba(0, 0, 0, 0.30),
            0 0 28px rgba(147, 197, 253, 0.22);
        }

        .ef-splash__title em {
          font-style: normal;
          color: #93c5fd;
        }

        .ef-splash__divider {
          width: min(48px, 12vw);
          height: 2px;
          border-radius: 999px;
          background: linear-gradient(90deg, #3b82f6, #93c5fd);
          transform-origin: center;
          animation: ef-divider-expand 500ms cubic-bezier(0.22, 1, 0.36, 1) 360ms both;
        }

        .ef-splash__subtitle {
          margin: 0;
          font-size: clamp(9px, 2vw, 11px);
          font-weight: 600;
          letter-spacing: 0.16em;
          text-align: center;
          text-transform: uppercase;
          color: rgba(186, 213, 250, 0.58);
        }

        /* ── Progress bar ──────────────────────────────────────── */

        .ef-splash__progress {
          width: min(160px, 40vw);
          height: 2px;
          overflow: hidden;
          border-radius: 999px;
          background: rgba(255, 255, 255, 0.10);
          margin-top: 8px;
          animation: ef-text-rise 600ms cubic-bezier(0.22, 1, 0.36, 1) 420ms both;
        }

        .ef-splash__progress-fill {
          width: 100%;
          height: 100%;
          transform-origin: left center;
          border-radius: inherit;
          background: linear-gradient(90deg, #3b82f6 0%, #93c5fd 100%);
          animation: ef-progress-fill ${HOLD_MS - 200}ms cubic-bezier(0.25, 0.1, 0.25, 1) 200ms both;
        }

        /* ── Responsive ────────────────────────────────────────── */

        @media (max-width: 380px) {
          .ef-splash__logo-wrap { width: clamp(80px, 26vw, 110px); }
          .ef-splash__body { gap: 16px; }
        }

        /* ── Reduced motion ────────────────────────────────────── */

        @media (prefers-reduced-motion: reduce) {
          .ef-splash *,
          .ef-splash *::before,
          .ef-splash *::after {
            animation-duration: 1ms !important;
            animation-iteration-count: 1 !important;
          }
        }
      `}</style>

      <div
        className={["ef-splash", phase === "exit" ? "ef-splash--exit" : ""].join(" ")}
        role="status"
        aria-label="Loading e-Formatter"
      >
        {/* Ambient orb */}
        <div className="ef-splash__orb" aria-hidden="true" />

        {/* Light ray */}
        <div className="ef-splash__ray" aria-hidden="true" />

        {/* Top accent rule */}
        <div className="ef-splash__rule" aria-hidden="true" />

        {/* Content */}
        <div className="ef-splash__body" aria-hidden="true">
          {/* Logo */}
          <div className="ef-splash__logo-wrap">
            <div className="ef-splash__logo-glow" />
            <img
              className="ef-splash__logo"
              src="/images/icon-512.png"
              alt=""
              width="140"
              height="140"
              draggable={false}
            />
          </div>

          {/* Text */}
          <div className="ef-splash__texts">
            <p className="ef-splash__eyebrow">Digital Manuscript Processing</p>
            <h1 className="ef-splash__title">
              <em>e-</em>Formatter
            </h1>
            <div className="ef-splash__divider" />
            <p className="ef-splash__subtitle">City College of Calamba · OVPREPQA</p>
          </div>

          {/* Progress */}
          <div className="ef-splash__progress" aria-hidden="true">
            <div className="ef-splash__progress-fill" />
          </div>
        </div>
      </div>
    </>
  );
}
