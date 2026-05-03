import { useEffect, useState } from "react";

const HOLD_MS = 1850;
const FADE_MS = 520;

interface Props {
  onFinish: () => void;
}

export default function SplashScreen({ onFinish }: Props) {
  const [phase, setPhase] = useState<"enter" | "visible" | "exit">("enter");

  useEffect(() => {
    const frame = requestAnimationFrame(() => setPhase("visible"));
    const exitTimer = window.setTimeout(() => setPhase("exit"), HOLD_MS);
    const finishTimer = window.setTimeout(onFinish, HOLD_MS + FADE_MS);

    return () => {
      cancelAnimationFrame(frame);
      window.clearTimeout(exitTimer);
      window.clearTimeout(finishTimer);
    };
  }, [onFinish]);

  return (
    <>
      <style>{`
        @keyframes ef-splash-card-in {
          from { opacity: 0; transform: translate3d(0, 22px, 0) scale(0.94); }
          to { opacity: 1; transform: translate3d(0, 0, 0) scale(1); }
        }

        @keyframes ef-splash-icon-float {
          0%, 100% { transform: translateY(0); }
          50% { transform: translateY(-7px); }
        }

        @keyframes ef-splash-border-glow {
          from {
            box-shadow:
              0 0 0 1px rgba(234, 213, 141, 0.48),
              0 0 26px rgba(119, 230, 232, 0.24),
              inset 0 0 28px rgba(119, 230, 232, 0.07);
          }
          to {
            box-shadow:
              0 0 0 1px rgba(234, 213, 141, 0.82),
              0 0 46px rgba(119, 230, 232, 0.38),
              inset 0 0 36px rgba(119, 230, 232, 0.12);
          }
        }

        @keyframes ef-splash-line-scan {
          from { transform: translateX(-40%); opacity: 0; }
          18% { opacity: 0.65; }
          to { transform: translateX(120%); opacity: 0; }
        }

        @keyframes ef-splash-progress {
          from { transform: scaleX(0); }
          to { transform: scaleX(1); }
        }

        @keyframes ef-splash-hex-drift {
          0%, 100% { transform: translate3d(0, 0, 0); opacity: 0.58; }
          50% { transform: translate3d(14px, -10px, 0); opacity: 0.84; }
        }

        .ef-splash {
          position: fixed;
          inset: 0;
          z-index: 99999;
          display: grid;
          place-items: center;
          overflow: hidden;
          min-height: 100svh;
          padding: max(24px, env(safe-area-inset-top)) max(18px, env(safe-area-inset-right)) max(24px, env(safe-area-inset-bottom)) max(18px, env(safe-area-inset-left));
          color: #f7f3df;
          font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
          background:
            radial-gradient(circle at 50% 28%, rgba(112, 230, 232, 0.27), transparent 34%),
            radial-gradient(circle at 12% 18%, rgba(234, 213, 141, 0.10), transparent 30%),
            linear-gradient(145deg, #08333d 0%, #062b35 46%, #020b13 100%);
          opacity: 0;
          user-select: none;
          -webkit-user-select: none;
          transition: opacity ${FADE_MS}ms cubic-bezier(0.4, 0, 0.2, 1);
        }

        .ef-splash--visible { opacity: 1; }
        .ef-splash--exit { opacity: 0; pointer-events: none; }

        .ef-splash__frame {
          position: absolute;
          inset: clamp(14px, 2.7vw, 36px);
          border-radius: clamp(22px, 4vw, 34px);
          border: 1px solid rgba(234, 213, 141, 0.58);
          pointer-events: none;
          animation: ef-splash-border-glow 1.8s ease-in-out infinite alternate;
        }

        .ef-splash__frame::before,
        .ef-splash__frame::after {
          content: "";
          position: absolute;
          left: 16px;
          right: 16px;
          height: 1px;
          background: linear-gradient(90deg, transparent, rgba(119, 230, 232, 0.33), transparent);
        }

        .ef-splash__frame::before { top: 30%; }
        .ef-splash__frame::after { bottom: 30%; }

        .ef-splash__glow {
          position: absolute;
          width: min(76vw, 760px);
          aspect-ratio: 1;
          border-radius: 999px;
          background: radial-gradient(circle, rgba(119, 230, 232, 0.26), transparent 64%);
          filter: blur(30px);
        }

        .ef-splash__hex {
          position: absolute;
          inset: 0;
          opacity: 0.58;
          background-image:
            linear-gradient(30deg, rgba(119, 230, 232, 0.08) 12%, transparent 12.5%, transparent 87%, rgba(119, 230, 232, 0.08) 87.5%, rgba(119, 230, 232, 0.08)),
            linear-gradient(150deg, rgba(119, 230, 232, 0.08) 12%, transparent 12.5%, transparent 87%, rgba(119, 230, 232, 0.08) 87.5%, rgba(119, 230, 232, 0.08)),
            linear-gradient(30deg, rgba(119, 230, 232, 0.08) 12%, transparent 12.5%, transparent 87%, rgba(119, 230, 232, 0.08) 87.5%, rgba(119, 230, 232, 0.08)),
            linear-gradient(150deg, rgba(119, 230, 232, 0.08) 12%, transparent 12.5%, transparent 87%, rgba(119, 230, 232, 0.08) 87.5%, rgba(119, 230, 232, 0.08));
          background-position: 0 0, 0 0, 48px 84px, 48px 84px;
          background-size: 96px 168px;
          mask-image: radial-gradient(circle at center, black 0%, black 44%, transparent 78%);
          -webkit-mask-image: radial-gradient(circle at center, black 0%, black 44%, transparent 78%);
          animation: ef-splash-hex-drift 9s ease-in-out infinite;
        }

        .ef-splash__scan {
          position: absolute;
          top: 22%;
          left: 0;
          width: 56%;
          height: 1px;
          background: linear-gradient(90deg, transparent, rgba(234, 213, 141, 0.70), rgba(119, 230, 232, 0.75), transparent);
          animation: ef-splash-line-scan 2.4s ease-in-out infinite;
        }

        .ef-splash__card {
          position: relative;
          z-index: 2;
          width: min(92vw, 520px);
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 18px;
          padding: clamp(26px, 6vw, 42px) clamp(20px, 5vw, 38px);
          border-radius: clamp(28px, 6vw, 42px);
          border: 1px solid rgba(234, 213, 141, 0.26);
          background:
            linear-gradient(180deg, rgba(255,255,255,0.08), rgba(255,255,255,0.02)),
            rgba(2, 11, 19, 0.36);
          backdrop-filter: blur(22px);
          -webkit-backdrop-filter: blur(22px);
          box-shadow: 0 28px 84px rgba(0,0,0,0.42), inset 0 1px 0 rgba(255,255,255,0.10);
          animation: ef-splash-card-in 680ms cubic-bezier(0.22, 1, 0.36, 1) both;
        }

        .ef-splash__icon-shell {
          position: relative;
          width: clamp(118px, 32vw, 166px);
          aspect-ratio: 1;
          display: grid;
          place-items: center;
          border-radius: 30%;
          background: rgba(6, 43, 53, 0.42);
        }

        .ef-splash__icon-shell::before {
          content: "";
          position: absolute;
          inset: -12px;
          border-radius: inherit;
          background: radial-gradient(circle, rgba(119, 230, 232, 0.38), transparent 64%);
          filter: blur(12px);
        }

        .ef-splash__icon {
          position: relative;
          width: 100%;
          height: 100%;
          border-radius: 30%;
          object-fit: cover;
          filter: drop-shadow(0 0 12px rgba(119, 230, 232, 0.40));
          animation: ef-splash-icon-float 2.4s ease-in-out infinite;
        }

        .ef-splash__kicker {
          margin: 0;
          font-size: clamp(10px, 2.4vw, 13px);
          font-weight: 800;
          letter-spacing: clamp(0.14em, 1.1vw, 0.24em);
          text-align: center;
          text-transform: uppercase;
          color: rgba(245, 249, 247, 0.90);
          text-shadow: 0 2px 10px rgba(0,0,0,0.45);
        }

        .ef-splash__title {
          margin: 0;
          display: flex;
          flex-wrap: wrap;
          justify-content: center;
          align-items: baseline;
          gap: 0.18em;
          font-size: clamp(36px, 10vw, 66px);
          line-height: 0.92;
          font-weight: 950;
          letter-spacing: -0.055em;
          text-align: center;
          color: #ead58f;
          text-transform: uppercase;
          text-shadow:
            0 2px 0 rgba(0,0,0,0.22),
            0 0 22px rgba(234, 213, 141, 0.28);
        }

        .ef-splash__title-prefix {
          color: #dffcf5;
          text-transform: none;
          letter-spacing: -0.04em;
        }

        .ef-splash__bracket {
          font-weight: 800;
          color: rgba(234, 213, 141, 0.74);
          letter-spacing: -0.10em;
        }

        .ef-splash__subline {
          margin: 0;
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.18em;
          text-align: center;
          text-transform: uppercase;
          color: rgba(164, 230, 228, 0.78);
        }

        .ef-splash__progress {
          width: min(190px, 48vw);
          height: 3px;
          margin-top: 4px;
          overflow: hidden;
          border-radius: 999px;
          background: rgba(255,255,255,0.12);
        }

        .ef-splash__progress-fill {
          width: 100%;
          height: 100%;
          transform-origin: left center;
          border-radius: inherit;
          background: linear-gradient(90deg, #77e6e8 0%, #ead58f 100%);
          animation: ef-splash-progress ${HOLD_MS - 180}ms cubic-bezier(0.25, 0, 0.35, 1) 160ms both;
        }

        .ef-splash__credit {
          position: absolute;
          bottom: max(26px, calc(env(safe-area-inset-bottom) + 16px));
          left: 50%;
          width: min(90vw, 520px);
          transform: translateX(-50%);
          margin: 0;
          font-size: 10px;
          font-weight: 600;
          letter-spacing: 0.04em;
          text-align: center;
          color: rgba(216, 255, 244, 0.42);
        }

        @media (max-width: 420px) {
          .ef-splash__card { width: min(92vw, 360px); }
          .ef-splash__title { gap: 0.10em; }
          .ef-splash__bracket { display: none; }
        }

        @media (prefers-reduced-motion: reduce) {
          .ef-splash,
          .ef-splash *,
          .ef-splash *::before,
          .ef-splash *::after {
            animation-duration: 1ms !important;
            animation-iteration-count: 1 !important;
            scroll-behavior: auto !important;
          }
        }
      `}</style>

      <div
        className={[
          "ef-splash",
          phase === "visible" ? "ef-splash--visible" : "",
          phase === "exit" ? "ef-splash--exit" : "",
        ].join(" ")}
        role="status"
        aria-label="Loading e-Formatter"
      >
        <div className="ef-splash__glow" aria-hidden="true" />
        <div className="ef-splash__hex" aria-hidden="true" />
        <div className="ef-splash__scan" aria-hidden="true" />
        <div className="ef-splash__frame" aria-hidden="true" />

        <section className="ef-splash__card" aria-hidden="true">
          <div className="ef-splash__icon-shell">
            <img
              className="ef-splash__icon"
              src="/images/icon-512.png"
              alt=""
              width="166"
              height="166"
              draggable={false}
            />
          </div>

          <div>
            <p className="ef-splash__kicker">Digital Manuscript Processing</p>
            <h1 className="ef-splash__title">
              <span className="ef-splash__bracket">&lt;</span>
              <span className="ef-splash__title-prefix">e-</span>
              <span>Formatter</span>
              <span className="ef-splash__bracket">&gt;</span>
            </h1>
          </div>

          <p className="ef-splash__subline">City College of Calamba · OVPREPQA</p>
          <div className="ef-splash__progress" aria-hidden="true">
            <div className="ef-splash__progress-fill" />
          </div>
        </section>

        <p className="ef-splash__credit">Developed by Railey Dela Peña</p>
      </div>
    </>
  );
}
