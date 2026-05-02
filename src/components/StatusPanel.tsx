import type { FormattingStandard } from "../constants";

type AiRunStatus = "not-used" | "running" | "success" | "failed";

interface Props {
  formattingStandard: FormattingStandard;
  selectedSections: string[];
  sectionLabels: Record<string, string>;
  file: File | null;
  enabledRules: string[];
  totalRules: number;
  processing: boolean;
  activeElapsedMs: number;
  lastRunMs: number | null;
  aiAssistEnabled: boolean;
  aiStatus: AiRunStatus;
  lastAiMs: number | null;
  aiError: string | null;
}

function StepIcon({
  done,
  active,
  num,
}: {
  done: boolean;
  active: boolean;
  num: number;
}) {
  const bg = done
    ? "#3b82f6"
    : active
      ? "rgba(255,255,255,0.25)"
      : "rgba(255,255,255,0.1)";

  return (
    <span
      className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-xs font-bold"
      style={{ background: bg, color: "#fff" }}
    >
      {done ? <i className="fa-solid fa-check text-[10px]" /> : num}
    </span>
  );
}

function stepClass(done: boolean, active: boolean, inactive: boolean) {
  let cls =
    "status-step flex items-center gap-3 rounded-2xl px-4 py-3 transition-all duration-300";
  if (done) cls += " status-step--done";
  else if (active) cls += " status-step--active";
  else if (inactive) cls += " status-step--inactive";
  return cls;
}

function formatDuration(ms: number | null): string {
  if (ms === null) return "--";
  if (ms >= 1000) return `${(ms / 1000).toFixed(2)}s`;
  return `${Math.max(0, Math.round(ms))}ms`;
}

export default function StatusPanel({
  formattingStandard,
  selectedSections,
  sectionLabels,
  file,
  enabledRules,
  totalRules,
  processing,
  activeElapsedMs,
  lastRunMs,
  aiAssistEnabled,
  aiStatus,
  lastAiMs,
  aiError,
}: Props) {
  const sectionDone =
    formattingStandard === "conference" ? true : selectedSections.length > 0;
  const fileDone = !!file;
  const activeRules = enabledRules.length;
  const rulesDone = formattingStandard === "conference" ? true : activeRules > 0;
  const allDone = sectionDone && fileDone && rulesDone;

  const headline = allDone
    ? "Ready to format!"
    : fileDone
      ? "Almost there…"
      : "Ready for upload";

  const subtext = allDone
    ? "All set — click Apply Formatting below to process your manuscript."
    : fileDone
      ? formattingStandard === "conference"
        ? "Conference rule is active. Upload and apply formatting."
        : "Check that all steps above are complete."
      : formattingStandard === "conference"
        ? "Choose a conference rule, then upload your manuscript."
        : "Select sections on the left, then upload your manuscript.";

  return (
    <div
      className="rounded-3xl p-6 text-white shadow-xl flex flex-col"
      style={{ background: "var(--status-bg)" }}
    >
      <p
        className="text-xs font-bold uppercase tracking-[0.2em]"
        style={{ color: "var(--accent-light)" }}
      >
        <i className="fa-solid fa-circle-info mr-1" /> Current Status
      </p>

      <h2 className="mt-2 text-lg font-bold sm:text-xl">{headline}</h2>
      <p
        className="mt-1.5 text-sm leading-6"
        style={{ color: "rgba(255,255,255,0.65)" }}
      >
        {subtext}
      </p>

      <div
        className="mt-4 rounded-2xl border px-3 py-2"
        style={{
          borderColor: "rgba(255,255,255,0.18)",
          background: "rgba(255,255,255,0.06)",
        }}
      >
        <p className="text-[11px] font-semibold">
          {processing
            ? `Run Time: ${formatDuration(activeElapsedMs)}`
            : `Last Run: ${formatDuration(lastRunMs)}`}
        </p>
        <p
          className="mt-1 text-[10px]"
          style={{ color: "rgba(255,255,255,0.68)" }}
        >
          {!aiAssistEnabled
            ? "AI Check: Off (local rule-based formatter)"
            : aiStatus === "running"
              ? "AI Check: Running..."
              : aiStatus === "success"
                ? `AI Check: Success in ${formatDuration(lastAiMs)}`
                : `AI Check: Failed${lastAiMs !== null ? ` after ${formatDuration(lastAiMs)}` : ""}. Local fallback used.`}
        </p>
        {aiError && aiStatus === "failed" && (
          <p className="mt-1 text-[10px]" style={{ color: "#fecaca" }}>
            {aiError}
          </p>
        )}
      </div>

      <div className="mt-5 space-y-2.5 flex-1">
        {/* Step 1: Section */}
        <div
          className={stepClass(sectionDone, !sectionDone, false)}
          style={{
            background: sectionDone ? undefined : "rgba(255,255,255,0.07)",
          }}
        >
          <StepIcon done={sectionDone} active={!sectionDone} num={1} />
          <div className="flex-1 min-w-0">
            <p className="text-xs font-semibold">Section selected</p>
            <p
              className="text-[11px] mt-0.5"
              style={{
                color: sectionDone
                  ? "rgba(255,255,255,0.6)"
                  : "rgba(255,255,255,0.5)",
              }}
            >
              {formattingStandard === "conference"
                ? "Conference mode uses full-document formatting."
                : sectionDone
                  ? selectedSections.map((v) => sectionLabels[v] ?? v).join(", ")
                  : "Choose a scope in the sidebar"}
            </p>
          </div>
        </div>

        {/* Step 2: File */}
        <div
          className={stepClass(
            fileDone,
            !fileDone && sectionDone,
            !sectionDone,
          )}
          style={{
            background: fileDone ? undefined : "rgba(255,255,255,0.07)",
          }}
        >
          <StepIcon done={fileDone} active={!fileDone && sectionDone} num={2} />
          <div className="flex-1 min-w-0 overflow-hidden">
            <p className="text-xs font-semibold">Manuscript uploaded</p>
            <p
              className="text-[11px] mt-0.5 truncate"
              style={{ color: "rgba(255,255,255,0.5)" }}
            >
              {fileDone ? file!.name : "No file chosen yet"}
            </p>
          </div>
        </div>

        {/* Step 3: Rules */}
        <div
          className={stepClass(rulesDone, !rulesDone, false)}
          style={{
            background: rulesDone ? undefined : "rgba(255,255,255,0.07)",
          }}
        >
          <StepIcon done={rulesDone} active={!rulesDone} num={3} />
          <div className="flex-1 min-w-0">
            <p className="text-xs font-semibold">Formatting rules</p>
            <p
              className="text-[11px] mt-0.5"
              style={{
                color: rulesDone
                  ? "rgba(255,255,255,0.6)"
                  : "rgba(255,255,255,0.5)",
              }}
            >
              {formattingStandard === "conference"
                ? "Conference rule set selected"
                : rulesDone
                  ? `${activeRules} of ${totalRules} rules active`
                  : "No rules enabled"}
            </p>
          </div>
        </div>

        {/* Step 4: Apply */}
        <div
          className={stepClass(allDone, false, !allDone)}
          style={{
            background: allDone ? undefined : "rgba(255,255,255,0.05)",
            opacity: allDone ? 1 : 0.45,
          }}
        >
          <StepIcon done={allDone} active={false} num={4} />
          <div className="flex-1 min-w-0">
            <p className="text-xs font-semibold">Apply formatting</p>
            <p
              className="text-[11px] mt-0.5"
              style={{
                color: allDone
                  ? "rgba(255,255,255,0.6)"
                  : "rgba(255,255,255,0.4)",
              }}
            >
              {allDone
                ? "Hit 'Apply Formatting' to run"
                : "Complete steps above first"}
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
