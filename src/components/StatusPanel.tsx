interface Props {
  selectedSections: string[];
  sectionLabels: Record<string, string>;
  file: File | null;
  enabledRules: string[];
  totalRules: number;
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

export default function StatusPanel({
  selectedSections,
  sectionLabels,
  file,
  enabledRules,
  totalRules,
}: Props) {
  const sectionDone = selectedSections.length > 0;
  const fileDone = !!file;
  const activeRules = enabledRules.length;
  const rulesDone = activeRules > 0;
  const allDone = sectionDone && fileDone && rulesDone;

  const headline = allDone
    ? "Ready to format!"
    : fileDone
      ? "Almost there…"
      : "Ready for upload";

  const subtext = allDone
    ? "All set — click Apply Formatting below to process your manuscript."
    : fileDone
      ? "Check that all steps above are complete."
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
              {sectionDone
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
          <div className="flex-1 min-w-0">
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
              {rulesDone
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
