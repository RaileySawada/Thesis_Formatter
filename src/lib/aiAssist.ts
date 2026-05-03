const TRUE_VALUES = new Set(["true", "1", "yes", "on"]);
const FALSE_VALUES = new Set(["false", "0", "no", "off"]);

export function isAiAssistEnabled(rawValue: unknown): boolean {
  const normalized = String(rawValue ?? "")
    .trim()
    .toLowerCase();

  if (FALSE_VALUES.has(normalized)) return false;
  if (TRUE_VALUES.has(normalized)) return true;

  // Default to enabled when no explicit flag is provided.
  return true;
}

