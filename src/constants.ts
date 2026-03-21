export const RULES_DEF: [string, string, string][] = [
  ["spacing", "fa-arrows-up-down", "Spacing"],
  ["indentation", "fa-indent", "Indentation"],
  ["alignment", "fa-align-left", "Alignment"],
  ["captions", "fa-image", "Figure / Table Captions"],
  ["continuation", "fa-rotate-right", "Continuation Labels"],
  ["borders", "fa-border-style", "Figure Borders"],
  ["pagination", "fa-file-lines", "Margins / Pagination"],
];

export interface ToastMsg {
  id: number;
  msg: string;
  type: "success" | "error";
}
