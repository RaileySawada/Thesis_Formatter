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

export type CitationStyle = "ieee" | "apa";

export const CITATION_STYLES: {
  value: CitationStyle;
  label: string;
  sub: string;
  icon: string;
}[] = [
  {
    value: "ieee",
    label: "IEEE",
    sub: "Numbered references [1]",
    icon: "fa-hashtag",
  },
  {
    value: "apa",
    label: "APA 7th",
    sub: "Author-date (Smith, 2024)",
    icon: "fa-user-pen",
  },
];
