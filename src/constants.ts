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
  actionLabel?: string;
  onAction?: () => void;
}

export type CitationStyle = "ieee" | "apa";
export type FormattingStandard = "ccc" | "conference";
export type ConferenceFormat = "acm" | "pubform";

export type ElementStyle = {
  fontFamily: string;
  fontSize: number; // in points
  lineSpacing: number; // multiplier (e.g. 1.0, 1.5, 2.0)
  alignment: "left" | "center" | "right" | "both";
  indentation: number; // in inches (e.g. 0, 0.5)
  bold?: boolean;
  italic?: boolean;
  textTransform?: "none" | "uppercase";
};

export type TableOptions = {
  fontFamily: string;
  fontSize: number;
  lineSpacing: number;
  alignment: "left" | "center" | "right" | "both";
  bold?: boolean;
  italic?: boolean;
};

export type FigureOptions = {
  borderWeight: number; // in points (e.g. 0.5, 1.0, 1.5)
  spacing: number; // line spacing multiplier
};

export type FormattingConfig = {
  body: ElementStyle;
  references: ElementStyle;
  headings: ElementStyle;
  titles: ElementStyle;
  chapterNumber: ElementStyle;
  tableCaption: ElementStyle;
  figureCaption: ElementStyle;
  tableContinuation: ElementStyle;
  appendixLetter: ElementStyle;
  appendixContinuation: ElementStyle;
  table: TableOptions;
  figure: FigureOptions;
  legends: ElementStyle;
};

export const DEFAULT_CONFIG_IEEE: FormattingConfig = {
  body: {
    fontFamily: "Garamond",
    fontSize: 12,
    lineSpacing: 2.0,
    alignment: "both",
    indentation: 0.5,
  },
  references: {
    fontFamily: "Garamond",
    fontSize: 11,
    lineSpacing: 1.0,
    alignment: "both",
    indentation: 0,
  },
  headings: {
    fontFamily: "Garamond",
    fontSize: 13,
    lineSpacing: 2.0,
    alignment: "left",
    indentation: 0,
    bold: true,
  },
  titles: {
    fontFamily: "Garamond",
    fontSize: 14,
    lineSpacing: 3.0,
    alignment: "center",
    indentation: 0,
    bold: true,
    textTransform: "uppercase",
  },
  chapterNumber: {
    fontFamily: "Garamond",
    fontSize: 14,
    lineSpacing: 2.0,
    alignment: "center",
    indentation: 0,
    bold: true,
  },
  tableCaption: {
    fontFamily: "Garamond",
    fontSize: 13,
    lineSpacing: 1.0,
    alignment: "left",
    indentation: 0,
  },
  figureCaption: {
    fontFamily: "Garamond",
    fontSize: 12,
    lineSpacing: 2.0,
    alignment: "center",
    indentation: 0,
  },
  tableContinuation: {
    fontFamily: "Garamond",
    fontSize: 12,
    lineSpacing: 1.0,
    alignment: "left",
    indentation: 0,
    italic: true,
  },
  appendixContinuation: {
    fontFamily: "Garamond",
    fontSize: 13,
    lineSpacing: 3.0,
    alignment: "left",
    indentation: 0,
    italic: true,
  },
  table: {
    fontFamily: "Arial",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "center",
  },
  figure: {
    borderWeight: 2.25,
    spacing: 1.0,
  },
  appendixLetter: {
    fontFamily: "Garamond",
    fontSize: 14,
    lineSpacing: 2.0,
    alignment: "center",
    indentation: 0,
    bold: true,
  },
  legends: {
    fontFamily: "Garamond",
    fontSize: 11,
    lineSpacing: 1.0,
    alignment: "left",
    indentation: 0,
  },
};

export const DEFAULT_CONFIG_APA: FormattingConfig = {
  ...DEFAULT_CONFIG_IEEE, // Currently sharing defaults as per user's identical lists
};

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

export const FORMATTING_STANDARDS: {
  value: FormattingStandard;
  label: string;
  sub: string;
  icon: string;
}[] = [
    {
      value: "ccc",
      label: "CCC",
      sub: "CCC formatting guidelines",
      icon: "fa-building-columns",
    },
    {
      value: "conference",
      label: "Conference",
      sub: "Conference-specific formatting rules",
      icon: "fa-users-rectangle",
    },
  ];

export const CONFERENCE_FORMATS: {
  value: ConferenceFormat;
  label: string;
  sub: string;
  icon: string;
  href: string;
  download: string;
}[] = [
    {
      value: "acm",
      label: "ACM Conference",
      sub: "ACM conference formatting file",
      icon: "fa-file-lines",
      href: "/conference/acm_formatting_guidelines.docx",
      download: "acm_conference_format.docx",
    },
    {
      value: "pubform",
      label: "Publication Form",
      sub: "Publication formatting file",
      icon: "fa-file-signature",
      href: "/conference/publication_formatting_guidelines.docx",
      download: "publication_form_format.docx",
    },
  ];

export type ConferenceTextAlignment = "left" | "center" | "right" | "both";

export interface ConferenceTextStyle {
  fontFamily: string;
  fontSize: number;
  lineSpacing: number;
  alignment: ConferenceTextAlignment;
  bold?: boolean;
  italic?: boolean;
  uppercase?: boolean;
  titleCase?: boolean;
}

export interface PublicationBodyStyle extends ConferenceTextStyle {
  spacingAfterPt: number;
  firstLineIndentCm: number;
}

export interface PublicationReferencesStyle extends ConferenceTextStyle {
  hangingIndentCm: number;
  ieeeStyle: boolean;
}

export interface PublicationFormattingConfig {
  heading1: ConferenceTextStyle;
  heading2: ConferenceTextStyle;
  body: PublicationBodyStyle;
  references: PublicationReferencesStyle;
}

export interface AcmFormattingConfig {
  title: ConferenceTextStyle;
  subtitle: ConferenceTextStyle;
  author: ConferenceTextStyle;
  heading: ConferenceTextStyle;
  body: ConferenceTextStyle;
  references: ConferenceTextStyle;
}

export interface ConferenceFormattingConfig {
  pubform: PublicationFormattingConfig;
  acm: AcmFormattingConfig;
}

export const DEFAULT_PUBLICATION_FORMATTING_CONFIG: PublicationFormattingConfig = {
  heading1: {
    fontFamily: "Times New Roman",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
    uppercase: true,
  },
  heading2: {
    fontFamily: "Times New Roman",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "left",
    bold: false,
    italic: true,
    titleCase: true,
  },
  body: {
    fontFamily: "Times New Roman",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingAfterPt: 6,
    firstLineIndentCm: 0.51,
  },
  references: {
    fontFamily: "Times New Roman",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    hangingIndentCm: 0.62,
    ieeeStyle: true,
  },
};

export const DEFAULT_ACM_FORMATTING_CONFIG: AcmFormattingConfig = {
  title: {
    fontFamily: "Times New Roman",
    fontSize: 18,
    lineSpacing: 1.15,
    alignment: "center",
    bold: true,
    italic: false,
  },
  subtitle: {
    fontFamily: "Times New Roman",
    fontSize: 11,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
  },
  author: {
    fontFamily: "Times New Roman",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "center",
    bold: true,
    italic: false,
  },
  heading: {
    fontFamily: "Times New Roman",
    fontSize: 12,
    lineSpacing: 1.05,
    alignment: "left",
    bold: true,
    italic: false,
  },
  body: {
    fontFamily: "Times New Roman",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "left",
    bold: false,
    italic: false,
  },
  references: {
    fontFamily: "Times New Roman",
    fontSize: 10,
    lineSpacing: 1.0,
    alignment: "left",
    bold: false,
    italic: false,
  },
};

export const DEFAULT_CONFERENCE_FORMATTING_CONFIG: ConferenceFormattingConfig = {
  pubform: DEFAULT_PUBLICATION_FORMATTING_CONFIG,
  acm: DEFAULT_ACM_FORMATTING_CONFIG,
};
