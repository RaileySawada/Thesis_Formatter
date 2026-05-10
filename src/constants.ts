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
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  firstLineIndentCm?: number;
  hangingIndentCm?: number;
  lineSpacingPt?: number;
  lineSpacingRule?: "auto" | "exact" | "atLeast";
  color?: string;
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
  authorAffiliation: ConferenceTextStyle;
  abstract: ConferenceTextStyle;
  concepts: ConferenceTextStyle;
  keywords: ConferenceTextStyle;
  referenceFormatLabel: ConferenceTextStyle;
  referenceFormatContent: ConferenceTextStyle;
  preliminaryFootnote: ConferenceTextStyle;
  heading: ConferenceTextStyle;
  heading1: ConferenceTextStyle;
  heading2: ConferenceTextStyle;
  heading3: ConferenceTextStyle;
  body: ConferenceTextStyle;
  tableCaption: ConferenceTextStyle;
  table: ConferenceTextStyle;
  footnote: ConferenceTextStyle;
  figure: ConferenceTextStyle;
  figureCaption: ConferenceTextStyle;
  equation: ConferenceTextStyle;
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
    fontFamily: "Linux Biolinum O",
    fontSize: 12,
    lineSpacing: 1.0,
    alignment: "left",
    bold: true,
    italic: false,
    lineSpacingPt: 18,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  subtitle: {
    fontFamily: "Linux Biolinum O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "left",
    bold: false,
    italic: false,
    spacingAfterPt: 18,
    lineSpacingPt: 16.65,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  author: {
    fontFamily: "Linux Biolinum O",
    fontSize: 11,
    lineSpacing: 1.0,
    alignment: "left",
    bold: false,
    italic: false,
    spacingBeforePt: 3,
    lineSpacingPt: 16,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  authorAffiliation: {
    fontFamily: "Linux Libertine O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "left",
    bold: false,
    italic: false,
    spacingBeforePt: 3,
    lineSpacingPt: 14.85,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  abstract: {
    fontFamily: "Linux Libertine O",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 10,
    lineSpacingPt: 12,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  concepts: {
    fontFamily: "Linux Libertine O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 7,
    lineSpacingPt: 13.5,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  keywords: {
    fontFamily: "Linux Libertine O",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 7,
    lineSpacingPt: 13.5,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  referenceFormatLabel: {
    fontFamily: "Linux Libertine O",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "both",
    bold: true,
    italic: false,
    spacingBeforePt: 8,
    lineSpacingPt: 9.6,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  referenceFormatContent: {
    fontFamily: "Linux Libertine O",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 1,
    lineSpacingPt: 12,
    lineSpacingRule: "exact",
    color: "000000",
  },
  preliminaryFootnote: {
    fontFamily: "Linux Libertine O",
    fontSize: 7,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 0,
    spacingAfterPt: 0,
    color: "000000",
  },
  heading: {
    fontFamily: "Linux Biolinum O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "left",
    bold: true,
    italic: false,
    uppercase: true,
    spacingBeforePt: 12,
    spacingAfterPt: 3,
    color: "000000",
  },
  heading1: {
    fontFamily: "Linux Biolinum O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "left",
    bold: true,
    italic: false,
    uppercase: true,
    spacingBeforePt: 12,
    spacingAfterPt: 3,
    color: "000000",
  },
  heading2: {
    fontFamily: "Linux Biolinum O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "left",
    bold: true,
    italic: false,
    titleCase: true,
    spacingBeforePt: 12,
    spacingAfterPt: 3,
    color: "000000",
  },
  heading3: {
    fontFamily: "Linux Biolinum O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: true,
    spacingBeforePt: 12,
    spacingAfterPt: 3,
    color: "000000",
  },
  body: {
    fontFamily: "Linux Libertine O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 0,
    spacingAfterPt: 0,
    firstLineIndentCm: 0.42,
    lineSpacingPt: 13.5,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  tableCaption: {
    fontFamily: "Linux Biolinum O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
    spacingBeforePt: 9,
    spacingAfterPt: 6,
    color: "000000",
  },
  table: {
    fontFamily: "Linux Libertine O",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    lineSpacingPt: 11,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  footnote: {
    fontFamily: "Times New Roman",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
    spacingBeforePt: 3,
    spacingAfterPt: 10,
    color: "000000",
  },
  figure: {
    fontFamily: "Linux Libertine O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
    spacingBeforePt: 6,
    spacingAfterPt: 10,
    lineSpacingPt: 11.25,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  figureCaption: {
    fontFamily: "Linux Biolinum O",
    fontSize: 8,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
    spacingBeforePt: 3,
    spacingAfterPt: 9,
    lineSpacingPt: 10,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
  equation: {
    fontFamily: "Linux Libertine O",
    fontSize: 9,
    lineSpacing: 1.0,
    alignment: "center",
    bold: false,
    italic: false,
    spacingBeforePt: 0,
    spacingAfterPt: 0,
    color: "000000",
  },
  references: {
    fontFamily: "Linux Libertine O",
    fontSize: 7,
    lineSpacing: 1.0,
    alignment: "both",
    bold: false,
    italic: false,
    spacingBeforePt: 0,
    spacingAfterPt: 3,
    hangingIndentCm: 0.63,
    lineSpacingPt: 8.4,
    lineSpacingRule: "atLeast",
    color: "000000",
  },
};

export const DEFAULT_CONFERENCE_FORMATTING_CONFIG: ConferenceFormattingConfig = {
  pubform: DEFAULT_PUBLICATION_FORMATTING_CONFIG,
  acm: DEFAULT_ACM_FORMATTING_CONFIG,
};
