export type TextRecord = {
  id: string;
  globalIndex: number;
  slideTextIndex: number;
  slideNumber: number;
  path: string;
  paragraphIndex: number;
  baseChildIndex: number;
  spans: TextSpan[];
  source: string;
  markedSource: string;
};

export type TextSpan = {
  childIndex: number;
  kind: "run" | "field" | "break";
  tagId: number | null;
  text: string;
};

export type ParseResult = {
  fileName: string;
  records: TextRecord[];
  slideCount: number;
  charCount: number;
};

export type TranslationMap = Map<string, string>;

export type ExportLayout = "preserve" | "wrap-shrink";

export type TranslationParseResult = {
  translations: TranslationMap;
  invalidLines: string[];
  formatIssues: Array<{ id: string; message: string }>;
  repairedIds: string[];
};
