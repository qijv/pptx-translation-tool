import JSZip from "jszip";
import type { ExportLayout, ParseResult, TextRecord, TextSpan, TranslationMap, TranslationParseResult } from "../types";
import { escapeText, parseRichTranslation } from "./richText";
import { browserTextMeasurer, fitTextBody, loadFitContext } from "./layoutFit";
import type { MeasureText } from "./layoutFit";

const slidePathPattern = /^ppt\/slides\/slide(\d+)\.xml$/;

const drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
const minBatchLineLimit = 1;
const minBatchPageLimit = 1;

export type CopyMode = "all" | "page" | "line";

function directElements(node: Element) {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1);
}

function isTextChild(node: Node): node is Element {
  return node.nodeType === 1 && (node as Element).namespaceURI === drawingNamespace
    && ["r", "fld", "br"].includes((node as Element).localName);
}

function styleSignature(element: Element): string {
  const attributes = Array.from(element.attributes)
    .filter((attribute) => attribute.namespaceURI !== "http://www.w3.org/2000/xmlns/")
    .map((attribute) => `${attribute.namespaceURI ?? ""}:${attribute.localName}=${JSON.stringify(attribute.value)}`).sort();
  return JSON.stringify([element.namespaceURI, element.localName, attributes, directElements(element).map(styleSignature)]);
}

function extractParagraph(paragraph: Element) {
  const groups: Array<TextSpan & { style: string }> = [];
  const styleLengths = new Map<string, number>();
  Array.from(paragraph.childNodes).forEach((child, childIndex) => {
    if (!isTextChild(child)) return;
    const kind = child.localName === "br" ? "break" : child.localName === "fld" ? "field" : "run";
    const text = kind === "break" ? "\n" : Array.from(child.getElementsByTagNameNS(drawingNamespace, "t"))
      .map((node) => node.textContent ?? "").join("");
    const properties = directElements(child).find((element) => element.localName === "rPr");
    const style = properties ? styleSignature(properties) : "";
    const previous = groups.at(-1);
    if (kind === "run") styleLengths.set(style, (styleLengths.get(style) ?? 0) + text.length);
    if (kind === "run" && previous?.kind === "run" && previous.style === style) previous.text += text;
    else groups.push({ childIndex, kind, text, style, tagId: null });
  });
  const baseStyle = [...styleLengths].sort((a, b) => b[1] - a[1])[0]?.[0];
  const baseChildIndex = groups.find((span) => span.kind === "run" && span.style === baseStyle)?.childIndex
    ?? groups[0]?.childIndex ?? 0;
  let tagId = 0;
  const spans: TextSpan[] = groups.map(({ style, ...span }) => ({
    ...span, tagId: span.kind !== "run" || style !== baseStyle ? ++tagId : null,
  }));
  const source = spans.map((span) => span.text).join("");
  const markedSource = spans.map((span) => {
    if (span.kind === "break") return `<br${span.tagId}/>`;
    const text = escapeText(span.text);
    return span.tagId === null ? text : `<s${span.tagId}>${text}</s${span.tagId}>`;
  }).join("");
  return { spans, source, markedSource, baseChildIndex };
}

export async function extractTextRecords(file: File): Promise<ParseResult> {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("当前版本仅支持 .pptx 文件，请先将 PPT 另存为 .pptx。");
  }

  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const slideEntries = Object.keys(zip.files)
    .map((path) => {
      const match = path.match(slidePathPattern);
      return match ? { path, slideNumber: Number(match[1]) } : null;
    })
    .filter((entry): entry is { path: string; slideNumber: number } => Boolean(entry))
    .sort((a, b) => a.slideNumber - b.slideNumber);

  const records: TextRecord[] = [];
  let globalIndex = 1;

  for (const entry of slideEntries) {
    const xml = await zip.file(entry.path)?.async("string");
    if (!xml) continue;

    const document = parseXml(xml, entry.path);
    const paragraphs = Array.from(document.getElementsByTagNameNS(drawingNamespace, "p"));
    let slideTextIndex = 1;

    paragraphs.forEach((paragraph, paragraphIndex) => {
      const { source, markedSource, spans, baseChildIndex } = extractParagraph(paragraph);
      if (!source.trim()) return;

      records.push({
        id: `${entry.slideNumber}-${slideTextIndex}`,
        globalIndex,
        slideTextIndex,
        slideNumber: entry.slideNumber,
        path: entry.path,
        paragraphIndex,
        baseChildIndex,
        spans,
        source,
        markedSource,
      });
      globalIndex += 1;
      slideTextIndex += 1;
    });
  }

  return {
    fileName: file.name,
    records,
    slideCount: slideEntries.length,
    charCount: getSourceCharCount(records),
  };
}

export function formatRecordsForAi(records: TextRecord[]): string {
  return formatRecordsGroupedBySlide(records);
}

export function formatRecordsGroupedBySlide(records: TextRecord[]): string {
  const lines: string[] = [];
  let activeSlideNumber: number | null = null;

  records.forEach((record) => {
    if (record.slideNumber !== activeSlideNumber) {
      if (lines.length > 0) lines.push("");
      lines.push(`第 ${record.slideNumber} 页`);
      activeSlideNumber = record.slideNumber;
    }
    lines.push(`[${record.id}] ${record.markedSource}`);
  });

  return lines.join("\n");
}

export function buildCopyBatches(records: TextRecord[], mode: CopyMode, limit: number): TextRecord[][] {
  if (mode === "all") return records.length > 0 ? [records] : [];
  if (mode === "page") return buildPageBatches(records, limit);
  return buildLineBatches(records, limit);
}

export function getBatchSummary(records: TextRecord[]) {
  const slides = new Set(records.map((record) => record.slideNumber));
  return {
    pages: slides.size,
    lines: records.length,
    chars: getSourceCharCount(records),
  };
}

export function getSourceCharCount(records: TextRecord[]) {
  return records.reduce((total, record) => total + record.source.length, 0);
}

function buildPageBatches(records: TextRecord[], maxPages: number): TextRecord[][] {
  if (records.length === 0) return [];

  const safeMaxPages = Math.max(minBatchPageLimit, Math.floor(maxPages));
  const slideNumbers = Array.from(new Set(records.map((record) => record.slideNumber))).sort((a, b) => a - b);
  const batches: TextRecord[][] = [];

  for (let index = 0; index < slideNumbers.length; index += safeMaxPages) {
    const slideSet = new Set(slideNumbers.slice(index, index + safeMaxPages));
    batches.push(records.filter((record) => slideSet.has(record.slideNumber)));
  }

  return batches;
}

function buildLineBatches(records: TextRecord[], maxLines: number): TextRecord[][] {
  if (records.length === 0) return [];

  const safeMaxLines = Math.max(minBatchLineLimit, Math.floor(maxLines));
  const batches: TextRecord[][] = [];

  for (let index = 0; index < records.length; index += safeMaxLines) {
    batches.push(records.slice(index, index + safeMaxLines));
  }

  return batches;
}

export function parseTranslations(input: string, records: TextRecord[] = []): TranslationParseResult {
  const translations: TranslationMap = new Map();
  const invalidLines: string[] = [];
  const formatIssues: TranslationParseResult["formatIssues"] = [];
  const repairedIds: string[] = [];
  const idAliases = buildIdAliases(records);
  const knownIds = new Set(records.map((record) => record.id));
  let activeId: string | null = null;

  input.split(/\r?\n/).forEach((rawLine, index) => {
    const line = rawLine.trimStart();
    if (!line.trim()) return;
    if (isIgnoredTranslationLine(line)) {
      activeId = null;
      return;
    }

    // A bare number can be正文, a year or a quantity. Only bracketed numeric
    // aliases or explicit page-text IDs count as record headers.
    const idPattern = "(?:S\\d+\\s*[-‐‑‒–—−]\\s*T\\d+|\\d+(?:\\s*[-‐‑‒–—−]\\s*\\d+)?)";
    const bracketed = line.match(new RegExp(`^[\\[［]\\s*(${idPattern})\\s*[\\]］](.*)$`, "i"));
    const explicit = bracketed ?? line.match(new RegExp(`^(${idPattern})\\s*[\\]］](.*)$`, "i"))
      ?? line.match(/^(S\d+\s*[-‐‑‒–—−]\s*T\d+|\d+\s*[-‐‑‒–—−]\s*\d+)(?=\s|[:：])(?:[ \t]*[:：][ \t]*|[ \t])(.*)$/i);
    if (!explicit) {
      invalidLines.push(line);
      formatIssues.push({
        id: activeId ?? `第 ${index + 1} 行`,
        message: `第 ${index + 1} 行没有有效编号：${line.slice(0, 80)}。若为续行，请合并到对应编号行；若为说明，请删除。`,
      });
      activeId = null;
      return;
    }
    const rawId = explicit[1];
    const id = rawId.replace(/[ \t]/g, "").replace(/[‐‑‒–—−]/g, "-");
    activeId = normalizeTranslationId(id, idAliases);
    if (records.length > 0 && !knownIds.has(activeId)) {
      invalidLines.push(line);
      formatIssues.push({ id: activeId, message: `第 ${index + 1} 行的编号不属于本次 PPT，请检查编号或删除该行` });
      activeId = null;
      return;
    }
    if (translations.has(activeId)) formatIssues.push({ id: activeId, message: "同一 ID 出现多次，请只保留一份译文" });
    const payload = explicit[2].replace(/^(?:[:：][ \t]*|[ \t])/, "");
    translations.set(activeId, payload);
    if (rawId !== id || line.startsWith("［") || (!bracketed && /[\]］]/.test(line))) repairedIds.push(activeId);
  });

  for (const record of records) {
    const translation = translations.get(record.id);
    if (translation === undefined) continue;
    const parsed = parseRichTranslation(translation, record);
    parsed.issues.forEach((message) => formatIssues.push({ id: record.id, message }));
    if (parsed.repaired && parsed.issues.length === 0 && !repairedIds.includes(record.id)) repairedIds.push(record.id);
    translations.set(record.id, parsed.normalized);
  }
  return { translations, invalidLines, formatIssues, repairedIds };
}

export function summarizeTranslationCoverage(records: TextRecord[], translations: TranslationMap) {
  const translated = records.filter((record) => translations.has(record.id)).length;
  return {
    total: records.length,
    translated,
    missing: records.length - translated,
  };
}

export async function buildTranslatedPptx(file: File, records: TextRecord[], translations: TranslationMap, layout: ExportLayout = "preserve", options: {
  onLayoutWarning?: (id: string, message: string) => void;
  measureText?: MeasureText;
} = {}) {
  // Revalidate here as well: callers must never export malformed markers as visible PPT text.
  const parsedById = new Map(records.filter((record) => translations.has(record.id)).map((record) => {
    const parsed = parseRichTranslation(translations.get(record.id)!, record);
    if (parsed.issues.length > 0) throw new Error(`[${record.id}] ${parsed.issues.join("；")}`);
    return [record.id, parsed] as const;
  }));
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const serializer = new XMLSerializer();
  const recordsByPath = groupRecordsByPath(records, translations);
  if (layout === "wrap-shrink" && typeof document !== "undefined") await document.fonts?.ready;
  const measureText = layout === "wrap-shrink" ? options.measureText ?? browserTextMeasurer() : undefined;

  for (const [path, slideRecords] of recordsByPath.entries()) {
    const xml = await zip.file(path)?.async("string");
    if (!xml) continue;

    const document = parseXml(xml, path);
    const paragraphs = Array.from(document.getElementsByTagNameNS(drawingNamespace, "p"));
    const translatedBodies = new Set<Element>();

    slideRecords.forEach((record) => {
      const parsed = parsedById.get(record.id);
      const paragraph = paragraphs[record.paragraphIndex];
      if (!parsed || !paragraph) throw new Error(`[${record.id}] 找不到原段落，请重新上传 PPTX。`);
      const children = Array.from(paragraph.childNodes);
      const originals = children.filter(isTextChild);
      const anchor = originals[0];
      if (!anchor) throw new Error(`[${record.id}] 找不到原文本，请重新上传 PPTX。`);
      for (const segment of parsed.segments) {
        const span = record.spans.find((item) => item.tagId === segment.tagId);
        const template = children[segment.tagId === null ? record.baseChildIndex : span!.childIndex] as Element;
        let run: Element;
        if (segment.isBreak || span?.kind === "field") {
          run = template.cloneNode(true) as Element;
        } else {
          run = document.createElementNS(drawingNamespace, "a:r");
          const properties = directElements(template).find((element) => element.localName === "rPr");
          if (properties) run.appendChild(properties.cloneNode(true));
          const text = document.createElementNS(drawingNamespace, "a:t");
          text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
          text.textContent = segment.text;
          run.appendChild(text);
        }
        paragraph.insertBefore(run, anchor);
      }
      originals.forEach((node) => paragraph.removeChild(node));
      if (paragraph.parentNode?.nodeType === 1) translatedBodies.add(paragraph.parentNode as Element);
    });

    const fitContext = layout === "wrap-shrink" ? await loadFitContext(zip, path) : undefined;
    translatedBodies.forEach((body) => {
      if (fitContext) {
        const result = fitTextBody(body, fitContext, measureText);
        if (result.warning) {
          const record = slideRecords.find((item) => paragraphs[item.paragraphIndex]?.parentNode === body)!;
          options.onLayoutWarning?.(record.id, result.warning);
          // Unsupported layouts retain all original settings, including existing font scaling.
          if (result.unsupported) return;
        }
      }
      applyExportLayout(body, layout);
    });

    zip.file(path, serializer.serializeToString(document));
  }

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

function applyExportLayout(body: Element, layout: ExportLayout) {
  if (layout === "preserve") return;
  const document = body.ownerDocument!;
  let properties = directElements(body).find((node) => node.namespaceURI === drawingNamespace && node.localName === "bodyPr");
  if (!properties) {
    properties = document.createElementNS(drawingNamespace, "a:bodyPr");
    body.insertBefore(properties, body.firstChild);
  }
  properties.setAttribute("wrap", "square");
  const fits = directElements(properties).filter((node) => node.namespaceURI === drawingNamespace
    && ["noAutofit", "normAutofit", "spAutoFit"].includes(node.localName));
  fits.forEach((node) => properties!.removeChild(node));
  // Write measured sizes explicitly so the editor cannot auto-shrink below our lower bound.
  const fit = document.createElementNS(drawingNamespace, "a:noAutofit");
  const following = directElements(properties).find((node) => node.namespaceURI === drawingNamespace
    && ["scene3d", "sp3d", "flatTx", "extLst"].includes(node.localName));
  properties.insertBefore(fit, following ?? null);
}

function groupRecordsByPath(records: TextRecord[], translations: TranslationMap) {
  const grouped = new Map<string, TextRecord[]>();

  records.forEach((record) => {
    if (!translations.has(record.id)) return;
    const existing = grouped.get(record.path) ?? [];
    existing.push(record);
    grouped.set(record.path, existing);
  });

  return grouped;
}

function parseXml(xml: string, filePath: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const parserError = document.getElementsByTagName("parsererror")[0];

  if (parserError) {
    throw new Error(`无法解析 ${filePath}，这个 PPTX 可能包含不标准的 XML。`);
  }

  return document;
}

function buildIdAliases(records: TextRecord[]) {
  const aliases = new Map<string, string>();

  records.forEach((record) => {
    aliases.set(record.id, record.id);
    aliases.set(String(record.globalIndex), record.id);
    aliases.set(`S${String(record.slideNumber).padStart(3, "0")}-T${String(record.globalIndex).padStart(4, "0")}`, record.id);
  });

  return aliases;
}

function normalizeTranslationId(id: string, aliases: Map<string, string>) {
  const oldIdMatch = id.match(/^S(\d+)-T(\d+)$/i);
  if (oldIdMatch) {
    const canonical = `S${oldIdMatch[1].padStart(3, "0")}-T${oldIdMatch[2].padStart(4, "0")}`;
    return aliases.get(canonical) ?? id;
  }
  return aliases.get(id) ?? id;
}

function isIgnoredTranslationLine(line: string) {
  return /^```(?:[a-z]*)\s*$/i.test(line)
    || /^[-#=\s]*第\s*\d+\s*页\s*[-#=]*$/.test(line)
    || /^[-#=\s]*Page\s+\d+\s*[-#=]*$/i.test(line);
}
