import type { TextRecord } from "../types";

export type TranslatedSegment = { tagId: number | null; text: string; isBreak?: boolean };

// Numeric entities keep literal angle brackets distinct from formatting markers.
export function escapeText(text: string) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&#60;").replace(/>/g, "&#62;")
    .replace(/＜/g, "&#65308;").replace(/＞/g, "&#65310;")
    .replace(/〈/g, "&#12296;").replace(/〉/g, "&#12297;");
}

function unescapeText(text: string) {
  return text.replace(/&#(?:60|x3c);|&lt;/gi, "<").replace(/&#(?:62|x3e);|&gt;/gi, ">")
    .replace(/&#(65308|65310|12296|12297);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

export function parseRichTranslation(text: string, record: TextRecord) {
  const expected = new Map(record.spans.filter((span) => span.tagId !== null).map((span) => [span.tagId!, span]));
  // Repair spelling of markers only; never invent a missing marker or its range.
  let normalized = text
    .replace(/&lt;(?=\s*\/?\s*(?:s|br)\s*\d)/gi, "<")
    .replace(/(<\s*\/?\s*(?:s|br)\s*\d+\s*\/?)&gt;/gi, "$1>")
    .replace(/[＜〈](?=\s*\/?\s*(?:s|br)\s*\d)/gi, "<")
    .replace(/(<\s*\/?\s*(?:s|br)\s*\d+\s*\/?)[＞〉]/gi, "$1>")
    .replace(/<\s*(\/?)\s*s\s*(\d+)\s*>/gi, (_, slash, id) => `<${slash}s${Number(id)}>`)
    .replace(/<\s*br\s*(\d+)\s*\/?\s*>/gi, (_, id) => `<br${Number(id)}/>`)
    .replace(/(?<!<)\/\s*s\s*(\d+)\s*>/gi, (_, id) => `</s${Number(id)}>`)
    .replace(/(^|[\s>])(?:\/\s*)?s\s*\d+\s*>/gi, (match, prefix) => {
      const marker = match.slice(prefix.length).replace(/\s/g, "").toLowerCase();
      return `${prefix}<${marker}`;
    })
    .replace(/<\s*(\/?)\s*s\s*(\d+)(?=\s|$)(?!\s*>)/gi, (_, slash, id) => `<${slash}s${Number(id)}>`);
  normalized = normalized.replace(/`(<\/?s\d+>|<br\d+\/>)`/g, "$1");
  const segments: TranslatedSegment[] = [];
  const seen = new Set<number>();
  const issues: string[] = [];
  let active: number | null = null;
  let position = 0;
  const addText = (value: string) => {
    if (!value) return;
    const previous = segments.at(-1);
    if (previous && !previous.isBreak && previous.tagId === active) previous.text += unescapeText(value);
    else segments.push({ tagId: active, text: unescapeText(value) });
  };
  for (const match of normalized.matchAll(/<(\/?s|br)(\d+)(\/?)>/g)) {
    addText(normalized.slice(position, match.index));
    position = match.index! + match[0].length;
    const id = Number(match[2]);
    const span = expected.get(id);
    if (!span || (match[1] === "br") !== (span.kind === "break")) {
      issues.push(`出现未知格式标记 ${match[0]}`);
      continue;
    }
    if (match[1] === "/s") {
      if (active !== id) issues.push(`格式标记 ${match[0]} 没有正确配对`);
      active = null;
    } else {
      if (active !== null) issues.push("格式标记不能嵌套或交叉");
      if (seen.has(id)) issues.push(`格式标记 ${id} 重复`);
      seen.add(id);
      if (span.kind === "break") segments.push({ tagId: id, text: "", isBreak: true });
      else {
        active = id;
        segments.push({ tagId: id, text: "" });
      }
    }
  }
  addText(normalized.slice(position));
  if (active !== null) issues.push(`缺少 </s${active}>`);
  for (const [id, span] of expected) {
    if (!seen.has(id)) issues.push(`缺少 ${span.kind === "break" ? `<br${id}/>` : `<s${id}>…</s${id}>`}`);
    else if (span.kind !== "break") {
      const content = segments.filter((segment) => segment.tagId === id).map((segment) => segment.text).join("");
      if (!content.trim() && span.text.trim()) issues.push(`标记 s${id} 内没有译文`);
      if (span.kind === "field" && content !== span.text) issues.push(`动态字段 s${id} 必须保留原值`);
    }
  }
  const remaining = normalized.replace(/<\/?s\d+>|<br\d+\/>/g, "");
  if (/<\s*\/?\s*(?:s|br)\s*\d/i.test(remaining)) issues.push("存在无法识别的格式标记");
  if (record.source.trim() && !segments.some((segment) => segment.text.trim())) issues.push("译文为空");
  return { segments, issues: [...new Set(issues)], normalized, repaired: normalized !== text };
}
