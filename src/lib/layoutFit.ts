import type JSZip from "jszip";

const a = "http://schemas.openxmlformats.org/drawingml/2006/main";
const p = "http://schemas.openxmlformats.org/presentationml/2006/main";
const emuPerPixel = 9525;
const children = (node?: Element | null) => node ? Array.from(node.childNodes).filter((item): item is Element => item.nodeType === 1) : [];
const child = (node: Element | null | undefined, name: string) => children(node).find((item) => item.namespaceURI === a && item.localName === name);
const desc = (node: Element | Document | undefined, ns: string, name: string) => node ? Array.from(node.getElementsByTagNameNS(ns, name)) : [];
const num = (value: string | null | undefined) => value != null && value !== "" && Number.isFinite(Number(value)) ? Number(value) : undefined;
const percent = (value: string | null | undefined) => value?.endsWith("%") ? (num(value.slice(0, -1)) ?? NaN) / 100 : (num(value) ?? 100000) / 100000;
const attr = (nodes: Array<Element | undefined>, name: string) => nodes.map((node) => node?.getAttribute(name)).find((value) => value != null && value !== "");
const property = (nodes: Array<Element | undefined>, name: string) => nodes.map((node) => child(node, name)).find(Boolean);

export type FontStyle = { size: number; family: string; bold: boolean; italic: boolean; spacing: number };
export type MeasureText = (text: string, style: FontStyle) => number;
export type FitResult = { scale: number; warning?: string; unsupported?: boolean };

export function browserTextMeasurer(): MeasureText | undefined {
  if (typeof document === "undefined") return;
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) return;
  const cache = new Map<string, number>();
  return (text, style) => {
    const font = `${style.italic ? "italic " : ""}${style.bold ? "bold " : ""}${style.size}px ${JSON.stringify(style.family)}, sans-serif`;
    const key = font + "|" + text;
    let width = cache.get(key);
    if (width === undefined) {
      context.font = font;
      width = context.measureText(text).width;
      cache.set(key, width);
    }
    return Math.max(0, width + style.spacing);
  };
}

// Read layout/master/theme only for adaptive export. Missing inherited facts cause a warning, not guessed sizes.
export async function loadFitContext(zip: JSZip, slidePath: string) {
  const read = async (path?: string) => {
    const xml = path && await zip.file(path)?.async("string");
    return xml ? new DOMParser().parseFromString(xml, "application/xml") : undefined;
  };
  const related = async (path: string | undefined, suffix: string) => {
    if (!path) return;
    const slash = path.lastIndexOf("/");
    const rels = await read(`${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`);
    const rel = rels && Array.from(rels.getElementsByTagName("Relationship")).find((node) => (node.getAttribute("Type") || "").endsWith(suffix) && node.getAttribute("TargetMode") !== "External");
    if (!rel) return;
    const target = rel.getAttribute("Target") || "";
    const parts: string[] = [];
    for (const part of (target.startsWith("/") ? target.slice(1) : path.slice(0, slash + 1) + target).split("/")) {
      if (part === "..") parts.pop();
      else if (part && part !== ".") parts.push(part);
    }
    return parts.join("/");
  };
  const layoutPath = await related(slidePath, "/slideLayout");
  const masterPath = await related(layoutPath, "/slideMaster");
  const themePath = await related(masterPath, "/theme");
  const [layout, master, theme, presentation] = await Promise.all([read(layoutPath), read(masterPath), read(themePath), read("ppt/presentation.xml")]);
  return { layout, master, theme, presentation };
}
export type FitContext = Awaited<ReturnType<typeof loadFitContext>>;

function placeholder(shape?: Element) { return desc(shape, p, "ph")[0]; }
function matchingShape(doc: Document | undefined, shape: Element | undefined) {
  const ph = placeholder(shape);
  if (!ph) return;
  const idx = ph.getAttribute("idx") || "0";
  const type = ph.getAttribute("type") || "body";
  return desc(doc, p, "sp").find((candidate) => {
    const match = placeholder(candidate);
    return match && (match.getAttribute("idx") || "0") === idx;
  }) ?? desc(doc, p, "sp").find((candidate) => (placeholder(candidate)?.getAttribute("type") || "body") === type && placeholder(candidate));
}
function bodyOf(shape?: Element) { return desc(shape, p, "txBody")[0]; }
function extent(shape?: Element) { return child(child(children(shape).find((node) => node.namespaceURI === p && node.localName === "spPr"), "xfrm"), "ext"); }

export function fitTextBody(body: Element, context: FitContext, measure?: MeasureText): FitResult {
  const shape = body.parentNode?.nodeType === 1 ? body.parentNode as Element : undefined;
  const layoutShape = matchingShape(context.layout, shape);
  const masterShape = matchingShape(context.master, layoutShape ?? shape);
  const layoutBody = bodyOf(layoutShape), masterBody = bodyOf(masterShape);
  const bodies = [body, layoutBody, masterBody];
  const bodyProperties = bodies.map((item) => child(item, "bodyPr"));
  const unsupported = (reason: string): FitResult => ({ scale: 1, unsupported: true, warning: `${reason}，保留字号，请在 PowerPoint / WPS 中检查。` });
  if (shape?.namespaceURI !== p || shape.localName !== "sp") return unsupported("表格或特殊文本容器暂不支持自动估算");
  if ((attr(bodyProperties, "vert") || "horz") !== "horz" || (num(attr(bodyProperties, "numCol")) ?? 1) > 1
    || bodies.some((item) => { const warp = child(child(item, "bodyPr"), "prstTxWarp"); return warp && warp.getAttribute("prst") !== "textNoShape"; })) return unsupported("竖排、分栏或艺术字暂不支持自动估算");
  const ext = extent(shape) ?? extent(layoutShape) ?? extent(masterShape);
  const width = (num(ext?.getAttribute("cx")) ?? 0) / emuPerPixel;
  const height = (num(ext?.getAttribute("cy")) ?? 0) / emuPerPixel;
  const innerWidth = width - (num(attr(bodyProperties, "lIns")) ?? 91440) / emuPerPixel - (num(attr(bodyProperties, "rIns")) ?? 91440) / emuPerPixel;
  const innerHeight = height - (num(attr(bodyProperties, "tIns")) ?? 45720) / emuPerPixel - (num(attr(bodyProperties, "bIns")) ?? 45720) / emuPerPixel;
  if (innerWidth <= 0 || innerHeight <= 0) return unsupported("缺少有效的文本框尺寸");
  if (!measure) return unsupported("当前环境无法测量字体");
  const originalFit = bodyProperties.map((node) => children(node).find((item) => item.namespaceURI === a
    && ["normAutofit", "noAutofit", "spAutoFit"].includes(item.localName))).find(Boolean);
  const originalScale = originalFit?.localName === "normAutofit" ? percent(originalFit.getAttribute("fontScale")) : 1;
  if (!Number.isFinite(originalScale) || originalScale <= 0 || originalScale > 1) return unsupported("原字号比例不受支持");
  const phType = attr([placeholder(shape), placeholder(layoutShape), placeholder(masterShape)], "type") || "body";
  const masterStyle = desc(context.master, p, "txStyles")[0];
  const masterList = children(masterStyle).find((node) => node.localName === (["title", "ctrTitle"].includes(phType) ? "titleStyle" : placeholder(shape) ? "bodyStyle" : "otherStyle"));
  const defaultStyle = desc(context.presentation, p, "defaultTextStyle")[0];
  const fontScheme = desc(context.theme, a, "fontScheme")[0];
  const themeFont = (typeface: string | undefined, cjk: boolean) => {
    if (!typeface?.startsWith("+")) return typeface;
    const group = child(fontScheme, typeface.startsWith("+mj") ? "majorFont" : "minorFont");
    return child(group, cjk ? "ea" : "latin")?.getAttribute("typeface")
      || children(group).find((node) => node.localName === "font" && node.getAttribute("script") === "Hans")?.getAttribute("typeface")
      || child(group, "latin")?.getAttribute("typeface");
  };
  type Glyph = { text: string; style: FontStyle };
  type Paragraph = { glyphs: Glyph[]; base: FontStyle; props: Array<Element | undefined> };
  const paragraphs: Paragraph[] = [];
  const resolvedRuns: Array<{ node: Element; size: number }> = [];
  for (const paragraph of children(body).filter((node) => node.namespaceURI === a && node.localName === "p")) {
    const props = child(paragraph, "pPr");
    const level = Math.min(8, Math.max(0, num(props?.getAttribute("lvl")) ?? 0));
    const listProperties = (list: Element | undefined) => [child(list, `lvl${level + 1}pPr`), child(list, "defPPr")];
    const paragraphProps = [props, ...listProperties(child(body, "lstStyle")),
      ...[layoutBody, masterBody].flatMap((item) => [child(desc(item, a, "p")[0], "pPr"), ...listProperties(child(item, "lstStyle"))]),
      ...listProperties(masterList), ...listProperties(defaultStyle)];
    if (property(paragraphProps, "tabLst") || attr(paragraphProps, "rtl") === "1") return unsupported("制表位或从右向左排版暂不支持自动估算");
    const defaults = paragraphProps.map((node) => child(node, "defRPr"));
    const glyphs: Glyph[] = [];
    let base: FontStyle | undefined;
    let previousStyle: FontStyle | undefined;
    for (const run of children(paragraph).filter((node) => node.namespaceURI === a && ["r", "fld", "br", "endParaRPr"].includes(node.localName))) {
      const properties = run.localName === "endParaRPr" ? run : child(run, "rPr");
      const styles = [properties, ...defaults];
      const inheritsPrevious = ["br", "endParaRPr"].includes(run.localName);
      const sz = num(attr(styles, "sz")) ?? (inheritsPrevious && previousStyle ? previousStyle.size / originalScale * 75 : undefined);
      if (!sz || sz <= 0) return unsupported("无法确定继承的字号");
      const text = run.localName === "br" ? "\n" : child(run, "t")?.textContent || "";
      if (text.includes("\t")) return unsupported("制表符暂不支持自动估算");
      const cjk = /[\u2e80-\u9fff\uac00-\ud7af]/u.test(text);
      const typeface = themeFont(property(styles, cjk ? "ea" : "latin")?.getAttribute("typeface")
        || property(styles, "latin")?.getAttribute("typeface") || undefined, cjk) || previousStyle?.family;
      if (!typeface) return unsupported("无法确定继承的字体");
      const style: FontStyle = { size: sz / 75 * originalScale, family: typeface, bold: attr(styles, "b") === "1", italic: attr(styles, "i") === "1", spacing: (num(attr(styles, "spc")) ?? 0) / 75 };
      base ??= style;
      previousStyle = style;
      resolvedRuns.push({ node: run, size: sz });
      for (const char of Array.from(text)) glyphs.push({ text: char, style });
    }
    if (!base) return unsupported("空段落的字号无法确定");
    paragraphs.push({ glyphs, base, props: paragraphProps });
  }
  if (!paragraphs.length) return { scale: 1 };
  const minPoints = Math.min(...resolvedRuns.map((run) => run.size / 100 * originalScale));
  const lowerBound = minPoints < 12 ? 1 : Math.min(1, Math.max(.7, 12 / minPoints));
  const spacing = (props: Array<Element | undefined>, name: string, line: number) => {
    const value = property(props, name);
    const pts = num(child(value, "spcPts")?.getAttribute("val"));
    const pct = child(value, "spcPct")?.getAttribute("val");
    return pts !== undefined ? pts / 75 : pct ? line * percent(pct) : name === "lnSpc" ? line : 0;
  };
  const fits = (scale: number) => {
    let total = 0;
    for (const paragraph of paragraphs) {
      const lineWidth = innerWidth - (num(attr(paragraph.props, "marL")) ?? 0) / emuPerPixel - (num(attr(paragraph.props, "marR")) ?? 0) / emuPerPixel;
      const indent = (num(attr(paragraph.props, "indent")) ?? 0) / emuPerPixel;
      const tokens: Glyph[][] = [];
      for (const glyph of paragraph.glyphs) {
        const word = /[\p{L}\p{N}'’]/u.test(glyph.text) && !/[\u2e80-\u9fff\uac00-\ud7af]/u.test(glyph.text);
        const last = tokens.at(-1);
        if (word && last && last.every((item) => /[\p{L}\p{N}'’]/u.test(item.text) && !/[\u2e80-\u9fff\uac00-\ud7af]/u.test(item.text))) last.push(glyph);
        else tokens.push([glyph]);
      }
      let used = Math.max(0, indent), lineSize = paragraph.base.size * scale, first = true;
      const finishLine = () => { total += spacing(paragraph.props, "lnSpc", lineSize * 1.22); used = 0; lineSize = paragraph.base.size * scale; first = false; };
      total += spacing(paragraph.props, "spcBef", lineSize * 1.22);
      for (const token of tokens) {
        if (token[0].text === "\n") { lineSize = Math.max(lineSize, token[0].style.size * scale); finishLine(); continue; }
        if (token[0].text === "\t") return false;
        const tokenWidth = token.reduce((sum, glyph) => sum + measure(glyph.text, { ...glyph.style, size: glyph.style.size * scale, spacing: glyph.style.spacing }), 0);
        const latinBreak = attr(paragraph.props, "latinLnBrk");
        if (["0", "false"].includes(latinBreak || "") && token.length > 1 && tokenWidth > lineWidth * .96) return false;
        if (used > (first ? Math.max(0, indent) : 0) && used + tokenWidth > lineWidth * .96) finishLine();
        if (/^\s+$/u.test(token.map((g) => g.text).join("")) && used === 0) continue;
        for (const glyph of token) {
          const w = measure(glyph.text, { ...glyph.style, size: glyph.style.size * scale });
          if (w > lineWidth * .96) return false;
          if (used + w > lineWidth * .96 && used > 0) finishLine();
          used += w;
          lineSize = Math.max(lineSize, glyph.style.size * scale);
        }
      }
      finishLine();
      total += spacing(paragraph.props, "spcAft", paragraph.base.size * scale * 1.22);
      if (total > innerHeight * .96) return false;
    }
    return total <= innerHeight * .96;
  };
  let scale = 1;
  while (scale > lowerBound && !fits(scale)) scale = Math.max(lowerBound, Math.round((scale - .01) * 10000) / 10000);
  const effectiveScale = scale * originalScale;
  if (effectiveScale < 1) {
    for (const run of resolvedRuns) {
      let properties = run.node.localName === "endParaRPr" ? run.node : child(run.node, "rPr");
      if (!properties) { properties = body.ownerDocument!.createElementNS(a, "a:rPr"); run.node.insertBefore(properties, run.node.firstChild); }
      properties.setAttribute("sz", String(Math.round(run.size * effectiveScale)));
    }
  }
  return { scale: effectiveScale, warning: fits(scale) ? undefined : "已达到字号下限，译文仍可能溢出，请检查此文本框。" };
}
