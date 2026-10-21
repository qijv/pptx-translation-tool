import assert from "node:assert/strict";
import { test } from "node:test";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { buildTranslatedPptx, extractTextRecords, formatRecordsForAi, parseTranslations, buildCopyBatches } from "../src/lib/pptxText";
import { buildPromptedCopyText } from "../src/lib/copyPrompt";
import { parseRichTranslation } from "../src/lib/richText";

Object.assign(globalThis, { DOMParser, XMLSerializer });
const ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
const plain = '<a:rPr lang="zh-CN"/>';
const bold = '<a:rPr lang="zh-CN" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>';
const run = (text: string, style = plain) => `<a:r>${style}<a:t xml:space="preserve">${text}</a:t></a:r>`;
const paragraph = (body: string) => `<a:p><a:pPr algn="ctr"/>${body}<a:endParaRPr lang="zh-CN"/></a:p>`;
async function fixture(body: string, extra = "") {
  const zip = new JSZip();
  zip.file("ppt/slides/slide1.xml", `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:cSld><a:txBody>${body}</a:txBody></p:cSld></p:sld>`);
  zip.file("ppt/slides/slide2.xml", `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:cSld><a:txBody>${extra}</a:txBody></p:cSld></p:sld>`);
  zip.file("untouched.bin", "UNCHANGED");
  return new File([await zip.generateAsync({ type: "uint8array" })], "test.pptx");
}
async function output(file: File, input: string) {
  const parsed = await extractTextRecords(file);
  const translations = parseTranslations(input, parsed.records);
  const blob = await buildTranslatedPptx(file, parsed.records, translations.translations);
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  return { zip, xml, document, parsed, translations };
}
function textOf(element: Element) {
  return Array.from(element.getElementsByTagNameNS(ns, "t")).map((item) => item.textContent).join("");
}

test("same paragraph stays one record, merging adjacent identical styles and retaining spaces", async () => {
  const file = await fixture(paragraph(run("Machine ") + run("inspection")));
  const { records } = await extractTextRecords(file);
  assert.equal(records.length, 1);
  assert.equal(records[0].markedSource, "Machine inspection");
  assert.equal(records[0].spans.length, 1);
  const { document } = await output(file, "[1-1] Machine inspection ");
  assert.equal(textOf(document as unknown as Element), "Machine inspection ");
});

test("equivalent style attribute order merges without discarding hyperlink properties", async () => {
  const first = '<a:rPr lang="en-US" b="1"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:rPr>';
  const second = '<a:rPr b="1" lang="en-US"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:rPr>';
  const file = await fixture(paragraph(run("Machine ", first) + run("inspection", second)));
  const { records } = await extractTextRecords(file);
  assert.equal(records[0].markedSource, "Machine inspection");
  const { document } = await output(file, "[1-1] Machine inspection");
  assert.equal(document.getElementsByTagNameNS(ns, "hlinkClick")[0].getAttribute("r:id"), "rId2");
});

test("feedback phrase is one paragraph; reorder translated emphasis while preserving XML style", async () => {
  const file = await fixture(paragraph(run("请检查 ") + run("红色按钮", bold)));
  const { records } = await extractTextRecords(file);
  assert.equal(records[0].markedSource, "请检查 <s1>红色按钮</s1>");
  const result = await output(file, "[1-1] <s1>red button</s1>: please check.");
  const runs = Array.from(result.document.getElementsByTagNameNS(ns, "r"));
  assert.equal(textOf(runs[0] as unknown as Element), "red button");
  assert.equal(runs[0].getElementsByTagNameNS(ns, "rPr")[0].getAttribute("b"), "1");
  assert.equal(runs[0].getElementsByTagNameNS(ns, "srgbClr")[0].getAttribute("val"), "FF0000");
  assert.equal(textOf(result.document as unknown as Element), "red button: please check.");
  assert.equal(result.document.getElementsByTagNameNS(ns, "pPr")[0].getAttribute("algn"), "ctr");
  assert.equal(await result.zip.file("untouched.bin")!.async("string"), "UNCHANGED");
  assert.ok(!result.xml.includes("&lt;s1"));
});

test("spaces between marked English words survive export without inserting guessed spaces", async () => {
  const file = await fixture(paragraph(run("检查机器 ") + run("按钮", bold)));
  const { document } = await output(file, "[1-1] Check the <s1>red button</s1>, please.");
  assert.equal(textOf(document as unknown as Element), "Check the red button, please.");
  const chinese = await output(file, "[1-1] 检查<s1>红色按钮</s1>。");
  assert.equal(textOf(chinese.document as unknown as Element), "检查红色按钮。");
});

test("paragraph edge spaces and leading minus are not mistaken for ID separators", async () => {
  const file = await fixture(paragraph(run("  -5 degrees  ")));
  const { records } = await extractTextRecords(file);
  const roundtrip = await output(file, formatRecordsForAi(records));
  assert.equal(textOf(roundtrip.document as unknown as Element), "  -5 degrees  ");
  assert.equal(parseTranslations("[1-1] -5 degrees", records).translations.get("1-1"), "-5 degrees");
  assert.ok(parseTranslations("[1-1] ", records).formatIssues.length > 0);
});

test("two different emphasis ranges may reorder independently", async () => {
  const blue = '<a:rPr i="1"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr>';
  const file = await fixture(paragraph(run("先检查设备，然后检查 ") + run("按钮", bold) + run("和 ") + run("电源", blue)));
  const { document } = await output(file, "[1-1] Inspect the equipment, then <s2>power</s2> and <s1>button</s1>.");
  const runs = Array.from(document.getElementsByTagNameNS(ns, "r"));
  const power = runs.find((item) => textOf(item as unknown as Element) === "power")!;
  const button = runs.find((item) => textOf(item as unknown as Element) === "button")!;
  assert.equal(power.getElementsByTagNameNS(ns, "srgbClr")[0].getAttribute("val"), "0000FF");
  assert.equal(button.getElementsByTagNameNS(ns, "srgbClr")[0].getAttribute("val"), "FF0000");
});

test("observed Google Translate output with stripped markers is rejected", async () => {
  const file = await fixture(paragraph(run("补充") + run("机床检测", bold) + run("说明")));
  const { records } = await extractTextRecords(file);
  assert.equal(records[0].markedSource, "补充<s1>机床检测</s1>说明");
  const parsed = parseTranslations("[1-1] Supplementary instructions for machine tool inspection.", records);
  assert.ok(parsed.formatIssues.some((item) => item.message.includes("缺少 <s1>")));
  await assert.rejects(buildTranslatedPptx(file, records, parsed.translations));
});

const repairs = [
  "Check < s 1 >red button< / s 1 >.",
  "Check ＜s1＞red button＜/s1＞.",
  "Check &lt;s1&gt;red button&lt;/s1&gt;.",
  "Check s1>red button/s1>.",
  "Check <s1 red button</s1 .",
  "Check `<s1>`red button`</s1>`.",
  "Check <S01>red button</S01>.",
];
for (const text of repairs) test(`repair marker spelling: ${text}`, async () => {
  const file = await fixture(paragraph(run("请检查 ") + run("红色按钮", bold)));
  const { records } = await extractTextRecords(file);
  const parsed = parseTranslations(`[1-1] ${text}`, records);
  assert.deepEqual(parsed.formatIssues, []);
  assert.deepEqual(parsed.repairedIds, ["1-1"]);
  assert.ok(parsed.translations.get("1-1")!.includes("<s1>"));
  assert.ok(parsed.translations.get("1-1")!.includes("</s1>"));
});

const failures = [
  "Check red button.",
  "Check <样式1>red button</样式1>.",
  "Check <s1>red button.",
  "Check <s1>red</s1> <s1>button</s1>.",
  "Check <s1><s1>red button</s1></s1>.",
  "Check <s1>red button</s2>.",
  "Check <s2>red button</s2>.",
  "Check <s1></s1>.",
];
for (const text of failures) test(`reject unsafe marker loss: ${text}`, async () => {
  const file = await fixture(paragraph(run("请检查 ") + run("红色按钮", bold)));
  const { records } = await extractTextRecords(file);
  const parsed = parseTranslations(`[1-1] ${text}`, records);
  assert.ok(parsed.formatIssues.length > 0);
  await assert.rejects(buildTranslatedPptx(file, records, parsed.translations));
});

test("paragraph, slide boundaries and partial translations remain independent", async () => {
  const file = await fixture(paragraph(run("one")) + paragraph(run("two")), paragraph(run("three")));
  const { records } = await extractTextRecords(file);
  assert.deepEqual(records.map((record) => record.id), ["1-1", "1-2", "2-1"]);
  assert.equal(buildCopyBatches(records, "page", 1).length, 2);
  assert.equal(buildCopyBatches(records, "line", 2).length, 2);
  const { document, zip } = await output(file, "[1-1] first");
  assert.equal(textOf(document as unknown as Element), "firsttwo");
  const originalZip = await JSZip.loadAsync(await file.arrayBuffer());
  assert.equal(await zip.file("ppt/slides/slide2.xml")!.async("string"), await originalZip.file("ppt/slides/slide2.xml")!.async("string"));
});

test("explicit breaks and dynamic fields retain their original nodes and attributes", async () => {
  const field = `<a:fld id="{FIELD-ID}" type="slidenum">${plain}<a:t>1</a:t></a:fld>`;
  const file = await fixture(paragraph(run("Page ") + field + '<a:br><a:rPr lang="zh-CN"/></a:br>' + run("subtitle")));
  const { records } = await extractTextRecords(file);
  assert.equal(records[0].markedSource, "Page <s1>1</s1><br2/>subtitle");
  const { document } = await output(file, "[1-1] Slide <s1>1</s1><br2>caption");
  assert.equal(document.getElementsByTagNameNS(ns, "br").length, 1);
  assert.equal(document.getElementsByTagNameNS(ns, "fld")[0].getAttribute("id"), "{FIELD-ID}");
  assert.equal(document.getElementsByTagNameNS(ns, "fld")[0].getAttribute("type"), "slidenum");
  assert.ok(parseTranslations("[1-1] Slide <s1>2</s1><br2/>caption", records).formatIssues.length > 0);
});

test("literal angle brackets, ampersands and marker-looking source text remain literal", async () => {
  const file = await fixture(paragraph(run("A &amp; B &lt;s1&gt; &lt; 3 ＜s2＞")));
  const { records } = await extractTextRecords(file);
  assert.equal(records[0].markedSource, "A &amp; B &#60;s1&#62; &#60; 3 &#65308;s2&#65310;");
  const { document } = await output(file, formatRecordsForAi(records));
  assert.equal(textOf(document as unknown as Element), "A & B <s1> < 3 ＜s2＞");
});

test("prompt explains non-translatable markers and supplies an actual example", async () => {
  const { records } = await extractTextRecords(await fixture(paragraph(run("请检查 ") + run("红色按钮", bold))));
  const prompt = buildPromptedCopyText({ records, fileName: "sample.pptx", targetLanguage: "英文", batchLabel: "全部" });
  assert.ok(prompt.includes("不要翻译或修改标记名称"));
  assert.ok(prompt.includes("Check the <s1>red button</s1>, please."));
  assert.ok(prompt.includes("词间空格"));
});

test("duplicate and unknown IDs are flagged; fences and page headings are accepted", async () => {
  const { records } = await extractTextRecords(await fixture(paragraph(run("original"))));
  assert.ok(parseTranslations("[1-1] first\n[1-1] second", records).formatIssues.length > 0);
  const parsed = parseTranslations("```text\n第 1 页\n[1-1] first\n[9-9] unknown\n```", records);
  assert.equal(parsed.translations.get("1-1"), "first");
  assert.equal(parsed.translations.size, 1);
  assert.equal(parsed.invalidLines.length, 1);
  assert.ok(parsed.formatIssues.some((issue) => issue.id === "9-9"));
});

test("numeric continuation cannot be silently dropped or mistaken for a numeric alias", async () => {
  const { records } = await extractTextRecords(await fixture(paragraph(run("first")) + paragraph(run("second"))));
  for (const continuation of ["2026 increased.", "2 percent growth"]) {
    const parsed = parseTranslations(`[1-1] Revenue in\n${continuation}\n[1-2] second`, records);
    assert.equal(parsed.translations.get("1-1"), "Revenue in");
    assert.equal(parsed.translations.get("1-2"), "second");
    assert.deepEqual(parsed.invalidLines, [continuation]);
    assert.ok(parsed.formatIssues.some((issue) => issue.message.includes("第 2 行")));
  }
});

test("ID typography variants match the correct record without entering the preceding paragraph", async () => {
  const { records } = await extractTextRecords(await fixture(paragraph(run("first")) + paragraph(run("second"))));
  for (const header of ["[1–2]", "［ 1 — 2 ］", "1-2]", "1-2"]) {
    const parsed = parseTranslations(`[1-1] first\n${header} second`, records);
    assert.equal(parsed.translations.get("1-1"), "first");
    assert.equal(parsed.translations.get("1-2"), "second");
    assert.deepEqual(parsed.formatIssues, []);
  }
  const invalid = parseTranslations("[1-1] first\n[one-two] second", records);
  assert.equal(invalid.translations.get("1-1"), "first");
  assert.ok(invalid.formatIssues.length > 0);
});

test("unnumbered AI explanations and ordinary continuation lines cannot enter translated text", async () => {
  const { records } = await extractTextRecords(await fixture(paragraph(run("first"))));
  for (const extra of ["Translation completed.", "翻译完成。", "second line", "第 1 页有一个图表"]) {
    const parsed = parseTranslations(`[1-1] translated\n${extra}`, records);
    assert.equal(parsed.translations.get("1-1"), "translated");
    assert.deepEqual(parsed.invalidLines, [extra]);
    assert.ok(parsed.formatIssues.length > 0);
  }
});

test("legacy IDs must match both original slide and global index", async () => {
  const { records } = await extractTextRecords(await fixture(paragraph(run("first"))));
  assert.equal(parseTranslations("[S1-T1] translated", records).translations.get("1-1"), "translated");
  const wrongPage = parseTranslations("[S999-T1] wrong", records);
  assert.equal(wrongPage.translations.size, 0);
  assert.ok(wrongPage.formatIssues.length > 0);
});

test("empty paragraphs are skipped and equivalent namespace prefixes are supported", async () => {
  const file = await fixture(paragraph(run(" ")) + paragraph(run("text")));
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
  zip.file("ppt/slides/slide1.xml", xml.replaceAll("a:", "d:").replace("xmlns:a=", "xmlns:d="));
  const alternate = new File([await zip.generateAsync({ type: "uint8array" })], "alternate.pptx");
  const { records } = await extractTextRecords(alternate);
  assert.equal(records.length, 1);
  assert.equal(records[0].paragraphIndex, 1);
  const parsed = parseRichTranslation("translated", records[0]);
  assert.equal(parsed.issues.length, 0);
});

test("export layout defaults preserve original body properties and geometry", async () => {
  const file = await fixture(paragraph(run("original", bold)));
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
  zip.file("ppt/slides/slide1.xml", xml.replace("<a:txBody>", '<a:txBody><a:bodyPr wrap="none" anchor="ctr" lIns="123"><a:spAutoFit/></a:bodyPr>'));
  const source = new File([await zip.generateAsync({ type: "uint8array" })], "layout.pptx");
  const parsed = await extractTextRecords(source);
  const blob = await buildTranslatedPptx(source, parsed.records, new Map([["1-1", "translated"]]));
  const result = await JSZip.loadAsync(await blob.arrayBuffer());
  const doc = new DOMParser().parseFromString(await result.file("ppt/slides/slide1.xml")!.async("string"), "application/xml");
  const props = doc.getElementsByTagNameNS(ns, "bodyPr")[0];
  assert.equal(props.getAttribute("wrap"), "none");
  assert.equal(props.getAttribute("anchor"), "ctr");
  assert.equal(props.getAttribute("lIns"), "123");
  assert.equal(props.getElementsByTagNameNS(ns, "spAutoFit").length, 1);
});

const explicitRun = (text: string, size = 2400, extras = "") => run(text, `<a:rPr sz="${size}" ${extras}><a:latin typeface="Arial"/></a:rPr>`);
async function fitFixture(body: string, width = 400, height = 120, props = '<a:noAutofit/>', extraShape = "") {
  const zip = new JSZip();
  zip.file("ppt/slides/slide1.xml", `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:ext cx="${width * 9525}" cy="${height * 9525}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr wrap="none">${props}</a:bodyPr><a:lstStyle/>${body}</p:txBody></p:sp>${extraShape}</p:spTree></p:cSld></p:sld>`);
  return new File([await zip.generateAsync({type:"uint8array"})], "fit.pptx");
}
const measureFit = (text: string, style: {size: number}) => text.length * style.size * .5;
async function fitOutput(file: File, text: string) {
  const {records} = await extractTextRecords(file);
  const parsed = parseTranslations(text,records);
  const warnings: Array<{id:string;message:string}> = [];
  const blob = await buildTranslatedPptx(file,records,parsed.translations,"wrap-shrink",{measureText:measureFit,onLayoutWarning:(id,message)=>warnings.push({id,message})});
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const doc = new DOMParser().parseFromString(await zip.file("ppt/slides/slide1.xml")!.async("string"),"application/xml");
  return {doc,warnings,zip};
}

test("adaptive export leaves short text sizes unchanged", async () => {
  const file = await fitFixture(paragraph(explicitRun("short")));
  const {doc,warnings} = await fitOutput(file,"[1-1] Short text");
  assert.deepEqual(warnings,[]);
  assert.equal(doc.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("sz"),"2400");
  assert.equal(doc.getElementsByTagNameNS(ns,"normAutofit").length,0);
  assert.equal(doc.getElementsByTagNameNS(ns,"noAutofit").length,1);
});

test("adaptive export fits long text by reducing size, retaining mixed styles and complete text", async () => {
  const file = await fitFixture(paragraph(explicitRun("请检查设备和 ") + explicitRun("红色按钮",2400,'b="1"')),300,115);
  const text = "Please inspect the <s1>red button</s1> before operating the machine.";
  const {doc,warnings} = await fitOutput(file,`[1-1] ${text}`);
  assert.deepEqual(warnings,[]);
  const runs=Array.from(doc.getElementsByTagNameNS(ns,"r"));
  const sizes=runs.map(r=>Number(r.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("sz")));
  assert.ok(sizes.every(size=>size<2400 && size>=1680));
  const boldRun=runs.find(r=>textOf(r as unknown as Element)==="red button")!;
  assert.equal(boldRun.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("b"),"1");
  assert.equal(textOf(doc as unknown as Element),"Please inspect the red button before operating the machine.");
});

test("adaptive export stops at 70 percent and reports overflow rather than deleting text", async () => {
  const file = await fitFixture(paragraph(explicitRun("原文")),180,45);
  const text="This is intentionally extremely long text with many words that cannot fit even at the permitted minimum font size.";
  const {doc,warnings}=await fitOutput(file,`[1-1] ${text}`);
  assert.equal(doc.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("sz"),"1680");
  assert.equal(warnings.length,1);
  assert.equal(warnings[0].id,"1-1");
  assert.ok(warnings[0].message.includes("字号下限"));
  assert.equal(textOf(doc as unknown as Element),text);
});

test("12 point floor and originally tiny fonts are respected", async () => {
  for(const size of [1400,1000]){
    const file=await fitFixture(paragraph(explicitRun("原文",size)),100,35);
    const {doc,warnings}=await fitOutput(file,"[1-1] Very long translated text cannot fit inside this small box.");
    assert.equal(doc.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("sz"),String(size===1400?1200:1000));
    assert.equal(warnings.length,1);
  }
});

test("unsupported dimensions retain original settings and return a specific warning", async () => {
  const file=await fixture(paragraph(run("原文")));
  const {doc,warnings}=await fitOutput(file,"[1-1] Translated text");
  assert.equal(warnings.length,1);
  assert.ok(warnings[0].message.includes("保留字号"));
  assert.equal(doc.getElementsByTagNameNS(ns,"bodyPr").length,0);
});

test("existing font scale is materialized without making text larger", async () => {
  const file=await fitFixture(paragraph(explicitRun("原文")),400,120,'<a:normAutofit fontScale="80000"/>');
  const {doc,warnings}=await fitOutput(file,"[1-1] Short text");
  assert.deepEqual(warnings,[]);
  assert.equal(doc.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("sz"),"1920");
  assert.equal(doc.getElementsByTagNameNS(ns,"normAutofit").length,0);
});

test("dynamic font size is inherited from the slide master and theme", async () => {
  const file=await fitFixture(paragraph(run("原文",'<a:rPr/>')),200,100);
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const xml=await zip.file("ppt/slides/slide1.xml")!.async("string");
  zip.file("ppt/slides/slide1.xml",xml.replace('<p:sp><p:spPr>','<p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>'));
  const relns='http://schemas.openxmlformats.org/package/2006/relationships';
  const reltype='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('ppt/slides/_rels/slide1.xml.rels',`<Relationships xmlns="${relns}"><Relationship Id="rId1" Type="${reltype}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout1.xml',`<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr></p:sp></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',`<Relationships xmlns="${relns}"><Relationship Id="rId1" Type="${reltype}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml',`<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',`<Relationships xmlns="${relns}"><Relationship Id="rId1" Type="${reltype}/theme" Target="../theme/theme1.xml"/></Relationships>`);
  zip.file('ppt/theme/theme1.xml',`<a:theme xmlns:a="${ns}"><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  const inherited=new File([await zip.generateAsync({type:"uint8array"})],"inherited.pptx");
  const {doc}=await fitOutput(inherited,"[1-1] Check the machine carefully before starting work.");
  const size=Number(doc.getElementsByTagNameNS(ns,"rPr")[0].getAttribute("sz"));
  assert.ok(size>=1680 && size<2400);
});

test("each translated text box fits independently and untranslated boxes remain unchanged", async () => {
  const original = await fitFixture(paragraph(explicitRun("first",2400)),300,145);
  const zip = await JSZip.loadAsync(await original.arrayBuffer());
  const xml=await zip.file("ppt/slides/slide1.xml")!.async("string");
  const shape = (text: string,size: number) => `<p:sp><p:spPr><a:xfrm><a:ext cx="3810000" cy="1143000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr>${paragraph(explicitRun(text,size))}</p:txBody></p:sp>`;
  zip.file("ppt/slides/slide1.xml",xml.replace('</p:spTree>',shape('second',1400)+shape('third',1600)+'</p:spTree>'));
  const file=new File([await zip.generateAsync({type:'uint8array'})],'independent.pptx');
  const {doc,warnings}=await fitOutput(file,'[1-1] Please inspect all equipment carefully before operating the machine.\n[1-2] Short');
  assert.deepEqual(warnings,[]);
  const bodies=Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main','txBody'));
  assert.ok(Number(bodies[0].getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'))<2400);
  assert.equal(bodies[1].getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'),'1400');
  assert.equal(bodies[2].getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'),'1600');
  assert.equal(bodies[2].getElementsByTagNameNS(ns,'bodyPr')[0].getAttribute('wrap'),'none');
  assert.equal(bodies[2].getElementsByTagNameNS(ns,'spAutoFit').length,1);
});

test("explicit breaks and fixed paragraph spacing count toward the fitting height", async () => {
  const first=paragraph(explicitRun('one')+'<a:br/>'+explicitRun('two')).replace('<a:pPr algn="ctr"/>','<a:pPr><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr>');
  const file=await fitFixture(first+paragraph(explicitRun('three')),500,120);
  const {doc,warnings}=await fitOutput(file,'[1-1] First line<br1/>second line\n[1-2] Third line');
  assert.deepEqual(warnings,[]);
  assert.ok(Number(doc.getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'))<2400);
  assert.equal(doc.getElementsByTagNameNS(ns,'br').length,1);
  assert.equal(doc.getElementsByTagNameNS(ns,'spcPts')[0].getAttribute('val'),'600');
});

test("local default paragraph style takes precedence over presentation defaults", async () => {
  const original = await fitFixture(paragraph(run("原文", '<a:rPr/>')),180,45);
  const zip = await JSZip.loadAsync(await original.arrayBuffer());
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  zip.file('ppt/slides/slide1.xml',xml.replace('<a:lstStyle/>','<a:lstStyle><a:defPPr><a:defRPr sz="2400"><a:latin typeface="Arial"/></a:defRPr></a:defPPr></a:lstStyle>'));
  zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:defaultTextStyle><a:defPPr><a:defRPr sz="1800"><a:latin typeface="Arial"/></a:defRPr></a:defPPr></p:defaultTextStyle></p:presentation>`);
  const file = new File([await zip.generateAsync({type:'uint8array'})],'default-style.pptx');
  const {doc,warnings} = await fitOutput(file,'[1-1] This intentionally long translated sentence cannot fit even at the minimum size.');
  assert.equal(doc.getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'),'1680');
  assert.equal(warnings.length,1);
});

test("inherited autofit is materialized while local fit settings override it", async () => {
  for (const localFit of ['', '<a:noAutofit/>', '<a:normAutofit fontScale="90000"/>']) {
    const original = await fitFixture(paragraph(explicitRun('原文')),400,120,localFit);
    const zip = await JSZip.loadAsync(await original.arrayBuffer());
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
    const ph = '<p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>';
    zip.file('ppt/slides/slide1.xml',xml.replace('<p:sp><p:spPr>',`<p:sp>${ph}<p:spPr>`));
    zip.file('ppt/slides/_rels/slide1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
    zip.file('ppt/slideLayouts/slideLayout1.xml',`<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${ns}"><p:sp>${ph}<p:txBody><a:bodyPr><a:normAutofit fontScale="80000"/></a:bodyPr></p:txBody></p:sp></p:sldLayout>`);
    const file = new File([await zip.generateAsync({type:'uint8array'})],'inherited-scale.pptx');
    const {doc,warnings} = await fitOutput(file,'[1-1] Short text');
    assert.deepEqual(warnings,[]);
    assert.equal(doc.getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'),localFit.includes('90000')?'2160':localFit?'2400':'1920');
    assert.equal(doc.getElementsByTagNameNS(ns,'normAutofit').length,0);
  }
});

test("title placeholder type inherited from layout selects title master style", async () => {
  const original = await fitFixture(paragraph(run('原文','<a:rPr/>')),400,65);
  const zip = await JSZip.loadAsync(await original.arrayBuffer());
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  const ph = '<p:nvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>';
  const pns = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const relns = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const reltype = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('ppt/slides/slide1.xml',xml.replace('<p:sp><p:spPr>',`<p:sp>${ph}<p:spPr>`));
  zip.file('ppt/slides/_rels/slide1.xml.rels',`<Relationships xmlns="${relns}"><Relationship Id="rId1" Type="${reltype}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout1.xml',`<p:sldLayout xmlns:p="${pns}" xmlns:a="${ns}"><p:sp>${ph.replace('idx="1"','idx="1" type="title"')}</p:sp></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',`<Relationships xmlns="${relns}"><Relationship Id="rId1" Type="${reltype}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml',`<p:sldMaster xmlns:p="${pns}" xmlns:a="${ns}"><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>`);
  const file = new File([await zip.generateAsync({type:'uint8array'})],'title-inherited.pptx');
  const {doc,warnings} = await fitOutput(file,'[1-1] Short');
  assert.deepEqual(warnings,[]);
  const size = Number(doc.getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'));
  assert.ok(size > 3200 && size < 4400);
});

test("prohibited word splitting reports a long word that cannot fit", async () => {
  for (const value of ['0','false']) {
    const body = paragraph(explicitRun('原文')).replace('<a:pPr algn="ctr"/>',`<a:pPr latinLnBrk="${value}"/>`);
    const file = await fitFixture(body,100,200);
    const {doc,warnings} = await fitOutput(file,'[1-1] ABCDEFGHIJ');
    assert.equal(warnings.length,1);
    assert.equal(doc.getElementsByTagNameNS(ns,'rPr')[0].getAttribute('sz'),'1680');
    assert.equal(doc.getElementsByTagNameNS(ns,'pPr')[0].getAttribute('latinLnBrk'),value);
    assert.equal(textOf(doc as unknown as Element),'ABCDEFGHIJ');
  }
});
