import type { TextRecord } from "../types";
import { formatRecordsForAi, getBatchSummary } from "./pptxText";
import type { CopyMode } from "./pptxText";

export function getDefaultLimit(mode: CopyMode) {
  if (mode === "page") return 1;
  if (mode === "line") return 80;
  return 1;
}

export function getLimitUnit(mode: CopyMode) {
  if (mode === "page") return "页";
  if (mode === "line") return "行";
  return "";
}

export function getBatchLabel(
  mode: CopyMode,
  batchIndex: number,
  batchCount: number,
  summary: ReturnType<typeof getBatchSummary>,
) {
  if (mode === "all") {
    return `全部内容（${summary.pages} 页 / ${summary.lines} 行 / ${summary.chars} 字符）`;
  }
  return `第 ${batchIndex + 1} / ${batchCount} 批（${summary.pages} 页 / ${summary.lines} 行 / ${summary.chars} 字符）`;
}

export function buildPromptedCopyText({
  records,
  fileName,
  targetLanguage,
  batchLabel,
}: {
  records: TextRecord[];
  fileName: string;
  targetLanguage: string;
  batchLabel: string;
}) {
  const language = targetLanguage.trim() || "目标语言";
  return [
    `请将以下 PPT 页面内容翻译成${language}。`,
    "要求：",
    "1. 保留方括号里的 ID，例如 [1-1]，不要修改 ID。",
    "2. 不要合并、删除或重新排序行。",
    "每个 ID 的译文必须放在同一行，不要在段内另起一行，也不要在前后追加完成提示或其他说明。",
    "3. 页标题可以保留；只输出翻译后的同格式文本，不要解释。",
    "4. 每个 ID 是一个完整段落，请结合全段上下文翻译，不要将格式片段分别翻译。",
    "5. <s1>…</s1> 等是格式标记。只翻译标记内的正文，不要翻译或修改标记名称、数字、尖括号和斜杠。",
    "6. 每对格式标记必须保留且只出现一次、正确配对，不得嵌套；可随对应词语调整在本段内的位置。不要新增标记。",
    "7. <br1/> 等是原文的换行标记，请原样保留一次。动态字段的标记及内容按下方要求保留原值。",
    "8. 保留正常词间空格，包括格式标记两侧的空格，不要把不同英文单词连在一起。不要将标记当作 HTML 渲染，不要给标记加反引号。",
    "9. &#60;、&#62;、&amp; 等是原文字符的转义，请原样保留，不要把它们改写成格式标记。",
    "示例：[1-1] 请检查 <s1>红色按钮</s1> → [1-1] Check the <s1>red button</s1>, please.",
    ...records.flatMap((record) => record.spans.filter((span) => span.kind === "field").map((span) =>
      `动态字段：[${record.id}] 的 <s${span.tagId}> 内必须原样保留 ${JSON.stringify(span.text)}。`)),
    "",
    `PPT 文件：${fileName}`,
    `复制范围：${batchLabel}`,
    "",
    "待翻译内容：",
    "```text",
    formatRecordsForAi(records),
    "```",
  ].join("\n");
}
