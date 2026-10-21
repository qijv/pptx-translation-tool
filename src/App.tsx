import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import {
  CheckCircle2,
  ClipboardCopy,
  Download,
  ExternalLink,
  FileText,
  Languages,
  Moon,
  Sun,
  Upload,
  Wand2,
} from "lucide-react";
import type { ExportLayout, ParseResult } from "./types";
import {
  buildTranslatedPptx,
  buildCopyBatches,
  extractTextRecords,
  formatRecordsForAi,
  getBatchSummary,
  parseTranslations,
  summarizeTranslationCoverage,
} from "./lib/pptxText";
import type { CopyMode } from "./lib/pptxText";
import { buildPromptedCopyText, getBatchLabel, getDefaultLimit } from "./lib/copyPrompt";
import { CopyDialog } from "./components/CopyDialog";
import { ExportOptions } from "./components/ExportOptions";

type Status = {
  tone: "idle" | "success" | "warning" | "error";
  message: string;
};

type Theme = "light" | "dark";
const THEME_STORAGE_KEY = "ppt-theme";
const THEME_PREF_VERSION_KEY = "ppt-theme-version";
const THEME_PREF_VERSION = "2";
// VITE_* values are public build-time configuration, never credentials.
const ANALYTICS_SITE = import.meta.env.VITE_ANALYTICS_SITE?.trim() ?? "";
const ANALYTICS_KEY = import.meta.env.VITE_ANALYTICS_KEY?.trim() ?? "";
const ANALYTICS_COLLECT_URL = import.meta.env.VITE_ANALYTICS_COLLECT_URL?.trim() ?? "";
const ANALYTICS_SUMMARY_URL = import.meta.env.VITE_ANALYTICS_SUMMARY_URL?.trim() ?? "";

type SiteStats = {
  online: number;
  today: number;
  total: number;
  visitors: number;
};

function App() {
  const [theme, setTheme] = useState<Theme>(() => {
    if (typeof window === "undefined") return "dark";
    return getStoredTheme();
  });
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [translationText, setTranslationText] = useState("");
  const [exportLayout, setExportLayout] = useState<ExportLayout>("preserve");
  const [layoutWarnings, setLayoutWarnings] = useState<Array<{ id: string; message: string }>>([]);
  const [exportStatus, setExportStatus] = useState<Status | null>(null);
  const [copyMode, setCopyMode] = useState<CopyMode>("all");
  const [batchLimit, setBatchLimit] = useState(1);
  const [batchIndex, setBatchIndex] = useState(0);
  const [targetLanguage, setTargetLanguage] = useState("英文");
  const [isCopyDialogOpen, setIsCopyDialogOpen] = useState(false);
  const [status, setStatus] = useState<Status>({ tone: "idle", message: "选择一个 .pptx 文件开始。" });
  const [isBusy, setIsBusy] = useState(false);
  const [siteStats, setSiteStats] = useState<SiteStats | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const uploadGenerationRef = useRef(0);
  const exportInProgressRef = useRef(false);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    setStoredTheme(theme);
  }, [theme]);

  useEffect(() => {
    trackPageView();
    loadSiteStats().then(setSiteStats).catch(() => {});
  }, []);

  const batches = useMemo(
    () => (parseResult ? buildCopyBatches(parseResult.records, copyMode, batchLimit) : []),
    [batchLimit, copyMode, parseResult],
  );
  const activeBatchIndex = batches.length > 0 ? Math.min(batchIndex, batches.length - 1) : 0;
  const activeBatch = batches[activeBatchIndex] ?? [];
  const extractedText = useMemo(() => (parseResult ? formatRecordsForAi(parseResult.records) : ""), [parseResult]);
  const activeBatchSummary = useMemo(() => getBatchSummary(activeBatch), [activeBatch]);
  const promptedCopyText = useMemo(
    () =>
      parseResult
        ? buildPromptedCopyText({
            records: activeBatch,
            fileName: parseResult.fileName,
            targetLanguage,
            batchLabel: getBatchLabel(copyMode, activeBatchIndex, batches.length, activeBatchSummary),
          })
        : "",
    [activeBatch, activeBatchIndex, activeBatchSummary, batches.length, copyMode, parseResult, targetLanguage],
  );
  const translationResult = useMemo(
    () => parseTranslations(translationText, parseResult?.records ?? []),
    [parseResult?.records, translationText],
  );
  const coverage = useMemo(
    () =>
      parseResult
        ? summarizeTranslationCoverage(parseResult.records, translationResult.translations)
        : { total: 0, translated: 0, missing: 0 },
    [parseResult, translationResult.translations],
  );

  async function handleFileChange(file: File | null) {
    if (!file || exportInProgressRef.current) return;
    const generation = ++uploadGenerationRef.current;
    setIsBusy(true);
    setExportStatus(null);
    setLayoutWarnings([]);
    setStatus({ tone: "idle", message: "正在解析 PPTX 文本..." });

    try {
      const result = await extractTextRecords(file);
      if (generation !== uploadGenerationRef.current) return;
      setSourceFile(file);
      setParseResult(result);
      setTranslationText("");
      setExportLayout("preserve");
      setLayoutWarnings([]);
      setExportStatus(null);
      setCopyMode("all");
      setBatchLimit(1);
      setBatchIndex(0);
      setIsCopyDialogOpen(false);
      setStatus({
        tone: result.records.length > 0 ? "success" : "warning",
        message:
          result.records.length > 0
            ? `解析完成：${result.slideCount} 页幻灯片，${result.records.length} 段文本。`
            : "没有在幻灯片中找到可替换文本。",
      });
    } catch (error) {
      if (generation !== uploadGenerationRef.current) return;
      setSourceFile(null);
      setParseResult(null);
      setStatus({ tone: "error", message: error instanceof Error ? error.message : "解析失败，请换一个文件重试。" });
    } finally {
      if (generation === uploadGenerationRef.current) setIsBusy(false);
    }
  }

  async function copyAiText() {
    if (!promptedCopyText) return;
    try {
      await navigator.clipboard.writeText(promptedCopyText);
      setIsCopyDialogOpen(false);
      setStatus({
        tone: "success",
        message: `已复制${getBatchLabel(copyMode, activeBatchIndex, batches.length, activeBatchSummary)}，内容已包含提示词。`,
      });
    } catch {
      setStatus({ tone: "error", message: "复制失败，请在弹窗预览框中手动全选复制。" });
    }
  }

  function changeCopyMode(nextMode: CopyMode) {
    setCopyMode(nextMode);
    setBatchLimit(getDefaultLimit(nextMode));
    setBatchIndex(0);
  }

  async function exportTranslatedPptx() {
    if (!sourceFile || !parseResult || exportInProgressRef.current) return;

    if (translationResult.formatIssues.length > 0) {
      setExportStatus({ tone: "warning", message: "译文有格式标记错误，请先修正下方提示后再导出。" });
      return;
    }

    if (coverage.translated === 0) {
      setExportStatus({ tone: "warning", message: "还没有识别到带 ID 的译文，请先粘贴翻译结果。" });
      return;
    }

    exportInProgressRef.current = true;
    const generation = uploadGenerationRef.current;
    setIsBusy(true);
    setExportStatus({ tone: "idle", message: "正在生成 PPTX，请稍候，完成前不能切换文件。" });
    setLayoutWarnings([]);

    try {
      const warnings: Array<{ id: string; message: string }> = [];
      const blob = await buildTranslatedPptx(sourceFile, parseResult.records, translationResult.translations, exportLayout, {
        onLayoutWarning: (id, message) => warnings.push({ id, message }),
      });
      if (generation !== uploadGenerationRef.current) return;
      setLayoutWarnings(warnings);
      downloadBlob(blob, makeOutputName(sourceFile.name));
      setExportStatus({
        tone: coverage.missing === 0 && warnings.length === 0 ? "success" : "warning",
        message:
          warnings.length > 0
            ? `已导出 PPTX，有 ${warnings.length} 个文本框需要检查排版${coverage.missing ? `，${coverage.missing} 段未翻译并保留原文` : ""}。`
            : coverage.missing === 0
            ? "已导出完整替换版 PPTX。"
            : `已导出 PPTX，还有 ${coverage.missing} 段没有译文，保留了原文。`,
      });
    } catch (error) {
      if (generation === uploadGenerationRef.current) setExportStatus({ tone: "error", message: error instanceof Error ? error.message : "导出失败，请重试。" });
    } finally {
      exportInProgressRef.current = false;
      if (generation === uploadGenerationRef.current) setIsBusy(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="hero">
        <div>
          <div className="brand">
            <Languages size={22} />
            <span>PPT 翻译替换工具</span>
          </div>
          <h1>提取、外部翻译、再批量写回 PPTX</h1>
          <p>
            文件只在本机浏览器中解析。复制带编号的文本给外部 AI，粘贴译文后导出替换版，原文件不会被修改。
          </p>
        </div>
        <div className="hero-actions">
          <a className="contact-action" href="https://naodong.store/" target="_blank" rel="noreferrer">
            <ExternalLink size={18} />
            联系我
          </a>
          <button
            className="theme-toggle"
            onClick={() => setTheme((current) => (current === "light" ? "dark" : "light"))}
            aria-label={theme === "light" ? "切换到暗色模式" : "切换到亮色模式"}
          >
            {theme === "light" ? <Moon size={18} /> : <Sun size={18} />}
            {theme === "light" ? "暗色" : "亮色"}
          </button>
          <button className="primary-action" onClick={() => fileInputRef.current?.click()} disabled={isBusy}>
            <Upload size={18} />
            选择 PPTX
          </button>
        </div>
        <input
          ref={fileInputRef}
          className="sr-only"
          type="file"
          disabled={isBusy}
          accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
          onChange={(event) => handleFileChange(event.target.files?.[0] ?? null)}
        />
      </section>

      <section className="guide">
        <div>
          <FileText size={20} />
          <h2>给外部 AI 的建议提示词</h2>
        </div>
        <p>
          按完整段落翻译，保留页内 ID 和格式标记，只翻译正文。复制文本会附带详细要求，可以分批翻译后追加粘贴回来。
        </p>
      </section>

      <section className="workflow">
        <Step number="1" title="上传 PPTX" active={!parseResult}>
          <DropZone isBusy={isBusy} fileName={sourceFile?.name} onFile={handleFileChange} />
          <StatusMessage status={status} />
        </Step>

        <Step number="2" title="复制给外部 AI" active={Boolean(parseResult)}>
          <div className="panel-toolbar">
            <Stats
              items={[
                ["幻灯片", parseResult?.slideCount ?? 0],
                ["文本段", parseResult?.records.length ?? 0],
                ["字符", parseResult?.charCount ?? 0],
              ]}
            />
            <button className="secondary-action" onClick={() => setIsCopyDialogOpen(true)} disabled={!parseResult || isBusy}>
              <ClipboardCopy size={17} />
              复制文本
            </button>
          </div>
          <textarea
            className="text-area"
            value={extractedText}
            readOnly
            placeholder="上传 PPTX 后会在这里显示完整解析结果。格式示例：第 1 页&#10;[1-1] 原文"
            aria-label="提取出的 PPT 文本"
          />
          <p className="hint">每个 ID 对应一个段落；格式标记用于保留字体、颜色等样式，请连同正文一起复制并保留。</p>
        </Step>

        <Step number="3" title="粘贴译文并导出" active={Boolean(parseResult)}>
          <div className="panel-toolbar">
            <Stats
              items={[
                ["已匹配", coverage.translated],
                ["未翻译", coverage.missing],
              ]}
            />
            <button className="primary-action compact" onClick={exportTranslatedPptx} disabled={!parseResult || isBusy || translationResult.formatIssues.length > 0}>
              <Download size={17} />
              导出 PPTX
            </button>
          </div>
          {exportStatus && <div className="export-feedback">
            <StatusMessage status={exportStatus} />
            {layoutWarnings.length > 0 && <details className="export-warning-details" open={layoutWarnings.length <= 3}>
              <summary>查看 {layoutWarnings.length} 处排版检查提示</summary>
              <ul>{layoutWarnings.map((item) => <li key={item.id}>第 {item.id.split("-")[0]} 页，文本 [{item.id}]：{item.message}</li>)}</ul>
            </details>}
          </div>}
          <ExportOptions value={exportLayout} onChange={(value) => { setExportLayout(value); setLayoutWarnings([]); setExportStatus(null); }} disabled={isBusy} />
          <textarea
            className="text-area"
            value={translationText}
            disabled={isBusy}
            onChange={(event) => { setTranslationText(event.target.value); setLayoutWarnings([]); setExportStatus(null); }}
            placeholder="第 1 页&#10;[1-1] Translated text&#10;[1-2] Translated text"
            aria-label="外部 AI 翻译后的文本"
          />
          {translationResult.invalidLines.length > 0 && (
            <p className="hint warning-text">有 {translationResult.invalidLines.length} 行无法匹配，请按下方提示修正后导出。</p>
          )}
          {translationResult.formatIssues.length > 0 && (
            <div className="hint warning-text" role="alert">
              <p>请先修正以下编号、正文或格式问题，再导出：</p>
              <p>部分翻译工具会删除格式标记；可复制附带提示词的原文，重新翻译对应段落。</p>
              <ul>{translationResult.formatIssues.slice(0, 8).map((issue, index) => (
                <li key={`${issue.id}-${index}`}>[{issue.id}] {issue.message}</li>
              ))}</ul>
              {translationResult.formatIssues.length > 8 && <p>另有 {translationResult.formatIssues.length - 8} 个问题。</p>}
            </div>
          )}
          {translationResult.repairedIds.length > 0 && (
            <p className="hint">已自动规范 {translationResult.repairedIds.length} 段译文中的编号或格式标记。</p>
          )}
        </Step>
      </section>

      {isCopyDialogOpen && parseResult && (
        <CopyDialog
          activeBatchIndex={activeBatchIndex}
          activeBatchSummary={activeBatchSummary}
          batchLimit={batchLimit}
          batchesLength={batches.length}
          copyMode={copyMode}
          copyText={promptedCopyText}
          onBatchIndexChange={setBatchIndex}
          onClose={() => setIsCopyDialogOpen(false)}
          onCopy={copyAiText}
          onLimitChange={(nextLimit) => {
            setBatchLimit(nextLimit);
            setBatchIndex(0);
          }}
          onModeChange={changeCopyMode}
          onTargetLanguageChange={setTargetLanguage}
          targetLanguage={targetLanguage}
        />
      )}
      <footer className="site-footer">
        <span>PPT 翻译工具</span>
        <span>在线 <strong>{siteStats?.online ?? "-"}</strong></span>
        <span>今日 <strong>{siteStats?.today ?? "-"}</strong></span>
        <span>总访问 <strong>{siteStats?.total ?? "-"}</strong></span>
        <span>访客 <strong>{siteStats?.visitors ?? "-"}</strong></span>
      </footer>
    </main>
  );
}

function Step({
  number,
  title,
  active,
  children,
}: {
  number: string;
  title: string;
  active: boolean;
  children: ReactNode;
}) {
  return (
    <article className={`step-card ${active ? "active" : ""}`}>
      <header className="step-header">
        <span>{number}</span>
        <h2>{title}</h2>
      </header>
      {children}
    </article>
  );
}

function DropZone({ isBusy, fileName, onFile }: { isBusy: boolean; fileName?: string; onFile: (file: File) => void }) {
  return (
    <label
      className="drop-zone"
      aria-disabled={isBusy}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        if (isBusy) return;
        const file = event.dataTransfer.files[0];
        if (file) onFile(file);
      }}
    >
      <input
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        disabled={isBusy}
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
        }}
      />
      <Wand2 size={24} />
      <strong>{fileName ?? "拖入或选择 PPTX"}</strong>
      <span>{isBusy ? "正在处理文件，请稍候再上传。" : "支持现代 PowerPoint 文件，不读取旧版 .ppt。"}</span>
    </label>
  );
}

function StatusMessage({ status }: { status: Status }) {
  return (
    <div className={`status ${status.tone}`} role="status" aria-live="polite">
      <CheckCircle2 size={17} />
      <span>{status.message}</span>
    </div>
  );
}

function Stats({ items }: { items: Array<[string, number]> }) {
  return (
    <div className="stats">
      {items.map(([label, value]) => (
        <span key={label}>
          <strong>{value}</strong>
          {label}
        </span>
      ))}
    </div>
  );
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function makeOutputName(fileName: string) {
  return fileName.replace(/\.pptx$/i, "-translated.pptx");
}

function getStoredTheme(): Theme {
  try {
    const themeVersion = window.localStorage.getItem(THEME_PREF_VERSION_KEY);
    if (themeVersion !== THEME_PREF_VERSION) return "dark";

    const storedTheme = window.localStorage.getItem(THEME_STORAGE_KEY);
    if (storedTheme === "dark" || storedTheme === "light") return storedTheme;
    return "dark";
  } catch {
    return "dark";
  }
}

function setStoredTheme(theme: Theme) {
  try {
    window.localStorage.setItem(THEME_STORAGE_KEY, theme);
    window.localStorage.setItem(THEME_PREF_VERSION_KEY, THEME_PREF_VERSION);
  } catch {
    // Theme persistence is a convenience; ignore storage failures.
  }
}

function getVisitorId() {
  const key = "ppt_analytics_visitor_id";
  try {
    const current = window.localStorage.getItem(key);
    if (current) return current;
    const next = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    window.localStorage.setItem(key, next);
    return next;
  } catch {
    return "";
  }
}

async function trackPageView() {
  if (!ANALYTICS_SITE || !ANALYTICS_KEY || !ANALYTICS_COLLECT_URL) return;
  try {
    await fetch(ANALYTICS_COLLECT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        site: ANALYTICS_SITE,
        key: ANALYTICS_KEY,
        host: window.location.host,
        path: window.location.pathname,
        full_path: window.location.pathname + window.location.search,
        referer: document.referrer,
        visitor_id: getVisitorId(),
        user_agent: navigator.userAgent,
      }),
      keepalive: true,
    });
  } catch {
    // Analytics is best-effort only.
  }
}

async function loadSiteStats(): Promise<SiteStats | null> {
  if (!ANALYTICS_SITE || !ANALYTICS_SUMMARY_URL) return null;
  const url = new URL(ANALYTICS_SUMMARY_URL);
  url.searchParams.set("site", ANALYTICS_SITE);
  const response = await fetch(url.toString());
  if (!response.ok) return null;
  const payload = await response.json();
  return payload?.ok ? payload.stats : null;
}

export default App;
