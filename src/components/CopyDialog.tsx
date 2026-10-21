import type { Dispatch, SetStateAction } from "react";
import { ClipboardCopy, X } from "lucide-react";
import { getBatchSummary } from "../lib/pptxText";
import type { CopyMode } from "../lib/pptxText";
import { getLimitUnit } from "../lib/copyPrompt";

type CopyDialogProps = {
  activeBatchIndex: number;
  activeBatchSummary: ReturnType<typeof getBatchSummary>;
  batchLimit: number;
  batchesLength: number;
  copyMode: CopyMode;
  copyText: string;
  onBatchIndexChange: Dispatch<SetStateAction<number>>;
  onClose: () => void;
  onCopy: () => void;
  onLimitChange: (nextLimit: number) => void;
  onModeChange: (nextMode: CopyMode) => void;
  onTargetLanguageChange: (nextLanguage: string) => void;
  targetLanguage: string;
};

export function CopyDialog({
  activeBatchIndex,
  activeBatchSummary,
  batchLimit,
  batchesLength,
  copyMode,
  copyText,
  onBatchIndexChange,
  onClose,
  onCopy,
  onLimitChange,
  onModeChange,
  onTargetLanguageChange,
  targetLanguage,
}: CopyDialogProps) {
  return (
    <div className="dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="copy-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="copy-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <h2 id="copy-dialog-title">复制给外部 AI</h2>
            <p>选择复制范围，内容会自动带上翻译提示词和代码块。</p>
          </div>
          <button className="icon-action" onClick={onClose} aria-label="关闭复制设置">
            <X size={18} />
          </button>
        </header>

        <label className="language-field">
          目标语言
          <input value={targetLanguage} onChange={(event) => onTargetLanguageChange(event.target.value)} />
        </label>

        <div className="copy-mode dialog-mode" aria-label="复制方式">
          <button className={copyMode === "all" ? "selected" : ""} onClick={() => onModeChange("all")}>
            全部复制
          </button>
          <button className={copyMode === "page" ? "selected" : ""} onClick={() => onModeChange("page")}>
            按页数
          </button>
          <button className={copyMode === "line" ? "selected" : ""} onClick={() => onModeChange("line")}>
            按行数
          </button>
        </div>

        {copyMode !== "all" && (
          <div className="batch-controls dialog-batch-controls">
            <label>
              每批
              <input
                type="number"
                min="1"
                step="1"
                value={batchLimit}
                onChange={(event) => onLimitChange(Number(event.target.value) || 1)}
              />
              {getLimitUnit(copyMode)}
            </label>
            <span className="batch-detail">
              本批 {activeBatchSummary.pages} 页 / {activeBatchSummary.lines} 行 / {activeBatchSummary.chars} 字符
            </span>
            <div className="batch-nav" aria-label="批次切换">
              <button
                className="ghost-action"
                onClick={() => onBatchIndexChange((current) => Math.max(0, current - 1))}
                disabled={activeBatchIndex === 0 || batchesLength === 0}
              >
                上一批
              </button>
              <span>{batchesLength === 0 ? "0 / 0" : `${activeBatchIndex + 1} / ${batchesLength}`}</span>
              <button
                className="ghost-action"
                onClick={() => onBatchIndexChange((current) => Math.min(batchesLength - 1, current + 1))}
                disabled={activeBatchIndex >= batchesLength - 1}
              >
                下一批
              </button>
            </div>
          </div>
        )}

        {copyMode === "all" && (
          <div className="batch-summary-line">
            全部内容：{activeBatchSummary.pages} 页 / {activeBatchSummary.lines} 行 / {activeBatchSummary.chars} 字符
          </div>
        )}

        <textarea className="dialog-preview" value={copyText} readOnly aria-label="将复制给外部 AI 的内容" />

        <footer className="dialog-actions">
          <button className="secondary-action" onClick={onClose}>
            取消
          </button>
          <button className="primary-action" onClick={onCopy} disabled={!copyText}>
            <ClipboardCopy size={17} />
            复制到剪贴板
          </button>
        </footer>
      </section>
    </div>
  );
}
