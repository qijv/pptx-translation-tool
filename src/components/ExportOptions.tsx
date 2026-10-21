import { useEffect, useRef, useState } from "react";
import { X } from "lucide-react";
import type { ExportLayout } from "../types";

const options: Array<{ value: ExportLayout; title: string; description: string }> = [
  { value: "preserve", title: "保留原格式（默认）", description: "不更改原字号、文本框和换行设置。译文变长时，可能超出文本框。" },
  { value: "wrap-shrink", title: "智能排版", description: "每个文本框单独估算：先自动换行，能放下就保留原字号，放不下才逐步调整字号。最低为原字号的 70%，且原本大于等于 12 磅的文字不缩到 12 磅以下；原本小于 12 磅的文字不再缩小。达到下限仍可能溢出，会在导出后提示。" },
];

export function ExportOptions({ value, onChange, disabled }: {
  value: ExportLayout; onChange: (value: ExportLayout) => void; disabled: boolean;
}) {
  const [detail, setDetail] = useState<ExportLayout | null>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);
  const detailOption = options.find((option) => option.value === detail);

  useEffect(() => {
    if (!detail) return;
    const dialog = dialogRef.current!;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    dialog.showModal();
    return () => {
      dialog.close();
      document.body.style.overflow = previousOverflow;
    };
  }, [detail]);

  return (
    <>
      <fieldset className="export-options" disabled={disabled}>
        <legend>译文内容排版</legend>
        <div className="export-option-grid">
          {options.map((option) => (
            <div key={option.value} className={`export-option ${value === option.value ? "selected" : ""}`}
              onClick={() => { if (!disabled) onChange(option.value); }}>
              <div className="layout-preview-image">
                <LayoutExample mode={option.value} />
                <span className="layout-hover-preview" aria-hidden="true"><LayoutExample mode={option.value} /></span>
              </div>
              <label className="export-option-title">
                <input type="radio" name="export-layout" value={option.value} checked={value === option.value}
                  aria-label={option.title} onChange={() => onChange(option.value)} />
                {option.value === "preserve" ? "保留原格式" : "智能排版"}
              </label>
              <div className="layout-card-footer">
                <span className="layout-default">{option.value === "preserve" ? "默认" : ""}</span>
                <button type="button" className="layout-detail-button" aria-label={`查看${option.title}详情`}
                  onClick={(event) => { event.stopPropagation(); setDetail(option.value); }}>详情</button>
              </div>
            </div>
          ))}
        </div>
        <p className="hint layout-help">点击卡片切换排版 · 悬浮图片放大 · 点击“详情”查看说明。重新打开页面或上传新文件恢复默认。</p>
      </fieldset>
      <dialog ref={dialogRef} className="layout-detail-dialog" aria-labelledby="layout-detail-title"
        onClose={() => setDetail(null)} onCancel={() => setDetail(null)}
        onClick={(event) => {
          if (event.target !== event.currentTarget) return;
          const rect = event.currentTarget.getBoundingClientRect();
          if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) setDetail(null);
        }}>
        {detailOption && <>
          <header className="dialog-header">
            <h2 id="layout-detail-title">{detailOption.title}</h2>
            <button type="button" className="icon-action" aria-label="关闭排版详情" onClick={() => setDetail(null)}><X size={18} /></button>
          </header>
          <LayoutExample mode={detailOption.value} />
          <p>{detailOption.description}</p>
          <p className="hint">图片是效果示意，并非当前 PPT 预览。设置作用于含译文的整个文本框（包括其中保留的原文），不扩大文本框。智能排版使用本机字体估算，特殊排版保留原设置并提示。实际效果受字体和 PowerPoint / WPS 影响，请检查导出的文件。</p>
          <footer className="dialog-actions">
            <button type="button" className="secondary-action" onClick={() => setDetail(null)}>关闭</button>
            <button type="button" className="primary-action" disabled={disabled} onClick={() => { onChange(detailOption.value); setDetail(null); }}>
              使用此排版
            </button>
          </footer>
        </>}
      </dialog>
    </>
  );
}

function LayoutExample({ mode }: { mode: ExportLayout }) {
  const lines = mode === "preserve" ? ["Supplementary machine inspection instructions"]
    : ["Supplementary machine", "inspection instructions"];
  return (
    <svg className="layout-example" viewBox="0 0 300 116" role="img"
      aria-label={mode === "preserve" ? "长译文超出原文本框右侧" : "译文缩小后两行放进原文本框"}>
      <rect width="300" height="116" rx="8" fill="#f3f6f4" />
      <text x="14" y="21" fill="#647269" fontSize="11">效果示意 · 同一段译文</text>
      <rect x="14" y="34" width="212" height="46" rx="3" fill="white" stroke="#438b69" strokeDasharray="4 3" />
      {lines.map((line, index) => (
        <text key={line} x="22" y={mode === "wrap-shrink" ? 53 + index * 18 : 60 + index * 25}
          fontSize={mode === "wrap-shrink" ? 14 : 16.5} fill="#172b20" fontFamily="Arial, sans-serif">{line}</text>
      ))}
      <text x="14" y="105" fill={mode === "wrap-shrink" ? "#24704d" : "#9b6500"} fontSize="11">
        {mode === "preserve" ? "可能向右溢出 →" : "按文本框调整 · 尽量放入框内"}
      </text>
    </svg>
  );
}
