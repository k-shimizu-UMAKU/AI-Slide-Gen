import { useState, useEffect, useRef } from "react";

// ─── PptxGenJS CDN 読み込み ──────────────────────────────────────
const loadPptxGen = () =>
  new Promise((resolve, reject) => {
    if (window.PptxGenJS) { resolve(window.PptxGenJS); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/PptxGenJS/3.12.0/pptxgen.bundle.js";
    s.onload = () => resolve(window.PptxGenJS);
    s.onerror = () => reject(new Error("PptxGenJS の読み込みに失敗しました"));
    document.head.appendChild(s);
  });

// ─── 用紙サイズ定義（インチ）────────────────────────────────────
const PAPER_SIZES = {
  A4:     { label: "A4",     w: 11.69, h: 8.27,  mm: "297×210mm" },
  A3:     { label: "A3",     w: 16.54, h: 11.69, mm: "420×297mm" },
  B5:     { label: "B5",     w: 10.12, h: 7.17,  mm: "257×182mm" },
  Letter: { label: "Letter", w: 11.0,  h: 8.5,   mm: '11×8.5"' },
};

const getLayout = (size, orientation) => {
  const p = PAPER_SIZES[size];
  return orientation === "landscape"
    ? { w: p.w, h: p.h, ratio: p.w / p.h }
    : { w: p.h, h: p.w, ratio: p.h / p.w };
};

// ─── カラーパレット ──────────────────────────────────────────────
const PAL = {
  bg:      "#F7F5F0",
  surface: "#FFFFFF",
  border:  "#E2DDD5",
  navy:    "#1B2B4B",
  accent:  "#C8602A",
  accentL: "#F0E4D8",
  muted:   "#8A8278",
  success: "#2A7A55",
  error:   "#B03030",
};

// ─── スライドテーマ（PPTX・プレビュー共通）──────────────────────
const THEMES = [
  { bg: "1B2B4B", title: "FFFFFF", body: "C8D4E8", accent: "C8602A" },
  { bg: "FFFFFF", title: "1B2B4B", body: "3A4A6B", accent: "C8602A" },
  { bg: "F7F5F0", title: "1B2B4B", body: "3A4A6B", accent: "C8602A" },
  { bg: "C8602A", title: "FFFFFF", body: "F5E0D0", accent: "1B2B4B" },
];

// ─── スライドプレビューカード ────────────────────────────────────
const SlideCard = ({ slide, index, total, ratio }) => {
  const isTitle = slide.type === "title";
  const t = THEMES[isTitle ? 0 : (index % THEMES.length)];
  return (
    <div style={{
      width: "100%", aspectRatio: `${ratio}`,
      background: `#${t.bg}`, borderRadius: "6px",
      border: `1px solid ${PAL.border}`, overflow: "hidden",
      position: "relative", display: "flex", flexDirection: "column",
      padding: isTitle ? "8%" : "5% 7%", boxSizing: "border-box",
      boxShadow: "0 4px 20px rgba(27,43,75,0.10)",
    }}>
      <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: "6px", background: `#${t.accent}` }} />
      <div style={{ position: "absolute", bottom: "3%", right: "3%", fontSize: "clamp(8px,1.2vw,11px)", color: `#${t.title}`, opacity: 0.3, fontFamily: "monospace" }}>
        {index + 1}/{total}
      </div>
      {isTitle ? (
        <div style={{ display: "flex", flexDirection: "column", justifyContent: "center", height: "100%", paddingLeft: "2%" }}>
          {slide.category && <div style={{ fontSize: "clamp(7px,1.1vw,11px)", color: `#${t.accent}`, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: "4%", fontWeight: 700 }}>{slide.category}</div>}
          <div style={{ fontSize: "clamp(14px,3.2vw,32px)", fontWeight: 800, color: `#${t.title}`, lineHeight: 1.15, marginBottom: "4%", fontFamily: "Georgia,serif" }}>{slide.title}</div>
          {slide.subtitle && <div style={{ fontSize: "clamp(9px,1.5vw,15px)", color: `#${t.body}`, opacity: 0.8, lineHeight: 1.5 }}>{slide.subtitle}</div>}
          {slide.url && <div style={{ marginTop: "5%", fontSize: "clamp(7px,1vw,11px)", color: `#${t.accent}`, fontFamily: "monospace", opacity: 0.8 }}>{slide.url}</div>}
        </div>
      ) : (
        <>
          {slide.section && <div style={{ fontSize: "clamp(6px,0.9vw,10px)", color: `#${t.accent}`, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "2%", fontWeight: 700, paddingLeft: "1%" }}>{slide.section}</div>}
          <div style={{ fontSize: "clamp(11px,2vw,20px)", fontWeight: 700, color: `#${t.title}`, marginBottom: "4%", fontFamily: "Georgia,serif", lineHeight: 1.25, paddingLeft: "1%" }}>{slide.title}</div>
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: "2%", paddingLeft: "1%" }}>
            {(slide.bullets || []).map((b, i) => (
              <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: "2.5%" }}>
                <span style={{ color: `#${t.accent}`, fontSize: "clamp(7px,1.1vw,11px)", flexShrink: 0, marginTop: "0.15em", fontWeight: 700 }}>▸</span>
                <span style={{ color: `#${t.body}`, fontSize: "clamp(8px,1.2vw,13px)", lineHeight: 1.45 }}>{b}</span>
              </div>
            ))}
          </div>
          {slide.highlight && (
            <div style={{ marginTop: "3%", padding: "2.5% 3%", background: `#${t.accent}18`, borderLeft: `3px solid #${t.accent}`, borderRadius: "3px", fontSize: "clamp(7px,1vw,11px)", color: `#${t.title}`, fontStyle: "italic" }}>{slide.highlight}</div>
          )}
        </>
      )}
    </div>
  );
};

// ─── 用紙サイズ設定パネル（共通）────────────────────────────────
const PaperSettings = ({ paperSize, setPaperSize, orientation, setOrientation }) => (
  <div style={{ display: "flex", flexWrap: "wrap", gap: "2rem", alignItems: "flex-start" }}>
    <div>
      <div style={{ fontSize: "11px", color: PAL.muted, marginBottom: "8px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em" }}>用紙サイズ</div>
      <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
        {Object.entries(PAPER_SIZES).map(([key, val]) => (
          <button key={key} onClick={() => setPaperSize(key)} style={{
            padding: "6px 14px", borderRadius: "20px", cursor: "pointer",
            fontSize: "13px", fontWeight: 600, fontFamily: "inherit",
            background: paperSize === key ? PAL.navy : PAL.bg,
            color: paperSize === key ? "#fff" : PAL.muted,
            border: `2px solid ${paperSize === key ? PAL.navy : PAL.border}`,
          }}>
            {val.label} <span style={{ fontSize: "10px", opacity: 0.7 }}>{val.mm}</span>
          </button>
        ))}
      </div>
    </div>
    <div>
      <div style={{ fontSize: "11px", color: PAL.muted, marginBottom: "8px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em" }}>向き</div>
      <div style={{ display: "flex", gap: "8px" }}>
        {[["landscape", "横", "26px", "18px"], ["portrait", "縦", "18px", "26px"]].map(([val, label, w, h]) => (
          <button key={val} onClick={() => setOrientation(val)} style={{
            display: "flex", flexDirection: "column", alignItems: "center", gap: "6px",
            padding: "10px 14px", borderRadius: "10px", cursor: "pointer",
            background: orientation === val ? PAL.navy : PAL.surface,
            border: `2px solid ${orientation === val ? PAL.navy : PAL.border}`,
          }}>
            <div style={{ width: w, height: h, background: orientation === val ? PAL.accent : PAL.bg, border: `2px solid ${orientation === val ? PAL.accent : PAL.muted}`, borderRadius: "2px" }} />
            <span style={{ fontSize: "11px", fontWeight: 600, color: orientation === val ? "#fff" : PAL.muted, fontFamily: "monospace" }}>{label}</span>
          </button>
        ))}
      </div>
    </div>
    <div style={{ marginLeft: "auto", alignSelf: "center" }}>
      <div style={{ padding: "6px 14px", background: PAL.accentL, borderRadius: "8px", fontSize: "13px", color: PAL.accent, fontWeight: 700, fontFamily: "monospace", whiteSpace: "nowrap" }}>
        {paperSize} / {orientation === "landscape" ? "横" : "縦"}
        <span style={{ opacity: 0.6, fontSize: "11px", marginLeft: "6px" }}>
          {orientation === "landscape"
            ? `${PAPER_SIZES[paperSize].w}" × ${PAPER_SIZES[paperSize].h}"`
            : `${PAPER_SIZES[paperSize].h}" × ${PAPER_SIZES[paperSize].w}"`}
        </span>
      </div>
    </div>
  </div>
);

// ─── スライドプレビューパネル（共通）────────────────────────────
const SlidePreviewPanel = ({ slides, current, setCurrent, layout, orientation, label, badge, onDownload, onReset, resetLabel, downloadName }) => (
  <div style={{ animation: "slideUp 0.4s ease" }}>
    <SectionCard label={label}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", flexWrap: "wrap", gap: "8px" }}>
        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
          {badge && <span style={{ padding: "3px 10px", background: PAL.accentL, borderRadius: "12px", fontSize: "12px", color: PAL.accent, fontWeight: 700 }}>{badge}</span>}
          <span style={{ padding: "3px 10px", background: PAL.bg, border: `1px solid ${PAL.border}`, borderRadius: "12px", fontSize: "12px", color: PAL.muted }}>
            {slides.length} スライド • {downloadName}
          </span>
        </div>
      </div>
      <SlideCard slide={slides[current]} index={current} total={slides.length} ratio={layout.ratio} />
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "12px", marginTop: "1rem" }}>
        <button onClick={() => setCurrent(c => Math.max(0, c - 1))} disabled={current === 0}
          style={{ background: PAL.bg, border: `1px solid ${PAL.border}`, borderRadius: "8px", padding: "0.5rem 1.1rem", cursor: "pointer", fontSize: "1.1rem", color: PAL.navy, opacity: current === 0 ? 0.3 : 1 }}>◀</button>
        <span style={{ color: PAL.muted, fontSize: "13px", fontFamily: "monospace", minWidth: "70px", textAlign: "center" }}>{current + 1} / {slides.length}</span>
        <button onClick={() => setCurrent(c => Math.min(slides.length - 1, c + 1))} disabled={current === slides.length - 1}
          style={{ background: PAL.bg, border: `1px solid ${PAL.border}`, borderRadius: "8px", padding: "0.5rem 1.1rem", cursor: "pointer", fontSize: "1.1rem", color: PAL.navy, opacity: current === slides.length - 1 ? 0.3 : 1 }}>▶</button>
      </div>
      <div style={{ display: "flex", gap: "8px", marginTop: "1rem", overflowX: "auto", paddingBottom: "6px" }}>
        {slides.map((sl, i) => (
          <div key={i} onClick={() => setCurrent(i)} style={{
            flexShrink: 0, width: orientation === "landscape" ? "90px" : "60px",
            aspectRatio: `${layout.ratio}`,
            background: `#${THEMES[sl.type === "title" ? 0 : (i % THEMES.length)].bg}`,
            borderRadius: "4px", border: `2px solid ${i === current ? PAL.accent : PAL.border}`,
            cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: "8px", color: `#${THEMES[sl.type === "title" ? 0 : (i % THEMES.length)].title}`,
            overflow: "hidden", padding: "3px", boxSizing: "border-box",
            textAlign: "center", lineHeight: 1.2,
          }}>
            {sl.title?.slice(0, 12)}{sl.title?.length > 12 ? "…" : ""}
          </div>
        ))}
      </div>
    </SectionCard>
    <SectionCard label="ダウンロード">
      <div style={{ display: "flex", justifyContent: "center", gap: "12px", flexWrap: "wrap" }}>
        <button onClick={onDownload} style={{ background: PAL.accent, color: "#fff", border: "none", borderRadius: "10px", padding: "0.8rem 2rem", fontSize: "1rem", fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
          ⬇ PowerPoint (.pptx) をダウンロード
        </button>
        <button onClick={onReset} style={{ background: "transparent", color: PAL.navy, border: `2px solid ${PAL.navy}`, borderRadius: "10px", padding: "0.7rem 1.5rem", fontSize: "0.95rem", fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
          ↺ {resetLabel}
        </button>
      </div>
      <p style={{ color: PAL.muted, fontSize: "0.8rem", marginTop: "0.75rem", textAlign: "center" }}>{downloadName} で出力されます</p>
    </SectionCard>
  </div>
);

// ─── PPTX 生成共通関数 ───────────────────────────────────────────
const buildPptx = async (slides, layout, fileName) => {
  const PptxGenJS = await loadPptxGen();
  const pptx = new PptxGenJS();
  const layoutName = `custom_${Date.now()}`;
  pptx.defineLayout({ name: layoutName, width: layout.w, height: layout.h });
  pptx.layout = layoutName;
  const W = layout.w, H = layout.h;

  slides.forEach((slide, idx) => {
    const s = pptx.addSlide();
    const t = THEMES[slide.type === "title" ? 0 : (idx % THEMES.length)];
    s.background = { color: t.bg };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.15, h: H, fill: { color: t.accent }, line: { color: t.accent } });

    if (slide.type === "title") {
      if (slide.category) s.addText(slide.category, { x: 0.4, y: H * 0.22, w: W - 0.8, h: 0.4, fontSize: 11, bold: true, color: t.accent, charSpacing: 3, fontFace: "Calibri" });
      s.addText(slide.title, { x: 0.4, y: H * 0.32, w: W - 0.8, h: H * 0.3, fontSize: Math.min(36, Math.floor(H * 5.5)), bold: true, color: t.title, fontFace: "Georgia", valign: "top" });
      if (slide.subtitle) s.addText(slide.subtitle, { x: 0.4, y: H * 0.63, w: W - 0.8, h: H * 0.15, fontSize: Math.min(16, Math.floor(H * 2.2)), color: t.body, fontFace: "Calibri" });
      if (slide.url) s.addText(slide.url, { x: 0.4, y: H - 0.5, w: W - 0.8, h: 0.35, fontSize: 10, color: t.accent, fontFace: "Courier New" });
    } else {
      const ctop = slide.section ? 0.7 : 0.55;
      if (slide.section) s.addText(slide.section, { x: 0.4, y: 0.25, w: W - 0.8, h: 0.35, fontSize: 10, bold: true, color: t.accent, charSpacing: 2, fontFace: "Calibri" });
      s.addText(slide.title, { x: 0.4, y: ctop, w: W - 0.8, h: 0.55, fontSize: Math.min(22, Math.floor(H * 3)), bold: true, color: t.title, fontFace: "Georgia", margin: 0 });
      s.addShape(pptx.ShapeType.rect, { x: 0.4, y: ctop + 0.6, w: 1.2, h: 0.04, fill: { color: t.accent }, line: { color: t.accent } });
      if (slide.bullets?.length) {
        const bH = slide.highlight ? H - ctop - 1.45 : H - ctop - 1.0;
        s.addText(
          slide.bullets.map((b, bi) => ({ text: b, options: { bullet: true, color: t.body, fontSize: Math.min(14, Math.floor(H * 2)), paraSpaceAfter: 4, breakLine: bi < slide.bullets.length - 1 } })),
          { x: 0.4, y: ctop + 0.75, w: W - 0.8, h: Math.max(0.5, bH), fontFace: "Calibri", valign: "top" }
        );
      }
      if (slide.highlight) {
        s.addShape(pptx.ShapeType.rect, { x: 0.4, y: H - 0.65, w: W - 0.8, h: 0.45, fill: { color: t.accent + "22" }, line: { color: t.accent, width: 1 } });
        s.addText(slide.highlight, { x: 0.55, y: H - 0.63, w: W - 1.1, h: 0.42, fontSize: 10, italic: true, color: t.title, fontFace: "Calibri", valign: "middle" });
      }
      if (slide.notes) s.addNotes(slide.notes);
    }
    s.addText(`${idx + 1} / ${slides.length}`, { x: W - 1.2, y: H - 0.32, w: 0.9, h: 0.25, fontSize: 9, align: "right", color: t.title + "55" });
  });

  await pptx.writeFile({ fileName });
};

// ─── Claude API ヘルパー ─────────────────────────────────────────
const callClaude = async (messages, system, useSearch = false, maxTokens = 4000) => {
  const body = { model: "claude-sonnet-4-20250514", max_tokens: maxTokens, system, messages };
  if (useSearch) body.tools = [{ type: "web_search_20250305", name: "web_search" }];
  let msgs = [...messages];
  while (true) {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...body, messages: msgs }),
    });
    if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || `APIエラー: ${res.status}`); }
    const data = await res.json();
    if (data.stop_reason === "tool_use") {
      msgs.push({ role: "assistant", content: data.content });
      msgs.push({ role: "user", content: data.content.filter(b => b.type === "tool_use").map(b => ({ type: "tool_result", tool_use_id: b.id, content: "検索完了" })) });
      continue;
    }
    return data.content.filter(b => b.type === "text").map(b => b.text).join("");
  }
};

// JSON安全パース（末尾切れ対策）
const safeParseJSON = (raw) => {
  let clean = raw.replace(/^```[\w]*\n?/, "").replace(/\n?```$/, "").trim();
  try { return JSON.parse(clean); } catch {
    const lb = Math.max(clean.lastIndexOf("]"), clean.lastIndexOf("}"));
    if (lb > 0) {
      clean = clean.slice(0, lb + 1);
      const opens = (clean.match(/\{/g)||[]).length - (clean.match(/\}/g)||[]).length;
      const arrs  = (clean.match(/\[/g)||[]).length - (clean.match(/\]/g)||[]).length;
      clean += "]".repeat(Math.max(0, arrs)) + "}".repeat(Math.max(0, opens));
    }
    return JSON.parse(clean);
  }
};

// ─── 共通UI部品 ─────────────────────────────────────────────────
const SectionCard = ({ label, children }) => (
  <div style={{ background: PAL.surface, border: `1px solid ${PAL.border}`, borderRadius: "14px", padding: "1.5rem", marginBottom: "1.25rem", boxShadow: "0 2px 12px rgba(27,43,75,0.06)" }}>
    <div style={{ fontSize: "11px", fontWeight: 700, color: PAL.muted, letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: "1rem" }}>{label}</div>
    {children}
  </div>
);

const Spinner = () => (
  <div style={{ display: "flex", justifyContent: "center" }}>
    <div style={{ width: "28px", height: "28px", border: `3px solid ${PAL.border}`, borderTop: `3px solid ${PAL.accent}`, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
  </div>
);

const ErrorBox = ({ msg }) => (
  <div style={{ marginTop: "1rem", padding: "0.75rem 1rem", background: "#FFF0F0", border: `1px solid ${PAL.error}44`, borderRadius: "8px", color: PAL.error, fontSize: "0.9rem" }}>
    ⚠ {msg}
  </div>
);

const Step = ({ label, done, active }) => (
  <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
    <div style={{ width: "22px", height: "22px", borderRadius: "50%", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", background: done ? PAL.success : active ? PAL.accent : PAL.border, color: done || active ? "#fff" : PAL.muted, fontSize: "11px", fontWeight: 700, transition: "all 0.3s" }}>
      {done ? "✓" : active ? "●" : "○"}
    </div>
    <span style={{ fontSize: "13px", color: active ? PAL.accent : done ? PAL.success : PAL.muted, fontWeight: active || done ? 600 : 400 }}>{label}</span>
  </div>
);

// ══════════════════════════════════════════════════════════════════
// タブ1: URL → 会社紹介資料
// ══════════════════════════════════════════════════════════════════
const UrlTab = ({ paperSize, setPaperSize, orientation, setOrientation }) => {
  const [url, setUrl] = useState("");
  const [status, setStatus] = useState("idle");
  const [slides, setSlides] = useState([]);
  const [current, setCurrent] = useState(0);
  const [error, setError] = useState("");
  const [steps, setSteps] = useState({ fetch: false, analyze: false, build: false });
  const [companyName, setCompanyName] = useState("");
  const layout = getLayout(paperSize, orientation);

  const analyze = async () => {
    const trimmedUrl = url.trim();
    if (!trimmedUrl) return;
    setStatus("processing"); setError(""); setSlides([]); setCurrent(0);
    setSteps({ fetch: false, analyze: false, build: false });
    try {
      setSteps({ fetch: true, analyze: false, build: false });
      const searchResult = await callClaude([{
        role: "user",
        content: `以下のURLの企業・サービスについてウェブ検索で詳細情報を収集してください。\nURL: ${trimmedUrl}\n\n企業名、事業内容、サービス・商品、強み・特徴、ターゲット顧客、主要な実績や数値などを収集してください。`,
      }], "あなたはウェブ調査の専門家です。指定URLの企業について詳細に調査し整理して報告してください。", true, 1000);

      setSteps({ fetch: true, analyze: true, build: false });
      const raw = await callClaude([{
        role: "user",
        content: `以下の調査結果を元にスライド用JSONを生成してください。\n\n【調査結果】\n${searchResult}\n\n【元のURL】${trimmedUrl}`,
      }], `あなたは企業情報を会社紹介スライドに変換する専門AIです。
以下のJSON形式のみで出力してください。コードフェンスや説明文は不要です。

{
  "companyName": "企業名",
  "slides": [
    { "type": "title", "title": "企業名またはキャッチコピー", "subtitle": "事業内容の一言説明", "category": "COMPANY PROFILE", "url": "ドメイン部分" },
    { "type": "content", "section": "OVERVIEW", "title": "会社概要", "bullets": ["基本情報","ミッション","主要事業"], "highlight": "強調数値（任意、なければnull）" },
    { "type": "content", "section": "SERVICES", "title": "主要サービス・事業", "bullets": ["サービス1","サービス2","サービス3"], "highlight": null },
    { "type": "content", "section": "STRENGTHS", "title": "強みと特徴", "bullets": ["強み1","強み2","強み3"], "highlight": null },
    { "type": "content", "section": "MARKET", "title": "ターゲット・市場", "bullets": ["顧客層","対象市場","解決課題"], "highlight": null },
    { "type": "content", "section": "CONTACT", "title": "お問い合わせ・詳細", "bullets": ["URL","サービス入口","次のアクション"], "highlight": null }
  ]
}
ルール: スライドは6〜8枚、bullets は3〜5項目、日本語で出力`, false, 4000);

      setSteps({ fetch: true, analyze: true, build: true });
      const parsed = safeParseJSON(raw);
      if (!parsed.slides?.length) throw new Error("スライドデータの形式が正しくありません");
      setCompanyName(parsed.companyName || "Company");
      setSlides(parsed.slides); setCurrent(0); setStatus("done");
    } catch (err) {
      setError(err.message || "解析に失敗しました"); setStatus("error");
    }
  };

  const reset = () => { setUrl(""); setSlides([]); setStatus("idle"); setCurrent(0); setError(""); setCompanyName(""); };
  const downloadName = `${paperSize} ${orientation === "landscape" ? "横" : "縦"}`;

  return (
    <div>
      <SectionCard label="用紙サイズと向きを選択">
        <PaperSettings paperSize={paperSize} setPaperSize={setPaperSize} orientation={orientation} setOrientation={setOrientation} />
      </SectionCard>
      <SectionCard label="URLを入力">
        <input
          value={url} onChange={e => setUrl(e.target.value)}
          onKeyDown={e => e.key === "Enter" && status !== "processing" && analyze()}
          placeholder="https://www.example.co.jp"
          disabled={status === "processing"}
          style={{ width: "100%", padding: "0.75rem 1rem", border: `2px solid ${PAL.border}`, borderRadius: "10px", fontSize: "1rem", color: PAL.navy, background: PAL.bg, outline: "none", fontFamily: "inherit", boxSizing: "border-box" }}
        />
        <p style={{ margin: "0.5rem 0 0", fontSize: "0.8rem", color: PAL.muted }}>会社・サービスのトップページURLを入力してください</p>
      </SectionCard>
      <SectionCard label="資料を生成">
        {status === "processing" ? (
          <div style={{ animation: "slideUp 0.3s ease", textAlign: "center" }}>
            <div style={{ display: "inline-flex", flexDirection: "column", gap: "12px", marginBottom: "1rem", textAlign: "left" }}>
              <Step label="ホームページを取得・検索中..." done={steps.fetch && steps.analyze} active={steps.fetch && !steps.analyze} />
              <Step label="AIが情報を分析中..." done={steps.analyze && steps.build} active={steps.analyze && !steps.build} />
              <Step label="スライドを構築中..." done={steps.build} active={steps.build && status === "processing"} />
            </div>
            <Spinner />
          </div>
        ) : (
          <button onClick={analyze} disabled={!url.trim() || status === "processing"}
            style={{ background: PAL.navy, color: "#fff", border: "none", borderRadius: "10px", padding: "0.8rem 2rem", fontSize: "1rem", fontWeight: 700, cursor: url.trim() ? "pointer" : "not-allowed", fontFamily: "inherit", opacity: url.trim() ? 1 : 0.4 }}>
            🔍 AIで会社紹介資料を生成する
          </button>
        )}
        {status === "error" && <ErrorBox msg={error} />}
      </SectionCard>
      {status === "done" && slides.length > 0 && (
        <SlidePreviewPanel
          slides={slides} current={current} setCurrent={setCurrent}
          layout={layout} orientation={orientation}
          label="プレビュー" badge={companyName} downloadName={downloadName}
          onDownload={() => buildPptx(slides, layout, `${companyName}_会社紹介.pptx`).catch(e => setError(e.message))}
          onReset={reset} resetLabel="別のURLで作り直す"
        />
      )}
    </div>
  );
};

// ══════════════════════════════════════════════════════════════════
// タブ2: PDF → スライド
// ══════════════════════════════════════════════════════════════════
const PdfTab = ({ paperSize, setPaperSize, orientation, setOrientation }) => {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("idle");
  const [slides, setSlides] = useState([]);
  const [current, setCurrent] = useState(0);
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [progress, setProgress] = useState("");
  const fileRef = useRef();
  const layout = getLayout(paperSize, orientation);

  const handleFile = (f) => {
    if (!f) return;
    if (f.type !== "application/pdf") { setError("PDFファイルのみ対応しています"); return; }
    setFile(f); setStatus("idle"); setSlides([]); setError(""); setCurrent(0);
  };

  const analyze = async () => {
    if (!file) return;
    setStatus("processing"); setError(""); setProgress("PDFを読み込み中...");
    try {
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = rej;
        r.readAsDataURL(file);
      });
      setProgress("Claude AI が内容を解析中...");
      const raw = await callClaude([{
        role: "user",
        content: [
          { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
          { type: "text", text: "このPDFの内容をプレゼンテーションスライドに変換してください。重要な情報を整理し、わかりやすいスライド構成にしてください。" }
        ]
      }], `あなたはPDFをプレゼンテーションスライドに変換する専門AIです。
以下のJSON形式のみで出力してください。コードフェンスや説明文は不要です。

{
  "title": "プレゼンテーション全体のタイトル",
  "slides": [
    { "type": "title", "title": "タイトル", "subtitle": "サブタイトルや日付" },
    { "type": "content", "title": "スライドのタイトル", "bullets": ["要点1","要点2","要点3"], "notes": "補足説明（任意）" }
  ]
}

ルール: 最初は必ずtype:"title"、スライドは6〜10枚、bullets は3〜5項目、日本語で出力`, false, 4000);

      setProgress("スライドを構築中...");
      const parsed = safeParseJSON(raw);
      if (!parsed.slides?.length) throw new Error("スライドデータの形式が正しくありません");
      setSlides(parsed.slides); setCurrent(0); setStatus("done");
    } catch (err) {
      setError(err.message || "解析に失敗しました"); setStatus("error");
    }
  };

  const reset = () => { setFile(null); setSlides([]); setStatus("idle"); setCurrent(0); setError(""); };
  const downloadName = `${paperSize} ${orientation === "landscape" ? "横" : "縦"}`;
  const fileName = file ? file.name.replace(/\.pdf$/i, "") + "_スライド.pptx" : "presentation.pptx";

  return (
    <div>
      <SectionCard label="用紙サイズと向きを選択">
        <PaperSettings paperSize={paperSize} setPaperSize={setPaperSize} orientation={orientation} setOrientation={setOrientation} />
      </SectionCard>
      <SectionCard label="PDFをアップロード">
        <div
          onClick={() => fileRef.current?.click()}
          onDrop={e => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]); }}
          onDragOver={e => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          style={{ border: `2px dashed ${dragging ? PAL.accent : PAL.border}`, borderRadius: "12px", padding: "2.5rem 2rem", textAlign: "center", cursor: "pointer", background: dragging ? `${PAL.accent}08` : PAL.bg }}>
          <div style={{ fontSize: "2.2rem", marginBottom: "0.5rem" }}>{file ? "📄" : "📂"}</div>
          <div style={{ color: file ? PAL.accent : PAL.muted, fontWeight: 600 }}>
            {file ? file.name : "ここにPDFをドラッグ＆ドロップ"}
          </div>
          <div style={{ color: PAL.muted, fontSize: "0.85rem", marginTop: "0.3rem" }}>
            {file ? `${(file.size / 1024).toFixed(0)} KB — クリックで変更` : "または クリックしてファイルを選択"}
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".pdf,application/pdf" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
      </SectionCard>
      <SectionCard label="AIで解析・変換">
        {status === "processing" ? (
          <div style={{ animation: "slideUp 0.3s ease", textAlign: "center" }}>
            <Spinner />
            <div style={{ color: PAL.accent, fontStyle: "italic", marginTop: "0.75rem" }}>{progress}</div>
            <div style={{ color: PAL.muted, fontSize: "0.85rem", marginTop: "0.3rem" }}>PDFの量によって1〜2分かかる場合があります</div>
          </div>
        ) : (
          <button onClick={analyze} disabled={!file || status === "processing"}
            style={{ background: PAL.navy, color: "#fff", border: "none", borderRadius: "10px", padding: "0.8rem 2rem", fontSize: "1rem", fontWeight: 700, cursor: file ? "pointer" : "not-allowed", fontFamily: "inherit", opacity: file ? 1 : 0.4 }}>
            ✦ AIでスライドに変換する
          </button>
        )}
        {status === "error" && <ErrorBox msg={error} />}
      </SectionCard>
      {status === "done" && slides.length > 0 && (
        <SlidePreviewPanel
          slides={slides} current={current} setCurrent={setCurrent}
          layout={layout} orientation={orientation}
          label="プレビュー" badge={null} downloadName={downloadName}
          onDownload={() => buildPptx(slides, layout, fileName).catch(e => setError(e.message))}
          onReset={reset} resetLabel="別のPDFを変換"
        />
      )}
    </div>
  );
};

// ══════════════════════════════════════════════════════════════════
// メインアプリ（タブ切り替え）
// ══════════════════════════════════════════════════════════════════
export default function App() {
  const [tab, setTab] = useState("url");
  const [paperSize, setPaperSize] = useState("A4");
  const [orientation, setOrientation] = useState("landscape");

  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Source+Sans+3:wght@400;600;700&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    loadPptxGen().catch(() => {});
  }, []);

  const tabs = [
    { id: "url", icon: "🏢", label: "URL → 会社紹介資料" },
    { id: "pdf", icon: "📄", label: "PDF → スライド変換" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: PAL.bg, fontFamily: "'Source Sans 3','Hiragino Sans',sans-serif", color: PAL.navy, padding: "2rem", boxSizing: "border-box" }}>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes slideUp { from { opacity:0; transform:translateY(12px); } to { opacity:1; transform:translateY(0); } }
        button { transition: opacity 0.15s, background 0.15s; }
        button:hover { opacity: 0.85; }
        input:focus { border-color: #C8602A !important; box-shadow: 0 0 0 3px rgba(200,96,42,0.12); }
      `}</style>

      <div style={{ maxWidth: "840px", margin: "0 auto" }}>
        {/* ヘッダー */}
        <div style={{ marginBottom: "2rem" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "0.3rem" }}>
            <div style={{ width: "38px", height: "38px", background: PAL.navy, borderRadius: "10px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "20px", flexShrink: 0 }}>✦</div>
            <h1 style={{ margin: 0, fontSize: "1.7rem", fontWeight: 800, color: PAL.navy, fontFamily: "'Playfair Display',Georgia,serif" }}>
              AI スライドジェネレーター
            </h1>
          </div>
          <p style={{ margin: 0, color: PAL.muted, fontSize: "0.9rem", paddingLeft: "50px" }}>
            URLまたはPDFから、プレゼンテーション資料を自動生成します
          </p>
        </div>

        {/* タブ */}
        <div style={{ display: "flex", gap: "8px", marginBottom: "1.5rem", background: PAL.surface, border: `1px solid ${PAL.border}`, borderRadius: "14px", padding: "6px", boxShadow: "0 2px 8px rgba(27,43,75,0.06)" }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              flex: 1, padding: "0.75rem 1rem", borderRadius: "10px", cursor: "pointer",
              border: "none", fontFamily: "inherit", fontSize: "0.95rem", fontWeight: 700,
              background: tab === t.id ? PAL.navy : "transparent",
              color: tab === t.id ? "#fff" : PAL.muted,
              display: "flex", alignItems: "center", justifyContent: "center", gap: "8px",
            }}>
              <span>{t.icon}</span><span>{t.label}</span>
            </button>
          ))}
        </div>

        {/* タブコンテンツ */}
        <div key={tab} style={{ animation: "slideUp 0.25s ease" }}>
          {tab === "url"
            ? <UrlTab paperSize={paperSize} setPaperSize={setPaperSize} orientation={orientation} setOrientation={setOrientation} />
            : <PdfTab paperSize={paperSize} setPaperSize={setPaperSize} orientation={orientation} setOrientation={setOrientation} />
          }
        </div>

        <div style={{ textAlign: "center", color: PAL.muted, fontSize: "0.78rem", marginTop: "2rem" }}>
          Powered by Claude API + Web Search ── APIキー不要
        </div>
      </div>
    </div>
  );
}
