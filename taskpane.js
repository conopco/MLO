/* ============================================================================
 * MLO 法令ペイン (Office.js / Word アドイン)
 * ----------------------------------------------------------------------------
 * Word で文字列を「選択（クリックでドラッグ選択）」すると、その中の
 *   「法令名 + 条・項・号」を解析し、e-Gov 法令API から条文を取得して
 *   タスクペインに表示します。クリックで e-Gov の該当法令ページも開けます。
 *
 *   出典 / API : https://laws.e-gov.go.jp/
 *   作成者(Author) : MLO   改訂(Revision) : MLO
 * ==========================================================================*/
"use strict";

/* ---------------------------------------------------------------------------
 * 設定
 *  CORS で直接アクセスがブロックされる場合は PROXY にプロキシURLを設定してください。
 *  （README の Cloudflare Worker サンプル参照。空文字なら直接アクセス）
 *    例: const PROXY = "https://your-worker.example.workers.dev/?url=";
 * ------------------------------------------------------------------------- */
const PROXY = "";
const API_V2 = "https://laws.e-gov.go.jp/api/2";
const API_V1 = "https://laws.e-gov.go.jp/api/1";
const WEB = "https://laws.e-gov.go.jp/law/";

/* セッション内キャッシュ */
const idCache = new Map();        // 正式法令名 -> 法令ID
const articleCache = new Map();   // key -> 条文テキスト

/* 略称・通称 -> 正式名称（企業法務で頻出のもの。必要に応じ追記可） */
const ALIAS = {
  "独占禁止法": "私的独占の禁止及び公正取引の確保に関する法律",
  "独禁法": "私的独占の禁止及び公正取引の確保に関する法律",
  "下請法": "下請代金支払遅延等防止法",
  "景品表示法": "不当景品類及び不当表示防止法",
  "景表法": "不当景品類及び不当表示防止法",
  "個人情報保護法": "個人情報の保護に関する法律",
  "特定商取引法": "特定商取引に関する法律",
  "特商法": "特定商取引に関する法律",
  "PL法": "製造物責任法",
  "金商法": "金融商品取引法",
  "労働者派遣法": "労働者派遣事業の適正な運営の確保及び派遣労働者の保護等に関する法律",
  "派遣法": "労働者派遣事業の適正な運営の確保及び派遣労働者の保護等に関する法律",
  "パートタイム・有期雇用労働法": "短時間労働者及び有期雇用労働者の雇用管理の改善等に関する法律",
  "一般法人法": "一般社団法人及び一般財団法人に関する法律",
  "公益法人認定法": "公益社団法人及び公益財団法人の認定等に関する法律",
  "電子契約法": "電子消費者契約に関する民法の特例に関する法律",
  "プロバイダ責任制限法": "特定電気通信役務提供者の損害賠償責任の制限及び発信者情報の開示に関する法律",
  "金融サービス提供法": "金融サービスの提供及び利用環境の整備等に関する法律",
  "犯収法": "犯罪による収益の移転防止に関する法律",
  "マイナンバー法": "行政手続における特定の個人を識別するための番号の利用等に関する法律",
  "番号法": "行政手続における特定の個人を識別するための番号の利用等に関する法律",
  "外為法": "外国為替及び外国貿易法",
  "労基法": "労働基準法",
  "道交法": "道路交通法"
};

/* ===========================================================================
 * Office 初期化
 * ======================================================================== */
Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;

  document.getElementById("status").textContent = "準備完了。本文を選択すると解析します。";

  // 選択変更イベント（デバウンス）
  let timer = null;
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => {
      clearTimeout(timer);
      timer = setTimeout(analyzeSelection, 300);
    }
  );

  document.getElementById("btnAnalyzeSel").addEventListener("click", analyzeSelection);
  document.getElementById("btnAnalyzeInput").addEventListener("click", () => {
    const t = document.getElementById("manualInput").value || "";
    handleText(t);
  });
});

/* 選択テキストを取得して解析 */
function analyzeSelection() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("選択テキストを取得できませんでした。");
      return;
    }
    handleText(res.value || "");
  });
}

/* ===========================================================================
 * テキスト処理 → 解析 → 表示
 * ======================================================================== */
async function handleText(text) {
  const results = document.getElementById("results");
  text = (text || "").trim();

  if (!text) {
    setStatus("法令参照を含む箇所を選択してください。");
    results.innerHTML = "";
    return;
  }

  const refs = parseReferences(text);
  if (refs.length === 0) {
    setStatus("選択範囲から法令の条文参照を検出できませんでした。");
    results.innerHTML = "";
    return;
  }

  setStatus(`${refs.length} 件の参照を検出。条文を取得中…`);
  results.innerHTML = "";

  for (const ref of refs) {
    const card = renderCardSkeleton(ref);
    results.appendChild(card);

    try {
      const lawId = await resolveLawId(ref.name);
      if (!lawId) {
        fillCardError(card, "法令IDを特定できませんでした（法令名の表記を確認してください）。");
        continue;
      }
      const body = await fetchArticle(lawId, ref.art, ref.branch, ref.para);
      fillCard(card, ref, lawId, body);
    } catch (e) {
      fillCardError(card, "取得に失敗しました（CORS の可能性。README のプロキシ設定をご確認ください）。");
    }
  }
  setStatus(`完了（${refs.length} 件）。`);
}

/* ===========================================================================
 * 参照解析（法令名 + 条・項・号。表記ゆれ対応）
 *   group: 1=法令名(任意) 2=条 3=条の枝番 4=項 5=号
 * ======================================================================== */
function buildRegex() {
  const NUM = "[0-9０-９一二三四五六七八九十百千〇零]+";
  const NAME =
    "(?:同法|本法|当該法律|[一-龥々〆ヶ・ー、A-Za-z0-9０-９]{1,40}?" +
    "(?:に関する法律|に関する法|施行令|施行規則|規則|条例|省令|政令|府令|府省令|勅令|令|法律|法))";
  const PAT =
    `(${NAME})?` +
    `\\s*(?:第\\s*)?(${NUM})\\s*条` +
    `(?:\\s*の\\s*(${NUM}))?` +
    `(?:\\s*(?:第\\s*)?(${NUM})\\s*項)?` +
    `(?:\\s*(?:第\\s*)?(${NUM})\\s*号)?`;
  return new RegExp(PAT, "g");
}

function parseReferences(text) {
  const re = buildRegex();
  const out = [];
  let lastLaw = "";
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index === re.lastIndex) re.lastIndex++; // 無限ループ防止

    let name = (m[1] || "").trim();
    const art = toNum(m[2]);
    let branch = m[3] ? String(toNum(m[3])) : "";
    const para = m[4] ? toNum(m[4]) : 0;
    const item = m[5] ? toNum(m[5]) : 0;

    if (!name || name === "同法" || name === "本法" || name === "当該法律") {
      name = lastLaw;
    } else {
      lastLaw = name;
    }
    if (!name || !art) continue;

    out.push({ raw: m[0].trim(), name, art, branch, para, item });
  }
  // 重複（同一引用）の除去
  const seen = new Set();
  return out.filter((r) => {
    const k = `${r.name}|${r.art}|${r.branch}|${r.para}|${r.item}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

/* ===========================================================================
 * 法令ID 解決（v2 /laws  法令名検索）
 * ======================================================================== */
async function resolveLawId(name) {
  name = (name || "").trim();
  if (!name) return "";
  const official = ALIAS[name] || name;
  if (idCache.has(official)) return idCache.get(official);

  const url = `${API_V2}/laws?law_title=${encodeURIComponent(official)}&limit=10`;
  let data;
  try {
    data = await apiFetchJson(url);
  } catch (e) {
    throw e;
  }

  const laws = (data && (data.laws || data.Laws)) || [];
  let best = "";
  let bestLen = Infinity;
  for (const l of laws) {
    const title =
      (l.revision_info && l.revision_info.law_title) ||
      (l.law_info && l.law_info.law_title) ||
      l.law_title || "";
    const id =
      (l.law_info && l.law_info.law_id) ||
      l.law_id || l.LawId || "";
    if (!id) continue;
    if (title === official) { best = id; bestLen = 0; break; }      // 完全一致優先
    if (title.startsWith(official) && title.length < bestLen) {     // 前方一致は最短
      best = id; bestLen = title.length;
    }
  }
  if (!best && laws.length > 0) {
    const l0 = laws[0];
    best = (l0.law_info && l0.law_info.law_id) || l0.law_id || l0.LawId || "";
  }
  idCache.set(official, best);
  return best;
}

/* ===========================================================================
 * 条文取得（v1 条文内容取得API。XML を解析）
 * ======================================================================== */
async function fetchArticle(lawId, art, branch, para) {
  const key = `${lawId}|${art}|${branch}|${para}`;
  if (articleCache.has(key)) return articleCache.get(key);

  let articleParam = "第" + toKanji(art) + "条";
  if (branch) articleParam += "の" + toKanji(parseInt(branch, 10));

  const url = `${API_V1}/articles;lawId=${lawId};article=${encodeURIComponent(articleParam)}`;
  const xmlText = await apiFetchText(url);

  let body = "";
  try {
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    const code = doc.getElementsByTagName("Code")[0];
    if (!code || code.textContent === "0") {
      const articles = doc.getElementsByTagName("Article");
      if (articles.length > 0) {
        if (para > 0) {
          const paras = articles[0].getElementsByTagName("Paragraph");
          if (paras.length >= para) body = collectSentences(paras[para - 1]);
        }
        if (!body) body = collectSentences(articles[0]);
      }
    }
  } catch (e) {
    body = "";
  }

  articleCache.set(key, body);
  return body;
}

function collectSentences(node) {
  const sents = node.getElementsByTagName("Sentence");
  let s = "";
  for (let i = 0; i < sents.length; i++) {
    const t = (sents[i].textContent || "").trim();
    if (t) s += t;
  }
  return s;
}

/* ===========================================================================
 * 通信（PROXY 設定があれば経由）
 * ======================================================================== */
function withProxy(url) {
  return PROXY ? PROXY + encodeURIComponent(url) : url;
}
async function apiFetchJson(url) {
  const r = await fetch(withProxy(url), { headers: { "Accept": "application/json" } });
  if (!r.ok) throw new Error("HTTP " + r.status);
  return await r.json();
}
async function apiFetchText(url) {
  const r = await fetch(withProxy(url));
  if (!r.ok) throw new Error("HTTP " + r.status);
  return await r.text();
}

/* ===========================================================================
 * 数値ユーティリティ（漢数字 / 算用 / 全角 に対応）
 * ======================================================================== */
function toNum(s) {
  if (!s) return 0;
  s = String(s).trim();
  const KAN = "一二三四五六七八九十百千〇零";
  let hasKanji = false;
  for (const ch of s) if (KAN.indexOf(ch) >= 0) { hasKanji = true; break; }

  if (!hasKanji) {
    let t = "";
    for (const ch of s) {
      if (ch >= "０" && ch <= "９") t += String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30);
      else if (ch >= "0" && ch <= "9") t += ch;
    }
    return t ? parseInt(t, 10) : 0;
  }

  let val = 0, tmp = 0;
  for (const ch of s) {
    switch (ch) {
      case "〇": case "零": tmp = tmp * 10; break;
      case "一": tmp = tmp * 10 + 1; break;
      case "二": tmp = tmp * 10 + 2; break;
      case "三": tmp = tmp * 10 + 3; break;
      case "四": tmp = tmp * 10 + 4; break;
      case "五": tmp = tmp * 10 + 5; break;
      case "六": tmp = tmp * 10 + 6; break;
      case "七": tmp = tmp * 10 + 7; break;
      case "八": tmp = tmp * 10 + 8; break;
      case "九": tmp = tmp * 10 + 9; break;
      case "十": val += (tmp === 0 ? 1 : tmp) * 10; tmp = 0; break;
      case "百": val += (tmp === 0 ? 1 : tmp) * 100; tmp = 0; break;
      case "千": val += (tmp === 0 ? 1 : tmp) * 1000; tmp = 0; break;
    }
  }
  return val + tmp;
}

function toKanji(n) {
  const k = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
  let res = "";
  const sen = Math.floor(n / 1000); n %= 1000;
  const hyaku = Math.floor(n / 100); n %= 100;
  const juu = Math.floor(n / 10); const ichi = n % 10;
  if (sen > 0) res += (sen === 1 ? "" : k[sen]) + "千";
  if (hyaku > 0) res += (hyaku === 1 ? "" : k[hyaku]) + "百";
  if (juu > 0) res += (juu === 1 ? "" : k[juu]) + "十";
  if (ichi > 0) res += k[ichi];
  return res || "〇";
}

/* ===========================================================================
 * 表示
 * ======================================================================== */
function citationText(ref) {
  let c = `${ref.name}第${ref.art}条`;
  if (ref.branch) c += `の${ref.branch}`;
  if (ref.para) c += `第${ref.para}項`;
  if (ref.item) c += `第${ref.item}号`;
  return c;
}

function renderCardSkeleton(ref) {
  const card = document.createElement("div");
  card.className = "card";
  card.innerHTML =
    `<div class="cite">${escapeHtml(citationText(ref))}</div>` +
    `<div class="body loading">取得中…</div>`;
  return card;
}

function fillCard(card, ref, lawId, body) {
  const bodyEl = card.querySelector(".body");
  bodyEl.classList.remove("loading");
  if (!body) {
    bodyEl.innerHTML =
      `<span class="muted">条文本文を取得できませんでした。下のリンクから e-Gov でご確認ください。</span>`;
  } else {
    bodyEl.textContent = body;
  }
  const link = document.createElement("a");
  link.className = "egov";
  link.href = WEB + lawId;
  link.target = "_blank";
  link.rel = "noopener";
  link.textContent = "e-Gov で開く ↗";
  card.appendChild(link);
}

function fillCardError(card, msg) {
  const bodyEl = card.querySelector(".body");
  bodyEl.classList.remove("loading");
  bodyEl.innerHTML = `<span class="muted">${escapeHtml(msg)}</span>`;
}

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
