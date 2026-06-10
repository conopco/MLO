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
const PROXY = "https://floral-silence-3dba.conopco.workers.dev/?url=";
const API_V2 = "https://laws.e-gov.go.jp/api/2";
const API_V1 = "https://laws.e-gov.go.jp/api/1";
const WEB = "https://laws.e-gov.go.jp/law/";

/* セッション内キャッシュ */
const idCache = new Map();        // 正式法令名 -> 法令ID
const articleCache = new Map();   // key -> 条文テキスト
let runSeq = 0;                   // 解析の実行世代（新しい解析が古い解析を無効化）

/* ---------------------------------------------------------------------------
 * 略称・通称 → 正式名称（弁護士業務・企業法務で扱う主要法令を網羅）
 *  ・値は「文字列」または「候補の配列」。配列の場合は先頭から順に e-Gov を検索し、
 *    最初にヒットしたものを採用します（改称の過渡期に新旧どちらでも当たるように）。
 *  ・正式名称がそのまま日常名称の法令（民法・会社法・特許法 等）は、本表に無くても
 *    法令名でそのまま検索されます（必要なら追記してください）。
 * ------------------------------------------------------------------------- */
const ALIAS = {
  /* ── 競争法・取引 ───────────────────────────── */
  "独占禁止法": "私的独占の禁止及び公正取引の確保に関する法律",
  "独禁法": "私的独占の禁止及び公正取引の確保に関する法律",
  // 下請法は2026年1月施行の改正で改称が進む過渡期。新旧双方を試行。
  "下請法": ["下請代金支払遅延等防止法", "製造委託等に係る中小受託事業者に対する代金の支払の遅延等の防止に関する法律"],
  "取適法": ["製造委託等に係る中小受託事業者に対する代金の支払の遅延等の防止に関する法律", "下請代金支払遅延等防止法"],
  "フリーランス新法": "特定受託事業者に係る取引の適正化等に関する法律",
  "フリーランス保護法": "特定受託事業者に係る取引の適正化等に関する法律",
  "フリーランス・事業者間取引適正化等法": "特定受託事業者に係る取引の適正化等に関する法律",

  /* ── 消費者保護・表示 ─────────────────────────── */
  "景品表示法": "不当景品類及び不当表示防止法",
  "景表法": "不当景品類及び不当表示防止法",
  "特定商取引法": "特定商取引に関する法律",
  "特商法": "特定商取引に関する法律",
  "割販法": "割賦販売法",
  "PL法": "製造物責任法",
  "製造物責任法": "製造物責任法",
  "公益通報者保護法": "公益通報者保護法",

  /* ── データ・IT・知財 ────────────────────────── */
  "個人情報保護法": "個人情報の保護に関する法律",
  "個情法": "個人情報の保護に関する法律",
  "マイナンバー法": "行政手続における特定の個人を識別するための番号の利用等に関する法律",
  "番号法": "行政手続における特定の個人を識別するための番号の利用等に関する法律",
  // プロバイダ責任制限法は2024年改正で改称（情プラ法）。新旧双方を試行。
  "プロバイダ責任制限法": ["特定電気通信による情報の流通によって発生する権利侵害等への対処に関する法律", "特定電気通信役務提供者の損害賠償責任の制限及び発信者情報の開示に関する法律"],
  "プロ責法": ["特定電気通信による情報の流通によって発生する権利侵害等への対処に関する法律", "特定電気通信役務提供者の損害賠償責任の制限及び発信者情報の開示に関する法律"],
  "情プラ法": "特定電気通信による情報の流通によって発生する権利侵害等への対処に関する法律",
  "情報流通プラットフォーム対処法": "特定電気通信による情報の流通によって発生する権利侵害等への対処に関する法律",
  "特定電子メール法": "特定電子メールの送信の適正化等に関する法律",
  "電子署名法": "電子署名及び認証業務に関する法律",
  "電子契約法": "電子消費者契約に関する民法の特例に関する法律",
  "不正アクセス禁止法": "不正アクセス行為の禁止等に関する法律",
  "不競法": "不正競争防止法",
  "半導体集積回路配置法": "半導体集積回路の回路配置に関する法律",

  /* ── 会社・組織・金融・証券 ───────────────────── */
  "一般法人法": "一般社団法人及び一般財団法人に関する法律",
  "一般社団・財団法人法": "一般社団法人及び一般財団法人に関する法律",
  "公益法人認定法": "公益社団法人及び公益財団法人の認定等に関する法律",
  "産業競争力強化法": "産業競争力強化法",
  "金商法": "金融商品取引法",
  "金融サービス提供法": ["金融サービスの提供及び利用環境の整備等に関する法律", "金融サービスの提供に関する法律", "金融商品の販売等に関する法律"],
  "金販法": ["金融サービスの提供及び利用環境の整備等に関する法律", "金融商品の販売等に関する法律"],
  "資金決済法": "資金決済に関する法律",
  "犯収法": "犯罪による収益の移転防止に関する法律",
  "出資法": "出資の受入れ、預り金及び金利等の取締りに関する法律",
  "投信法": "投資信託及び投資法人に関する法律",
  "社債株式振替法": "社債、株式等の振替に関する法律",
  "振替法": "社債、株式等の振替に関する法律",

  /* ── 労働 ───────────────────────────────── */
  "労基法": "労働基準法",
  "労契法": "労働契約法",
  "労組法": "労働組合法",
  "労調法": "労働関係調整法",
  "最賃法": "最低賃金法",
  "賃確法": "賃金の支払の確保等に関する法律",
  "安衛法": "労働安全衛生法",
  "労働安全衛生法": "労働安全衛生法",
  "労働者派遣法": "労働者派遣事業の適正な運営の確保及び派遣労働者の保護等に関する法律",
  "派遣法": "労働者派遣事業の適正な運営の確保及び派遣労働者の保護等に関する法律",
  "パートタイム・有期雇用労働法": "短時間労働者及び有期雇用労働者の雇用管理の改善等に関する法律",
  "パート有期法": "短時間労働者及び有期雇用労働者の雇用管理の改善等に関する法律",
  "男女雇用機会均等法": "雇用の分野における男女の均等な機会及び待遇の確保等に関する法律",
  "均等法": "雇用の分野における男女の均等な機会及び待遇の確保等に関する法律",
  "育児介護休業法": "育児休業、介護休業等育児又は家族介護を行う労働者の福祉に関する法律",
  "育介法": "育児休業、介護休業等育児又は家族介護を行う労働者の福祉に関する法律",
  "高年齢者雇用安定法": "高年齢者等の雇用の安定等に関する法律",
  "障害者雇用促進法": "障害者の雇用の促進等に関する法律",
  "労働施策総合推進法": "労働施策の総合的な推進並びに労働者の雇用の安定及び職業生活の充実等に関する法律",
  "パワハラ防止法": "労働施策の総合的な推進並びに労働者の雇用の安定及び職業生活の充実等に関する法律",

  /* ── 手続・倒産・紛争解決 ─────────────────────── */
  "民訴法": "民事訴訟法",
  "ADR法": "裁判外紛争解決手続の利用の促進に関する法律",
  "ADR促進法": "裁判外紛争解決手続の利用の促進に関する法律",

  /* ── 不動産・登記・建設 ──────────────────────── */
  "借地借家法": "借地借家法",
  "区分所有法": "建物の区分所有等に関する法律",
  "マンション法": "建物の区分所有等に関する法律",
  "宅建業法": "宅地建物取引業法",

  /* ── 国際取引・規制・環境 ─────────────────────── */
  "外為法": "外国為替及び外国貿易法",
  "経済安全保障推進法": "経済施策を一体的に講ずることによる安全保障の確保の推進に関する法律",
  "経済安保推進法": "経済施策を一体的に講ずることによる安全保障の確保の推進に関する法律",
  "化審法": "化学物質の審査及び製造等の規制に関する法律",
  "廃棄物処理法": "廃棄物の処理及び清掃に関する法律",
  "廃掃法": "廃棄物の処理及び清掃に関する法律",

  /* ── その他 ─────────────────────────────── */
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

  document.getElementById("btnAnalyzeSel").addEventListener("click", () => analyzeSelection(true));
  document.getElementById("btnAnalyzeInput").addEventListener("click", () => {
    const t = document.getElementById("manualInput").value || "";
    handleText(t, "", true);   // 手動入力は前後の文脈なし
  });
});

/* 選択テキスト＋「選択より前の本文」を取得して解析
 *   「同法」「法」「（同法）施行令／施行規則」の解決に、文書内の直前の法令名を用いる */
function analyzeSelection(fromButton) {
  // Word API が使える場合：選択前の本文も読み取り、直前の法令名を文書から探す
  if (typeof Word !== "undefined" && Word.run) {
    Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const selText = sel.text || "";

      // ① 選択開始位置より前の本文を取得
      let preceding = "";
      try {
        const before =
          context.document.body.getRange("Start").expandTo(sel.getRange("Start"));
        before.load("text");
        await context.sync();
        preceding = before.text || "";
      } catch (e) {
        preceding = "";
      }

      // ② 取れなかった場合は本文全体を文脈として使う（直前の法令名の特定用）
      if (!preceding) {
        try {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          let bodyText = body.text || "";
          // 選択箇所より後ろの法令名を拾わないよう、選択テキストの直前までに限定
          const idx = selText ? bodyText.indexOf(selText) : -1;
          preceding = idx > 0 ? bodyText.slice(0, idx) : bodyText;
        } catch (e2) {
          preceding = "";
        }
      }

      handleText(selText, preceding, fromButton);
    }).catch(() => {
      // フォールバック：共通APIで選択テキストのみ
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) handleText(res.value || "", "", fromButton);
        else setStatus("選択テキストを取得できませんでした。");
      });
    });
    return;
  }
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("選択テキストを取得できませんでした。");
      return;
    }
    handleText(res.value || "", "", fromButton);
  });
}

/* ===========================================================================
 * テキスト処理 → 解析 → 表示
 *   precedingText : 選択範囲より前の本文（直前の法令名の特定に使用）
 * ======================================================================== */
async function handleText(text, precedingText, fromButton) {
  const results = document.getElementById("results");
  text = (text || "").trim();

  // 選択解除（空選択）時は、直前の条文表示を維持する（次の条文を選択するまで消さない）
  if (!text) {
    if (fromButton) setStatus("条文参照を含む箇所を選択してから実行してください。");
    return;
  }

  const seedLaw = seedLawFromPreceding(precedingText || "");
  const refs = parseReferences(text, seedLaw);
  if (refs.length === 0) {
    // 法令参照を含まない選択：直前の表示は維持し、必要時のみヒントを出す
    if (/(?:同法|^法|[^一-龥]法)\s*(?:第\s*)?[0-9０-９一二三四五六七八九十百千]+\s*条/.test(text) && !seedLaw) {
      setStatus("「同法／法」の基準となる法令名が文書から見つかりませんでした。法令名を含めて選択してください。");
    } else if (fromButton) {
      setStatus("選択範囲から法令の条文参照を検出できませんでした。");
    }
    return;  // results は消さない
  }

  // 新たな解析を開始。これ以降、より新しい解析が始まったら本処理は中断する
  const mySeq = ++runSeq;

  setStatus(`${refs.length} 件の参照を検出。条文を取得中…`);
  results.innerHTML = "";  // 新たに条文を検出したときのみ置き換え

  for (const ref of refs) {
    if (mySeq !== runSeq) return;        // 後発の解析に取って代わられた → 中断
    const card = renderCardSkeleton(ref);
    results.appendChild(card);

    try {
      const lawId = await resolveLawId(ref.name);
      if (mySeq !== runSeq) return;       // await 後の失効チェック
      if (!lawId) {
        fillCardError(card, "法令IDを特定できませんでした（法令名の表記を確認してください）。");
        continue;
      }
      const body = await fetchArticle(lawId, ref.art, ref.branch, ref.para);
      if (mySeq !== runSeq) return;       // await 後の失効チェック
      fillCard(card, ref, lawId, body);
    } catch (e) {
      if (mySeq !== runSeq) return;
      fillCardError(card, "取得に失敗しました。通信状況をご確認のうえ、再度選択してお試しください（繰り返す場合は README のプロキシ設定をご確認ください）。");
    }
  }
  if (mySeq === runSeq) setStatus(`完了（${refs.length} 件）。`);
}

/* ===========================================================================
 * 参照解析（参照語 + 法令名 + 施行令/施行規則 + 条・項・号。表記ゆれ対応）
 *   group: 1=参照語(同法/法 等) 2=法律名 3=施行令/施行規則 4=条 5=枝番 6=項 7=号
 * ======================================================================== */
function buildRegex() {
  const NUM = "[0-9０-９一二三四五六七八九十百千〇零]+";
  // 参照語（長いものを先に）：「同法」「法」など → 直前の法令名を指す
  const REF = "当該法律の|当該法律|同法|本法|当該|同|本|法";
  // 法律名を構成する文字：第・条・項・号・数字（算用/全角/漢数字）は除外し、
  // 条番号を法律名として誤って飲み込まないようにする
  const NAMECH = "(?:(?![第条項号一二三四五六七八九十百千〇零])[一-龥々〆ヶ]|[・ーA-Za-z])";
  // 法律名（…法／…法律／…に関する法律／…条例／政令・省令等）。施行令/規則はここに含めない
  const LAWBASE =
    NAMECH + "{1,40}?(?:に関する法律|に関する法|法律|法|条例|政令|省令|府令|勅令)";
  // 施行令／施行規則
  const ENFORCE = "施行令|施行規則|施行細則";

  const PAT =
    `(?:(${REF})\\s*)?` +                     // 1: 参照語
    `(${LAWBASE})?` +                         // 2: 法律名
    `\\s*(${ENFORCE})?` +                     // 3: 施行令/施行規則
    `\\s*(?:第\\s*)?(${NUM})\\s*条` +          // 4: 条
    `(?:\\s*の\\s*(${NUM}))?` +               // 5: 枝番
    `(?:\\s*(?:第\\s*)?(${NUM})\\s*項)?` +     // 6: 項
    `(?:\\s*(?:第\\s*)?(${NUM})\\s*号)?`;      // 7: 号
  return new RegExp(PAT, "g");
}

function parseReferences(text, seedLaw) {
  const re = buildRegex();
  const out = [];
  let lastLaw = (seedLaw || "").trim();   // 直前の法令名（選択前の本文から引き継ぎ）
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index === re.lastIndex) re.lastIndex++; // 無限ループ防止

    const base = (m[2] || "").trim();        // 明示の法律名
    const enforce = (m[3] || "").trim();     // 施行令/施行規則
    const art = toNum(m[4]);
    let branch = m[5] ? String(toNum(m[5])) : "";
    const para = m[6] ? toNum(m[6]) : 0;
    const item = m[7] ? toNum(m[7]) : 0;

    // 検索する法令名を決定
    let name = "";
    if (base) {
      lastLaw = base;                                // 後続の「同法」用に基準法を更新
      name = enforce ? base + enforce : base;        // 明示の○○法（施行令/規則）
    } else if (enforce) {
      // 「施行令」「同施行令」「同法施行令」/ 規則 → 直前の法律名＋施行令/規則
      name = lastLaw ? lastLaw + enforce : "";
    } else {
      // 「同法」「法」「本法」「当該…」/ 名称なし → 直前の法律名
      name = lastLaw;
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

/* 選択範囲より前の本文から「直前の法令名」を推定
 *   1) 「○○法…第n条」の形で引用された直近の法律名（高精度）
 *   2) 無ければ、条を伴わずに登場した直近の法律名（誤検出語は除外） */
function seedLawFromPreceding(text) {
  if (!text) return "";
  let last = "";

  const re = buildRegex();
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index === re.lastIndex) re.lastIndex++;
    const base = (m[2] || "").trim();
    if (base) last = base;
  }
  if (last) return last;

  const re2 = /((?:(?![第条項号一二三四五六七八九十百千〇零])[一-龥々〆ヶ]|[・ーA-Za-z]){1,40}?(?:に関する法律|法律|法))/g;
  let m2;
  while ((m2 = re2.exec(text)) !== null) {
    const cand = m2[1].trim();
    const before = m2.index > 0 ? text.charAt(m2.index - 1) : "";
    if (DENY_NAMES.has(cand)) continue;
    // 法令番号（例：平成十七年法律第八十六号／令和6年法律第25号）の断片を除外
    if (cand === "法律" || /年法律$/.test(cand)) continue;
    if (/[0-9０-９第年]/.test(before)) continue;  // 直前が数字・第・年なら番号の一部とみなす
    last = cand;
  }
  return last;
}

/* 「○○法」の形だが法令名ではない一般語（誤検出の除外用） */
const DENY_NAMES = new Set([
  "方法", "用法", "手法", "文法", "作法", "寸法", "語法", "技法", "話法", "製法",
  "療法", "兵法", "魔法", "無法", "違法", "合法", "適法", "不法", "立法", "司法",
  "公法", "私法", "実体法", "手続法", "成文法", "不文法", "現行法", "旧法", "新法",
  "国内法", "国際法", "慣習法", "判例法", "自然法", "用法", "奏法", "書法", "算法"
]);

/* 中核法令の法令ID（同名包含が多く検索が不安定な法令の取り違えを防ぐため直指定）
 *  ・キーは「解決後の正式名称」。ALIASで正式名称に正規化された候補もここで一致する。
 *  ・施行令／施行規則などの派生法令は含めず、検索（searchLawIdByTitle）に委ねる。 */
const KNOWN_IDS = {
  "会社法": "417AC0000000086",
  "民法": "129AC0000000089",
  "商法": "132AC0000000048",
  "刑法": "140AC0000000045",
  "民事訴訟法": "408AC0000000109",
  "民事執行法": "354AC0000000004",
  "破産法": "416AC0000000075",
  "労働基準法": "322AC0000000049",
  "労働契約法": "419AC0000000128",
  "労働組合法": "324AC0000000174",
  "金融商品取引法": "323AC0000000025",
  "個人情報の保護に関する法律": "415AC0000000057",
  "行政手続法": "405AC0000000088",
  "道路交通法": "335AC0000000105",
  "特許法": "334AC0000000121",
  "実用新案法": "334AC0000000123",
  "意匠法": "334AC0000000125",
  "商標法": "334AC0000000127",
  "著作権法": "345AC0000000048",
  "不正競争防止法": "405AC0000000047",
  "下請代金支払遅延等防止法": "331AC0000000120",
  "私的独占の禁止及び公正取引の確保に関する法律": "322AC0000000054"
};

/* ===========================================================================
 * 法令ID 解決（KNOWN_IDS 直指定 → v2 /laws 検索）
 * ======================================================================== */
async function resolveLawId(name) {
  name = (name || "").trim();
  if (!name) return "";
  if (idCache.has(name)) return idCache.get(name);

  // 候補の正式名称リスト（ALIASの配列/文字列＋生の入力名）
  const mapped = ALIAS[name];
  let candidates = mapped ? (Array.isArray(mapped) ? mapped.slice() : [mapped]) : [];
  candidates.push(name);
  candidates = [...new Set(candidates)];

  let id = "";
  // ① 中核法令はID直指定を最優先（検索の取り違え防止）
  for (const cand of candidates) {
    if (KNOWN_IDS[cand]) { id = KNOWN_IDS[cand]; break; }
  }
  // ② それ以外は e-Gov 検索
  if (!id) {
    for (const cand of candidates) {
      id = await searchLawIdByTitle(cand);
      if (id) break;
    }
  }
  if (id) idCache.set(name, id);   // 成功時のみキャッシュ（失敗は再試行を許容）
  return id;
}

/* 法令名（正式名称）から e-Gov v2 /laws で法令IDを検索 */
async function searchLawIdByTitle(title) {
  title = (title || "").trim();
  if (!title) return "";
  if (idCache.has("T:" + title)) return idCache.get("T:" + title);

  const url = `${API_V2}/laws?law_title=${encodeURIComponent(title)}&limit=50`;
  const data = await apiFetchJson(url);  // 失敗時は例外を上位へ

  const titleOf = (l) =>
    (l.revision_info && l.revision_info.law_title) ||
    (l.law_info && l.law_info.law_title) || l.law_title || "";
  const idOf = (l) =>
    (l.law_info && l.law_info.law_id) || l.law_id || l.LawId || "";

  // q の派生法令（○○施行令／施行規則／施行細則／の施行に伴う…／の一部を改正…）か
  const isDerivative = (t) => {
    if (t === title) return false;
    if (!t.startsWith(title)) return false;
    const rest = t.slice(title.length);
    return /^(施行令|施行規則|施行細則|施行法|の施行|の一部)/.test(rest);
  };

  const laws = (data && (data.laws || data.Laws)) || [];
  let best = "", bestLen = Infinity;
  for (const l of laws) {
    const t = titleOf(l), id = idOf(l);
    if (!id) continue;
    if (t === title) { best = id; bestLen = 0; break; }            // 完全一致が最優先
    if (t.startsWith(title) && !isDerivative(t) && t.length < bestLen) {
      best = id; bestLen = t.length;                               // 前方一致（派生を除く）の最短
    }
  }
  // 前方一致も無ければ「包含」一致のうち最短（派生は除外。最後の手段）
  if (!best) {
    for (const l of laws) {
      const t = titleOf(l), id = idOf(l);
      if (id && !isDerivative(t) && t.indexOf(title) >= 0 && t.length < bestLen) {
        best = id; bestLen = t.length;
      }
    }
  }
  if (best) idCache.set("T:" + title, best);   // ヒット時のみキャッシュ
  return best;
}

/* ===========================================================================
 * 条文取得（v1 条文内容取得API。XML を解析）
 * ======================================================================== */
async function fetchArticle(lawId, art, branch, para) {
  const key = `${lawId}|${art}|${branch}|${para}`;
  if (articleCache.has(key)) return articleCache.get(key);

  // article パラメータは「数字」を第一候補に、「第○条」漢数字を予備として試行する。
  // （公式ドキュメントの例： article=1 ／ 仕様書の例： article=第十一条）
  const attempts = [];
  if (branch) {
    attempts.push(`${art}の${branch}`);
    attempts.push(`第${toKanji(art)}条の${toKanji(parseInt(branch, 10))}`);
  } else {
    attempts.push(`${art}`);
    attempts.push(`第${toKanji(art)}条`);
  }

  let body = "";
  let hadResponse = false;   // 正常に応答（HTTP 200）を受け取れたか
  let networkErr = false;    // 通信失敗があったか
  for (const ap of attempts) {
    const url = `${API_V1}/articles;lawId=${lawId};article=${encodeURIComponent(ap)}`;
    let xmlText;
    try {
      xmlText = await apiFetchText(url);
      hadResponse = true;
    } catch (e) {
      networkErr = true;
      continue; // この形式は通信失敗。次の候補へ
    }
    body = parseArticleBody(xmlText, para);
    if (body) break;
  }

  if (body) {
    articleCache.set(key, body);   // 成功時のみキャッシュ
    return body;
  }
  // 通信失敗が原因で空のときは「失敗を固定化」しないよう、キャッシュせず例外に
  if (networkErr && !hadResponse) throw new Error("条文取得に失敗（通信）");
  // 正常応答だが該当条文が無い場合も、念のためキャッシュせず（再試行を許容）
  return "";
}

/* 条文内容取得APIの応答XMLから本文を抽出
 *  ※ 応答には <Article> が2つ含まれる：
 *     (1) <ApplData> 直下の <Article>34</Article>（条番号のメタ情報・中身なし）
 *     (2) <LawContents> 配下の <Article Num="34">（条文本文）
 *     → (2) を選ぶ必要がある。 */
function parseArticleBody(xmlText, para) {
  try {
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (doc.getElementsByTagName("parsererror").length > 0) return "";

    const code = doc.getElementsByTagName("Code")[0];
    if (code && code.textContent.trim() !== "0") return "";

    // 本文の <Article> は LawContents 配下にある（無ければ全体から探索）
    const lc = doc.getElementsByTagName("LawContents")[0];
    const scope = lc || doc;

    // Num属性を持つ、または <Sentence> を含む <Article> を本文として採用
    let contentArticle = null;
    const arts = scope.getElementsByTagName("Article");
    for (let i = 0; i < arts.length; i++) {
      if (arts[i].getAttribute("Num") !== null ||
          arts[i].getElementsByTagName("Sentence").length > 0) {
        contentArticle = arts[i];
        break;
      }
    }
    if (!contentArticle) return "";

    // 表示範囲：項を指定 → その項すべて（号を含む）／ 条のみ → その条すべて
    if (para > 0) {
      const ps = childElems(contentArticle, "Paragraph");
      let target =
        ps.find((p) => p.getAttribute("Num") === String(para)) || ps[para - 1];
      if (target) {
        const lines = [];
        renderParagraph(target, lines);
        const t = lines.join("\n").trim();
        if (t) return t;
      }
      // 指定の項が見つからない場合は条全体にフォールバック
    }
    return renderArticle(contentArticle).trim();
  } catch (e) {
    return "";
  }
}

/* ---- XML 構造レンダリング（項番号・号番号・枝番号を保持） ---------------- */

// 条：見出し＋全項（各項に号・細分を含む）
function renderArticle(art) {
  const lines = [];
  const title = textOf(firstChild(art, "ArticleTitle"));
  const cap = textOf(firstChild(art, "ArticleCaption"));
  let head = title || "";
  if (cap) head += (head ? "　" : "") + cap;
  if (head) lines.push(head);
  for (const p of childElems(art, "Paragraph")) renderParagraph(p, lines);
  return lines.join("\n");
}

// 項：項番号（２、３…。第1項は番号表記なし）＋本文＋号
function renderParagraph(p, lines) {
  const num = textOf(firstChild(p, "ParagraphNum"));   // "２" 等。第1項は空
  const cap = textOf(firstChild(p, "ParagraphCaption"));
  const sent = sentenceText(firstChild(p, "ParagraphSentence"));
  let line = num ? num + "　" : "　";                  // 第1項は字下げのみ
  if (cap) line += cap + "　";
  line += sent;
  lines.push(line);
  for (const it of childElems(p, "Item")) renderItem(it, lines, 1);
}

// 号：号番号（一、二…）＋本文。イロハ等の細分は再帰
function renderItem(it, lines, depth) {
  const title = textOf(firstChild(it, "ItemTitle"));
  const sent = sentenceText(firstChild(it, "ItemSentence"));
  lines.push(indent(depth) + (title ? title + "　" : "") + sent);
  renderSubitems(it, lines, depth + 1, 1);
}

// 細分（Subitem1=イロハ、Subitem2=(1)(2)… など）を再帰表示
function renderSubitems(parent, lines, depth, level) {
  const tag = "Subitem" + level;
  for (const si of childElems(parent, tag)) {
    const title = textOf(firstChild(si, tag + "Title"));
    const sent = sentenceText(firstChild(si, tag + "Sentence"));
    lines.push(indent(depth) + (title ? title + "　" : "") + sent);
    renderSubitems(si, lines, depth + 1, level + 1);
  }
}

function indent(d) { return "　".repeat(Math.max(0, d)); }

// *Sentence ラッパ配下の <Sentence> を連結（Column=表項目は全角スペース区切り）
function sentenceText(wrapper) {
  if (!wrapper) return "";
  const sents = wrapper.getElementsByTagName("Sentence");
  const parts = [];
  for (let i = 0; i < sents.length; i++) {
    const t = (sents[i].textContent || "").trim();
    if (t) parts.push(t);
  }
  return parts.join(sents.length > 1 ? "　" : "");
}

/* 直下要素ヘルパ（getElementsByTagName は子孫まで拾うため、構造保持には直下走査を使う） */
function childElems(el, tag) {
  const out = [];
  if (!el) return out;
  const ch = el.children || [];
  for (let i = 0; i < ch.length; i++) if (ch[i].tagName === tag) out.push(ch[i]);
  return out;
}
function firstChild(el, tag) {
  const a = childElems(el, tag);
  return a.length ? a[0] : null;
}
function textOf(el) { return el ? (el.textContent || "").trim() : ""; }

/* ===========================================================================
 * 通信（PROXY 設定があれば経由）
 * ======================================================================== */
function withProxy(url) {
  return PROXY ? PROXY + encodeURIComponent(url) : url;
}
// タイムアウト付き fetch（応答待ちで固まらないように。既定10秒）
async function fetchWithTimeout(url, opts, ms) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms || 10000);
  try {
    return await fetch(url, Object.assign({ signal: ctrl.signal }, opts || {}));
  } finally {
    clearTimeout(timer);
  }
}
async function apiFetchJson(url) {
  const r = await fetchWithTimeout(withProxy(url), { headers: { "Accept": "application/json" } });
  if (!r.ok) throw new Error("HTTP " + r.status);
  return await r.json();
}
async function apiFetchText(url) {
  const r = await fetchWithTimeout(withProxy(url));
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
  // 該当条文へジャンプ：e-Gov のアンカー #Mp-At_{条数}（枝番は _{枝番}）
  const anchor = "#Mp-At_" + ref.art + (ref.branch ? "_" + ref.branch : "");
  const link = document.createElement("a");
  link.className = "egov";
  link.href = WEB + lawId + anchor;
  link.target = "_blank";
  link.rel = "noopener";
  link.textContent = "e-Gov で開く（該当条文へ） ↗";
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
