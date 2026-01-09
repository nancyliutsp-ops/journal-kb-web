let table;
let warningSet = new Set();

function normIssn(v) {
  if (v === null || v === undefined) return "";
  let s = String(v).trim().toUpperCase();
  if (!s || s === "NAN") return "";
  s = s.replace(/\s+/g, "");
  if (/^\d{7}[\dX]$/.test(s)) s = s.slice(0, 4) + "-" + s.slice(4);
  return s;
}

// 你Excel可能的列名（自动适配）
const COLS = {
  issn: ["ISSN","issn"],
  eissn: ["eISSN","EISSN","e-ISSN","E-ISSN","eissn"],
  region: ["REGION","Region","地区","region"],
  publisher: ["PUBLISHER","Publisher","出版商","publisher"],
  subject: ["Category","category","学科","学科分类","Subject","subject"],
  cas: ["中科院分区","CAS分区","CAS Partition","cas_partition","cas"],
  jifq: ["JIF Quartile","JIF四分位","JIF分区","Quartile","quartile"]
};

function pickCol(row, list) {
  for (const k of list) if (Object.prototype.hasOwnProperty.call(row, k)) return k;
  return null;
}

function casScore(cas) {
  if (!cas) return null;
  const s = String(cas).toLowerCase();
  if (s.includes("top")) return 1; // 当作1区
  const m = s.match(/[1-4]/);
  return m ? parseInt(m[0], 10) : null;
}
function jifScore(q) {
  if (!q) return null;
  const s = String(q).toUpperCase();
  if (s.includes("Q1")) return 1;
  if (s.includes("Q2")) return 2;
  if (s.includes("Q3")) return 3;
  if (s.includes("Q4")) return 4;
  return null;
}
function isMismatch(cas, q) {
  const cs = casScore(cas);
  const js = jifScore(q);
  if (cs === null || js === null) return false;
  return Math.abs(cs - js) >= 2;
}
function isInWarning(issn, eissn) {
  const a = normIssn(issn);
  const b = normIssn(eissn);
  return (a && warningSet.has(a)) || (b && warningSet.has(b));
}

async function safeFetch(path) {
  const res = await fetch(path, { cache: "no-store" });
  if (!res.ok) throw new Error(`${path} 读取失败（HTTP ${res.status}）`);
  return res;
}

async function loadWarningList() {
  try {
    const res = await fetch("data/cas_warning.json", { cache: "no-store" });
    if (!res.ok) return; // 没有这个文件也允许
    const arr = await res.json();
    warningSet = new Set((arr || []).map(normIssn).filter(Boolean));
  } catch {
    // ignore
  }
}

async function loadExcel() {
  const res = await safeFetch("data/journals.xlsx");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function ensureStatsPanels() {
  if (document.getElementById("statsWrap")) return;

  const wrap = document.createElement("div");
  wrap.id = "statsWrap";
  wrap.style.margin = "12px 0 14px";
  wrap.innerHTML = `
    <div style="display:flex; gap:12px; flex-wrap:wrap;">
      <div style="flex:1; min-width:260px; border:1px solid #eee; border-radius:10px; padding:12px;">
        <div style="font-weight:600; margin-bottom:8px;">按地区统计</div>
        <div id="statRegion"></div>
      </div>
      <div style="flex:1; min-width:260px; border:1px solid #eee; border-radius:10px; padding:12px;">
        <div style="font-weight:600; margin-bottom:8px;">按出版商统计</div>
        <div id="statPublisher"></div>
      </div>
      <div style="flex:1; min-width:260px; border:1px solid #eee; border-radius:10px; padding:12px;">
        <div style="font-weight:600; margin-bottom:8px;">按学科统计</div>
        <div id="statSubject"></div>
      </div>
    </div>
    <div style="margin:10px 0 0; display:flex; gap:10px; flex-wrap:wrap;">
      <button id="btnAll">显示全部</button>
      <button id="btnWarning">只看中科院预警</button>
      <button id="btnMismatch">只看分区差异</button>
      <button id="btnRisk">只看风险（预警或差异）</button>
    </div>
    <div style="color:#666; font-size:12px; margin-top:6px;">
      点击统计项可筛选表格；风险期刊会自动标红（“风险原因”列）。
    </div>
  `;

  const tableEl = document.getElementById("journalTable");
  tableEl.parentNode.insertBefore(wrap, tableEl);

  document.getElementById("btnAll").onclick = () => table.search("").draw();
  document.getElementById("btnWarning").onclick = () => table.search("中科院预警").draw();
  document.getElementById("btnMismatch").onclick = () => table.search("分区差异").draw();
  document.getElementById("btnRisk").onclick = () => table.search("中科院预警|分区差异", true, false).draw();
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, m => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;" }[m]));
}

function groupCount(rows, key) {
  const m = new Map();
  for (const r of rows) {
    const v = (r[key] ?? "").toString().trim() || "(空)";
    m.set(v, (m.get(v) || 0) + 1);
  }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]);
}

function renderStat(containerId, pairs, onClick, limit=12) {
  const el = document.getElementById(containerId);
  el.innerHTML = pairs.slice(0,limit).map(([k,c]) => `
    <div style="display:flex; justify-content:space-between; gap:8px; padding:4px 0; border-bottom:1px dashed #eee;">
      <a href="#" data-k="${encodeURIComponent(k)}" style="text-decoration:none;">${escapeHtml(k)}</a>
      <span>${c}</span>
    </div>
  `).join("") + (pairs.length>limit ? `<div style="color:#888;font-size:12px;padding-top:6px;">仅显示前${limit}项</div>` : "");

  el.querySelectorAll("a[data-k]").forEach(a=>{
    a.addEventListener("click",(ev)=>{
      ev.preventDefault();
      onClick(decodeURIComponent(a.getAttribute("data-k")));
    });
  });
}

function renderTable(rows) {
  if (!rows.length) return;

  const origCols = Object.keys(rows[0]).filter(k=>!k.startsWith("__"));
  const columns = [
    { title:"ID", data:"__id" },
    { title:"风险原因", data:"__risk" },
    { title:"ISSN(规范)", data:"__issn" },
    { title:"eISSN(规范)", data:"__eissn" },
    { title:"地区(统计)", data:"__region" },
    { title:"出版商(统计)", data:"__publisher" },
    { title:"学科(统计)", data:"__subject" },
    ...origCols.map(k=>({ title:k, data:k }))
  ];

  if (table) table.destroy();
  table = $("#journalTable").DataTable({
    data: rows,
    columns,
    pageLength: 25,
    deferRender: true,
    autoWidth: false,
    createdRow: (row, data) => {
      if (data.__risk) row.style.backgroundColor = "#fff1f2";
      if (data.__risk.includes("中科院预警")) row.querySelectorAll("td")[1].style.fontWeight = "700";
    }
  });
}

function enrich(rows) {
  const sample = rows[0] || {};
  const cIssn = pickCol(sample, COLS.issn);
  const cEissn = pickCol(sample, COLS.eissn);
  const cRegion = pickCol(sample, COLS.region);
  const cPublisher = pickCol(sample, COLS.publisher);
  const cSubject = pickCol(sample, COLS.subject);
  const cCas = pickCol(sample, COLS.cas);
  const cJifq = pickCol(sample, COLS.jifq);

  return rows.map((r, i) => {
    const issn = cIssn ? r[cIssn] : "";
    const eissn = cEissn ? r[cEissn] : "";
    const cas = cCas ? r[cCas] : "";
    const jifq = cJifq ? r[cJifq] : "";

    const reasons = [];
    if (isInWarning(issn, eissn)) reasons.push("中科院预警");
    if (isMismatch(cas, jifq)) reasons.push("分区差异");

    r.__id = i + 1;
    r.__issn = normIssn(issn);
    r.__eissn = normIssn(eissn);
    r.__region = cRegion ? r[cRegion] : "";
    r.__publisher = cPublisher ? r[cPublisher] : "";
    r.__subject = cSubject ? r[cSubject] : "";
    r.__risk = reasons.join("；");
    return r;
  });
}

async function init() {
  try {
    window.setLoadStatus?.("", "正在加载预警名单…");
    await loadWarningList();

    window.setLoadStatus?.("", "正在加载 Excel…");
    const raw = await loadExcel();
    if (!raw.length) {
      window.showError?.("Excel 解析结果为空：请确认 journals.xlsx 第一个工作表有数据。");
      return;
    }

    const rows = enrich(raw);

    ensureStatsPanels();
    const regionPairs = groupCount(rows, "__region");
    const pubPairs = groupCount(rows, "__publisher");
    const subPairs = groupCount(rows, "__subject");

    renderStat("statRegion", regionPairs, (k)=>table.search(k).draw());
    renderStat("statPublisher", pubPairs, (k)=>table.search(k).draw());
    renderStat("statSubject", subPairs, (k)=>table.search(k).draw());

    renderTable(rows);
    window.setLoadStatus?.("ok", `加载完成：${rows.length} 条`);
  } catch (e) {
    const msg = (e && e.message) ? e.message : String(e);
    window.showError?.(`
      <div><b>加载失败原因：</b>${escapeHtml(msg)}</div>
      <div style="margin-top:8px;">
        <div>请逐项检查：</div>
        <ol style="margin:6px 0 0 18px;">
          <li>浏览器打开 <code>/data/journals.xlsx</code> 是否能下载（否则说明文件未放对路径/大小写不一致）</li>
          <li>GitHub Pages 是否已启用（仓库 Settings → Pages：main / root）</li>
          <li>若你在内网/国内网络，CDN 可能被拦：我已换成 staticfile，但若仍失败我再给“本地化库文件”方案</li>
        </ol>
      </div>
    `);
    console.error(e);
  }
}

window.addEventListener("load", init);
