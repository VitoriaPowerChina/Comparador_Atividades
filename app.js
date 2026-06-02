// ════════════════════════════════════════════════════════════════════
//  STATE
// ════════════════════════════════════════════════════════════════════
const state = {
  baselineSheets: {}, subSheets: {},
  baselineRows: [], subRows: [],
  comparisonRows: [], blockComparisonRows: [],
  baselineSelected: new Set(), subSelected: new Set()
};

const $ = id => document.getElementById(id);

// ════════════════════════════════════════════════════════════════════
//  UTILS
// ════════════════════════════════════════════════════════════════════
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function normText(v) {
  return String(v ?? "").normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function dateToStr(d) {
  if (!(d instanceof Date) || isNaN(d)) return "—";
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
}

function daysDiff(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date) || isNaN(a) || isNaN(b)) return null;
  return Math.round(
    (Date.UTC(b.getFullYear(), b.getMonth(), b.getDate())
    - Date.UTC(a.getFullYear(), a.getMonth(), a.getDate())) / 86400000
  );
}

function parseDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number") {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return isNaN(d) ? null : d;
  }
  const t = String(v).trim();
  const m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    const dt = new Date(Number(y), Number(mo) - 1, Number(d));
    return isNaN(dt) ? null : dt;
  }
  const dt = new Date(t);
  return isNaN(dt) ? null : dt;
}

function getIndentLevel(text) {
  const s = String(text ?? "");
  return Math.round((s.length - s.trimStart().length) / 3);
}

function uniqueSorted(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "es", { sensitivity: "base" })
  );
}

function isPbBlock(t) {
  const n = normText(t);
  return /^pb[\-\s]?n[\-\s]?\d+/.test(n) || /^pb\s?\d+/.test(n) || /^power block\s?\d+/.test(n) || /^bess/.test(n);
}

function isCtBlock(t) {
  const n = normText(t);
  return /^ct\s*\d+$/.test(n);
}

function isBlock(t) {
  return isPbBlock(t) || isCtBlock(t);
}

// ════════════════════════════════════════════════════════════════════
//  DETECÇÃO DE COLUNAS
// ════════════════════════════════════════════════════════════════════
function detectColumns(rows) {
  const cols = rows.length ? Object.keys(rows[0]) : [];
  const pick = patterns => cols.find(c => patterns.some(p => normText(c).includes(normText(p)))) || "";
  return {
    activity: pick(["Task Name", "Nombre de tarea", "actividad", "activity", "task", "nombre", "name", "descripcion"]),
    pv:       pick(["pv", "block", "bloque", "frente", "codigo", "EDT", "wbs"]),
    start:    pick(["Comienzo de linea base", "Comienzo de l", "Comienzo", "Start", "Inicio", "fecha inicio", "comienzo"]),
    finish:   pick(["Fin de linea base", "Fin de l", "Fin", "Finish", "End", "fecha fin", "termino"]),
    hierarchy: pick(["EDT", "WBS", "outline", "estructura", "edt"])
  };
}

// ════════════════════════════════════════════════════════════════════
//  LEITURA DE ARQUIVO
// ════════════════════════════════════════════════════════════════════
function workbookToSheets(buf) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheets = {};
  wb.SheetNames.forEach(n => {
    sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "" });
  });
  return sheets;
}

function xmlProjectToSheets(text) {
  const xml = new DOMParser().parseFromString(text, "application/xml");
  const rows = [...xml.getElementsByTagName("Task")].map(t => ({
    UID: tc(t, "UID"), ID: tc(t, "ID"), Name: tc(t, "Name"),
    Start: tc(t, "Start"), Finish: tc(t, "Finish"),
    OutlineLevel: tc(t, "OutlineLevel"), WBS: tc(t, "WBS")
  })).filter(r => r.Name);
  return { ProjectXML: rows };
}

function tc(p, tag) {
  const e = p.getElementsByTagName(tag)[0];
  return e ? e.textContent : "";
}

async function loadFile(file, kind) {
  const statusEl = kind === "baseline" ? $("baselineStatus") : $("subStatus");
  statusEl.innerHTML = "Lendo…";
  try {
    const ext = file.name.toLowerCase().split(".").pop();
    let sheets = {};
    if (ext === "xlsx" || ext === "xls") {
      if (!window.XLSX) throw new Error("Biblioteca SheetJS não carregou (verifique a internet).");
      sheets = workbookToSheets(await file.arrayBuffer());
    } else if (ext === "xml") {
      sheets = xmlProjectToSheets(await file.text());
    } else {
      throw new Error("Formato não suportado.");
    }
    if (kind === "baseline") {
      state.baselineSheets = sheets;
      fillSelect($("baselineSheet"), Object.keys(sheets), "Escolha a aba");
    } else {
      state.subSheets = sheets;
      fillSelect($("subSheet"), Object.keys(sheets), "Escolha a aba");
    }
    statusEl.innerHTML = `<span class="oktxt">${escapeHtml(file.name)}</span> — ${Object.keys(sheets).length} aba(s) carregada(s).`;
    $("suggestBtn").disabled = false;
  } catch (err) {
    statusEl.innerHTML = `<span class="err">Erro: ${escapeHtml(err.message || String(err))}</span>`;
  }
}

function fillSelect(sel, items, ph) {
  sel.innerHTML = `<option value="">${ph}</option>`;
  items.forEach(i => {
    const o = document.createElement("option");
    o.value = o.textContent = i;
    sel.appendChild(o);
  });
}

// ════════════════════════════════════════════════════════════════════
//  PREPARAÇÃO DE LINHAS
//  - parentBlock: calculado em passagem linear (O(n)) com currentBlock tracker
//  - isSummary: marcado na mesma passagem via EDT e indentação
//  - activityKey: chave única = "bloco||nome" — evita colapsar atividades
//    homônimas de blocos diferentes (bug "tudo CT22")
// ════════════════════════════════════════════════════════════════════
function getPreparedRows(kind) {
  const rawRows = kind === "baseline" ? state.baselineRows : state.subRows;
  const p = kind === "baseline" ? "baseline" : "sub";
  const actCol    = $(p + "ActivityCol").value;
  const pvCol     = $(p + "PvCol").value;
  const startCol  = $(p + "StartCol").value;
  const finishCol = $(p + "FinishCol").value;
  const hierCol   = $(p + "HierarchyCol").value;

  if (!rawRows.length || !actCol) return [];

  // ── Passagem 1: montar objetos com parentBlock e activityKey ──────
  let currentBlock = "";
  const full = rawRows.map((r, idx) => {
    const actRaw     = String(r[actCol] ?? "");
    const actTrimmed = actRaw.trim();
    const level      = getIndentLevel(actRaw);
    const hierarchy  = String(hierCol ? (r[hierCol] ?? "") : "").trim();

    if (level <= 1 && isBlock(actTrimmed)) currentBlock = actTrimmed;

    const parentBlock  = level > 1 ? currentBlock : "";
    const activityKey  = parentBlock ? `${parentBlock}||${actTrimmed}` : actTrimmed;

    return {
      raw: r,
      activity: actRaw.trimEnd(),
      activityTrimmed: actTrimmed,
      activityKey,
      pv: String(r[pvCol] ?? "").trim(),
      hierarchy,
      indentLevel: level,
      start:  parseDate(r[startCol]),
      finish: parseDate(r[finishCol]),
      rowIdx: idx,
      parentBlock,
      isSummary: false
    };
  });

  // ── Passagem 2: marcar resumos ────────────────────────────────────
  const hierSet = new Set(full.map(r => r.hierarchy).filter(Boolean));
  full.forEach((r, i) => {
    if (r.indentLevel === 0) { r.isSummary = true; return; }
    if (r.indentLevel === 1) { r.isSummary = true; return; }
    if (r.hierarchy && hierSet.has(r.hierarchy)) {
      if ([...hierSet].some(h => h !== r.hierarchy && h.startsWith(r.hierarchy + "."))) {
        r.isSummary = true; return;
      }
    }
    if (i < full.length - 1 && full[i + 1].indentLevel > r.indentLevel) {
      r.isSummary = true;
    }
  });

  return full.filter(r => r.activityTrimmed && r.start && r.finish);
}

// ════════════════════════════════════════════════════════════════════
//  PREPARAÇÃO DE ABA
// ════════════════════════════════════════════════════════════════════
function prepareSheet(kind) {
  const sheetName = kind === "baseline" ? $("baselineSheet").value : $("subSheet").value;
  const rows = kind === "baseline"
    ? (state.baselineSheets[sheetName] || [])
    : (state.subSheets[sheetName] || []);
  const cols = rows.length ? Object.keys(rows[0]) : [];
  const detected = detectColumns(rows);
  const prefix = kind === "baseline" ? "baseline" : "sub";

  if (kind === "baseline") state.baselineRows = rows;
  else state.subRows = rows;

  const colKeys  = ["ActivityCol", "PvCol", "StartCol", "FinishCol", "HierarchyCol"];
  const detKeys  = ["activity", "pv", "start", "finish", "hierarchy"];
  colKeys.forEach((suffix, i) => {
    const sel = $(prefix + suffix);
    const cur = sel.options[0]?.textContent || "";
    fillSelect(sel, cols, cur);
    sel.value = detected[detKeys[i]] || "";
  });
  refreshActivities();
}

// ════════════════════════════════════════════════════════════════════
//  CHECKLISTS
// ════════════════════════════════════════════════════════════════════
function getVisibleRows(kind) {
  const allRows = getPreparedRows(kind);
  const hideSum = $("hideSummaryRows").checked;
  const blockF  = kind === "baseline" ? $("baselineBlockSel").value : $("subBlockSel").value;
  return allRows.filter(r => {
    if (hideSum && r.isSummary) return false;
    if (blockF && r.parentBlock !== blockF) return false;
    return true;
  });
}

function renderChecklist(kind) {
  const box      = kind === "baseline" ? $("baselineChecklist") : $("subChecklist");
  const search   = normText(kind === "baseline" ? $("baselineSearch").value : $("subSearch").value);
  const selected = kind === "baseline" ? state.baselineSelected : state.subSelected;

  const visible = getVisibleRows(kind)
    .filter(r => !search || normText(r.activityTrimmed).includes(search));

  // Dedup por activityKey (bloco||nome) — preserva a primeira ocorrência
  const seen = new Set();
  const uniqRows = visible.filter(r => seen.has(r.activityKey) ? false : seen.add(r.activityKey));

  if (!uniqRows.length) {
    box.innerHTML = `<div class="checkitem"><span class="small">Nenhuma atividade encontrada.</span></div>`;
    return;
  }

  // Ordena por bloco (numérico: CT2 < CT10) depois por nome
  uniqRows.sort((a, b) => {
    const bp = a.parentBlock.localeCompare(b.parentBlock, "es", { sensitivity: "base", numeric: true });
    return bp !== 0 ? bp : a.activityTrimmed.localeCompare(b.activityTrimmed, "es", { sensitivity: "base", numeric: true });
  });

  box.innerHTML = uniqRows.map(r => {
    const checked = selected.has(r.activityKey) ? "checked" : "";
    return `<label class="checkitem">
      <input type="checkbox" data-kind="${kind}" value="${escapeHtml(r.activityKey)}" ${checked}/>
      ${r.parentBlock ? `<span class="block-badge">${escapeHtml(r.parentBlock)}</span>` : ""}
      <span>${escapeHtml(r.activityTrimmed)}</span>
    </label>`;
  }).join("");

  box.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener("change", () => {
      const set = cb.dataset.kind === "baseline" ? state.baselineSelected : state.subSelected;
      if (cb.checked) set.add(cb.value); else set.delete(cb.value);
      updateCounts();
    });
  });
}

function updateCounts() {
  $("baselineCount").textContent = `${state.baselineSelected.size} sel.`;
  $("subCount").textContent      = `${state.subSelected.size} sel.`;
  $("compareBtn").disabled       = !(state.baselineSelected.size && state.subSelected.size);
}

function refreshActivities() {
  const baseAll = getPreparedRows("baseline");
  const subAll  = getPreparedRows("sub");

  const bBlocks = uniqueSorted([...new Set(
    baseAll.filter(r => r.indentLevel <= 1 && isPbBlock(r.activityTrimmed)).map(r => r.activityTrimmed)
  )]);
  const sBlocks = uniqueSorted([...new Set(
    subAll.filter(r => r.indentLevel <= 1 && isCtBlock(r.activityTrimmed)).map(r => r.activityTrimmed)
  )]);

  fillSelect($("baselineBlockSel"), bBlocks, "Todos os blocos PB-N");
  fillSelect($("subBlockSel"), sBlocks, "Todos os CTs");

  const baseKeys = new Set(baseAll.map(r => r.activityKey));
  const subKeys  = new Set(subAll.map(r => r.activityKey));
  state.baselineSelected = new Set([...state.baselineSelected].filter(x => baseKeys.has(x)));
  state.subSelected      = new Set([...state.subSelected].filter(x => subKeys.has(x)));

  renderChecklist("baseline");
  renderChecklist("sub");
  updateCounts();

  $("blockInfoBox").innerHTML = `
    <div><strong style="color:var(--cyan)">Linha base:</strong> ${bBlocks.length} blocos PB-N detectados</div>
    <div style="margin-top:4px"><strong style="color:var(--emerald)">Subcontrata:</strong> ${sBlocks.length} CTs detectados</div>
    ${bBlocks.length === 0 && subAll.length === 0
      ? `<div class="err" style="margin-top:6px">Nenhum bloco encontrado — verifique a coluna de atividade.</div>`
      : ""}
  `;
  $("compareBlockBtn").disabled   = !(bBlocks.length || sBlocks.length);
  $("compareSuggestBtn").disabled = !(baseAll.length && subAll.length);
}

// ════════════════════════════════════════════════════════════════════
//  SIMILARIDADE / ALIASES
// ════════════════════════════════════════════════════════════════════
function similarity(a, b) {
  const sa = new Set(normText(a).split(" ").filter(Boolean));
  const sb = new Set(normText(b).split(" ").filter(Boolean));
  if (!sa.size || !sb.size) return 0;
  let inter = 0;
  for (const x of sa) if (sb.has(x)) inter++;
  return inter / new Set([...sa, ...sb]).size;
}

function aliasScore(a, b) {
  const na = normText(a), nb = normText(b);
  const aliases = [
    ["excavacion trincheras",        "trenches excavation"],
    ["excavacion de trincheras",     "trench excavation"],
    ["tendido cable solar",          "solar cable laying"],
    ["tendido de cable",             "cable laying"],
    ["tendido cable",                "cable laying"],
    ["instalacion malla de tierra",  "grounding cables"],
    ["malla de tierra",              "grounding cables"],
    ["canalizacion hdpe",            "hdpe ducts"],
    ["ductos hdpe",                  "hdpe ducts"],
    ["tapado",                       "backfill"],
    ["relleno",                      "backfill"],
    ["backfilling",                  "backfill"],
    ["conexion",                     "connection"],
    ["conexiones",                   "connections"],
    ["cable connection",             "conexiones"],
    ["obras civiles",                "civil activities"],
    ["civil works",                  "civil activities"],
    ["obras electricas",             "electrical activities"],
    ["electrical works",             "electrical activities"],
    ["cableado",                     "cable laying"],
    ["instalacion de cable",         "cable laying"],
    ["camara",                       "chamber"],
    ["camaras",                      "chambers"],
    ["pozos",                        "chambers"],
    ["segregacion de material",      "material segregation"],
    ["trazado de trinchera",         "trench layout"],
    ["fiber optic",                  "fibra optica"],
    ["baja tension",                 "low voltage"],
    ["lv cable",                     "tendido cable baja tension"]
  ];
  for (const [x, y] of aliases) {
    if ((na.includes(x) && nb.includes(y)) || (na.includes(y) && nb.includes(x))) return 0.95;
  }
  return similarity(a, b);
}

// ════════════════════════════════════════════════════════════════════
//  COMPARAÇÃO DE ATIVIDADES
// ════════════════════════════════════════════════════════════════════
function compareBySequence(base, sub) {
  const a = [...base].sort((x, y) => (x.start - y.start) || (x.finish - y.finish));
  const b = [...sub].sort((x, y) => (x.start - y.start) || (x.finish - y.finish));
  const n = Math.max(a.length, b.length);
  return Array.from({ length: n }, (_, i) => makeActivityRow(i + 1, a[i] ?? null, b[i] ?? null));
}

function compareByPv(base, sub) {
  const mapA = new Map(base.map(x => [normText(x.pv || x.activityTrimmed + "|" + dateToStr(x.start)), x]));
  const mapB = new Map(sub.map(x  => [normText(x.pv || x.activityTrimmed + "|" + dateToStr(x.start)), x]));
  const keys = [...new Set([...mapA.keys(), ...mapB.keys()])];
  return keys.map((k, i) => makeActivityRow(i + 1, mapA.get(k) ?? null, mapB.get(k) ?? null));
}

function makeActivityRow(order, a, b) {
  const diffStart  = a?.start  && b?.start  ? daysDiff(a.start,  b.start)  : null;
  const diffFinish = a?.finish && b?.finish ? daysDiff(a.finish, b.finish) : null;
  const obs = [];
  if (!a) obs.push("sem item na LB");
  if (!b) obs.push("sem item na SUB");
  if (diffStart  != null) obs.push(diffStart  > 0 ? "SUB começa depois"   : diffStart  < 0 ? "SUB começa antes"   : "mesmo início");
  if (diffFinish != null) obs.push(diffFinish > 0 ? "SUB termina depois"  : diffFinish < 0 ? "SUB termina antes"  : "mesmo fim");
  return {
    ordem: order,
    blockBase: a?.parentBlock || "", actBase: a?.activityTrimmed || "",
    iniBase: a?.start || null, fimBase: a?.finish || null,
    blockSub: b?.parentBlock  || "", actSub:  b?.activityTrimmed || "",
    iniSub:  b?.start || null, fimSub:  b?.finish || null,
    diffStart, diffFinish, obs: obs.join("; ")
  };
}

function renderComparison(rows) {
  state.comparisonRows = rows;
  $("comparisonBody").innerHTML = rows.length
    ? rows.map(r => `
      <tr>
        <td>${r.ordem}</td>
        <td>${r.blockBase ? `<span class="pill" style="color:var(--cyan)">${escapeHtml(r.blockBase)}</span>` : "—"}</td>
        <td>${escapeHtml(r.actBase  || "—")}</td>
        <td>${escapeHtml(dateToStr(r.iniBase))}</td>
        <td>${escapeHtml(dateToStr(r.fimBase))}</td>
        <td>${r.blockSub ? `<span class="pill" style="color:var(--emerald)">${escapeHtml(r.blockSub)}</span>` : "—"}</td>
        <td>${escapeHtml(r.actSub  || "—")}</td>
        <td>${escapeHtml(dateToStr(r.iniSub))}</td>
        <td>${escapeHtml(dateToStr(r.fimSub))}</td>
        <td>${delayBadge(r.diffStart)}</td>
        <td>${delayBadge(r.diffFinish)}</td>
        <td>${escapeHtml(r.obs || "—")}</td>
      </tr>`).join("")
    : `<tr><td colspan="12" class="small" style="padding:20px;text-align:center">Nenhum registro encontrado.</td></tr>`;
  switchTab("activities");
}

function delayBadge(d) {
  if (d == null) return `<span style="color:#475569">—</span>`;
  const cls = d > 0 ? "bad" : d < 0 ? "good" : "neutral";
  return `<span class="delay ${cls}">${d > 0 ? "+" : ""}${d}d</span>`;
}

function runComparison() {
  if (!state.baselineSelected.size || !state.subSelected.size) {
    alert("Marque pelo menos uma atividade de cada lado."); return;
  }
  const baseRows = getPreparedRows("baseline").filter(r => state.baselineSelected.has(r.activityKey));
  const subRows  = getPreparedRows("sub").filter(r => state.subSelected.has(r.activityKey));
  if (!baseRows.length || !subRows.length) {
    alert("As atividades selecionadas não foram encontradas.\nVerifique a aba e as colunas."); return;
  }
  const rows = $("compareMode").value === "pv"
    ? compareByPv(baseRows, subRows)
    : compareBySequence(baseRows, subRows);
  renderComparison(rows);
  updateCards(baseRows, subRows);
  updateParallel(baseRows, subRows);
  updateReport(baseRows, subRows);
  $("exportBtn").disabled = false;
  $("tagBox").innerHTML = `
    <div class="tag">LB: ${state.baselineSelected.size === 1 ? [...state.baselineSelected][0] : state.baselineSelected.size + " atividades"}</div>
    <div class="tag">SUB: ${state.subSelected.size === 1 ? [...state.subSelected][0] : state.subSelected.size + " atividades"}</div>
    <div class="tag warn">${escapeHtml($("compareMode").selectedOptions[0].textContent)}</div>
  `;
  $("comparisonSubtitle").textContent =
    `Comparação de atividades — modo ${$("compareMode").selectedOptions[0].textContent.toLowerCase()}.`;
}

// ════════════════════════════════════════════════════════════════════
//  COMPARAÇÃO POR BLOCO
// ════════════════════════════════════════════════════════════════════
function buildBlockMap(allRows) {
  const map = new Map();
  for (const r of allRows) {
    if (r.indentLevel <= 1 && isBlock(r.activityTrimmed)) {
      if (!map.has(r.activityTrimmed)) {
        map.set(r.activityTrimmed, { name: r.activityTrimmed, start: r.start, finish: r.finish, activities: [] });
      }
    } else if (r.parentBlock) {
      if (!map.has(r.parentBlock)) {
        map.set(r.parentBlock, { name: r.parentBlock, start: null, finish: null, activities: [] });
      }
      map.get(r.parentBlock).activities.push(r);
    }
  }
  for (const blk of map.values()) {
    const acts = blk.activities.filter(r => r.start && r.finish);
    if (acts.length) {
      blk.start  = new Date(Math.min(...acts.map(r => +r.start)));
      blk.finish = new Date(Math.max(...acts.map(r => +r.finish)));
    }
  }
  return [...map.values()].filter(b => b.start && b.finish);
}

function compareByBlock() {
  const baseAll    = getPreparedRows("baseline");
  const subAll     = getPreparedRows("sub");
  const baseBlocks = buildBlockMap(baseAll).sort((a, b) => a.start - b.start);
  const subBlocks  = buildBlockMap(subAll).sort((a, b) => a.start - b.start);

  if (!baseBlocks.length && !subBlocks.length) {
    alert("Nenhum bloco (PB-N / CT) detectado.\nVerifique se a coluna de atividade está correta.");
    return;
  }

  const n = Math.max(baseBlocks.length, subBlocks.length);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const a = baseBlocks[i] ?? null;
    const b = subBlocks[i]  ?? null;
    const durA = a?.start && a?.finish ? daysDiff(a.start, a.finish) : null;
    const durB = b?.start && b?.finish ? daysDiff(b.start, b.finish) : null;
    rows.push({
      i: i + 1,
      nameA: a?.name || "", startA: a?.start || null, finishA: a?.finish || null, durA,
      nameB: b?.name || "", startB: b?.start || null, finishB: b?.finish || null, durB,
      diffStart:  a?.start  && b?.start  ? daysDiff(a.start,  b.start)  : null,
      diffFinish: a?.finish && b?.finish ? daysDiff(a.finish, b.finish) : null,
      diffDur: durA != null && durB != null ? durB - durA : null
    });
  }
  state.blockComparisonRows = rows;

  $("blockComparisonBody").innerHTML = rows.map(r => `
    <tr>
      <td>${r.i}</td>
      <td><strong style="color:var(--cyan)">${escapeHtml(r.nameA || "—")}</strong></td>
      <td>${escapeHtml(dateToStr(r.startA))}</td>
      <td>${escapeHtml(dateToStr(r.finishA))}</td>
      <td>${r.durA != null ? r.durA + "d" : "—"}</td>
      <td><strong style="color:var(--emerald)">${escapeHtml(r.nameB || "—")}</strong></td>
      <td>${escapeHtml(dateToStr(r.startB))}</td>
      <td>${escapeHtml(dateToStr(r.finishB))}</td>
      <td>${r.durB != null ? r.durB + "d" : "—"}</td>
      <td>${delayBadge(r.diffStart)}</td>
      <td>${delayBadge(r.diffFinish)}</td>
      <td>${delayBadge(r.diffDur)}</td>
    </tr>`).join("");

  renderGantt(rows);

  const allBaseActs = baseBlocks.flatMap(b => b.activities);
  const allSubActs  = subBlocks.flatMap(b => b.activities);
  updateCards(allBaseActs, allSubActs);
  updateParallel(allBaseActs, allSubActs);
  updateReport(allBaseActs, allSubActs);

  $("tagBox").innerHTML = `
    <div class="tag">LB: ${baseBlocks.length} blocos PB-N</div>
    <div class="tag">SUB: ${subBlocks.length} CTs</div>
    <div class="tag warn">Comparação por bloco</div>
  `;
  $("comparisonSubtitle").textContent =
    `${baseBlocks.length} blocos PB-N (LB) vs ${subBlocks.length} CTs (SUB), ordenados por data de início.`;
  $("exportBtn").disabled = false;
  switchTab("blocks");
}

// ════════════════════════════════════════════════════════════════════
//  GANTT
// ════════════════════════════════════════════════════════════════════
function renderGantt(rows) {
  if (!rows.length) { $("ganttContent").innerHTML = `<div class="small">Sem dados.</div>`; return; }
  const allDates = rows.flatMap(r => [r.startA, r.finishA, r.startB, r.finishB].filter(Boolean));
  if (!allDates.length) return;
  const gMin  = new Date(Math.min(...allDates));
  const gMax  = new Date(Math.max(...allDates));
  const total = Math.max(1, daysDiff(gMin, gMax));
  const pL = (d, ref) => d ? Math.max(0, (daysDiff(ref, d) / total * 100)) : 0;
  const pW = (s, f)   => s && f ? Math.max(0.3, (daysDiff(s, f) / total * 100)) : 0;

  $("ganttContent").innerHTML = rows.map(r => `
    <div class="gantt-block">
      <div class="g-label">
        <span style="color:var(--cyan);font-size:11px">${escapeHtml(r.nameA || "—")}</span>
        <span style="color:var(--emerald);font-size:11px">${escapeHtml(r.nameB || "—")}</span>
      </div>
      <div class="gantt-track">
        ${r.startA ? `<div class="gantt-bar base" style="left:${pL(r.startA, gMin).toFixed(1)}%;width:${pW(r.startA, r.finishA).toFixed(1)}%"></div>` : ""}
        ${r.startB ? `<div class="gantt-bar sub"  style="left:${pL(r.startB, gMin).toFixed(1)}%;width:${pW(r.startB, r.finishB).toFixed(1)}%"></div>` : ""}
      </div>
    </div>`).join("");
}

// ════════════════════════════════════════════════════════════════════
//  CARDS / TIMELINE / PARECER / PARALELO
// ════════════════════════════════════════════════════════════════════
function peakParallel(rows) {
  const ev = [];
  rows.forEach(r => {
    if (!r.start || !r.finish) return;
    const s = Date.UTC(r.start.getFullYear(),  r.start.getMonth(),  r.start.getDate());
    const f = Date.UTC(r.finish.getFullYear(), r.finish.getMonth(), r.finish.getDate()) + 86400000;
    ev.push({ t: s, d: 1 }, { t: f, d: -1 });
  });
  if (!ev.length) return 0;
  ev.sort((a, b) => a.t - b.t || b.d - a.d);
  let cur = 0, max = 0;
  for (const e of ev) { cur += e.d; if (cur > max) max = cur; }
  return max;
}

function updateCards(base, sub) {
  const vb = base.filter(r => r.start && r.finish);
  const vs = sub.filter(r  => r.start && r.finish);
  const minBase = vb.length ? new Date(Math.min(...vb.map(x => +x.start)))  : null;
  const maxBase = vb.length ? new Date(Math.max(...vb.map(x => +x.finish))) : null;
  const minSub  = vs.length ? new Date(Math.min(...vs.map(x => +x.start)))  : null;
  const maxSub  = vs.length ? new Date(Math.max(...vs.map(x => +x.finish))) : null;
  $("cardBaseStart").textContent = dateToStr(minBase);
  $("cardSubStart").textContent  = dateToStr(minSub);
  $("cardBaseEnd").textContent   = dateToStr(maxBase);
  $("cardSubEnd").textContent    = dateToStr(maxSub);
  $("cardBaseStartSub").textContent = $("cardBaseEndSub").textContent = `${vb.length} atividades`;
  $("cardSubStartSub").textContent  = $("cardSubEndSub").textContent  = `${vs.length} atividades`;
  $("baseTimelineLabel").textContent = `${dateToStr(minBase)} → ${dateToStr(maxBase)}`;
  $("subTimelineLabel").textContent  = `${dateToStr(minSub)} → ${dateToStr(maxSub)}`;
  if (minBase && maxBase && minSub && maxSub) {
    const gMin  = new Date(Math.min(+minBase, +minSub));
    const gMax  = new Date(Math.max(+maxBase, +maxSub));
    const total = Math.max(1, daysDiff(gMin, gMax));
    const bOff  = Math.max(0, daysDiff(gMin, minBase));
    const bW    = Math.max(0.5, daysDiff(minBase, maxBase));
    const sOff  = Math.max(0, daysDiff(gMin, minSub));
    const sW    = Math.max(0.5, daysDiff(minSub, maxSub));
    $("baseTimelineFill").style.cssText = `left:${(bOff/total*100).toFixed(1)}%;width:${(bW/total*100).toFixed(1)}%;height:100%;background:rgba(34,211,238,.75);border-radius:999px;position:absolute`;
    $("subTimelineFill").style.cssText  = `left:${(sOff/total*100).toFixed(1)}%;width:${(sW/total*100).toFixed(1)}%;height:100%;background:rgba(52,211,153,.78);border-radius:999px;position:absolute`;
    const oStart  = new Date(Math.max(+minBase, +minSub));
    const oEnd    = new Date(Math.min(+maxBase, +maxSub));
    const overlap = daysDiff(oStart, oEnd);
    $("overlapInfo").innerHTML = overlap > 0
      ? `<span class="pill" style="color:var(--amber)">Sobreposição de períodos: ${overlap} dias</span>`
      : `<span class="pill" style="color:#475569">Sem sobreposição de períodos</span>`;
  }
}

function updateParallel(base, sub) {
  const pb = peakParallel(base), ps = peakParallel(sub);
  $("parallelBox").innerHTML = `
    <div class="peak-box">
      <div><div style="font-size:13px">Linha base</div><div class="small">pico de frentes paralelas</div></div>
      <div style="font-size:28px;font-weight:800;color:var(--cyan)">${pb}</div>
    </div>
    <div class="peak-box" style="margin-bottom:0">
      <div><div style="font-size:13px">Subcontrata</div><div class="small">pico de frentes paralelas</div></div>
      <div style="font-size:28px;font-weight:800;color:var(--emerald)">${ps}</div>
    </div>`;
}

function updateReport(base, sub) {
  const vb = base.filter(r => r.start && r.finish);
  const vs = sub.filter(r  => r.start && r.finish);
  if (!vb.length && !vs.length) { $("reportBox").innerHTML = "Sem dados suficientes."; return; }
  const minBase = vb.length ? new Date(Math.min(...vb.map(x => +x.start)))  : null;
  const maxBase = vb.length ? new Date(Math.max(...vb.map(x => +x.finish))) : null;
  const minSub  = vs.length ? new Date(Math.min(...vs.map(x => +x.start)))  : null;
  const maxSub  = vs.length ? new Date(Math.max(...vs.map(x => +x.finish))) : null;
  const ds = minBase && minSub ? daysDiff(minBase, minSub) : null;
  const df = maxBase && maxSub ? daysDiff(maxBase, maxSub) : null;
  const durBase = minBase && maxBase ? daysDiff(minBase, maxBase) : null;
  const durSub  = minSub  && maxSub  ? daysDiff(minSub,  maxSub)  : null;
  const startTxt = ds == null ? "início não mensurável"
    : ds > 0 ? `<strong class="a">${ds} dias depois</strong> da LB`
    : ds < 0 ? `<strong class="e">${Math.abs(ds)} dias antes</strong> da LB`
    : `<strong class="e">mesma data da LB</strong>`;
  const finTxt = df == null ? "fim não mensurável"
    : df > 0 ? `<strong class="a">${df} dias depois</strong> da LB`
    : df < 0 ? `<strong class="e">${Math.abs(df)} dias antes</strong> da LB`
    : `<strong class="e">mesma data da LB</strong>`;
  const verdict = (ds ?? 0) > 0 || (df ?? 0) > 0
    ? `<strong class="a">⚠ O cronograma da subcontrata precisa de ajuste.</strong>`
    : `<strong class="e">✓ O cronograma está alinhado com a linha base.</strong>`;
  $("reportBox").innerHTML = `
    <p><strong class="c">Linha base:</strong> ${dateToStr(minBase)} → ${dateToStr(maxBase)}${durBase != null ? ` (${durBase} dias)` : ""} · pico de ${peakParallel(vb)} frentes</p>
    <p><strong class="e">Subcontrata:</strong> ${dateToStr(minSub)} → ${dateToStr(maxSub)}${durSub != null ? ` (${durSub} dias)` : ""} · pico de ${peakParallel(vs)} frentes</p>
    <p><strong class="a">Início SUB:</strong> ${startTxt}</p>
    <p><strong class="a">Fim SUB:</strong> ${finTxt}</p>
    <p style="margin-top:8px">${verdict}</p>`;
}

// ════════════════════════════════════════════════════════════════════
//  SUGESTÃO DE MATCH
// ════════════════════════════════════════════════════════════════════
function suggestMatch() {
  const baseActs = getVisibleRows("baseline").map(r => r.activityTrimmed);
  const subActs  = getVisibleRows("sub").map(r => r.activityTrimmed);
  if (!baseActs.length || !subActs.length) {
    $("suggestionBox").textContent = "Carregue e leia as abas primeiro.";
    switchTab("suggest"); return;
  }
  const sug = baseActs.map(lb => {
    let best = "", bs = -1;
    for (const s of subActs) { const sc = aliasScore(lb, s); if (sc > bs) { best = s; bs = sc; } }
    return { lb, sub: best, score: bs };
  }).filter(x => x.score >= 0.2).sort((a, b) => b.score - a.score).slice(0, 20);

  $("suggestionBox").innerHTML = sug.length
    ? `<div style="margin-bottom:8px;font-weight:600">Top ${sug.length} correspondências sugeridas:</div>
       ${sug.map(s => `
       <div style="border:1px solid var(--line);background:#020617;border-radius:10px;padding:8px 10px;margin-bottom:6px">
         <div style="font-size:12px">
           <span class="pill" style="color:var(--cyan)">${escapeHtml(s.lb)}</span>
           <span style="color:var(--cyan);margin:0 6px">↔</span>
           <span class="pill" style="color:var(--emerald)">${escapeHtml(s.sub)}</span>
         </div>
         <div class="small" style="margin-top:4px;font-size:11px">Confiança: ${(s.score * 100).toFixed(0)}%</div>
       </div>`).join("")}`
    : "Nenhuma sugestão encontrada (score mínimo 20%).";

  if (state.baselineSelected.size) {
    const subRows = getPreparedRows("sub");
    for (const key of state.baselineSelected) {
      const actName = key.includes("||") ? key.split("||")[1] : key;
      let bestKey = "", bs = -1;
      for (const r of subRows) {
        const sc = aliasScore(actName, r.activityTrimmed);
        if (sc > bs) { bestKey = r.activityKey; bs = sc; }
      }
      if (bs >= 0.2 && bestKey) state.subSelected.add(bestKey);
    }
    renderChecklist("sub"); updateCounts();
  }
  switchTab("suggest");
}

// ════════════════════════════════════════════════════════════════════
//  EXPORT CSV
// ════════════════════════════════════════════════════════════════════
function exportCsv() {
  const useBlock = state.comparisonRows.length === 0 && state.blockComparisonRows.length > 0;
  const rows = useBlock ? state.blockComparisonRows : state.comparisonRows;
  if (!rows.length) { alert("Sem dados para exportar."); return; }
  const header = useBlock
    ? ["#","bloco_lb","inicio_lb","fim_lb","dur_lb_d","bloco_sub","inicio_sub","fim_sub","dur_sub_d","delta_inicio_d","delta_fim_d","delta_dur_d"]
    : ["#","bloco_lb","atividade_lb","inicio_lb","fim_lb","bloco_sub","atividade_sub","inicio_sub","fim_sub","delta_inicio_d","delta_fim_d","observacao"];
  const vals = r => useBlock
    ? [r.i,r.nameA,dateToStr(r.startA),dateToStr(r.finishA),r.durA??"",r.nameB,dateToStr(r.startB),dateToStr(r.finishB),r.durB??"",r.diffStart??"",r.diffFinish??"",r.diffDur??""]
    : [r.ordem,r.blockBase,r.actBase,dateToStr(r.iniBase),dateToStr(r.fimBase),r.blockSub,r.actSub,dateToStr(r.iniSub),dateToStr(r.fimSub),r.diffStart??"",r.diffFinish??"",r.obs];
  const csv = [header, ...rows.map(r => vals(r).map(v => `"${String(v ?? "").replaceAll('"', '""')}"`))].map(l => l.join(",")).join("\n");
  const url = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" }));
  const a = Object.assign(document.createElement("a"), { href: url, download: "comparacao_cronogramas.csv" });
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

// ════════════════════════════════════════════════════════════════════
//  BULK SELECT / TABS
// ════════════════════════════════════════════════════════════════════
function bulkSelect(kind, mode) {
  const q = normText(kind === "baseline" ? $("baselineSearch").value : $("subSearch").value);
  const seen = new Set();
  const visible = getVisibleRows(kind)
    .filter(r => !q || normText(r.activityTrimmed).includes(q))
    .filter(r => seen.has(r.activityKey) ? false : seen.add(r.activityKey))
    .map(r => r.activityKey);
  const set = kind === "baseline" ? state.baselineSelected : state.subSelected;
  if (mode === "all")   visible.forEach(x => set.add(x));
  if (mode === "clear") visible.forEach(x => set.delete(x));
  renderChecklist(kind); updateCounts();
}

function switchTab(tab) {
  ["activities", "blocks", "gantt", "suggest"].forEach(t => {
    const cap = t.charAt(0).toUpperCase() + t.slice(1);
    const v = $("view" + cap), tb = $("tab" + cap);
    if (v)  v.style.display = t === tab ? "block" : "none";
    if (tb) tb.classList.toggle("active", t === tab);
  });
}

// ════════════════════════════════════════════════════════════════════
//  EVENTOS
// ════════════════════════════════════════════════════════════════════
$("baselineFile").addEventListener("change", e => { const f = e.target.files[0]; if (f) loadFile(f, "baseline"); });
$("subFile").addEventListener("change",      e => { const f = e.target.files[0]; if (f) loadFile(f, "sub"); });
$("loadSheetsBtn").addEventListener("click", () => {
  if ($("baselineSheet").value) prepareSheet("baseline");
  if ($("subSheet").value)      prepareSheet("sub");
});
$("compareBtn").addEventListener("click",       runComparison);
$("compareBlockBtn").addEventListener("click",  compareByBlock);
$("compareSuggestBtn").addEventListener("click", suggestMatch);
$("suggestBtn").addEventListener("click",       suggestMatch);
$("exportBtn").addEventListener("click",        exportCsv);
$("baselineSelectAll").addEventListener("click", () => bulkSelect("baseline", "all"));
$("baselineClear").addEventListener("click",     () => bulkSelect("baseline", "clear"));
$("subSelectAll").addEventListener("click",      () => bulkSelect("sub", "all"));
$("subClear").addEventListener("click",          () => bulkSelect("sub", "clear"));
$("baselineSearch").addEventListener("input", () => renderChecklist("baseline"));
$("subSearch").addEventListener("input",      () => renderChecklist("sub"));
$("baselineBlockSel").addEventListener("change", () => renderChecklist("baseline"));
$("subBlockSel").addEventListener("change",      () => renderChecklist("sub"));
$("hideSummaryRows").addEventListener("change",   refreshActivities);
["baselineActivityCol","baselinePvCol","baselineStartCol","baselineFinishCol","baselineHierarchyCol",
 "subActivityCol","subPvCol","subStartCol","subFinishCol","subHierarchyCol"].forEach(id => {
  $(id).addEventListener("change", refreshActivities);
});
