/* ======================================================================
   Bitácora de Compras — Navimag
   App vanilla JS: parseo Excel, almacenamiento local, dashboard, informes PDF
   ====================================================================== */

const COLOR = {
  teal: "#22B8A6", tealDim: "#144B45",
  amber: "#F2A93B", amberDim: "#4D3A16",
  red: "#E5595F", redDim: "#4A1F22",
  blue: "#5AA9E6", purple: "#B08AE0", muted: "#7E9499",
};

const ESTADOS = {
  ACTIVA: "Solicitud activa",
  APROBACION: "En aprobación",
  PENDIENTE_OC: "Aprobada · sin OC",
  PARCIAL: "OC parcial",
  OC_CREADA: "OC creada",
};

const ESTADO_COLOR = {
  [ESTADOS.ACTIVA]: COLOR.blue,
  [ESTADOS.APROBACION]: COLOR.amber,
  [ESTADOS.PENDIENTE_OC]: COLOR.red,
  [ESTADOS.PARCIAL]: COLOR.amber,
  [ESTADOS.OC_CREADA]: COLOR.teal,
};

const SOURCES = [
  { key: "taller", label: "Taller", color: COLOR.teal },
  { key: "esperanza", label: "Esperanza", color: COLOR.blue },
  { key: "dalka", label: "Dalka", color: COLOR.amber },
];

const MAX_DETAIL = 8;
const DIAS_ALERTA = 10;
const STORAGE_KEY = "bitacora_compras_v2";

/* ---------------------------------------------------------------------- */
/* Helpers                                                                  */
/* ---------------------------------------------------------------------- */
function normalize(s) {
  if (s === null || s === undefined) return "";
  return String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\./g, "").replace(/\s+/g, " ").trim();
}

function findIdx(headers, name, occurrence = 0) {
  const target = normalize(name);
  let count = 0;
  for (let i = 0; i < headers.length; i++) {
    if (normalize(headers[i]) === target) { if (count === occurrence) return i; count++; }
  }
  return -1;
}

function toDateSafe(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function fmtDate(d) {
  if (!d) return "—";
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt)) return "—";
  return dt.toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric" });
}

function fmtMoney(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return "$" + Math.round(n).toLocaleString("es-CL");
}

function diffDays(a, b) {
  if (!a || !b) return null;
  return Math.round((b - a) / 86400000);
}

function isoWeekLabel(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-S${String(weekNo).padStart(2, "0")}`;
}

function escapeHtml(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/* ---------------------------------------------------------------------- */
/* Parsing                                                                  */
/* ---------------------------------------------------------------------- */
function parseWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!raw.length) throw new Error("El archivo no tiene datos.");

  const headers = raw[0];
  const idx = {
    descMaterial: findIdx(headers, "descripcion del material", 0),
    monto: findIdx(headers, "monto neto", 0),
    solped: findIdx(headers, "numero solicitud de pedido", 0),
    posicion: findIdx(headers, "posicion", 0),
    ocNumero: findIdx(headers, "numero del documento de compras", 0),
    proveedor: findIdx(headers, "proveedor", 0),
    centro: findIdx(headers, "centro", 0),
    fechaSolicitud: findIdx(headers, "fecha de solicitud", 0),
    solicitante: findIdx(headers, "nombre del solicitante", 0),
    responsableMod: findIdx(headers, "nombre del responsable que modifico", 0),
    grupoCompras: findIdx(headers, "grupo de compras", 0),
    statusSolped: findIdx(headers, "status trat solped", 0),
    claseDocOC: findIdx(headers, "clase de documento", 1),
    fechaDocCompras: findIdx(headers, "fecha del documento de compras", 0),
    fechaCreacionOC: findIdx(headers, "fecha de creacion", 0),
    fechaLiberacion: findIdx(headers, "fecha de liberacion", 0),
    fechaRecepcionMI: findIdx(headers, "fecha registro de mi", 0),
    moneda: findIdx(headers, "moneda", 0),
  };

  if (idx.solped === -1 || idx.fechaSolicitud === -1) {
    throw new Error("No se reconocen las columnas esperadas (Número Solicitud de Pedido / Fecha de Solicitud). Verifica que sea el reporte estándar de solicitudes.");
  }

  const rows = [];
  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    if (!row || row.every((c) => c === null || c === "")) continue;
    const solped = row[idx.solped];
    if (!solped) continue;

    const ocNumero = idx.ocNumero !== -1 ? row[idx.ocNumero] : null;
    const fechaSolicitud = toDateSafe(idx.fechaSolicitud !== -1 ? row[idx.fechaSolicitud] : null);
    const fechaOC = toDateSafe(
      (idx.fechaDocCompras !== -1 && row[idx.fechaDocCompras]) ? row[idx.fechaDocCompras] :
      (idx.fechaCreacionOC !== -1 ? row[idx.fechaCreacionOC] : null)
    );
    const statusSolped = idx.statusSolped !== -1 ? row[idx.statusSolped] : null;

    let estado;
    const tieneOC = !!(ocNumero && String(ocNumero).trim());
    const statusNorm = normalize(statusSolped);
    if (tieneOC) estado = ESTADOS.OC_CREADA;
    else if (statusNorm.includes("activa")) estado = ESTADOS.ACTIVA;
    else if (statusNorm.includes("proceso")) estado = ESTADOS.APROBACION;
    else estado = ESTADOS.PENDIENTE_OC;

    const hoy = new Date();
    const dias = tieneOC ? diffDays(fechaSolicitud, fechaOC) : diffDays(fechaSolicitud, hoy);

    const fechaLiberacion = toDateSafe(idx.fechaLiberacion !== -1 ? row[idx.fechaLiberacion] : null);
    const fechaRecepcionMI = toDateSafe(idx.fechaRecepcionMI !== -1 ? row[idx.fechaRecepcionMI] : null);
    const tieneRecepcion = !!fechaRecepcionMI;

    const diasAprobacion = fechaLiberacion ? diffDays(fechaSolicitud, fechaLiberacion) : null;
    const diasGeneracionOC = (fechaOC && (fechaLiberacion || fechaSolicitud)) ? diffDays(fechaLiberacion || fechaSolicitud, fechaOC) : null;
    const diasRecepcion = (fechaOC && fechaRecepcionMI) ? diffDays(fechaOC, fechaRecepcionMI) : null;
    const diasTotalCompleto = fechaRecepcionMI ? diffDays(fechaSolicitud, fechaRecepcionMI) : null;

    rows.push({
      solped: String(solped).trim(),
      posicion: idx.posicion !== -1 ? row[idx.posicion] : null,
      material: (idx.descMaterial !== -1 ? row[idx.descMaterial] : null) || "Sin descripción",
      monto: idx.monto !== -1 ? (parseFloat(row[idx.monto]) || 0) : 0,
      proveedor: idx.proveedor !== -1 ? row[idx.proveedor] : null,
      centro: idx.centro !== -1 ? row[idx.centro] : null,
      fechaSolicitud: fechaSolicitud ? fechaSolicitud.toISOString() : null,
      solicitante: (idx.solicitante !== -1 ? row[idx.solicitante] : null) || (idx.responsableMod !== -1 ? row[idx.responsableMod] : null) || "—",
      grupoCompras: idx.grupoCompras !== -1 ? row[idx.grupoCompras] : null,
      statusSolped,
      claseDocOC: idx.claseDocOC !== -1 ? row[idx.claseDocOC] : null,
      ocNumero: tieneOC ? String(ocNumero).trim() : null,
      fechaOC: fechaOC ? fechaOC.toISOString() : null,
      fechaLiberacion: fechaLiberacion ? fechaLiberacion.toISOString() : null,
      fechaRecepcionMI: fechaRecepcionMI ? fechaRecepcionMI.toISOString() : null,
      tieneRecepcion,
      moneda: idx.moneda !== -1 ? row[idx.moneda] : "CLP",
      estado,
      dias,
      diasAprobacion,
      diasGeneracionOC,
      diasRecepcion,
      diasTotalCompleto,
    });
  }
  return rows;
}

/* rows here always have fechaSolicitud/fechaOC as ISO strings or null.
   These helpers work off Date objects, so callers convert when needed. */
function rd(row, field) { return row[field] ? new Date(row[field]) : null; }

function groupBySolped(rows) {
  const map = new Map();
  for (const r of rows) {
    if (!map.has(r.solped)) {
      map.set(r.solped, {
        solped: r.solped, lineas: [], montoTotal: 0, fechaSolicitud: rd(r, "fechaSolicitud"),
        solicitante: r.solicitante, materiales: [], conOC: 0, sinOC: 0, proveedores: new Set(), centro: r.centro,
      });
    }
    const g = map.get(r.solped);
    g.lineas.push(r);
    g.montoTotal += r.monto || 0;
    g.materiales.push(r.material);
    if (r.ocNumero) g.conOC++; else g.sinOC++;
    if (r.proveedor) g.proveedores.add(r.proveedor);
    const fs = rd(r, "fechaSolicitud");
    if (fs && (!g.fechaSolicitud || fs < g.fechaSolicitud)) g.fechaSolicitud = fs;
  }
  const hoy = new Date();
  return Array.from(map.values()).map((g) => {
    let estado;
    if (g.sinOC === 0) estado = ESTADOS.OC_CREADA;
    else if (g.conOC > 0) estado = ESTADOS.PARCIAL;
    else {
      const worst = g.lineas.find((l) => l.estado !== ESTADOS.OC_CREADA);
      estado = worst ? worst.estado : ESTADOS.PENDIENTE_OC;
    }
    const ultimaFechaOC = g.lineas.reduce((acc, l) => { const fo = rd(l, "fechaOC"); return (fo && (!acc || fo > acc)) ? fo : acc; }, null);
    const dias = estado === ESTADOS.OC_CREADA ? diffDays(g.fechaSolicitud, ultimaFechaOC) : diffDays(g.fechaSolicitud, hoy);
    return {
      ...g, estado, dias, numLineas: g.lineas.length,
      material: g.materiales[0] + (g.materiales.length > 1 ? ` (+${g.materiales.length - 1} más)` : ""),
      proveedor: Array.from(g.proveedores)[0] || null,
    };
  });
}

function avg(arr) {
  const vals = arr.filter((d) => d !== null && d !== undefined);
  return vals.length ? Math.round(vals.reduce((a, b) => a + b, 0) / vals.length) : null;
}

function computeSummary(rows) {
  const grouped = groupBySolped(rows);
  const totalSolped = grouped.length;
  const conOC = grouped.filter((g) => g.estado === ESTADOS.OC_CREADA).length;
  const sinOC = totalSolped - conOC;
  const atrasadas = grouped.filter((g) => g.estado !== ESTADOS.OC_CREADA && g.dias !== null && g.dias > DIAS_ALERTA).length;
  const sinRecepcion = rows.filter((r) => r.ocNumero && !r.tieneRecepcion).length;
  const avgAprobacion = avg(rows.map((r) => r.diasAprobacion));
  const avgGeneracionOC = avg(rows.map((r) => r.diasGeneracionOC));
  const avgRecepcion = avg(rows.map((r) => r.diasRecepcion));
  const avgTotalCompleto = avg(rows.map((r) => r.diasTotalCompleto));
  return { totalSolped, totalLineas: rows.length, conOC, sinOC, atrasadas, sinRecepcion, avgAprobacion, avgGeneracionOC, avgRecepcion, avgTotalCompleto };
}

/* ---------------------------------------------------------------------- */
/* Estado global + persistencia (localStorage)                             */
/* ---------------------------------------------------------------------- */
let STATE = {
  sources: { taller: { snapshots: [] }, esperanza: { snapshots: [] }, dalka: { snapshots: [] } },
  activeTab: "taller",
  selectedIds: {},
  loading: false,
  error: "",
};

let UI = { search: "", estadoFilter: "TODOS", view: "solped", sortDias: false };
let CHARTS = {}; // canvasId -> Chart.js instance
let CURRENT_CTX = null; // { rows, grouped, accent } for the active tab render

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      STATE.sources = { ...STATE.sources, ...parsed };
    }
  } catch (e) { console.warn("No se pudo leer almacenamiento local", e); }
}

function persist() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(STATE.sources)); } catch (e) { console.warn("No se pudo guardar", e); }
}

/* ---------------------------------------------------------------------- */
/* Render: tabs                                                            */
/* ---------------------------------------------------------------------- */
function renderTabs() {
  const tabsEl = document.getElementById("tabs");
  const allTabs = [...SOURCES, { key: "consolidado", label: "Consolidado", color: COLOR.purple }];
  tabsEl.innerHTML = allTabs.map((t) => {
    const isActive = STATE.activeTab === t.key;
    const count = t.key !== "consolidado" ? (STATE.sources[t.key]?.snapshots?.length || 0) : null;
    return `<button class="tab-btn ${isActive ? "active" : ""}" data-action="set-tab" data-tab="${t.key}" style="border-bottom-color:${isActive ? t.color : "transparent"}">
      <span style="color:${isActive ? t.color : COLOR.muted}">●</span> ${t.label} ${count !== null ? `<span class="tab-count">(${count})</span>` : ""}
    </button>`;
  }).join("");
}

/* ---------------------------------------------------------------------- */
/* Render: upload box                                                       */
/* ---------------------------------------------------------------------- */
function uploadBoxHtml(sourceKey, sourceLabel, compact) {
  if (compact) {
    return `<label class="upload-box-compact" data-dropzone="${sourceKey}">
      ⭱ ${STATE.loading ? "Procesando…" : `Cargar semana nueva de ${sourceLabel}`}
      <input type="file" accept=".xlsx,.xls,.csv" style="display:none" data-action="upload-file" data-source="${sourceKey}" />
    </label>`;
  }
  return `<div class="upload-screen">
    <label class="upload-box" data-dropzone="${sourceKey}">
      <div style="font-size:22px;color:${COLOR.muted}">⭱</div>
      <h2>${STATE.loading ? "Procesando archivo…" : `Sube el primer reporte de ${sourceLabel}`}</h2>
      <span class="muted-dark" style="font-size:12px">Formato Excel (.xlsx, .xls, .csv)</span>
      <input type="file" accept=".xlsx,.xls,.csv" style="display:none" data-action="upload-file" data-source="${sourceKey}" />
    </label>
    ${STATE.error ? `<div class="error-box">${escapeHtml(STATE.error)}</div>` : ""}
  </div>`;
}

/* ---------------------------------------------------------------------- */
/* Render: KPI cards                                                        */
/* ---------------------------------------------------------------------- */
function kpiCardHtml(label, value, sub, accent) {
  return `<div class="kpi-card" style="--accent:${accent}">
    <div class="kpi-label">${label}</div>
    <div class="kpi-value">${value}</div>
    ${sub ? `<div class="kpi-sub">${sub}</div>` : ""}
  </div>`;
}

function badgeHtml(estado) {
  const color = ESTADO_COLOR[estado] || COLOR.muted;
  return `<span class="badge" style="color:${color};background:${color}1E;border:1px solid ${color}40">
    <span class="dot" style="background:${color}"></span>${estado}
  </span>`;
}

/* ---------------------------------------------------------------------- */
/* Render: full source dashboard                                           */
/* ---------------------------------------------------------------------- */
function renderSourceDashboard(source, snapshot, snapshots) {
  const rows = snapshot.rows;
  const grouped = groupBySolped(rows);
  const summary = snapshot.summary;
  CURRENT_CTX = { rows, grouped, accent: source.color, source, snapshot, snapshots };

  const viewingBar = `<div class="viewing-bar">
    <div style="font-size:12.5px" class="muted">
      Viendo semana <span class="mono" style="color:var(--text)">${snapshot.weekLabel}</span> · ${escapeHtml(snapshot.fileName)}
      ${snapshot.id !== snapshots[0].id ? `<span style="color:${COLOR.amber}"> (histórica)</span>` : ""}
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <button class="btn primary" data-action="generate-pdf" data-source="${source.key}">⬇ Generar informe PDF</button>
      ${uploadBoxHtml(source.key, source.label, true)}
    </div>
  </div>`;

  const kpis = `<div class="kpi-row">
    ${kpiCardHtml("Solicitudes", summary.totalSolped, `${summary.totalLineas} líneas de material`, COLOR.blue)}
    ${kpiCardHtml("OC creadas", `${summary.conOC} (${summary.totalSolped ? Math.round(100 * summary.conOC / summary.totalSolped) : 0}%)`, `${summary.sinOC} aún sin OC`, source.color)}
    ${kpiCardHtml("Pendientes atrasadas", summary.atrasadas, `&gt; ${DIAS_ALERTA} días sin OC`, COLOR.red)}
    ${kpiCardHtml("Repuestos por recepcionar", summary.sinRecepcion, "OC creada, sin ingreso a bodega", COLOR.amber)}
    ${kpiCardHtml("Tiempo aprobación", summary.avgAprobacion !== null ? `${summary.avgAprobacion} d` : "—", "Solicitud → liberación", COLOR.blue)}
    ${kpiCardHtml("Tiempo generación OC", summary.avgGeneracionOC !== null ? `${summary.avgGeneracionOC} d` : "—", "Liberación → OC", source.color)}
    ${kpiCardHtml("Tiempo recepción", summary.avgRecepcion !== null ? `${summary.avgRecepcion} d` : "—", "OC → ingreso a bodega", COLOR.purple)}
  </div>`;

  const charts = `<div class="panel-row">
    <div class="panel"><div class="panel-header"><span class="panel-title">Estado de las solicitudes</span></div><div class="chart-box"><canvas id="chartEstado"></canvas></div></div>
    <div class="panel"><div class="panel-header"><span class="panel-title">Solicitudes por mes</span></div><div class="chart-box"><canvas id="chartMensual"></canvas></div></div>
    <div class="panel"><div class="panel-header"><span class="panel-title">Tiempos por etapa (promedio, días)</span></div><div class="chart-box"><canvas id="chartTiempos"></canvas></div></div>
  </div>`;

  const tablePanel = `<div class="panel">
    <div class="panel-header">
      <span class="panel-title" id="tableTitle">Solicitudes (agrupadas)</span>
      <div class="controls">
        <div class="search-wrap"><span class="muted-dark" style="position:absolute;left:8px;top:6px;font-size:11px">🔍</span>
          <input type="text" class="with-icon" id="searchInput" placeholder="Buscar…" value="${escapeHtml(UI.search)}" /></div>
        <select id="estadoSelect">
          <option value="TODOS">Todos los estados</option>
          ${Object.values(ESTADOS).map((e) => `<option value="${e}" ${UI.estadoFilter === e ? "selected" : ""}>${e}</option>`).join("")}
        </select>
        <button class="btn" id="viewToggle" data-action="toggle-view">${UI.view === "solped" ? "▤ Ver por línea" : "▥ Ver por solicitud"}</button>
        <button class="btn ${UI.sortDias ? "active" : ""}" id="sortToggle" data-action="toggle-sort">↕ Ordenar por días</button>
      </div>
    </div>
    <div class="table-scroll">
      <table>
        <thead><tr>
          <th>SOLPED</th><th>Material</th><th>Solicitante</th><th>F. Solicitud</th><th>Estado</th><th>Nº OC</th><th class="right">Días</th>
        </tr></thead>
        <tbody id="tableBody"></tbody>
      </table>
      <div id="tableEmptyNote"></div>
    </div>
  </div>`;

  const pendientesPanel = pendientesPanelHtml(grouped);
  const sinRecepcionPanel = sinRecepcionPanelHtml(rows);
  const historySection = historySectionHtml(snapshots, source.color);

  document.getElementById("content").innerHTML = viewingBar + kpis + charts + tablePanel + pendientesPanel + sinRecepcionPanel + historySection;
  renderCharts(grouped, rows, summary, source.color);
  renderTableBody();
  renderHistoryChart(snapshots, source.color);
  wireDropzones();
}

/* ---------------------------------------------------------------------- */
/* Solicitudes pendientes (todas, ordenadas por días de mayor a menor)      */
/* ---------------------------------------------------------------------- */
function getPendientesGrouped(grouped) {
  return grouped.filter((g) => g.estado !== ESTADOS.OC_CREADA).sort((a, b) => (b.dias ?? -1) - (a.dias ?? -1));
}

function pendientesPanelHtml(grouped) {
  const list = getPendientesGrouped(grouped);
  return `<div class="panel" style="margin-top:16px">
    <div class="panel-header"><span class="panel-title">⏳ Solicitudes pendientes (${list.length}) — mayor a menor retraso</span></div>
    <div class="table-scroll"><table>
      <thead><tr><th>SOLPED</th><th>Material</th><th>Solicitante</th><th>F. Solicitud</th><th>Estado</th><th class="right">Días</th></tr></thead>
      <tbody>${list.length ? list.map((g) => `<tr>
        <td class="mono" style="color:${COLOR.blue}">${escapeHtml(g.solped)}</td>
        <td style="max-width:260px">${escapeHtml(g.material)}</td>
        <td class="muted">${escapeHtml(g.solicitante)}</td>
        <td class="mono muted">${fmtDate(g.fechaSolicitud)}</td>
        <td>${badgeHtml(g.estado)}</td>
        <td class="mono right" style="color:${g.dias > DIAS_ALERTA ? COLOR.red : "var(--muted)"}">${g.dias !== null ? g.dias : "—"}</td>
      </tr>`).join("") : `<tr><td colspan="6" class="empty-note">No hay solicitudes pendientes — todo con OC creada.</td></tr>`}</tbody>
    </table></div>
  </div>`;
}

/* ---------------------------------------------------------------------- */
/* OC liberadas sin recepción de repuesto (seguimiento a la entrega)        */
/* ---------------------------------------------------------------------- */
function getSinRecepcionRows(rows) {
  const hoy = new Date();
  return rows.filter((r) => r.ocNumero && !r.tieneRecepcion)
    .map((r) => ({ ...r, diasDesdeOC: diffDays(rd(r, "fechaOC"), hoy) }))
    .sort((a, b) => (b.diasDesdeOC ?? -1) - (a.diasDesdeOC ?? -1));
}

function sinRecepcionPanelHtml(rows) {
  const list = getSinRecepcionRows(rows);
  return `<div class="panel" style="margin-top:16px">
    <div class="panel-header"><span class="panel-title">📦 OC liberadas sin recepción de repuesto (${list.length})</span></div>
    <div class="table-scroll"><table>
      <thead><tr><th>SOLPED</th><th>Material</th><th>Proveedor</th><th>Nº OC</th><th>Fecha OC</th><th class="right">Días desde OC</th></tr></thead>
      <tbody>${list.length ? list.map((r) => `<tr class="${r.diasDesdeOC > DIAS_ALERTA ? "alert-row" : ""}">
        <td class="mono" style="color:${COLOR.blue}">${escapeHtml(r.solped)}${r.posicion ? `-${r.posicion}` : ""}</td>
        <td style="max-width:260px">${escapeHtml(r.material)}</td>
        <td class="muted mono">${escapeHtml(r.proveedor || "—")}</td>
        <td class="mono muted">${escapeHtml(r.ocNumero)}</td>
        <td class="mono muted">${fmtDate(rd(r, "fechaOC"))}</td>
        <td class="mono right" style="color:${r.diasDesdeOC > DIAS_ALERTA ? COLOR.red : "var(--muted)"}">${r.diasDesdeOC !== null ? r.diasDesdeOC : "—"}</td>
      </tr>`).join("") : `<tr><td colspan="6" class="empty-note">Todos los repuestos con OC ya fueron recepcionados.</td></tr>`}</tbody>
    </table></div>
  </div>`;
}

/* ---------------------------------------------------------------------- */
/* Table body (re-rendered independently to preserve input focus)          */
/* ---------------------------------------------------------------------- */
function getFilteredList() {
  if (!CURRENT_CTX) return [];
  const { rows, grouped } = CURRENT_CTX;
  const source = UI.view === "solped" ? grouped : rows;
  const q = normalize(UI.search);
  let list = source.filter((r) => {
    if (UI.estadoFilter !== "TODOS" && r.estado !== UI.estadoFilter) return false;
    if (!q) return true;
    return normalize(r.solped).includes(q) || normalize(r.material).includes(q) || normalize(r.solicitante || "").includes(q) || normalize(r.proveedor || "").includes(q);
  });
  list = UI.sortDias
    ? [...list].sort((a, b) => (b.dias ?? -1) - (a.dias ?? -1))
    : [...list].sort((a, b) => (rowDate(b) || 0) - (rowDate(a) || 0));
  return list;
}
function rowDate(r) { return r.fechaSolicitud instanceof Date ? r.fechaSolicitud : (r.fechaSolicitud ? new Date(r.fechaSolicitud) : null); }

function renderTableBody() {
  const list = getFilteredList();
  const tbody = document.getElementById("tableBody");
  const titleEl = document.getElementById("tableTitle");
  if (titleEl) titleEl.textContent = UI.view === "solped" ? "Solicitudes (agrupadas)" : "Líneas de detalle";
  if (!tbody) return;
  tbody.innerHTML = list.slice(0, 150).map((r, i) => {
    const isAlert = r.estado !== ESTADOS.OC_CREADA && r.dias !== null && r.dias > DIAS_ALERTA;
    const fecha = fmtDate(rowDate(r));
    const solpedLabel = r.solped + (UI.view === "linea" && r.posicion ? `-${r.posicion}` : "");
    const ocLabel = r.ocNumero || (UI.view === "solped" ? `${r.conOC}/${r.numLineas}` : "—");
    return `<tr class="${isAlert ? "alert-row" : ""}">
      <td class="mono" style="color:${COLOR.blue}">${escapeHtml(solpedLabel)}</td>
      <td style="max-width:260px">${escapeHtml(r.material)}</td>
      <td class="muted">${escapeHtml(r.solicitante)}</td>
      <td class="mono muted">${fecha}</td>
      <td>${badgeHtml(r.estado)}</td>
      <td class="mono muted">${escapeHtml(ocLabel)}</td>
      <td class="mono right" style="color:${isAlert ? COLOR.red : "var(--muted)"};font-weight:${isAlert ? 700 : 400}">${r.dias !== null ? r.dias : "—"}</td>
    </tr>`;
  }).join("");
  const note = document.getElementById("tableEmptyNote");
  if (note) {
    if (list.length === 0) note.innerHTML = `<div class="empty-note">Sin resultados para los filtros aplicados.</div>`;
    else if (list.length > 150) note.innerHTML = `<div class="empty-note">Mostrando 150 de ${list.length} resultados. Usa la búsqueda para acotar.</div>`;
    else note.innerHTML = "";
  }
}

/* ---------------------------------------------------------------------- */
/* Charts (Chart.js)                                                       */
/* ---------------------------------------------------------------------- */
function destroyChart(id) { if (CHARTS[id]) { CHARTS[id].destroy(); delete CHARTS[id]; } }

function chartDefaults() {
  return {
    plugins: { legend: { labels: { color: COLOR.muted, font: { size: 11 } } }, tooltip: { backgroundColor: "#0D1618", borderColor: "#1F2E32", borderWidth: 1, titleColor: "#E7F1F0", bodyColor: "#E7F1F0" } },
    scales: {},
  };
}

function renderCharts(grouped, rows, summary, accent) {
  ["chartEstado", "chartMensual", "chartTiempos"].forEach(destroyChart);

  // Estado doughnut
  const estadoCounts = {};
  grouped.forEach((g) => { estadoCounts[g.estado] = (estadoCounts[g.estado] || 0) + 1; });
  const estadoLabels = Object.keys(estadoCounts);
  const ctxEstado = document.getElementById("chartEstado");
  if (ctxEstado) {
    CHARTS.chartEstado = new Chart(ctxEstado, {
      type: "doughnut",
      data: { labels: estadoLabels, datasets: [{ data: estadoLabels.map((l) => estadoCounts[l]), backgroundColor: estadoLabels.map((l) => ESTADO_COLOR[l] || COLOR.muted), borderColor: "#111B1E", borderWidth: 2 }] },
      options: { ...chartDefaults(), cutout: "58%" },
    });
  }

  // Monthly stacked bar
  const monthMap = new Map();
  grouped.forEach((g) => {
    if (!g.fechaSolicitud) return;
    const key = `${g.fechaSolicitud.getFullYear()}-${String(g.fechaSolicitud.getMonth() + 1).padStart(2, "0")}`;
    if (!monthMap.has(key)) monthMap.set(key, { conOC: 0, sinOC: 0 });
    const m = monthMap.get(key);
    if (g.estado === ESTADOS.OC_CREADA) m.conOC++; else m.sinOC++;
  });
  const months = Array.from(monthMap.keys()).sort();
  const ctxMensual = document.getElementById("chartMensual");
  if (ctxMensual) {
    CHARTS.chartMensual = new Chart(ctxMensual, {
      type: "bar",
      data: {
        labels: months,
        datasets: [
          { label: "Con OC", data: months.map((m) => monthMap.get(m).conOC), backgroundColor: accent, stack: "s" },
          { label: "Sin OC", data: months.map((m) => monthMap.get(m).sinOC), backgroundColor: COLOR.red, stack: "s" },
        ],
      },
      options: { ...chartDefaults(), scales: { x: { stacked: true, ticks: { color: COLOR.muted, font: { size: 10.5 } }, grid: { color: "#1F2E32" } }, y: { stacked: true, ticks: { color: COLOR.muted, precision: 0 }, grid: { color: "#1F2E32" } } } },
    });
  }

  // Tiempos por etapa (promedio en días)
  const etapas = ["Aprobación", "Generación OC", "Recepción"];
  const valores = [summary.avgAprobacion, summary.avgGeneracionOC, summary.avgRecepcion];
  const ctxTiempos = document.getElementById("chartTiempos");
  if (ctxTiempos) {
    CHARTS.chartTiempos = new Chart(ctxTiempos, {
      type: "bar",
      data: { labels: etapas, datasets: [{ label: "Días promedio", data: valores.map((v) => v ?? 0), backgroundColor: [COLOR.blue, accent, COLOR.purple] }] },
      options: { ...chartDefaults(), plugins: { ...chartDefaults().plugins, legend: { display: false } }, scales: { x: { ticks: { color: COLOR.muted }, grid: { color: "#1F2E32" } }, y: { ticks: { color: COLOR.muted, precision: 0 }, grid: { color: "#1F2E32" } } } },
    });
  }
}

/* ---------------------------------------------------------------------- */
/* Historial semanal                                                       */
/* ---------------------------------------------------------------------- */
function historySectionHtml(snapshots, accent) {
  return `<div class="panel" style="margin-top:16px">
    <div class="panel-header"><span class="panel-title">📅 Historial semanal</span></div>
    <div class="chart-box small"><canvas id="chartHistorial"></canvas></div>
    <div class="table-scroll" style="margin-top:10px">
      <table>
        <thead><tr><th>Semana</th><th>Fecha carga</th><th>Archivo</th><th>Solicitudes</th><th>Con OC</th><th>Backlog</th><th>Atrasadas</th><th></th></tr></thead>
        <tbody>${snapshots.map((s, i) => historyRowHtml(s, snapshots[i + 1])).join("")}</tbody>
      </table>
    </div>
  </div>`;
}

function historyRowHtml(s, prev) {
  const delta = prev ? s.summary.totalSolped - prev.summary.totalSolped : null;
  const isSelected = CURRENT_CTX && CURRENT_CTX.snapshot && CURRENT_CTX.snapshot.id === s.id;
  return `<tr class="${isSelected ? "selected-row" : ""}">
    <td class="mono">${s.weekLabel}</td>
    <td class="muted">${fmtDate(new Date(s.date))}</td>
    <td class="muted-dark" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(s.fileName)}</td>
    <td class="mono">${s.summary.totalSolped} ${delta !== null && delta !== 0 ? `<span style="color:${delta > 0 ? COLOR.blue : "var(--muted-dark)"};font-size:10.5px">(${delta > 0 ? "+" : ""}${delta})</span>` : ""}</td>
    <td class="mono" style="color:${COLOR.teal}">${s.summary.conOC}</td>
    <td class="mono" style="color:${s.summary.sinOC > 0 ? COLOR.red : "var(--muted)"}">${s.summary.sinOC}</td>
    <td class="mono" style="color:${s.summary.atrasadas > 0 ? COLOR.amber : "var(--muted)"}">${s.summary.atrasadas}</td>
    <td>
      ${s.rows ? `<button class="icon-btn" data-action="select-snapshot" data-id="${s.id}" title="Ver detalle de esta semana">👁</button>` : `<span class="muted-dark" title="Solo resumen">👁</span>`}
      <button class="icon-btn" data-action="delete-snapshot" data-id="${s.id}" title="Eliminar snapshot">🗑</button>
    </td>
  </tr>`;
}

function renderHistoryChart(snapshots, accent) {
  destroyChart("chartHistorial");
  const ordered = [...snapshots].reverse();
  const ctx = document.getElementById("chartHistorial");
  if (!ctx) return;
  CHARTS.chartHistorial = new Chart(ctx, {
    type: "line",
    data: {
      labels: ordered.map((s) => s.weekLabel),
      datasets: [
        { label: "Solicitudes", data: ordered.map((s) => s.summary.totalSolped), borderColor: accent, backgroundColor: accent, tension: 0.25 },
        { label: "Backlog", data: ordered.map((s) => s.summary.sinOC), borderColor: COLOR.red, backgroundColor: COLOR.red, tension: 0.25 },
        { label: "Atrasadas", data: ordered.map((s) => s.summary.atrasadas), borderColor: COLOR.amber, backgroundColor: COLOR.amber, borderDash: [4, 3], tension: 0.25 },
      ],
    },
    options: { ...chartDefaults(), scales: { x: { ticks: { color: COLOR.muted, font: { size: 10.5 } }, grid: { color: "#1F2E32" } }, y: { ticks: { color: COLOR.muted, precision: 0 }, grid: { color: "#1F2E32" } } } },
  });
}

/* ---------------------------------------------------------------------- */
/* Consolidado                                                             */
/* ---------------------------------------------------------------------- */
function renderConsolidado() {
  CURRENT_CTX = null;
  const activeSources = SOURCES.filter((s) => STATE.sources[s.key]?.snapshots?.length);
  if (!activeSources.length) {
    document.getElementById("content").innerHTML = `<div class="empty-note">Aún no hay datos cargados en ninguna fuente. Ve a las pestañas Taller, Esperanza o Dalka y sube el primer reporte de cada una.</div>`;
    return;
  }
  const latest = activeSources.map((s) => ({ ...s, summary: STATE.sources[s.key].snapshots[0].summary, weekLabel: STATE.sources[s.key].snapshots[0].weekLabel }));
  const combined = latest.reduce((acc, s) => ({
    totalSolped: acc.totalSolped + s.summary.totalSolped, conOC: acc.conOC + s.summary.conOC, sinOC: acc.sinOC + s.summary.sinOC,
    atrasadas: acc.atrasadas + s.summary.atrasadas, sinRecepcion: acc.sinRecepcion + (s.summary.sinRecepcion || 0),
  }), { totalSolped: 0, conOC: 0, sinOC: 0, atrasadas: 0, sinRecepcion: 0 });

  const kpis = `<div class="kpi-row">
    ${kpiCardHtml("Solicitudes (todas)", combined.totalSolped, `${activeSources.length}/3 fuentes cargadas`, COLOR.blue)}
    ${kpiCardHtml("OC creadas", `${combined.conOC} (${combined.totalSolped ? Math.round(100 * combined.conOC / combined.totalSolped) : 0}%)`, `${combined.sinOC} sin OC`, COLOR.teal)}
    ${kpiCardHtml("Pendientes atrasadas", combined.atrasadas, `&gt; ${DIAS_ALERTA} días sin OC`, COLOR.red)}
    ${kpiCardHtml("Repuestos por recepcionar", combined.sinRecepcion, "OC creada, sin ingreso a bodega", COLOR.amber)}
  </div>`;

  const actionBar = `<div class="viewing-bar"><div></div><button class="btn primary" data-action="generate-pdf" data-source="consolidado">⬇ Generar informe PDF consolidado</button></div>`;

  const charts = `<div class="panel-row">
    <div class="panel"><div class="panel-header"><span class="panel-title">Con OC / Sin OC por fuente</span></div><div class="chart-box"><canvas id="chartConsEstado"></canvas></div></div>
    <div class="panel"><div class="panel-header"><span class="panel-title">Evolución semanal por fuente</span></div><div class="chart-box"><canvas id="chartConsHistorial"></canvas></div></div>
  </div>`;

  const table = `<div class="panel"><div class="panel-header"><span class="panel-title">Resumen por fuente (última semana cargada)</span></div>
    <div class="table-scroll"><table>
      <thead><tr><th>Fuente</th><th>Semana</th><th>Solicitudes</th><th>Con OC</th><th>Sin OC</th><th>Atrasadas</th><th>Sin recepción</th></tr></thead>
      <tbody>${latest.map((s) => `<tr>
        <td style="font-weight:600;color:${s.color}">${s.label}</td>
        <td class="mono muted">${s.weekLabel}</td>
        <td class="mono">${s.summary.totalSolped}</td>
        <td class="mono" style="color:${COLOR.teal}">${s.summary.conOC}</td>
        <td class="mono" style="color:${s.summary.sinOC ? COLOR.red : "var(--muted)"}">${s.summary.sinOC}</td>
        <td class="mono" style="color:${s.summary.atrasadas ? COLOR.amber : "var(--muted)"}">${s.summary.atrasadas}</td>
        <td class="mono" style="color:${s.summary.sinRecepcion ? COLOR.amber : "var(--muted)"}">${s.summary.sinRecepcion ?? "—"}</td>
      </tr>`).join("")}</tbody>
    </table></div>
  </div>`;

  document.getElementById("content").innerHTML = actionBar + kpis + charts + table;

  destroyChart("chartConsEstado"); destroyChart("chartConsHistorial");
  const ctx1 = document.getElementById("chartConsEstado");
  if (ctx1) {
    CHARTS.chartConsEstado = new Chart(ctx1, {
      type: "bar",
      data: { labels: latest.map((s) => s.label), datasets: [
        { label: "Con OC", data: latest.map((s) => s.summary.conOC), backgroundColor: COLOR.teal, stack: "s" },
        { label: "Sin OC", data: latest.map((s) => s.summary.sinOC), backgroundColor: COLOR.red, stack: "s" },
      ] },
      options: { ...chartDefaults(), scales: { x: { stacked: true, ticks: { color: COLOR.muted }, grid: { color: "#1F2E32" } }, y: { stacked: true, ticks: { color: COLOR.muted, precision: 0 }, grid: { color: "#1F2E32" } } } },
    });
  }
  const weekMap = new Map();
  activeSources.forEach((s) => {
    STATE.sources[s.key].snapshots.forEach((snap) => {
      if (!weekMap.has(snap.weekLabel)) weekMap.set(snap.weekLabel, {});
      weekMap.get(snap.weekLabel)[s.key] = snap.summary.totalSolped;
    });
  });
  const weeks = Array.from(weekMap.keys()).sort();
  const ctx2 = document.getElementById("chartConsHistorial");
  if (ctx2) {
    CHARTS.chartConsHistorial = new Chart(ctx2, {
      type: "line",
      data: { labels: weeks, datasets: activeSources.map((s) => ({ label: s.label, data: weeks.map((w) => weekMap.get(w)[s.key] ?? null), borderColor: s.color, backgroundColor: s.color, spanGaps: true, tension: 0.25 })) },
      options: { ...chartDefaults(), scales: { x: { ticks: { color: COLOR.muted, font: { size: 10.5 } }, grid: { color: "#1F2E32" } }, y: { ticks: { color: COLOR.muted, precision: 0 }, grid: { color: "#1F2E32" } } } },
    });
  }
}

/* ---------------------------------------------------------------------- */
/* Master render                                                           */
/* ---------------------------------------------------------------------- */
function render() {
  renderTabs();
  if (STATE.activeTab === "consolidado") {
    renderConsolidado();
    return;
  }
  const source = SOURCES.find((s) => s.key === STATE.activeTab);
  const snapshots = STATE.sources[source.key]?.snapshots || [];
  if (!snapshots.length) {
    document.getElementById("content").innerHTML = uploadBoxHtml(source.key, source.label, false);
    wireDropzones();
    return;
  }
  const selId = STATE.selectedIds[source.key] ?? snapshots[0].id;
  const selected = snapshots.find((s) => s.id === selId && s.rows) || snapshots[0];
  if (STATE.error) {
    // show inline error above dashboard by prefixing content after render
  }
  renderSourceDashboard(source, selected, snapshots);
  if (STATE.error) {
    const bar = document.querySelector(".viewing-bar");
    if (bar) bar.insertAdjacentHTML("afterend", `<div class="error-inline">${escapeHtml(STATE.error)}</div>`);
  }
}

/* ---------------------------------------------------------------------- */
/* File handling                                                           */
/* ---------------------------------------------------------------------- */
function handleFileForSource(sourceKey, file) {
  STATE.loading = true; STATE.error = ""; render();
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array", cellDates: true });
      const rows = parseWorkbook(wb);
      if (!rows.length) throw new Error("No se encontraron filas de solicitudes en el archivo.");
      const summary = computeSummary(rows);
      const now = new Date();
      const entry = { id: now.getTime(), date: now.toISOString(), weekLabel: isoWeekLabel(now), fileName: file.name, summary, rows };

      const existing = STATE.sources[sourceKey]?.snapshots || [];
      let updated = [entry, ...existing];
      updated = updated.map((s, i) => (i < MAX_DETAIL ? s : { ...s, rows: undefined }));
      STATE.sources[sourceKey] = { snapshots: updated };
      STATE.selectedIds[sourceKey] = entry.id;
      persist();
      STATE.error = "";
    } catch (err) {
      STATE.error = err.message || "No se pudo procesar el archivo.";
    } finally {
      STATE.loading = false;
      render();
    }
  };
  reader.onerror = () => { STATE.error = "No se pudo leer el archivo."; STATE.loading = false; render(); };
  reader.readAsArrayBuffer(file);
}

function deleteSnapshot(sourceKey, id) {
  const updated = (STATE.sources[sourceKey]?.snapshots || []).filter((s) => s.id !== Number(id));
  const rehydrated = updated.map((s, i) => (i < MAX_DETAIL ? s : { ...s, rows: undefined }));
  STATE.sources[sourceKey] = { snapshots: rehydrated };
  persist();
  render();
}

/* ---------------------------------------------------------------------- */
/* Informe PDF                                                             */
/* ---------------------------------------------------------------------- */
function pdfHeader(doc, subtitle) {
  doc.setFontSize(16); doc.setTextColor(20, 30, 32);
  doc.text("Bitácora de Compras — Informe de Avance", 40, 44);
  doc.setFontSize(10); doc.setTextColor(100, 110, 115);
  doc.text(`Navimag · ${subtitle}`, 40, 60);
  doc.text(`Generado el ${new Date().toLocaleString("es-CL")}`, 40, 73);
  doc.setDrawColor(210, 210, 210); doc.line(40, 82, 555, 82);
  return 100;
}

function sectionTitle(doc, text, y) {
  if (y > 740) { doc.addPage(); y = 40; }
  doc.setFontSize(11.5); doc.setTextColor(20, 30, 32);
  doc.text(text, 40, y);
  return y + 10;
}

function addFooters(doc) {
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(8.5); doc.setTextColor(150, 150, 150);
    doc.text(`Bitácora de Compras · Navimag`, 40, 820);
    doc.text(`Página ${i} de ${pageCount}`, 500, 820);
  }
}

function buildSourceSection(doc, source, snapshot, snapshots, y) {
  const grouped = groupBySolped(snapshot.rows);
  const summary = snapshot.summary;

  y = sectionTitle(doc, `${source.label} — resumen ejecutivo (semana ${snapshot.weekLabel})`, y);
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "grid",
    head: [["Indicador", "Valor"]],
    body: [
      ["Solicitudes totales", String(summary.totalSolped)],
      ["Líneas de material", String(summary.totalLineas)],
      ["OC creadas", `${summary.conOC} (${summary.totalSolped ? Math.round(100 * summary.conOC / summary.totalSolped) : 0}%)`],
      ["Sin OC (backlog)", String(summary.sinOC)],
      ["Solicitudes pendientes atrasadas (> " + DIAS_ALERTA + " días)", String(summary.atrasadas)],
      ["Repuestos con OC sin recepcionar", String(summary.sinRecepcion)],
      ["Tiempo promedio de aprobación (solicitud → liberación)", summary.avgAprobacion !== null ? `${summary.avgAprobacion} días` : "—"],
      ["Tiempo promedio de generación de OC (liberación → OC)", summary.avgGeneracionOC !== null ? `${summary.avgGeneracionOC} días` : "—"],
      ["Tiempo promedio de recepción (OC → ingreso a bodega)", summary.avgRecepcion !== null ? `${summary.avgRecepcion} días` : "—"],
      ["Tiempo total promedio (solicitud → recepción, ciclo completo)", summary.avgTotalCompleto !== null ? `${summary.avgTotalCompleto} días` : "—"],
    ],
    headStyles: { fillColor: [17, 27, 30], textColor: 255, fontSize: 9 },
    styles: { fontSize: 8.5, textColor: [40, 50, 52] },
    columnStyles: { 0: { cellWidth: 300 } },
  });
  y = doc.lastAutoTable.finalY + 22;

  // Distribución de estados
  y = sectionTitle(doc, "Distribución de estados", y);
  const estadoCounts = {};
  grouped.forEach((g) => { estadoCounts[g.estado] = (estadoCounts[g.estado] || 0) + 1; });
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "grid",
    head: [["Estado", "Cantidad", "%"]],
    body: Object.entries(estadoCounts).map(([e, c]) => [e, String(c), `${Math.round(100 * c / summary.totalSolped)}%`]),
    headStyles: { fillColor: [17, 27, 30], textColor: 255, fontSize: 9 },
    styles: { fontSize: 8.5, textColor: [40, 50, 52] },
  });
  y = doc.lastAutoTable.finalY + 22;

  // Solicitudes pendientes (todas, mayor a menor retraso)
  const pendientes = getPendientesGrouped(grouped);
  y = sectionTitle(doc, `Solicitudes pendientes (${pendientes.length}) — mayor a menor retraso`, y);
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "grid",
    head: [["SOLPED", "Material", "Solicitante", "F. Solicitud", "Estado", "Días"]],
    body: pendientes.length ? pendientes.map((g) => [g.solped, g.material, g.solicitante, fmtDate(g.fechaSolicitud), g.estado, g.dias !== null ? String(g.dias) : "—"]) : [["—", "Sin solicitudes pendientes", "", "", "", ""]],
    headStyles: { fillColor: [74, 31, 34], textColor: 255, fontSize: 9 },
    styles: { fontSize: 8, textColor: [40, 50, 52] },
    columnStyles: { 1: { cellWidth: 160 } },
  });
  y = doc.lastAutoTable.finalY + 22;

  // OC liberadas sin recepción de repuesto
  const sinRecepcion = getSinRecepcionRows(snapshot.rows);
  y = sectionTitle(doc, `OC liberadas sin recepción de repuesto (${sinRecepcion.length})`, y);
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "grid",
    head: [["SOLPED", "Material", "Proveedor", "Nº OC", "Fecha OC", "Días desde OC"]],
    body: sinRecepcion.length ? sinRecepcion.map((r) => [r.solped + (r.posicion ? `-${r.posicion}` : ""), r.material, r.proveedor || "—", r.ocNumero, fmtDate(rd(r, "fechaOC")), r.diasDesdeOC !== null ? String(r.diasDesdeOC) : "—"]) : [["—", "Todo recepcionado", "", "", "", ""]],
    headStyles: { fillColor: [77, 58, 22], textColor: 255, fontSize: 9 },
    styles: { fontSize: 8, textColor: [40, 50, 52] },
    columnStyles: { 1: { cellWidth: 150 } },
  });
  y = doc.lastAutoTable.finalY + 22;

  // Historial semanal
  if (snapshots.length > 1) {
    y = sectionTitle(doc, "Historial semanal", y);
    doc.autoTable({
      startY: y, margin: { left: 40, right: 40 }, theme: "grid",
      head: [["Semana", "Fecha carga", "Solicitudes", "Con OC", "Sin OC", "Atrasadas"]],
      body: snapshots.map((s) => [s.weekLabel, fmtDate(new Date(s.date)), String(s.summary.totalSolped), String(s.summary.conOC), String(s.summary.sinOC), String(s.summary.atrasadas)]),
      headStyles: { fillColor: [17, 27, 30], textColor: 255, fontSize: 9 },
      styles: { fontSize: 8.5, textColor: [40, 50, 52] },
    });
    y = doc.lastAutoTable.finalY + 22;
  }

  // Detalle completo
  y = sectionTitle(doc, `Detalle completo de solicitudes (${grouped.length})`, y);
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "striped",
    head: [["SOLPED", "Material", "Solicitante", "F. Solicitud", "Estado", "Nº OC / avance", "Días"]],
    body: grouped.sort((a, b) => (b.fechaSolicitud || 0) - (a.fechaSolicitud || 0)).map((g) => [
      g.solped, g.material, g.solicitante, fmtDate(g.fechaSolicitud), g.estado, `${g.conOC}/${g.numLineas}`, g.dias !== null ? String(g.dias) : "—",
    ]),
    headStyles: { fillColor: [17, 27, 30], textColor: 255, fontSize: 8.5 },
    styles: { fontSize: 7.5, textColor: [40, 50, 52] },
    columnStyles: { 1: { cellWidth: 160 } },
  });
  return doc.lastAutoTable.finalY + 30;
}

function generateSourceReport(sourceKey) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const source = SOURCES.find((s) => s.key === sourceKey);
  const snapshots = STATE.sources[sourceKey].snapshots;
  const snapshot = snapshots.find((s) => s.rows) || snapshots[0];
  let y = pdfHeader(doc, `Informe de ${source.label} — semana ${snapshot.weekLabel}`);
  buildSourceSection(doc, source, snapshot, snapshots, y);
  addFooters(doc);
  doc.save(`Informe_${source.label}_${snapshot.weekLabel}.pdf`);
}

function generateConsolidadoReport() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const activeSources = SOURCES.filter((s) => STATE.sources[s.key]?.snapshots?.length);
  let y = pdfHeader(doc, "Informe consolidado — Taller, Esperanza y Dalka");

  // Resumen combinado
  const latest = activeSources.map((s) => ({ ...s, summary: STATE.sources[s.key].snapshots[0].summary, weekLabel: STATE.sources[s.key].snapshots[0].weekLabel }));
  const combined = latest.reduce((acc, s) => ({
    totalSolped: acc.totalSolped + s.summary.totalSolped, conOC: acc.conOC + s.summary.conOC, sinOC: acc.sinOC + s.summary.sinOC,
    atrasadas: acc.atrasadas + s.summary.atrasadas, sinRecepcion: acc.sinRecepcion + (s.summary.sinRecepcion || 0),
  }), { totalSolped: 0, conOC: 0, sinOC: 0, atrasadas: 0, sinRecepcion: 0 });

  y = sectionTitle(doc, "Resumen consolidado", y);
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "grid",
    head: [["Indicador", "Valor"]],
    body: [
      ["Solicitudes totales (todas las fuentes)", String(combined.totalSolped)],
      ["OC creadas", `${combined.conOC} (${combined.totalSolped ? Math.round(100 * combined.conOC / combined.totalSolped) : 0}%)`],
      ["Sin OC (backlog)", String(combined.sinOC)],
      ["Pendientes atrasadas (> " + DIAS_ALERTA + " días)", String(combined.atrasadas)],
      ["Repuestos con OC sin recepcionar", String(combined.sinRecepcion)],
    ],
    headStyles: { fillColor: [17, 27, 30], textColor: 255, fontSize: 9 },
    styles: { fontSize: 8.5, textColor: [40, 50, 52] },
  });
  y = doc.lastAutoTable.finalY + 22;

  y = sectionTitle(doc, "Resumen por fuente", y);
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 }, theme: "grid",
    head: [["Fuente", "Semana", "Solicitudes", "Con OC", "Sin OC", "Atrasadas", "Sin recepción"]],
    body: latest.map((s) => [s.label, s.weekLabel, String(s.summary.totalSolped), String(s.summary.conOC), String(s.summary.sinOC), String(s.summary.atrasadas), String(s.summary.sinRecepcion ?? "—")]),
    headStyles: { fillColor: [17, 27, 30], textColor: 255, fontSize: 9 },
    styles: { fontSize: 8.5, textColor: [40, 50, 52] },
  });
  y = doc.lastAutoTable.finalY + 30;

  // Secciones detalladas por fuente (cada una en páginas nuevas)
  activeSources.forEach((source) => {
    doc.addPage();
    y = 40;
    const snapshots = STATE.sources[source.key].snapshots;
    const snapshot = snapshots.find((s) => s.rows) || snapshots[0];
    y = buildSourceSection(doc, source, snapshot, snapshots, y);
  });

  addFooters(doc);
  doc.save(`Informe_Consolidado_Navimag_${isoWeekLabel(new Date())}.pdf`);
}

/* ---------------------------------------------------------------------- */
/* Eventos                                                                  */
/* ---------------------------------------------------------------------- */
function wireDropzones() {
  document.querySelectorAll("[data-dropzone]").forEach((el) => {
    el.addEventListener("dragover", (e) => { e.preventDefault(); el.classList.add("dragover"); });
    el.addEventListener("dragleave", () => el.classList.remove("dragover"));
    el.addEventListener("drop", (e) => {
      e.preventDefault(); el.classList.remove("dragover");
      const file = e.dataTransfer.files?.[0];
      if (file) handleFileForSource(el.dataset.dropzone, file);
    });
  });
}

document.getElementById("tabs").addEventListener("click", (e) => {
  const btn = e.target.closest("[data-action='set-tab']");
  if (!btn) return;
  STATE.activeTab = btn.dataset.tab;
  UI = { search: "", estadoFilter: "TODOS", view: "solped", sortDias: false };
  render();
});

document.getElementById("content").addEventListener("click", (e) => {
  const genBtn = e.target.closest("[data-action='generate-pdf']");
  if (genBtn) {
    const src = genBtn.dataset.source;
    if (src === "consolidado") generateConsolidadoReport(); else generateSourceReport(src);
    return;
  }
  const selBtn = e.target.closest("[data-action='select-snapshot']");
  if (selBtn) { STATE.selectedIds[STATE.activeTab] = Number(selBtn.dataset.id); render(); return; }
  const delBtn = e.target.closest("[data-action='delete-snapshot']");
  if (delBtn) { deleteSnapshot(STATE.activeTab, delBtn.dataset.id); return; }
  const viewBtn = e.target.closest("[data-action='toggle-view']");
  if (viewBtn) { UI.view = UI.view === "solped" ? "linea" : "solped"; renderTableBody(); viewBtn.textContent = UI.view === "solped" ? "▤ Ver por línea" : "▥ Ver por solicitud"; return; }
  const sortBtn = e.target.closest("[data-action='toggle-sort']");
  if (sortBtn) { UI.sortDias = !UI.sortDias; sortBtn.classList.toggle("active", UI.sortDias); renderTableBody(); return; }
});

document.getElementById("content").addEventListener("input", (e) => {
  if (e.target.id === "searchInput") { UI.search = e.target.value; renderTableBody(); }
});

document.getElementById("content").addEventListener("change", (e) => {
  if (e.target.id === "estadoSelect") { UI.estadoFilter = e.target.value; renderTableBody(); }
  if (e.target.matches("[data-action='upload-file']")) {
    const file = e.target.files?.[0];
    if (file) handleFileForSource(e.target.dataset.source, file);
  }
});

/* ---------------------------------------------------------------------- */
/* Init                                                                     */
/* ---------------------------------------------------------------------- */
loadFromStorage();
render();
