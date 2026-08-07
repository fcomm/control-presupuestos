import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";
import { useCollection } from "./useCollection";

/* ----------------------------------------------------------------------
   TOKENS
---------------------------------------------------------------------- */
const T = {
  bg: "#0E1416",
  panel: "#161F23",
  panelAlt: "#1C282D",
  border: "#2B3940",
  borderSoft: "#222E33",
  text: "#E7EDEF",
  textDim: "#8FA0A8",
  textFaint: "#5C6E76",
  amber: "#E3A23C",
  amberDim: "#7A5A26",
  teal: "#3FA796",
  red: "#D9645C",
  blue: "#5B8FB0",
  fontUI: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
  fontMono: "ui-monospace, SFMono-Regular, 'SF Mono', Consolas, 'Liberation Mono', monospace",
};

const RUBROS = [
  { rubro: "Materiales e Insumos", categorias: ["Material eléctrico","Material mecánico","Ferretería","Instrumentación","Tubería y conexiones","Consumibles industriales","Consumibles de oficina","Papelería","EPP","Importación y logística","Aranceles e impuestos de importación","Permisos regulatorios de importación"] },
  { rubro: "Equipos y Herramientas", categorias: ["Herramienta menor","Herramienta especializada","Equipos de medición","Equipos de cómputo","Equipos de comunicación","Maquinaria","Mobiliario","Importación y logística","Aranceles e impuestos de importación","Permisos regulatorios de importación"] },
  { rubro: "Productos Químicos", categorias: ["Productos químicos de operación","Reactivos","Solventes","Aditivos","Químicos de laboratorio","Importación y logística","Aranceles e impuestos de importación","Permisos regulatorios de importación"] },
  { rubro: "Servicios de Mantenimiento", categorias: ["Mantenimiento vehicular","Mantenimiento de equipos","Mantenimiento de instalaciones","Calibración","Reparaciones"] },
  { rubro: "Servicios Operativos", categorias: ["Arrendamientos","Maniobras","Transporte y logística","Fletes","Servicios de campo","Laboratorios","Servicios especializados"] },
  { rubro: "Servicios Administrativos", categorias: ["Mensajería","Limpieza","Vigilancia","Capacitación","Consultoría","Honorarios profesionales","Reclutamiento"] },
  { rubro: "Servicios Básicos", categorias: ["Energía eléctrica","Agua","Telefonía fija","Telefonía móvil","Internet","Gas","Recolección de residuos"] },
  { rubro: "Tecnologías de Información", categorias: ["Software","Licencias","Suscripciones","Hosting","Telecomunicaciones","Desarrollo de software","Servicios en la nube","Ciberseguridad"] },
  { rubro: "Vehículos", categorias: ["Combustible","Casetas","Refacciones","Accesorios","Seguros","Verificación","Tenencias","Placas","Permisos vehiculares","Llantas","Lubricantes"] },
  { rubro: "Viajes y Viáticos", categorias: ["Hospedaje","Boletos de avión","Transporte terrestre","Alimentación","Renta de vehículos","Viáticos diversos"] },
  { rubro: "Promoción e Imagen", categorias: ["Publicidad","Merchandising","Eventos","Material publicitario","Uniformes corporativos"] },
  { rubro: "Cumplimiento Legal", categorias: ["Permisos regulatorios","Licencias","Certificaciones","Estudios","Trámites gubernamentales","Auditorías"] },
  { rubro: "Gastos Financieros e Impuestos", categorias: ["Comisiones bancarias","Derechos","Impuestos","Gastos notariales","Fianzas"] },
  { rubro: "Nómina y Personal", categorias: ["Sueldos y salarios","Prestaciones de ley","Prestaciones adicionales","Cuotas IMSS/Infonavit","Impuesto sobre nómina (ISN)","Finiquitos y liquidaciones","Bonos y comisiones"] },
  { rubro: "Otros", categorias: ["Gastos extraordinarios","Donativos","Diversos"] },
];

const UNIDADES_BASE = ["OSB", "CTM", "ISE"];

const MONEDAS = ["MXN", "USD"];
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const uid = () => (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2, 10));

const MES_ABR = {
  Enero: "ENE", Febrero: "FEB", Marzo: "MAR", Abril: "ABR", Mayo: "MAY", Junio: "JUN",
  Julio: "JUL", Agosto: "AGO", Septiembre: "SEP", Octubre: "OCT", Noviembre: "NOV", Diciembre: "DIC",
};

function folioPrefix(unidad, mes, anio) {
  const abr = MES_ABR[mes] || String(mes || "").slice(0, 3).toUpperCase() || "GEN";
  const yy = String(anio || new Date().getFullYear()).slice(-2);
  return `${unidad}-${abr}${yy}-`;
}

// Returns the next free NNN suffix for a given prefix, based on folios already in use.
function nextFolioNumber(prefix, existingFolios) {
  let max = 0;
  existingFolios.forEach((f) => {
    if (f && f.toUpperCase().startsWith(prefix)) {
      const n = parseInt(f.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return max + 1;
}

function autoFolio(unidad, mes, anio, existingFolios) {
  const prefix = folioPrefix(unidad, mes, anio);
  return prefix + String(nextFolioNumber(prefix, existingFolios)).padStart(3, "0");
}

const money = (n, moneda = "MXN") =>
  (moneda === "USD" ? "$" : "$") +
  (Number(n) || 0).toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) +
  (moneda === "USD" ? " USD" : "");

/* ----------------------------------------------------------------------
   PRORRATEO
---------------------------------------------------------------------- */
// Given a partida's "proyecto" marker and the unidad's proyecto config,
// return [{proyecto, fraccion}] fractions summing to 1.
function resolverProrrateo(marcador, proyectosUnidad) {
  if (!marcador) return [];
  const directo = proyectosUnidad.find((p) => p.nombre === marcador);
  if (directo) return [{ proyecto: directo.nombre, fraccion: 1 }];

  if (marcador === "Todos") {
    const total = proyectosUnidad.reduce((s, p) => s + Number(p.pct || 0), 0) || 1;
    return proyectosUnidad.map((p) => ({ proyecto: p.nombre, fraccion: Number(p.pct || 0) / total }));
  }
  // "<Grupo> Gral" e.g. "Desh Gral" -> group = "Desh"
  const m = /^(.*) Gral$/.exec(marcador);
  if (m) {
    const grupo = m[1];
    const miembros = proyectosUnidad.filter((p) => p.grupo === grupo);
    const total = miembros.reduce((s, p) => s + Number(p.pct || 0), 0) || 1;
    return miembros.map((p) => ({ proyecto: p.nombre, fraccion: Number(p.pct || 0) / total }));
  }
  return [];
}

function marcadoresDisponibles(proyectosUnidad) {
  const grupos = [...new Set(proyectosUnidad.map((p) => p.grupo))].filter(Boolean);
  const marcadores = grupos.filter((g) => proyectosUnidad.filter((p) => p.grupo === g).length > 1).map((g) => `${g} Gral`);
  return [...proyectosUnidad.map((p) => p.nombre), ...marcadores, "Todos"];
}

/* ----------------------------------------------------------------------
   IMPORTACION DESDE EXCEL (hojas RawData-*)
---------------------------------------------------------------------- */
const UNIDAD_KEYS = ["OSB", "CTM", "ISE"];

function normHeader(h) {
  return String(h ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\n/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
}

function findCol(headers, ...keywordSets) {
  const norm = headers.map(normHeader);
  for (const keywords of keywordSets) {
    const idx = norm.findIndex((h) => keywords.every((k) => h.includes(k)));
    if (idx !== -1) return idx;
  }
  return -1;
}

function findExactCol(headers, candidates) {
  const norm = headers.map(normHeader);
  for (const cand of candidates) {
    const idx = norm.findIndex((h) => h === cand);
    if (idx !== -1) return idx;
  }
  return -1;
}

// Parses an ArrayBuffer (.xlsx) and returns { rows, sheetsFound, sheetsIgnored }
// rows: array of partida objects ready to merge into state.
function parsePresupuestoWorkbook(arrayBuffer, options = {}) {
  const { anioDefault = new Date().getFullYear(), existingPartidas = [] } = options;
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const rows = [];
  const sheetsFound = [];
  const sheetsIgnored = [];
  const counters = {}; // prefix -> highest number in use so far (seed from existing + this batch)

  const registerFolio = (prefix, folio) => {
    if (!(prefix in counters)) counters[prefix] = nextFolioNumber(prefix, existingPartidas.map((p) => p.folio)) - 1;
    if (folio && folio.toUpperCase().startsWith(prefix)) {
      const n = parseInt(folio.slice(prefix.length), 10);
      if (!isNaN(n) && n > counters[prefix]) counters[prefix] = n;
    }
  };
  const assignFolio = (prefix) => {
    if (!(prefix in counters)) counters[prefix] = nextFolioNumber(prefix, existingPartidas.map((p) => p.folio)) - 1;
    counters[prefix] += 1;
    return prefix + String(counters[prefix]).padStart(3, "0");
  };

  wb.SheetNames.forEach((sheetName) => {
    const m = /^RawData-(.+)$/i.exec(sheetName.trim());
    if (!m) return;
    const unidad = m[1].toUpperCase();
    if (!UNIDAD_KEYS.includes(unidad)) {
      sheetsIgnored.push(sheetName);
      return;
    }
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    if (!aoa.length) return;
    const headers = aoa[0];

    const col = {
      mes: findCol(headers, ["mes"]),
      anio: findCol(headers, ["ano"], ["anio"]),
      smi: findCol(headers, ["smi"]),
      concepto: findCol(headers, ["concepto"]),
      subtotalMXN: findCol(headers, ["sub total", "mxn"]),
      subtotalUSD: findCol(headers, ["sub total", "usd"]),
      totalMXN: findCol(headers, ["total", "mxn"]),
      moneda: findCol(headers, ["moneda"]),
      proveedor: findCol(headers, ["proveedor"]),
      proyecto: findCol(headers, ["proyecto"]),
      rubro: findCol(headers, ["rubro"]),
      categoria: findCol(headers, ["categoria"]),
      folio: findExactCol(headers, ["id", "folio", "a partida", "partida", "no. partida"]),
    };

    let count = 0;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const concepto = col.concepto !== -1 ? row[col.concepto] : null;
      if (concepto === null || concepto === undefined || String(concepto).trim() === "") continue;

      const moneda = col.moneda !== -1 && row[col.moneda] ? String(row[col.moneda]).trim().toUpperCase() : "MXN";
      let monto = 0;
      if (moneda === "USD" && col.subtotalUSD !== -1 && typeof row[col.subtotalUSD] === "number") {
        monto = row[col.subtotalUSD];
      } else if (col.subtotalMXN !== -1 && typeof row[col.subtotalMXN] === "number") {
        monto = row[col.subtotalMXN];
      } else if (col.totalMXN !== -1 && typeof row[col.totalMXN] === "number") {
        monto = row[col.totalMXN];
      }

      const mes = (col.mes !== -1 && row[col.mes]) ? String(row[col.mes]).trim() : "";
      const anio = (col.anio !== -1 && row[col.anio]) ? Number(row[col.anio]) : anioDefault;
      const prefix = folioPrefix(unidad, mes, anio);
      let folio = (col.folio !== -1 && row[col.folio] !== null && row[col.folio] !== undefined) ? String(row[col.folio]).trim() : "";
      if (folio) registerFolio(prefix, folio); else folio = assignFolio(prefix);

      rows.push({
        id: uid(),
        unidad,
        mes,
        anio,
        smi: (col.smi !== -1 && row[col.smi]) ? String(row[col.smi]).trim() : "",
        concepto: String(concepto).trim(),
        rubro: (col.rubro !== -1 && row[col.rubro]) ? String(row[col.rubro]).trim() : "Otros",
        categoria: (col.categoria !== -1 && row[col.categoria]) ? String(row[col.categoria]).trim() : "Diversos",
        proyecto: (col.proyecto !== -1 && row[col.proyecto]) ? String(row[col.proyecto]).trim() : "",
        monto_estimado: monto,
        moneda,
        folio,
      });
      count++;
    }
    sheetsFound.push({ sheetName, unidad, count });
  });

  return { rows, sheetsFound, sheetsIgnored };
}

function toISODate(val) {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val)) return val.toISOString().slice(0, 10);
  const s = String(val).trim();
  const m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/.exec(s);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  const m2 = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/.exec(s);
  if (m2) {
    const [, y, mo, d] = m2;
    return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return s;
}

// Parses a workbook looking for a "transacciones reales" sheet — identified by
// having Día + Importe + A Partida columns, regardless of sheet name.
// Matches each row to a partida via its `folio` field (and unit prefix of the folio).
function parseTransaccionesWorkbook(arrayBuffer, partidas) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const rows = [];
  const sheetsFound = [];
  let matched = 0, unmatched = 0;

  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    if (!aoa.length) return;
    const headers = aoa[0];

    const hasImporte = findCol(headers, ["importe"]) !== -1;
    const hasFolio = findCol(headers, ["a partida"]) !== -1 || findExactCol(headers, ["partida", "folio"]) !== -1;
    const hasDia = findExactCol(headers, ["dia", "fecha"]) !== -1 || findCol(headers, ["dia"]) !== -1;
    if (!(hasImporte && hasFolio && hasDia)) return;

    const col = {
      dia: findExactCol(headers, ["dia", "fecha"]) !== -1 ? findExactCol(headers, ["dia", "fecha"]) : findCol(headers, ["dia"]),
      smi: findCol(headers, ["smi"]),
      solicitante: findCol(headers, ["solicitante"]),
      area: findExactCol(headers, ["area"]),
      proveedor: findCol(headers, ["nombre", "denominacion"], ["razon", "social"], ["proveedor"]),
      concepto: findCol(headers, ["concepto"]),
      importe: findCol(headers, ["importe"]),
      moneda: findCol(headers, ["moneda"]),
      folio: findCol(headers, ["a partida"]) !== -1 ? findCol(headers, ["a partida"]) : findExactCol(headers, ["partida", "folio"]),
    };

    let count = 0;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const importeRaw = col.importe !== -1 ? row[col.importe] : null;
      const importe = typeof importeRaw === "number" ? importeRaw : parseFloat(String(importeRaw || "").replace(/[^0-9.-]/g, ""));
      if (!importe || isNaN(importe)) continue;
      const folioRaw = col.folio !== -1 ? row[col.folio] : null;
      const folio = folioRaw ? String(folioRaw).trim() : "";
      if (!folio) continue;

      const prefixMatch = /^([A-Za-z]+)-/.exec(folio);
      const unidad_detectada = prefixMatch && UNIDAD_KEYS.includes(prefixMatch[1].toUpperCase()) ? prefixMatch[1].toUpperCase() : null;

      const partida = partidas.find(
        (p) => p.folio && p.folio.trim().toUpperCase() === folio.toUpperCase() &&
               (!unidad_detectada || p.unidad === unidad_detectada)
      );
      if (partida) matched++; else unmatched++;

      rows.push({
        id: uid(),
        partida_id: partida ? partida.id : "",
        folio_original: folio,
        unidad_detectada: partida ? partida.unidad : unidad_detectada,
        dia: col.dia !== -1 ? toISODate(row[col.dia]) : "",
        solicitante: (col.solicitante !== -1 && row[col.solicitante]) ? String(row[col.solicitante]).trim() : "",
        area: (col.area !== -1 && row[col.area]) ? String(row[col.area]).trim() : "",
        proveedor: (col.proveedor !== -1 && row[col.proveedor]) ? String(row[col.proveedor]).trim() : "",
        concepto_detallado: (col.concepto !== -1 && row[col.concepto]) ? String(row[col.concepto]).trim() : "",
        importe,
        moneda: (col.moneda !== -1 && row[col.moneda]) ? String(row[col.moneda]).trim().toUpperCase() : "MXN",
      });
      count++;
    }
    if (count) sheetsFound.push({ sheetName, count });
  });

  return { rows, sheetsFound, matched, unmatched };
}

/* ----------------------------------------------------------------------
   SMALL UI PRIMITIVES
---------------------------------------------------------------------- */
function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12, color: T.textDim }}>
      <span style={{ textTransform: "uppercase", letterSpacing: "0.06em", fontSize: 10.5 }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  background: T.panelAlt,
  border: `1px solid ${T.border}`,
  borderRadius: 6,
  color: T.text,
  padding: "8px 10px",
  fontSize: 13,
  fontFamily: T.fontUI,
  outline: "none",
};

function TextInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}
function Select(props) {
  return <select {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}

function Button({ children, variant = "primary", ...rest }) {
  const styles = {
    primary: { background: T.amber, color: "#1A1206", border: `1px solid ${T.amber}` },
    ghost: { background: "transparent", color: T.textDim, border: `1px solid ${T.border}` },
    danger: { background: "transparent", color: T.red, border: `1px solid ${T.red}44` },
  };
  return (
    <button
      {...rest}
      style={{
        ...styles[variant],
        borderRadius: 6,
        padding: "8px 14px",
        fontSize: 12.5,
        fontWeight: 600,
        cursor: "pointer",
        fontFamily: T.fontUI,
        letterSpacing: "0.01em",
      }}
    >
      {children}
    </button>
  );
}

function Pill({ children, tone = "dim" }) {
  const tones = {
    dim: { color: T.textDim, background: "#22303680", border: `1px solid ${T.border}` },
    amber: { color: T.amber, background: "#E3A23C1A", border: `1px solid ${T.amberDim}` },
    teal: { color: T.teal, background: "#3FA7961A", border: `1px solid #3FA79644` },
    red: { color: T.red, background: "#D9645C1A", border: `1px solid #D9645C44` },
  };
  return (
    <span style={{
      ...tones[tone], padding: "2px 8px", borderRadius: 999, fontSize: 10.5,
      fontFamily: T.fontMono, letterSpacing: "0.02em", whiteSpace: "nowrap",
    }}>
      {children}
    </span>
  );
}

/* Industrial gauge — the signature element */
function Gauge({ pct, size = 92 }) {
  const clamped = Math.max(0, Math.min(pct, 150));
  const angle = -120 + (clamped / 150) * 240; // -120..120 deg sweep over 0..150%
  const color = pct > 105 ? T.red : pct > 90 ? T.amber : T.teal;
  const r = size / 2 - 8;
  const cx = size / 2, cy = size / 2;
  const arc = (a0, a1, radius) => {
    const toXY = (a) => [cx + radius * Math.sin((a * Math.PI) / 180), cy - radius * Math.cos((a * Math.PI) / 180)];
    const [x0, y0] = toXY(a0);
    const [x1, y1] = toXY(a1);
    const large = a1 - a0 > 180 ? 1 : 0;
    return `M ${x0} ${y0} A ${radius} ${radius} 0 ${large} 1 ${x1} ${y1}`;
  };
  const needle = (() => {
    const a = angle;
    const x = cx + (r - 6) * Math.sin((a * Math.PI) / 180);
    const y = cy - (r - 6) * Math.cos((a * Math.PI) / 180);
    return { x, y };
  })();
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      <path d={arc(-120, 120, r)} stroke={T.border} strokeWidth={7} fill="none" strokeLinecap="round" />
      <path d={arc(-120, Math.min(angle, 120), r)} stroke={color} strokeWidth={7} fill="none" strokeLinecap="round" />
      <line x1={cx} y1={cy} x2={needle.x} y2={needle.y} stroke={T.text} strokeWidth={2} />
      <circle cx={cx} cy={cy} r={3.5} fill={T.text} />
      <text x={cx} y={cy + size / 2 - 10} textAnchor="middle" fontSize={13} fontFamily={T.fontMono} fill={color} fontWeight={700}>
        {pct.toFixed(0)}%
      </text>
    </svg>
  );
}

/* ----------------------------------------------------------------------
   TABS: DASHBOARD
---------------------------------------------------------------------- */
function Dashboard({ unidad, unidades, partidas, transacciones }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const idsPartidas = new Set(partidasUnidad.map((p) => p.id));
  const transUnidad = transacciones.filter((t) => idsPartidas.has(t.partida_id));

  const porProyecto = useMemo(() => {
    const map = {};
    proyectosUnidad.forEach((p) => { map[p.nombre] = { proyecto: p.nombre, presupuestado: 0, ejecutado: 0 }; });

    partidasUnidad.forEach((partida) => {
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!map[proyecto]) map[proyecto] = { proyecto, presupuestado: 0, ejecutado: 0 };
        map[proyecto].presupuestado += (Number(partida.monto_estimado) || 0) * fraccion;
      });
    });

    transUnidad.forEach((t) => {
      const partida = partidasUnidad.find((p) => p.id === t.partida_id);
      if (!partida) return;
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!map[proyecto]) map[proyecto] = { proyecto, presupuestado: 0, ejecutado: 0 };
        map[proyecto].ejecutado += (Number(t.importe) || 0) * fraccion;
      });
    });

    return Object.values(map);
  }, [proyectosUnidad, partidasUnidad, transUnidad]);

  const porRubro = useMemo(() => {
    const map = {};
    partidasUnidad.forEach((p) => {
      map[p.rubro] = map[p.rubro] || { rubro: p.rubro, presupuestado: 0, ejecutado: 0 };
      map[p.rubro].presupuestado += Number(p.monto_estimado) || 0;
    });
    transUnidad.forEach((t) => {
      const partida = partidasUnidad.find((p) => p.id === t.partida_id);
      if (!partida) return;
      map[partida.rubro] = map[partida.rubro] || { rubro: partida.rubro, presupuestado: 0, ejecutado: 0 };
      map[partida.rubro].ejecutado += Number(t.importe) || 0;
    });
    return Object.values(map).sort((a, b) => b.presupuestado - a.presupuestado);
  }, [partidasUnidad, transUnidad]);

  const porMes = useMemo(() => {
    const map = {};
    MESES.forEach((m) => { map[m] = { mes: m, presupuestado: 0, ejecutado: 0 }; });
    partidasUnidad.forEach((p) => {
      if (map[p.mes]) map[p.mes].presupuestado += Number(p.monto_estimado) || 0;
    });
    transUnidad.forEach((t) => {
      const partida = partidasUnidad.find((p) => p.id === t.partida_id);
      if (!partida || !map[partida.mes]) return;
      map[partida.mes].ejecutado += Number(t.importe) || 0;
    });
    return Object.values(map).filter((m) => m.presupuestado > 0 || m.ejecutado > 0);
  }, [partidasUnidad, transUnidad]);

  const totalPresupuestado = partidasUnidad.reduce((s, p) => s + (Number(p.monto_estimado) || 0), 0);
  const totalEjecutado = transUnidad.reduce((s, t) => s + (Number(t.importe) || 0), 0);
  const totalPct = totalPresupuestado ? (totalEjecutado / totalPresupuestado) * 100 : 0;

  const COLORS = [T.amber, T.teal, T.blue, T.red, "#8B6FB0", "#B0955B", "#5BA0B0", "#B05B7A"];

  if (!partidasUnidad.length) {
    return (
      <EmptyState
        title="Sin partidas registradas"
        body={`Todavía no hay partidas presupuestales para ${unidad}. Agrega la primera en la pestaña Partidas.`}
      />
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <KpiCard label="Presupuestado" value={money(totalPresupuestado)} />
        <KpiCard label="Ejecutado" value={money(totalEjecutado)} accent={totalPct > 100 ? T.red : T.teal} />
        <KpiCard label="Disponible" value={money(totalPresupuestado - totalEjecutado)} accent={totalPresupuestado - totalEjecutado < 0 ? T.red : T.text} />
        <div style={{ ...panelStyle, display: "flex", alignItems: "center", gap: 14, padding: "12px 20px" }}>
          <Gauge pct={totalPct} />
          <div>
            <div style={{ fontSize: 10.5, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em" }}>Avance global</div>
            <div style={{ fontSize: 12, color: T.textFaint, maxWidth: 140 }}>% ejecutado sobre presupuesto de {unidad}</div>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 20 }}>
        <Panel title="Presupuesto vs. ejecutado por proyecto" subtitle="Gastos compartidos ya prorrateados según su marcador">
          <ResponsiveContainer width="100%" height={Math.max(220, porProyecto.length * 34)}>
            <BarChart data={porProyecto} layout="vertical" margin={{ left: 8, right: 16 }}>
              <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} horizontal={false} />
              <XAxis type="number" tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
              <YAxis type="category" dataKey="proyecto" width={110} tick={{ fill: T.textDim, fontSize: 11 }} stroke={T.border} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
              <Bar dataKey="presupuestado" name="Presupuestado" fill={T.border} radius={[0,3,3,0]} />
              <Bar dataKey="ejecutado" name="Ejecutado" fill={T.amber} radius={[0,3,3,0]} />
            </BarChart>
          </ResponsiveContainer>
        </Panel>

        <Panel title="Distribución por rubro" subtitle="Sobre el total presupuestado">
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={porRubro} dataKey="presupuestado" nameKey="rubro" innerRadius={55} outerRadius={90} paddingAngle={2}>
                {porRubro.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} stroke={T.panel} strokeWidth={2} />)}
              </Pie>
              <Tooltip content={<ChartTooltip />} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 4 }}>
            {porRubro.slice(0, 8).map((r, i) => (
              <div key={r.rubro} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 10.5, color: T.textDim }}>
                <span style={{ width: 8, height: 8, borderRadius: 2, background: COLORS[i % COLORS.length] }} />
                {r.rubro}
              </div>
            ))}
          </div>
        </Panel>
      </div>

      {porMes.length > 1 && (
        <Panel title="Tendencia mensual" subtitle="Presupuestado vs. ejecutado">
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={porMes}>
              <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} />
              <XAxis dataKey="mes" tick={{ fill: T.textFaint, fontSize: 10 }} stroke={T.border} />
              <YAxis tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
              <Line type="monotone" dataKey="presupuestado" name="Presupuestado" stroke={T.textFaint} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="ejecutado" name="Ejecutado" stroke={T.amber} strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </Panel>
      )}

      <Panel title="Avance por proyecto" subtitle="Detalle con marcador de prorrateo aplicado">
        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                {["Proyecto","Presupuestado","Ejecutado","Disponible","Avance"].map((h) => (
                  <th key={h} style={thStyle}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {porProyecto.map((r) => {
                const pct = r.presupuestado ? (r.ejecutado / r.presupuestado) * 100 : 0;
                const tone = pct > 100 ? "red" : pct > 85 ? "amber" : "teal";
                return (
                  <tr key={r.proyecto}>
                    <td style={tdStyle}>{r.proyecto}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(r.presupuestado)}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(r.ejecutado)}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, color: r.presupuestado - r.ejecutado < 0 ? T.red : T.text }}>
                      {money(r.presupuestado - r.ejecutado)}
                    </td>
                    <td style={tdStyle}><Pill tone={tone}>{pct.toFixed(0)}%</Pill></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Panel>
    </div>
  );
}

function ChartTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 6, padding: "8px 12px", fontSize: 11.5 }}>
      <div style={{ color: T.text, marginBottom: 4, fontWeight: 600 }}>{label ?? payload[0]?.payload?.rubro ?? payload[0]?.payload?.proyecto}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, fontFamily: T.fontMono }}>{p.name}: {money(p.value)}</div>
      ))}
    </div>
  );
}

function KpiCard({ label, value, accent = T.text }) {
  return (
    <div style={{ ...panelStyle, padding: "14px 20px", minWidth: 160 }}>
      <div style={{ fontSize: 10.5, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
      <div style={{ fontSize: 21, fontFamily: T.fontMono, color: accent, marginTop: 4, fontWeight: 600 }}>{value}</div>
    </div>
  );
}

function Panel({ title, subtitle, children, right }) {
  return (
    <div style={panelStyle}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11, color: T.textFaint, marginTop: 2 }}>{subtitle}</div>}
        </div>
        {right}
      </div>
      {children}
    </div>
  );
}

function EmptyState({ title, body }) {
  return (
    <div style={{ ...panelStyle, textAlign: "center", padding: "48px 24px" }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: T.text, marginBottom: 6 }}>{title}</div>
      <div style={{ fontSize: 12.5, color: T.textDim, maxWidth: 420, margin: "0 auto" }}>{body}</div>
    </div>
  );
}

const panelStyle = { background: T.panel, border: `1px solid ${T.border}`, borderRadius: 10, padding: 18 };

// Sortable table header + generic multi-type comparator (used by Partidas & Transacciones tables)
function SortableTh({ label, sortKey, sort, setSort, width }) {
  const active = sort.key === sortKey;
  return (
    <th
      style={{ ...thStyle, cursor: "pointer", userSelect: "none", width, color: active ? T.amber : thStyle.color }}
      onClick={() => setSort((s) => (s.key === sortKey ? { key: sortKey, dir: s.dir === "asc" ? "desc" : "asc" } : { key: sortKey, dir: "asc" }))}
    >
      {label}{active ? (sort.dir === "asc" ? " ▲" : " ▼") : ""}
    </th>
  );
}

function sortRows(rows, sort, extractors = {}) {
  if (!sort.key) return rows;
  const get = extractors[sort.key] || ((r) => r[sort.key]);
  const sorted = [...rows].sort((a, b) => {
    let va = get(a), vb = get(b);
    if (va == null) va = "";
    if (vb == null) vb = "";
    if (typeof va === "string") va = va.toLowerCase();
    if (typeof vb === "string") vb = vb.toLowerCase();
    if (va < vb) return sort.dir === "asc" ? -1 : 1;
    if (va > vb) return sort.dir === "asc" ? 1 : -1;
    return 0;
  });
  return sorted;
}

const tableStyle = { width: "100%", borderCollapse: "collapse", fontSize: 12.5 };
const thStyle = { textAlign: "left", padding: "8px 10px", color: T.textFaint, fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.05em", borderBottom: `1px solid ${T.border}` };
const tdStyle = { padding: "9px 10px", borderBottom: `1px solid ${T.borderSoft}`, color: T.text };

/* ----------------------------------------------------------------------
   TABS: PARTIDAS
---------------------------------------------------------------------- */
function ImportarExcelPanel({ partidas, partidasApi }) {
  const inputRef = useRef(null);
  const [preview, setPreview] = useState(null); // { rows, sheetsFound, sheetsIgnored, fileName }
  const [modo, setModo] = useState("agregar"); // 'agregar' | 'reemplazar'
  const [anioObjetivo, setAnioObjetivo] = useState(new Date().getFullYear());
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [importing, setImporting] = useState(false);

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(""); setStatus(""); setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const { rows, sheetsFound, sheetsIgnored } = parsePresupuestoWorkbook(buf, { anioDefault: anioObjetivo, existingPartidas: partidas });
      if (!rows.length) {
        setError('No encontré filas con Concepto en hojas "RawData-OSB / RawData-CTM / RawData-ISE" dentro de este archivo.');
        return;
      }
      setPreview({ rows, sheetsFound, sheetsIgnored, fileName: file.name });
    } catch (err) {
      setError("No pude leer el archivo. Verifica que sea un .xlsx válido.");
    } finally {
      if (inputRef.current) inputRef.current.value = "";
    }
  };

  const confirmar = async () => {
    if (!preview) return;
    setImporting(true);
    try {
      const unidadesImportadas = [...new Set(preview.rows.map((r) => r.unidad))];
      if (modo === "reemplazar") {
        await partidasApi.removeWhere("unidad", unidadesImportadas);
      }
      await partidasApi.bulkInsert(preview.rows);
      setStatus(`Importadas ${preview.rows.length} partidas (${unidadesImportadas.join(", ")}).`);
      setPreview(null);
    } catch (err) {
      setError("Ocurrió un error al importar en Supabase: " + (err.message || err));
    } finally {
      setImporting(false);
    }
  };

  return (
    <Panel title="Carga masiva desde Excel" subtitle='Sube el libro con hojas "RawData-OSB", "RawData-CTM" y/o "RawData-ISE" — se detectan automáticamente'>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <input ref={inputRef} type="file" accept=".xlsx" onChange={onFile}
          style={{ fontSize: 12, color: T.textDim }} />
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: T.textDim }}>
          Año objetivo (si el archivo no trae columna Año)
          <TextInput type="number" value={anioObjetivo} onChange={(e) => setAnioObjetivo(Number(e.target.value))} style={{ width: 80 }} />
        </label>
        {status && <Pill tone="teal">{status}</Pill>}
      </div>
      {error && <div style={{ marginTop: 10, fontSize: 12, color: T.red }}>{error}</div>}

      {preview && (
        <div style={{ marginTop: 16, borderTop: `1px solid ${T.borderSoft}`, paddingTop: 14 }}>
          <div style={{ fontSize: 12.5, color: T.text, marginBottom: 8 }}>
            <strong>{preview.fileName}</strong> — {preview.rows.length} partidas detectadas
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
            {preview.sheetsFound.map((s) => (
              <Pill key={s.sheetName} tone="amber">{s.sheetName}: {s.count} filas</Pill>
            ))}
            {preview.sheetsIgnored.map((s) => (
              <Pill key={s} tone="dim">{s} (ignorada)</Pill>
            ))}
          </div>

          <div style={{ overflowX: "auto", maxHeight: 220, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["Unidad","Mes","Concepto","Rubro","Proyecto","Folio","Monto"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.slice(0, 12).map((r) => (
                  <tr key={r.id}>
                    <td style={tdStyle}>{r.unidad}</td>
                    <td style={tdStyle}>{r.mes}</td>
                    <td style={tdStyle}>{r.concepto}</td>
                    <td style={tdStyle}>{r.rubro}</td>
                    <td style={tdStyle}>{r.proyecto}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{r.folio || "—"}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(r.monto_estimado, r.moneda)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {preview.rows.length > 12 && (
            <div style={{ fontSize: 11, color: T.textFaint, marginTop: 6 }}>… y {preview.rows.length - 12} más</div>
          )}

          <div style={{ display: "flex", alignItems: "center", gap: 16, marginTop: 14, flexWrap: "wrap" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textDim }}>
              <input type="radio" checked={modo === "agregar"} onChange={() => setModo("agregar")} />
              Agregar a lo existente
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textDim }}>
              <input type="radio" checked={modo === "reemplazar"} onChange={() => setModo("reemplazar")} />
              Reemplazar partidas de las unidades detectadas
            </label>
            <Button onClick={confirmar} disabled={importing}>{importing ? "Importando…" : "Confirmar importación"}</Button>
            <Button variant="ghost" onClick={() => setPreview(null)}>Cancelar</Button>
          </div>
        </div>
      )}
    </Panel>
  );
}

function PartidasTab({ unidad, unidades, partidas, partidasApi }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadores = marcadoresDisponibles(proyectosUnidad);
  const anioDefault = (() => {
    const anios = partidas.filter((p) => p.unidad === unidad).map((p) => p.anio).filter(Boolean);
    return anios.length ? Math.max(...anios) : new Date().getFullYear();
  })();
  const blank = { unidad, mes: "Agosto", anio: anioDefault, smi: "", concepto: "", rubro: RUBROS[0].rubro, categoria: RUBROS[0].categorias[0], proyecto: marcadores[0] || "", monto_estimado: "", moneda: "MXN", folio: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);

  const categoriasDisponibles = RUBROS.find((r) => r.rubro === form.rubro)?.categorias || [];
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);

  const [filtros, setFiltros] = useState({ texto: "", mes: "Todos", rubro: "Todos", proyecto: "Todos" });
  const rubrosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.rubro).filter(Boolean))].sort();
  const proyectosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.proyecto).filter(Boolean))].sort();
  const mesesDisponiblesFiltro = MESES.filter((m) => partidasUnidad.some((p) => p.mes === m));

  const partidasFiltradas = partidasUnidad.filter((p) => {
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const enTexto = [p.concepto, p.folio, p.smi, p.categoria].some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    if (filtros.mes !== "Todos" && p.mes !== filtros.mes) return false;
    if (filtros.rubro !== "Todos" && p.rubro !== filtros.rubro) return false;
    if (filtros.proyecto !== "Todos" && p.proyecto !== filtros.proyecto) return false;
    return true;
  });
  const filtrosActivos = filtros.texto.trim() || filtros.mes !== "Todos" || filtros.rubro !== "Todos" || filtros.proyecto !== "Todos";
  const limpiarFiltros = () => setFiltros({ texto: "", mes: "Todos", rubro: "Todos", proyecto: "Todos" });

  const [sort, setSort] = useState({ key: null, dir: "asc" });
  const partidasOrdenadas = sortRows(partidasFiltradas, sort, {
    mes: (r) => MESES.indexOf(r.mes),
    monto_estimado: (r) => Number(r.monto_estimado) || 0,
    anio: (r) => Number(r.anio) || 0,
  });

  const submit = async (e) => {
    e.preventDefault();
    if (!form.concepto || !form.monto_estimado) return;
    let folio = (form.folio || "").trim();
    if (!folio) {
      const existingFolios = partidas.filter((p) => p.id !== editId).map((p) => p.folio);
      folio = autoFolio(unidad, form.mes, form.anio, existingFolios);
    }
    const { id, ...rest } = form;
    const record = { ...rest, folio, unidad };
    setSaving(true);
    try {
      if (editId) {
        await partidasApi.update(editId, record);
        setEditId(null);
      } else {
        await partidasApi.insert({ ...record, id: uid() });
      }
      setForm({ ...blank, anio: anioDefault, proyecto: marcadores[0] || "" });
    } catch (err) {
      alert("No se pudo guardar la partida: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };

  const startEdit = (p) => { setForm(p); setEditId(p.id); };
  const remove = (id) => partidasApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <ImportarExcelPanel partidas={partidas} partidasApi={partidasApi} />

      <Panel title={editId ? "Editar partida" : "Nueva partida presupuestal"} subtitle="Una fila por partida — puede recibir varias transacciones reales">
        <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
          <Field label="Mes">
            <Select value={form.mes} onChange={(e) => setForm({ ...form, mes: e.target.value })}>
              {MESES.map((m) => <option key={m}>{m}</option>)}
            </Select>
          </Field>
          <Field label="Año">
            <TextInput type="number" value={form.anio} onChange={(e) => setForm({ ...form, anio: Number(e.target.value) })} style={{ width: 90 }} />
          </Field>
          <Field label="SMI">
            <TextInput value={form.smi} onChange={(e) => setForm({ ...form, smi: e.target.value })} placeholder="Opcional" />
          </Field>
          <Field label="Folio de partida (liga a transacciones)">
            <TextInput value={form.folio} onChange={(e) => setForm({ ...form, folio: e.target.value })} placeholder="Vacío = se asigna automático" />
          </Field>
          <Field label="Rubro">
            <Select value={form.rubro} onChange={(e) => setForm({ ...form, rubro: e.target.value, categoria: RUBROS.find(r=>r.rubro===e.target.value).categorias[0] })}>
              {RUBROS.map((r) => <option key={r.rubro}>{r.rubro}</option>)}
            </Select>
          </Field>
          <Field label="Categoría">
            <Select value={form.categoria} onChange={(e) => setForm({ ...form, categoria: e.target.value })}>
              {categoriasDisponibles.map((c) => <option key={c}>{c}</option>)}
            </Select>
          </Field>
          <Field label="Concepto">
            <TextInput style={{ gridColumn: "span 2" }} value={form.concepto} onChange={(e) => setForm({ ...form, concepto: e.target.value })} placeholder="Ej. Servicio energía eléctrica base admtva" />
          </Field>
          <Field label="Proyecto / marcador de prorrateo">
            <Select value={form.proyecto} onChange={(e) => setForm({ ...form, proyecto: e.target.value })}>
              {marcadores.length === 0 && <option value="">Sin proyectos configurados — ve a Catálogo</option>}
              {marcadores.map((m) => <option key={m}>{m}</option>)}
            </Select>
          </Field>
          <Field label="Monto estimado">
            <TextInput type="number" step="0.01" value={form.monto_estimado} onChange={(e) => setForm({ ...form, monto_estimado: e.target.value })} placeholder="0.00" />
          </Field>
          <Field label="Moneda">
            <Select value={form.moneda} onChange={(e) => setForm({ ...form, moneda: e.target.value })}>
              {MONEDAS.map((m) => <option key={m}>{m}</option>)}
            </Select>
          </Field>
          <div style={{ gridColumn: "span 4", fontSize: 10.5, color: T.textFaint, marginTop: -6 }}>
            Formato automático del folio: UNIDAD-MESAÑO-### (ej. {unidad}-{(MES_ABR[form.mes]||"MES")}{String(form.anio).slice(-2)}-045)
          </div>
          <div style={{ gridColumn: "span 4", display: "flex", gap: 10, marginTop: 4 }}>
            <Button type="submit" disabled={saving}>{saving ? "Guardando…" : editId ? "Guardar cambios" : "Agregar partida"}</Button>
            {editId && <Button type="button" variant="ghost" onClick={() => { setEditId(null); setForm(blank); }}>Cancelar</Button>}
          </div>
        </form>
      </Panel>

      <Panel
        title={`Partidas de ${unidad}`}
        subtitle={filtrosActivos ? `${partidasFiltradas.length} de ${partidasUnidad.length} registradas` : `${partidasUnidad.length} registradas`}
      >
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${T.borderSoft}` }}>
          <Field label="Buscar">
            <TextInput
              value={filtros.texto}
              onChange={(e) => setFiltros({ ...filtros, texto: e.target.value })}
              placeholder="Concepto, folio, SMI, categoría…"
              style={{ width: 220 }}
            />
          </Field>
          <Field label="Mes">
            <Select value={filtros.mes} onChange={(e) => setFiltros({ ...filtros, mes: e.target.value })} style={{ width: 130 }}>
              <option>Todos</option>
              {mesesDisponiblesFiltro.map((m) => <option key={m}>{m}</option>)}
            </Select>
          </Field>
          <Field label="Rubro">
            <Select value={filtros.rubro} onChange={(e) => setFiltros({ ...filtros, rubro: e.target.value })} style={{ width: 200 }}>
              <option>Todos</option>
              {rubrosDisponiblesFiltro.map((r) => <option key={r}>{r}</option>)}
            </Select>
          </Field>
          <Field label="Proyecto">
            <Select value={filtros.proyecto} onChange={(e) => setFiltros({ ...filtros, proyecto: e.target.value })} style={{ width: 180 }}>
              <option>Todos</option>
              {proyectosDisponiblesFiltro.map((p) => <option key={p}>{p}</option>)}
            </Select>
          </Field>
          {filtrosActivos && <Button variant="ghost" onClick={limpiarFiltros}>Limpiar filtros</Button>}
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                <SortableTh label="Mes" sortKey="mes" sort={sort} setSort={setSort} />
                <SortableTh label="Año" sortKey="anio" sort={sort} setSort={setSort} />
                <SortableTh label="Concepto" sortKey="concepto" sort={sort} setSort={setSort} />
                <SortableTh label="Rubro" sortKey="rubro" sort={sort} setSort={setSort} />
                <SortableTh label="Categoría" sortKey="categoria" sort={sort} setSort={setSort} />
                <SortableTh label="Proyecto" sortKey="proyecto" sort={sort} setSort={setSort} />
                <SortableTh label="Folio" sortKey="folio" sort={sort} setSort={setSort} />
                <SortableTh label="Monto" sortKey="monto_estimado" sort={sort} setSort={setSort} />
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {partidasOrdenadas.map((p) => (
                <tr key={p.id}>
                  <td style={tdStyle}>{p.mes}</td>
                  <td style={tdStyle}>{p.anio}</td>
                  <td style={tdStyle}>{p.concepto}</td>
                  <td style={tdStyle}><Pill>{p.rubro}</Pill></td>
                  <td style={{ ...tdStyle, color: T.textDim }}>{p.categoria}</td>
                  <td style={tdStyle}>{p.proyecto}</td>
                  <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{p.folio || "—"}</td>
                  <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(p.monto_estimado, p.moneda)}</td>
                  <td style={tdStyle}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <Button variant="ghost" onClick={() => startEdit(p)}>Editar</Button>
                      <Button variant="danger" onClick={() => remove(p.id)}>Eliminar</Button>
                    </div>
                  </td>
                </tr>
              ))}
              {!partidasUnidad.length && (
                <tr><td colSpan={9} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin partidas aún</td></tr>
              )}
              {partidasUnidad.length > 0 && !partidasFiltradas.length && (
                <tr><td colSpan={9} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna partida coincide con estos filtros</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Panel>
    </div>
  );
}

/* ----------------------------------------------------------------------
   TABS: TRANSACCIONES
---------------------------------------------------------------------- */
function ImportarTransaccionesPanel({ partidas, transaccionesApi }) {
  const inputRef = useRef(null);
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [importing, setImporting] = useState(false);

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(""); setStatus(""); setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const { rows, sheetsFound, matched, unmatched } = parseTransaccionesWorkbook(buf, partidas);
      if (!rows.length) {
        setError('No encontré una hoja con columnas "Día", "Importe" y "A Partida" en este archivo.');
        return;
      }
      setPreview({ rows, sheetsFound, matched, unmatched, fileName: file.name });
    } catch (err) {
      setError("No pude leer el archivo. Verifica que sea un .xlsx válido.");
    } finally {
      if (inputRef.current) inputRef.current.value = "";
    }
  };

  const confirmar = async () => {
    if (!preview) return;
    setImporting(true);
    try {
      const toInsert = preview.rows.map((r) => ({ ...r, partida_id: r.partida_id || null }));
      await transaccionesApi.bulkInsert(toInsert);
      const vinculadas = preview.rows.filter((r) => r.partida_id).length;
      setStatus(`Importadas ${preview.rows.length} transacciones (${vinculadas} vinculadas, ${preview.rows.length - vinculadas} sin partida).`);
      setPreview(null);
    } catch (err) {
      setError("Ocurrió un error al importar en Supabase: " + (err.message || err));
    } finally {
      setImporting(false);
    }
  };

  const asignarPartida = (rowId, partida_id) => {
    setPreview((prev) => ({
      ...prev,
      rows: prev.rows.map((r) => (r.id === rowId ? { ...r, partida_id } : r)),
    }));
  };

  const vinculadasActuales = preview ? preview.rows.filter((r) => r.partida_id).length : 0;
  const sinVincularActuales = preview ? preview.rows.length - vinculadasActuales : 0;

  return (
    <Panel title="Carga masiva de transacciones reales" subtitle='Sube el registro de pagos (columnas Día, Solicitante, Área, Proveedor, Importe, A Partida) — se vincula por el folio de la partida'>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <input ref={inputRef} type="file" accept=".xlsx" onChange={onFile} style={{ fontSize: 12, color: T.textDim }} />
        {status && <Pill tone="teal">{status}</Pill>}
      </div>
      {error && <div style={{ marginTop: 10, fontSize: 12, color: T.red }}>{error}</div>}

      {preview && (
        <div style={{ marginTop: 16, borderTop: `1px solid ${T.borderSoft}`, paddingTop: 14 }}>
          <div style={{ fontSize: 12.5, color: T.text, marginBottom: 8 }}>
            <strong>{preview.fileName}</strong> — {preview.rows.length} transacciones detectadas
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
            <Pill tone="teal">{vinculadasActuales} vinculadas a una partida</Pill>
            {sinVincularActuales > 0 && <Pill tone="red">{sinVincularActuales} sin vincular</Pill>}
          </div>
          {sinVincularActuales > 0 && (
            <div style={{ fontSize: 11.5, color: T.amber, marginBottom: 12 }}>
              Para las filas sin folio coincidente, elige la partida correcta en el selector de esa fila antes de confirmar — o déjalas sin vincular e impórtalas igual, podrás asignarlas después.
            </div>
          )}
          <div style={{ overflowX: "auto", maxHeight: 340, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["Día","Folio","Proveedor","Importe","Partida"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.map((r) => {
                  const candidatas = r.unidad_detectada ? partidas.filter((p) => p.unidad === r.unidad_detectada) : partidas;
                  return (
                    <tr key={r.id}>
                      <td style={tdStyle}>{r.dia}</td>
                      <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{r.folio_original}</td>
                      <td style={tdStyle}>{r.proveedor}</td>
                      <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(r.importe, r.moneda)}</td>
                      <td style={tdStyle}>
                        <Select value={r.partida_id} onChange={(e) => asignarPartida(r.id, e.target.value)} style={{ minWidth: 220 }}>
                          <option value="">— Sin vincular —</option>
                          {candidatas.map((p) => (
                            <option key={p.id} value={p.id}>{p.folio ? `${p.folio} · ` : ""}{p.concepto}</option>
                          ))}
                        </Select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <Button onClick={confirmar} disabled={importing}>{importing ? "Importando…" : "Confirmar importación"}</Button>
            <Button variant="ghost" onClick={() => setPreview(null)}>Cancelar</Button>
          </div>
        </div>
      )}
    </Panel>
  );
}

function TransaccionesTab({ unidad, partidas, transacciones, transaccionesApi }) {
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const blank = { partida_id: partidasUnidad[0]?.id || "", dia: "", solicitante: "", area: "", proveedor: "", concepto_detallado: "", importe: "", moneda: "MXN" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const transUnidad = transacciones.filter((t) => partidasUnidad.some((p) => p.id === t.partida_id));
  const sinVincular = transacciones.filter((t) => !t.partida_id && t.unidad_detectada === unidad);

  const [filtros, setFiltros] = useState({ texto: "", area: "Todas", vinculo: "Todas" });
  const [sort, setSort] = useState({ key: "dia", dir: "desc" });

  const partidaDe = (t) => partidasUnidad.find((p) => p.id === t.partida_id);
  const areasDisponibles = [...new Set(transUnidad.map((t) => t.area).filter(Boolean))].sort();

  const transFiltradas = transUnidad.filter((t) => {
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const partida = partidaDe(t);
      const enTexto = [t.proveedor, t.concepto_detallado, t.solicitante, t.area, partida?.folio, partida?.concepto]
        .some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    if (filtros.area !== "Todas" && t.area !== filtros.area) return false;
    if (filtros.vinculo === "Vinculadas" && !t.partida_id) return false;
    if (filtros.vinculo === "Sin vincular" && t.partida_id) return false;
    return true;
  });
  const filtrosActivos = filtros.texto.trim() || filtros.area !== "Todas" || filtros.vinculo !== "Todas";
  const limpiarFiltros = () => setFiltros({ texto: "", area: "Todas", vinculo: "Todas" });

  const transOrdenadas = sortRows(transFiltradas, sort, {
    importe: (r) => Number(r.importe) || 0,
    proveedor: (r) => (r.proveedor || "").toLowerCase(),
    dia: (r) => r.dia || "",
    partida: (r) => (partidaDe(r)?.concepto || "").toLowerCase(),
  });

  const submit = async (e) => {
    e.preventDefault();
    if (!form.partida_id || !form.importe) return;
    const { id, ...rest } = form;
    setSaving(true);
    try {
      if (editId) {
        await transaccionesApi.update(editId, rest);
        setEditId(null);
      } else {
        await transaccionesApi.insert({ ...rest, id: uid() });
      }
      setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" });
    } catch (err) {
      alert("No se pudo guardar la transacción: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };
  const startEdit = (t) => { setForm(t); setEditId(t.id); };
  const remove = (id) => transaccionesApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));

  if (!partidasUnidad.length) {
    return <EmptyState title="Primero crea partidas" body={`Registra al menos una partida de ${unidad} en la pestaña Partidas antes de capturar transacciones reales.`} />;
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <ImportarTransaccionesPanel partidas={partidas} transaccionesApi={transaccionesApi} />

      <Panel title={editId ? "Editar transacción" : "Nueva transacción real"} subtitle="Se vincula a una partida — una partida puede tener varias">
        <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
          <Field label="Partida">
            <Select style={{ gridColumn: "span 2" }} value={form.partida_id} onChange={(e) => setForm({ ...form, partida_id: e.target.value })}>
              {partidasUnidad.map((p) => <option key={p.id} value={p.id}>{p.mes} · {p.concepto} ({p.proyecto})</option>)}
            </Select>
          </Field>
          <Field label="Día">
            <TextInput type="date" value={form.dia} onChange={(e) => setForm({ ...form, dia: e.target.value })} />
          </Field>
          <Field label="Solicitante">
            <TextInput value={form.solicitante} onChange={(e) => setForm({ ...form, solicitante: e.target.value })} />
          </Field>
          <Field label="Área">
            <TextInput value={form.area} onChange={(e) => setForm({ ...form, area: e.target.value })} />
          </Field>
          <Field label="Proveedor / razón social">
            <TextInput value={form.proveedor} onChange={(e) => setForm({ ...form, proveedor: e.target.value })} />
          </Field>
          <Field label="Concepto de pago (detallado)">
            <TextInput style={{ gridColumn: "span 2" }} value={form.concepto_detallado} onChange={(e) => setForm({ ...form, concepto_detallado: e.target.value })} />
          </Field>
          <Field label="Importe">
            <TextInput type="number" step="0.01" value={form.importe} onChange={(e) => setForm({ ...form, importe: e.target.value })} placeholder="0.00" />
          </Field>
          <Field label="Moneda">
            <Select value={form.moneda} onChange={(e) => setForm({ ...form, moneda: e.target.value })}>
              {MONEDAS.map((m) => <option key={m}>{m}</option>)}
            </Select>
          </Field>
          <div style={{ gridColumn: "span 4", display: "flex", gap: 10, marginTop: 4 }}>
            <Button type="submit" disabled={saving}>{saving ? "Guardando…" : editId ? "Guardar cambios" : "Registrar transacción"}</Button>
            {editId && <Button type="button" variant="ghost" onClick={() => { setEditId(null); setForm(blank); }}>Cancelar</Button>}
          </div>
        </form>
      </Panel>

      <Panel
        title={`Transacciones de ${unidad}`}
        subtitle={filtrosActivos ? `${transFiltradas.length} de ${transUnidad.length} registradas` : `${transUnidad.length} registradas`}
      >
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${T.borderSoft}` }}>
          <Field label="Buscar">
            <TextInput
              value={filtros.texto}
              onChange={(e) => setFiltros({ ...filtros, texto: e.target.value })}
              placeholder="Proveedor, concepto, folio, área…"
              style={{ width: 220 }}
            />
          </Field>
          <Field label="Área">
            <Select value={filtros.area} onChange={(e) => setFiltros({ ...filtros, area: e.target.value })} style={{ width: 170 }}>
              <option>Todas</option>
              {areasDisponibles.map((a) => <option key={a}>{a}</option>)}
            </Select>
          </Field>
          <Field label="Vínculo">
            <Select value={filtros.vinculo} onChange={(e) => setFiltros({ ...filtros, vinculo: e.target.value })} style={{ width: 150 }}>
              <option>Todas</option>
              <option>Vinculadas</option>
              <option>Sin vincular</option>
            </Select>
          </Field>
          {filtrosActivos && <Button variant="ghost" onClick={limpiarFiltros}>Limpiar filtros</Button>}
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                <SortableTh label="Día" sortKey="dia" sort={sort} setSort={setSort} />
                <SortableTh label="Partida" sortKey="partida" sort={sort} setSort={setSort} />
                <SortableTh label="Proveedor" sortKey="proveedor" sort={sort} setSort={setSort} />
                <th style={thStyle}>Concepto</th>
                <SortableTh label="Importe" sortKey="importe" sort={sort} setSort={setSort} />
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {transOrdenadas.map((t) => {
                const partida = partidaDe(t);
                return (
                  <tr key={t.id}>
                    <td style={tdStyle}>{t.dia || "—"}</td>
                    <td style={tdStyle}>{partida ? `${partida.concepto}` : <span style={{ color: T.red }}>Partida eliminada</span>}</td>
                    <td style={tdStyle}>{t.proveedor}</td>
                    <td style={{ ...tdStyle, color: T.textDim }}>{t.concepto_detallado}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</td>
                    <td style={tdStyle}>
                      <div style={{ display: "flex", gap: 6 }}>
                        <Button variant="ghost" onClick={() => startEdit(t)}>Editar</Button>
                        <Button variant="danger" onClick={() => remove(t.id)}>Eliminar</Button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {!transUnidad.length && (
                <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin transacciones aún</td></tr>
              )}
              {transUnidad.length > 0 && !transFiltradas.length && (
                <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna transacción coincide con estos filtros</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Panel>

      {sinVincular.length > 0 && (
        <Panel title="Transacciones importadas sin partida vinculada" subtitle={`${sinVincular.length} en ${unidad} — su folio no coincidió con ninguna partida`}>
          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["Día","Folio","Proveedor","Concepto","Importe",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {sinVincular.map((t) => (
                  <tr key={t.id}>
                    <td style={tdStyle}>{t.dia || "—"}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.red }}>{t.folio_original}</td>
                    <td style={tdStyle}>{t.proveedor}</td>
                    <td style={{ ...tdStyle, color: T.textDim }}>{t.concepto_detallado}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</td>
                    <td style={tdStyle}>
                      <Button variant="ghost" onClick={() => startEdit(t)}>Vincular</Button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11, color: T.textFaint, marginTop: 10 }}>
            Clic en "Vincular" para abrir el formulario de arriba, elegir la partida correcta en el selector y guardar.
          </div>
        </Panel>
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------
   TABS: CATALOGO
---------------------------------------------------------------------- */
function CatalogoTab({ unidad, unidades, proyectosApi }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const [nuevo, setNuevo] = useState({ nombre: "", grupo: "", pct: "" });
  const [drafts, setDrafts] = useState({}); // id -> { field: value } — edición local antes de confirmar en blur

  const draftValue = (p, field) => (drafts[p.id]?.[field] !== undefined ? drafts[p.id][field] : p[field]);
  const setDraft = (id, field, val) => setDrafts((d) => ({ ...d, [id]: { ...d[id], [field]: val } }));
  const commitDraft = (id, field) => {
    if (drafts[id]?.[field] === undefined) return;
    const val = field === "pct" ? Number(drafts[id][field]) || 0 : drafts[id][field];
    updateProyecto(id, field, val);
  };

  const updateProyecto = (id, field, val) => {
    proyectosApi.update(id, { [field]: val }).catch((err) => alert("No se pudo actualizar: " + (err.message || err)));
  };
  const removeProyecto = (id) => {
    proyectosApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };
  const addProyecto = () => {
    if (!nuevo.nombre) return;
    proyectosApi.insert({ id: uid(), unidad, nombre: nuevo.nombre, grupo: nuevo.grupo || "General", pct: Number(nuevo.pct) || 0 })
      .catch((err) => alert("No se pudo agregar: " + (err.message || err)));
    setNuevo({ nombre: "", grupo: "", pct: "" });
  };

  const totalPct = proyectosUnidad.reduce((s, p) => s + Number(p.pct || 0), 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel
        title={`Proyectos y % de prorrateo — ${unidad}`}
        subtitle="El % se usa para repartir gastos compartidos marcados como “Todos” o “<Grupo> Gral”"
        right={<Pill tone={Math.abs(totalPct - 100) < 0.01 ? "teal" : "red"}>{totalPct.toFixed(1)}% asignado</Pill>}
      >
        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>{["Proyecto","Grupo","% Administrativos",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {proyectosUnidad.map((p) => (
                <tr key={p.id}>
                  <td style={tdStyle}><TextInput value={draftValue(p, "nombre")} onChange={(e) => setDraft(p.id, "nombre", e.target.value)} onBlur={() => commitDraft(p.id, "nombre")} /></td>
                  <td style={tdStyle}><TextInput value={draftValue(p, "grupo")} onChange={(e) => setDraft(p.id, "grupo", e.target.value)} onBlur={() => commitDraft(p.id, "grupo")} placeholder="Desh / Prod / IMP" /></td>
                  <td style={tdStyle}><TextInput type="number" value={draftValue(p, "pct")} onChange={(e) => setDraft(p.id, "pct", e.target.value)} onBlur={() => commitDraft(p.id, "pct")} style={{ width: 90 }} /></td>
                  <td style={tdStyle}><Button variant="danger" onClick={() => removeProyecto(p.id)}>Eliminar</Button></td>
                </tr>
              ))}
              <tr>
                <td style={tdStyle}><TextInput placeholder="Nuevo proyecto" value={nuevo.nombre} onChange={(e) => setNuevo({ ...nuevo, nombre: e.target.value })} /></td>
                <td style={tdStyle}><TextInput placeholder="Grupo" value={nuevo.grupo} onChange={(e) => setNuevo({ ...nuevo, grupo: e.target.value })} /></td>
                <td style={tdStyle}><TextInput type="number" placeholder="%" value={nuevo.pct} onChange={(e) => setNuevo({ ...nuevo, pct: e.target.value })} style={{ width: 90 }} /></td>
                <td style={tdStyle}><Button onClick={addProyecto}>Agregar</Button></td>
              </tr>
            </tbody>
          </table>
        </div>
        {Math.abs(totalPct - 100) > 0.01 && proyectosUnidad.length > 0 && (
          <div style={{ marginTop: 10, fontSize: 11.5, color: T.amber }}>
            El % total no suma 100 — el prorrateo bajo el marcador “Todos” se renormaliza automáticamente, pero conviene que sí sumen 100 para que el % represente lo real.
          </div>
        )}
      </Panel>

      <Panel title="Catálogo de rubros y categorías" subtitle="Referencia — igual al acordado">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "4px 24px" }}>
          {RUBROS.map((r) => (
            <details key={r.rubro} style={{ borderBottom: `1px solid ${T.borderSoft}`, padding: "8px 0" }}>
              <summary style={{ cursor: "pointer", fontSize: 12.5, color: T.text, fontWeight: 600 }}>{r.rubro}</summary>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 8 }}>
                {r.categorias.map((c) => <Pill key={c}>{c}</Pill>)}
              </div>
            </details>
          ))}
        </div>
      </Panel>
    </div>
  );
}

/* ----------------------------------------------------------------------
   ROOT APP
---------------------------------------------------------------------- */
export default function App() {
  const proyectosApi = useCollection("proyectos");
  const partidasApi = useCollection("partidas");
  const transaccionesApi = useCollection("transacciones");
  const [unidad, setUnidad] = useState("CTM");
  const [tab, setTab] = useState("dashboard");

  const unidades = useMemo(() => {
    const map = {};
    UNIDADES_BASE.forEach((u) => { map[u] = { proyectos: [] }; });
    proyectosApi.rows.forEach((p) => {
      if (!map[p.unidad]) map[p.unidad] = { proyectos: [] };
      map[p.unidad].proyectos.push(p);
    });
    return map;
  }, [proyectosApi.rows]);

  const partidas = partidasApi.rows;
  const transacciones = transaccionesApi.rows;
  const ready = proyectosApi.ready && partidasApi.ready && transaccionesApi.ready;

  const TABS = [
    { id: "dashboard", label: "Dashboard" },
    { id: "partidas", label: "Partidas" },
    { id: "transacciones", label: "Transacciones" },
    { id: "catalogo", label: "Catálogo" },
  ];

  return (
    <div style={{ background: T.bg, minHeight: "100%", fontFamily: T.fontUI, color: T.text, padding: 24 }}>
      <style>{`
        * { box-sizing: border-box; }
        table { font-family: ${T.fontUI}; }
        details > summary { list-style: none; }
        details > summary::-webkit-details-marker { display: none; }
        details > summary:before { content: '▸ '; color: ${T.textFaint}; }
        details[open] > summary:before { content: '▾ '; }
        select, input { color-scheme: dark; }
        ::selection { background: ${T.amberDim}; }
      `}</style>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 22, flexWrap: "wrap", gap: 14 }}>
        <div>
          <div style={{ fontSize: 10.5, color: T.amber, letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>
            Control de presupuestos · prototipo
          </div>
          <div style={{ fontSize: 24, fontWeight: 700 }}>Panel de gasto</div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <span style={{ fontSize: 11, color: T.textDim }}>Unidad de negocio</span>
          <div style={{ display: "flex", background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8, padding: 3 }}>
            {Object.keys(unidades).map((u) => (
              <button
                key={u}
                onClick={() => setUnidad(u)}
                style={{
                  padding: "7px 16px", borderRadius: 6, border: "none", cursor: "pointer",
                  background: unidad === u ? T.amber : "transparent",
                  color: unidad === u ? "#1A1206" : T.textDim,
                  fontWeight: 700, fontSize: 12.5, fontFamily: T.fontMono,
                }}
              >
                {u}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${T.border}`, marginBottom: 22 }}>
        {TABS.map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            style={{
              background: "transparent", border: "none", cursor: "pointer",
              padding: "10px 16px", fontSize: 12.5, fontWeight: 600,
              color: tab === t.id ? T.text : T.textFaint,
              borderBottom: tab === t.id ? `2px solid ${T.amber}` : "2px solid transparent",
              marginBottom: -1,
            }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {!ready ? (
        <div style={{ color: T.textDim, fontSize: 13 }}>Cargando datos compartidos…</div>
      ) : (
        <>
          {tab === "dashboard" && <Dashboard unidad={unidad} unidades={unidades} partidas={partidas} transacciones={transacciones} />}
          {tab === "partidas" && <PartidasTab unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} />}
          {tab === "transacciones" && <TransaccionesTab unidad={unidad} partidas={partidas} transacciones={transacciones} transaccionesApi={transaccionesApi} />}
          {tab === "catalogo" && <CatalogoTab unidad={unidad} unidades={unidades} proyectosApi={proyectosApi} />}
        </>
      )}
    </div>
  );
}
