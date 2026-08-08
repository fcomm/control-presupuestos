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
  bg: "#F5F6F8",
  panel: "#FFFFFF",
  panelAlt: "#ECEEF1",
  border: "#DCDFE3",
  borderSoft: "#E7E9EC",
  text: "#232A31",
  textDim: "#5B6B79",
  textFaint: "#8B99A6",
  accent: "#3E5C76",
  accentDim: "#2F4A63",
  accentBg: "#DCE4EC",
  amber: "#B8791C",
  amberDim: "#8A5A16",
  teal: "#1E8F73",
  red: "#C0483F",
  blue: "#3E6E8E",
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
const ZONAS = ["Queretaro", "Poza Rica", "Paraiso", "Altamira", "Cerro Azul", "CDMX", "Guaymas", "Torreon", "Rosarito", "Agua Dulce", "Cotaxtla"];

// Catálogos oficiales del SAT (CFDI)
const FORMAS_PAGO = [
  { value: "01", label: "01 - Efectivo" },
  { value: "02", label: "02 - Cheque nominativo" },
  { value: "03", label: "03 - Transferencia electrónica de fondos" },
  { value: "04", label: "04 - Tarjeta de crédito" },
  { value: "05", label: "05 - Monedero electrónico" },
  { value: "06", label: "06 - Dinero electrónico" },
  { value: "08", label: "08 - Vales de despensa" },
  { value: "12", label: "12 - Dación en pago" },
  { value: "13", label: "13 - Pago por subrogación" },
  { value: "14", label: "14 - Pago por consignación" },
  { value: "15", label: "15 - Condonación" },
  { value: "17", label: "17 - Compensación" },
  { value: "23", label: "23 - Novación" },
  { value: "24", label: "24 - Confusión" },
  { value: "25", label: "25 - Remisión de deuda" },
  { value: "26", label: "26 - Prescripción o caducidad" },
  { value: "27", label: "27 - A satisfacción del acreedor" },
  { value: "28", label: "28 - Tarjeta de débito" },
  { value: "29", label: "29 - Tarjeta de servicios" },
  { value: "30", label: "30 - Aplicación de anticipos" },
  { value: "31", label: "31 - Intermediario pagos" },
  { value: "99", label: "99 - Por definir" },
];
const METODOS_PAGO = [
  { value: "PUE", label: "PUE - Pago en una sola exhibición" },
  { value: "PPD", label: "PPD - Pago en parcialidades o diferido" },
];

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const uid = () => (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2, 10));

// ----------------------------------------------------------------------
// VERSIÓN — súbela cada vez que cambies este archivo. Formato MAJOR.MINOR.PATCH:
// MINOR = feature nueva, PATCH = fix/ajuste menor. Se muestra en el header de
// la app y debe ir en el nombre del archivo que se comparte (App-v1.5.0.jsx).
// ----------------------------------------------------------------------
const APP_VERSION = "1.26.0";
const CHANGELOG = [
  { v: "1.26.0", desc: "Reporte de Pagos: agrega columna Día y filtro por rango de fechas (Desde/Hasta) — sirve para un solo día o un rango" },
  { v: "1.25.2", desc: "Quita el campo 'Días de crédito' por completo (formulario, tabla, carga masiva, Reporte de Pagos)" },
  { v: "1.25.1", desc: "Carga de proveedores reconoce columna 'Compañía' (OSB/CTM/ISE) para decidir a quién pertenece cada fila, sin depender del nombre de la hoja" },
  { v: "1.25.0", desc: "Cambio de modelo: un proveedor puede tener varias cuentas bancarias (tabla separada); Transacciones ahora liga a la cuenta específica usada" },
  { v: "1.24.0", desc: "Carga masiva de Proveedores (hojas OSB/CTM/ISE), actualiza por RFC o Nombre si ya existe" },
  { v: "1.23.0", desc: "Nueva pestaña Reporte de Pagos; Proveedor ahora liga al catálogo real; agrega Área, Folio Compra SAE, Folio Factura, Forma/Método de Pago (catálogos SAT) y Días de Crédito" },
  { v: "1.22.0", desc: "Nuevo catálogo de Proveedores por compañía (Nombre, RFC, Id SAE, Banco, CLABE, No. Cuenta, Divisa) en la pestaña Catálogo" },
  { v: "1.21.1", desc: "Fix crítico: el agrupamiento agrupaba todo como 'Sin dato'; ahora el encabezado de cada grupo también muestra el nombre del campo" },
  { v: "1.21.0", desc: "Columnas de Partidas y Transacciones se pueden reordenar arrastrando el encabezado — se recuerda entre visitas" },
  { v: "1.20.0", desc: "Columnas de Partidas y Transacciones ahora se pueden ajustar de ancho arrastrando el borde — se recuerda entre visitas" },
  { v: "1.19.0", desc: "Agrupamiento rediseñado estilo Airtable: panel desplegable con campo+dirección+quitar por nivel, añadir subgrupo, contraer/expandir todo" },
  { v: "1.18.0", desc: "Partidas: columna Proyecto siempre visible aunque agrupes por ella; Transacciones: nueva columna Folio (código de la partida vinculada)" },
  { v: "1.17.3", desc: "Transacciones: botón 'Borrar todas' de una unidad (vinculadas + sin vincular), para reimportar limpio sin duplicar" },
  { v: "1.17.2", desc: "Fix: filas de transacciones sin folio ya no se pierden al importar — entran como 'sin vincular'; detecta unidad por nombre de hoja como respaldo" },
  { v: "1.17.1", desc: "Catálogo de Zonas (11 zonas) — el campo Zona en Transacciones ya es un selector" },
  { v: "1.17.0", desc: "Transacciones: agrega Zona (reemplaza a Área en captura/tabla/filtro/agrupamiento; Área queda en la BD sin usarse)" },
  { v: "1.16.1", desc: "Transacciones: Proyecto ahora selecciona del catálogo (mismos marcadores que Partidas); Status es Pagado/No Pagado" },
  { v: "1.16.0", desc: "Transacciones: agrega columnas Proyecto y Status (base de datos, importador, formulario, tabla, agrupamiento)" },
  { v: "1.15.2", desc: "Transacciones: botón para eliminar en lote las 'sin vincular' (recuperación tras un link roto)" },
  { v: "1.15.1", desc: "Fix crítico: la carga masiva ahora actualiza por folio en vez de duplicar o borrar — ya no rompe el vínculo con transacciones existentes" },
  { v: "1.15.0", desc: "Partidas: captura TODAS las columnas del Excel en 'extra' (visible en modal), aunque no se usen todavía" },
  { v: "1.14.1", desc: "Partidas: reordena filtros a Mes → Proyecto → Rubro" },
  { v: "1.14.0", desc: "Transacciones: filtros de Mes y Proyecto (vía partida vinculada); Zona queda pendiente" },
  { v: "1.13.0", desc: "Transacciones: columna Partida editable directo en la tabla (y en 'sin vincular')" },
  { v: "1.12.0", desc: "Fix: campos que se salían del modal (Partida, Concepto); agrega filtro de Mes/Proyecto para elegir partida en Transacciones" },
  { v: "1.11.2", desc: "Dashboard: reordena secciones — Resumen presupuestado sube, Avance por proyecto baja al final" },
  { v: "1.11.1", desc: "Fix: selector de mes ahora se cierra al hacer clic afuera (o en 'Cerrar')" },
  { v: "1.11.0", desc: "Dashboard: selector de mes ahora permite elegir varios meses a la vez" },
  { v: "1.10.0", desc: "Dashboard: selector de mes (filtra todo excepto Tendencia mensual), se quitó el gauge de avance global" },
  { v: "1.9.0", desc: "Dashboard: tabla dinámica Proyecto→Rubro→Concepto x Mes (como pivot de Excel)" },
  { v: "1.8.0", desc: "Transacciones: mismo patrón que Partidas (popup, agrupamiento 3 niveles, orden)" },
  { v: "1.7.0", desc: "Nuevo tema claro (Gris Suave) — antes era oscuro" },
  { v: "1.6.0", desc: "Tercer nivel de agrupamiento + sangría con color por nivel" },
  { v: "1.5.0", desc: "Código de versión visible en la interfaz" },
  { v: "1.4.0", desc: "Agrupamiento tipo Airtable en tabla de Partidas" },
  { v: "1.3.0", desc: "Tabla de Partidas arriba + popup para nueva/editar" },
  { v: "1.2.1", desc: "Fix: prorrateo ya no pierde montos sin proyecto en Catálogo" },
  { v: "1.2.0", desc: "Filtros y ordenamiento en Partidas y Transacciones" },
  { v: "1.1.0", desc: "Filtros en tabla de Partidas" },
  { v: "1.0.1", desc: "Configuración de Wrangler (Cloudflare Workers Static Assets)" },
  { v: "1.0.0", desc: "Setup inicial: Vite + React + Supabase" },
];

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

// Captura TODAS las columnas del Excel que no tengan ya su propio campo en la
// partida (rubro, categoria, proyecto, folio, etc.) dentro de un objeto libre —
// así no se pierde nada, aunque la app todavía no tenga pantalla para esos campos.
function buildExtra(row, headers, promotedIdx) {
  const extra = {};
  headers.forEach((h, i) => {
    if (!h || promotedIdx.has(i)) return;
    const key = String(h).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
    if (!key) return;
    let val = row[i];
    if (val === null || val === undefined || val === "") return;
    if (val instanceof Date) val = val.toISOString().slice(0, 10);
    extra[key] = val;
  });
  return extra;
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
    const total = proyectosUnidad.reduce((s, p) => s + Number(p.pct || 0), 0);
    if (total > 0) return proyectosUnidad.map((p) => ({ proyecto: p.nombre, fraccion: Number(p.pct || 0) / total }));
  }
  // "<Grupo> Gral" e.g. "Desh Gral" -> group = "Desh"
  const m = /^(.*) Gral$/.exec(marcador);
  if (m) {
    const grupo = m[1];
    const miembros = proyectosUnidad.filter((p) => p.grupo === grupo);
    const total = miembros.reduce((s, p) => s + Number(p.pct || 0), 0);
    if (total > 0) return miembros.map((p) => ({ proyecto: p.nombre, fraccion: Number(p.pct || 0) / total }));
  }
  // Fallback: el proyecto/marcador no está dado de alta en el Catálogo (falta % o
  // aún no se ha configurado esa unidad). En vez de perder el monto en silencio,
  // se muestra tal cual bajo su propio nombre, sin prorratear.
  return [{ proyecto: marcador, fraccion: 1 }];
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

      const promotedIdx = new Set(
        [col.mes, col.anio, col.smi, col.concepto, col.subtotalMXN, col.subtotalUSD, col.totalMXN, col.moneda, col.proyecto, col.rubro, col.categoria, col.folio]
          .filter((i) => i !== -1)
      );

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
        extra: buildExtra(row, headers, promotedIdx),
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
// Lee un libro de Excel buscando hojas literalmente llamadas OSB/CTM/ISE con
// columnas de proveedores (Nombre, RFC, Id SAE, Banco, CLABE, No. Cuenta, Divisa,
// Días de crédito). Marca cada fila como actualización (mismo RFC, o mismo
// Nombre si no hay RFC, dentro de esa unidad) o alta nueva.
function parseProveedoresWorkbook(arrayBuffer, existingProveedores = []) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const rows = [];
  const sheetsFound = [];
  let actualizaciones = 0, nuevas = 0, sinCompania = 0;

  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    if (!aoa.length) return;
    const headers = aoa[0];

    const col = {
      compania: findCol(headers, ["compania"], ["compañia"], ["unidad"]),
      nombre: findCol(headers, ["nombre"]),
      rfc: findExactCol(headers, ["rfc"]),
      idSae: findCol(headers, ["id", "sae"]),
      banco: findExactCol(headers, ["banco"]),
      clabe: findExactCol(headers, ["clabe"]),
      cuenta: findCol(headers, ["numero", "cuenta"], ["no.", "cuenta"], ["cuenta"]),
      divisa: findExactCol(headers, ["divisa"]),
    };
    if (col.nombre === -1) return;

    // Si la hoja se llama literalmente OSB/CTM/ISE, sirve de respaldo cuando
    // una fila no trae (o no reconoce) la columna Compañía.
    const sheetUnidad = UNIDAD_KEYS.includes(sheetName.trim().toUpperCase()) ? sheetName.trim().toUpperCase() : null;

    let count = 0;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const nombre = row[col.nombre] ? String(row[col.nombre]).trim() : "";
      if (!nombre) continue;

      const companiaRaw = (col.compania !== -1 && row[col.compania]) ? String(row[col.compania]).trim().toUpperCase() : "";
      const unidad = UNIDAD_KEYS.includes(companiaRaw) ? companiaRaw : sheetUnidad;
      if (!unidad) { sinCompania++; continue; }

      const rfc = (col.rfc !== -1 && row[col.rfc]) ? String(row[col.rfc]).trim().toUpperCase() : "";
      const existente = existingProveedores.find((p) => {
        if (p.unidad !== unidad) return false;
        if (rfc && p.rfc) return p.rfc.trim().toUpperCase() === rfc;
        if (!rfc && !p.rfc) return (p.nombre || "").trim().toUpperCase() === nombre.toUpperCase();
        return false;
      });
      if (existente) actualizaciones++; else nuevas++;

      rows.push({
        id: uid(),
        unidad,
        nombre,
        rfc,
        id_sae: (col.idSae !== -1 && row[col.idSae]) ? String(row[col.idSae]).trim() : "",
        _existenteId: existente ? existente.id : null,
        _cuenta: {
          banco: (col.banco !== -1 && row[col.banco]) ? String(row[col.banco]).trim() : "",
          clabe: (col.clabe !== -1 && row[col.clabe]) ? String(row[col.clabe]).trim() : "",
          numero_cuenta: (col.cuenta !== -1 && row[col.cuenta]) ? String(row[col.cuenta]).trim() : "",
          divisa: (col.divisa !== -1 && row[col.divisa]) ? String(row[col.divisa]).trim().toUpperCase() : "MXN",
        },
      });
      count++;
    }
    if (count) sheetsFound.push({ sheetName, count });
  });

  return { rows, sheetsFound, nuevas, actualizaciones, sinCompania };
}

function parseTransaccionesWorkbook(arrayBuffer, partidas, proveedores = []) {
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
      proyecto: findExactCol(headers, ["proyecto"]),
      zona: findExactCol(headers, ["zona"]),
      area: findExactCol(headers, ["area"]),
      proveedor: findCol(headers, ["nombre", "denominacion"], ["razon", "social"], ["proveedor"]),
      concepto: findCol(headers, ["concepto"]),
      importe: findCol(headers, ["importe"]),
      moneda: findCol(headers, ["moneda"]),
      status: findExactCol(headers, ["status", "estatus"]),
      folioCompraSae: findCol(headers, ["folio", "compra"]),
      folioFactura: findCol(headers, ["folio", "factura"]),
      formaPago: findExactCol(headers, ["forma de pago", "forma pago"]),
      metodoPago: findExactCol(headers, ["metodo de pago", "metodo pago"]),
      folio: findCol(headers, ["a partida"]) !== -1 ? findCol(headers, ["a partida"]) : findExactCol(headers, ["partida", "folio"]),
    };

    // Si el nombre de la hoja es literalmente OSB/CTM/ISE, lo usamos como respaldo
    // para saber la unidad de una fila cuando no trae folio (o el folio no lo delata).
    const sheetUnidad = UNIDAD_KEYS.includes(sheetName.trim().toUpperCase()) ? sheetName.trim().toUpperCase() : null;

    let count = 0;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const importeRaw = col.importe !== -1 ? row[col.importe] : null;
      const importe = typeof importeRaw === "number" ? importeRaw : parseFloat(String(importeRaw || "").replace(/[^0-9.-]/g, ""));
      if (!importe || isNaN(importe)) continue;
      const folioRaw = col.folio !== -1 ? row[col.folio] : null;
      const folio = folioRaw ? String(folioRaw).trim() : "";

      const prefixMatch = folio ? /^([A-Za-z]+)-/.exec(folio) : null;
      const unidad_detectada = (prefixMatch && UNIDAD_KEYS.includes(prefixMatch[1].toUpperCase())) ? prefixMatch[1].toUpperCase() : sheetUnidad;

      const partida = folio ? partidas.find(
        (p) => p.folio && p.folio.trim().toUpperCase() === folio.toUpperCase() &&
               (!unidad_detectada || p.unidad === unidad_detectada)
      ) : null;
      if (partida) matched++; else unmatched++;

      const proveedorNombre = (col.proveedor !== -1 && row[col.proveedor]) ? String(row[col.proveedor]).trim() : "";
      const proveedorMatch = proveedorNombre ? proveedores.find(
        (pv) => pv.nombre && pv.nombre.trim().toUpperCase() === proveedorNombre.toUpperCase() &&
                (!unidad_detectada || pv.unidad === unidad_detectada)
      ) : null;

      rows.push({
        id: uid(),
        partida_id: partida ? partida.id : "",
        folio_original: folio,
        unidad_detectada: partida ? partida.unidad : unidad_detectada,
        dia: col.dia !== -1 ? toISODate(row[col.dia]) : "",
        solicitante: (col.solicitante !== -1 && row[col.solicitante]) ? String(row[col.solicitante]).trim() : "",
        proyecto: (col.proyecto !== -1 && row[col.proyecto]) ? String(row[col.proyecto]).trim() : "",
        zona: (col.zona !== -1 && row[col.zona]) ? String(row[col.zona]).trim() : "",
        area: (col.area !== -1 && row[col.area]) ? String(row[col.area]).trim() : "",
        proveedor: proveedorNombre,
        proveedor_id: proveedorMatch ? proveedorMatch.id : "",
        concepto_detallado: (col.concepto !== -1 && row[col.concepto]) ? String(row[col.concepto]).trim() : "",
        importe,
        moneda: (col.moneda !== -1 && row[col.moneda]) ? String(row[col.moneda]).trim().toUpperCase() : "MXN",
        status: (col.status !== -1 && row[col.status]) ? String(row[col.status]).trim() : "",
        folio_compra_sae: (col.folioCompraSae !== -1 && row[col.folioCompraSae]) ? String(row[col.folioCompraSae]).trim() : "",
        folio_factura: (col.folioFactura !== -1 && row[col.folioFactura]) ? String(row[col.folioFactura]).trim() : "",
        forma_pago: (col.formaPago !== -1 && row[col.formaPago]) ? String(row[col.formaPago]).trim() : "",
        metodo_pago: (col.metodoPago !== -1 && row[col.metodoPago]) ? String(row[col.metodoPago]).trim() : "",
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
function Field({ label, children, style }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12, color: T.textDim, minWidth: 0, ...style }}>
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
  width: "100%",
  boxSizing: "border-box",
  minWidth: 0,
};

function TextInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}
function Select(props) {
  return <select {...props} style={{ ...inputStyle, ...(props.style || {}) }} />;
}

function Button({ children, variant = "primary", ...rest }) {
  const styles = {
    primary: { background: T.accent, color: "#FFFFFF", border: `1px solid ${T.accent}` },
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
    dim: { color: T.textDim, background: T.panelAlt, border: `1px solid ${T.border}` },
    accent: { color: T.accentDim, background: T.accentBg, border: `1px solid ${T.accent}55` },
    amber: { color: T.amberDim, background: "#B8791C1A", border: `1px solid ${T.amberDim}` },
    teal: { color: T.teal, background: "#1E8F731A", border: `1px solid #1E8F7355` },
    red: { color: T.red, background: "#C0483F1A", border: `1px solid #C0483F55` },
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

  const mesesDisponibles = MESES.filter((m) => partidasUnidad.some((p) => p.mes === m));
  const [mesesSeleccionados, setMesesSeleccionados] = useState([]);
  const partidasFiltradasMes = mesesSeleccionados.length ? partidasUnidad.filter((p) => mesesSeleccionados.includes(p.mes)) : partidasUnidad;
  const mesLabel = mesesSeleccionados.length ? ` · ${mesesSeleccionados.join(", ")}` : "";
  const idsPartidasFiltradas = new Set(partidasFiltradasMes.map((p) => p.id));
  const transFiltradasMes = transUnidad.filter((t) => idsPartidasFiltradas.has(t.partida_id));

  const porProyecto = useMemo(() => {
    const map = {};
    proyectosUnidad.forEach((p) => { map[p.nombre] = { proyecto: p.nombre, presupuestado: 0, ejecutado: 0 }; });

    partidasFiltradasMes.forEach((partida) => {
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!map[proyecto]) map[proyecto] = { proyecto, presupuestado: 0, ejecutado: 0 };
        map[proyecto].presupuestado += (Number(partida.monto_estimado) || 0) * fraccion;
      });
    });

    transFiltradasMes.forEach((t) => {
      const partida = partidasFiltradasMes.find((p) => p.id === t.partida_id);
      if (!partida) return;
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!map[proyecto]) map[proyecto] = { proyecto, presupuestado: 0, ejecutado: 0 };
        map[proyecto].ejecutado += (Number(t.importe) || 0) * fraccion;
      });
    });

    return Object.values(map);
  }, [proyectosUnidad, partidasFiltradasMes, transFiltradasMes]);

  const porRubro = useMemo(() => {
    const map = {};
    partidasFiltradasMes.forEach((p) => {
      map[p.rubro] = map[p.rubro] || { rubro: p.rubro, presupuestado: 0, ejecutado: 0 };
      map[p.rubro].presupuestado += Number(p.monto_estimado) || 0;
    });
    transFiltradasMes.forEach((t) => {
      const partida = partidasFiltradasMes.find((p) => p.id === t.partida_id);
      if (!partida) return;
      map[partida.rubro] = map[partida.rubro] || { rubro: partida.rubro, presupuestado: 0, ejecutado: 0 };
      map[partida.rubro].ejecutado += Number(t.importe) || 0;
    });
    return Object.values(map).sort((a, b) => b.presupuestado - a.presupuestado);
  }, [partidasFiltradasMes, transFiltradasMes]);

  // La tendencia mensual siempre usa TODOS los meses (sin filtrar), es la única
  // vista pensada justo para comparar entre meses.
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

  const totalPresupuestado = partidasFiltradasMes.reduce((s, p) => s + (Number(p.monto_estimado) || 0), 0);
  const totalEjecutado = transFiltradasMes.reduce((s, t) => s + (Number(t.importe) || 0), 0);

  const COLORS = [T.accent, T.teal, T.blue, T.amber, T.red, "#8B6FB0", "#B0955B", "#5BA0B0"];

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
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 12 }}>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          <KpiCard label="Presupuestado" value={money(totalPresupuestado)} />
          <KpiCard label="Ejecutado" value={money(totalEjecutado)} accent={totalPresupuestado && totalEjecutado / totalPresupuestado > 1 ? T.red : T.teal} />
          <KpiCard label="Disponible" value={money(totalPresupuestado - totalEjecutado)} accent={totalPresupuestado - totalEjecutado < 0 ? T.red : T.text} />
        </div>
        <Field label="Mes">
          <MesMultiSelect mesesDisponibles={mesesDisponibles} seleccionados={mesesSeleccionados} onChange={setMesesSeleccionados} />
        </Field>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 20 }}>
        <Panel title="Presupuesto vs. ejecutado por proyecto" subtitle={`Gastos compartidos ya prorrateados según su marcador${mesLabel}`}>
          <ResponsiveContainer width="100%" height={Math.max(220, porProyecto.length * 34)}>
            <BarChart data={porProyecto} layout="vertical" margin={{ left: 8, right: 16 }}>
              <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} horizontal={false} />
              <XAxis type="number" tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
              <YAxis type="category" dataKey="proyecto" width={110} tick={{ fill: T.textDim, fontSize: 11 }} stroke={T.border} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
              <Bar dataKey="presupuestado" name="Presupuestado" fill={T.border} radius={[0,3,3,0]} />
              <Bar dataKey="ejecutado" name="Ejecutado" fill={T.accent} radius={[0,3,3,0]} />
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
              <Line type="monotone" dataKey="ejecutado" name="Ejecutado" stroke={T.accent} strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </Panel>
      )}

      <ResumenPivotPanel partidasUnidad={partidasFiltradasMes} />

      <Panel title="Avance por proyecto" subtitle={`Detalle con marcador de prorrateo aplicado${mesLabel}`}>
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

function ResumenPivotPanel({ partidasUnidad }) {
  const partidasMXN = partidasUnidad.filter((p) => (p.moneda || "MXN") === "MXN");
  const meses = MESES.filter((m) => partidasMXN.some((p) => p.mes === m));
  const [collapsed, setCollapsed] = useState(new Set());
  const toggle = (path) => setCollapsed((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });

  if (!partidasMXN.length) return null;

  const pivot = pivotearPorMes(partidasMXN, ["proyecto", "rubro", "concepto"], meses);
  const totalGeneral = meses.reduce((acc, m) => { acc[m] = 0; return acc; }, {});
  let granTotal = 0;
  partidasMXN.forEach((p) => {
    const v = Number(p.monto_estimado) || 0;
    if (totalGeneral[p.mes] !== undefined) totalGeneral[p.mes] += v;
    granTotal += v;
  });

  return (
    <Panel title="Resumen presupuestado por proyecto y rubro" subtitle="Solo montos en MXN — clic en una fila para expandir/colapsar">
      <div style={{ overflowX: "auto" }}>
        <table style={tableStyle}>
          <thead>
            <tr>
              <th style={thStyle}>Proyecto</th>
              {meses.map((m) => <th key={m} style={{ ...thStyle, textAlign: "right" }}>{m}</th>)}
              <th style={{ ...thStyle, textAlign: "right" }}>Total general</th>
            </tr>
          </thead>
          <tbody>
            {buildPivotTrs(pivot, "", collapsed, toggle, meses, 0)}
            <tr style={{ background: T.panelAlt }}>
              <td style={{ ...tdStyle, fontWeight: 700 }}>Total general</td>
              {meses.map((m) => <td key={m} style={{ ...tdStyle, fontFamily: T.fontMono, textAlign: "right", fontWeight: 700 }}>{money(totalGeneral[m])}</td>)}
              <td style={{ ...tdStyle, fontFamily: T.fontMono, textAlign: "right", fontWeight: 700 }}>{money(granTotal)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      {partidasUnidad.length > partidasMXN.length && (
        <div style={{ fontSize: 11, color: T.textFaint, marginTop: 10 }}>
          {partidasUnidad.length - partidasMXN.length} partida(s) en USD no se incluyen en esta tabla (no se suman monedas distintas sin tipo de cambio).
        </div>
      )}
    </Panel>
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

// Panel desplegable de agrupamiento estilo Airtable: botón "Agrupar por" que abre
// un panel con un renglón por nivel (campo + dirección + quitar), botón para agregar
// subgrupo, y accesos para contraer/expandir todo. `options` = [{value,label}].
// `value` = [{field, dir}]. `groupedTree` (opcional) habilita Contraer/Expandir todo.
function GroupByControl({ options, value, onChange, maxLevels = 3, groupedTree, collapsed, setCollapsed }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const onClickFuera = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClickFuera);
    return () => document.removeEventListener("mousedown", onClickFuera);
  }, []);

  const usedFields = value.map((v) => v.field);
  const opcionesPara = (fieldActual) => options.filter((o) => o.value && (o.value === fieldActual || !usedFields.includes(o.value)));
  const primeraDisponible = () => options.find((o) => o.value && !usedFields.includes(o.value));

  const setLevel = (i, patch) => onChange(value.map((v, idx) => (idx === i ? { ...v, ...patch } : v)));
  const removeLevel = (i) => onChange(value.filter((_, idx) => idx !== i));
  const addLevel = () => {
    const next = primeraDisponible();
    if (next) onChange([...value, { field: next.value, dir: "asc" }]);
  };

  const label = value.length === 0
    ? "Agrupar por"
    : `Agrupado por ${options.find((o) => o.value === value[0].field)?.label || value[0].field}${value.length > 1 ? ` +${value.length - 1}` : ""}`;

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <Button variant={value.length ? "primary" : "ghost"} onClick={() => setOpen((o) => !o)}>{label}</Button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 6px)", left: 0, zIndex: 50,
          background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8,
          padding: 14, minWidth: 340, boxShadow: "0 8px 24px rgba(35,42,49,0.14)",
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: T.text }}>Agrupar por</span>
            {groupedTree && value.length > 0 && (
              <div style={{ display: "flex", gap: 10 }}>
                <button
                  type="button"
                  onClick={() => setCollapsed(new Set(collectGroupPaths(groupedTree)))}
                  style={{ background: "none", border: "none", color: T.accent, fontSize: 11, cursor: "pointer", padding: 0 }}
                >
                  Contraer todo
                </button>
                <button
                  type="button"
                  onClick={() => setCollapsed(new Set())}
                  style={{ background: "none", border: "none", color: T.accent, fontSize: 11, cursor: "pointer", padding: 0 }}
                >
                  Expandir todos
                </button>
              </div>
            )}
          </div>

          {value.length === 0 && (
            <div style={{ fontSize: 11.5, color: T.textFaint, marginBottom: 10 }}>Sin agrupar todavía.</div>
          )}

          {value.map((lvl, i) => (
            <div key={i} style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 8 }}>
              <Select value={lvl.field} onChange={(e) => setLevel(i, { field: e.target.value })} style={{ flex: 1 }}>
                {opcionesPara(lvl.field).map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
              </Select>
              <Select value={lvl.dir} onChange={(e) => setLevel(i, { dir: e.target.value })} style={{ width: 90 }}>
                <option value="asc">A → Z</option>
                <option value="desc">Z → A</option>
              </Select>
              <button
                type="button"
                onClick={() => removeLevel(i)}
                title="Quitar nivel"
                style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 6, color: T.textFaint, cursor: "pointer", width: 30, height: 30, fontSize: 13 }}
              >
                🗑
              </button>
            </div>
          ))}

          {value.length < maxLevels && primeraDisponible() && (
            <button
              type="button"
              onClick={addLevel}
              style={{ background: "none", border: "none", color: T.accent, fontSize: 12, cursor: "pointer", padding: "4px 0", fontWeight: 600 }}
            >
              + Añadir subgrupo
            </button>
          )}
        </div>
      )}
    </div>
  );
}


function MesMultiSelect({ mesesDisponibles, seleccionados, onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const onClickFuera = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClickFuera);
    return () => document.removeEventListener("mousedown", onClickFuera);
  }, []);

  const label = seleccionados.length === 0 ? "Todos los meses" : seleccionados.length === 1 ? seleccionados[0] : `${seleccionados.length} meses`;
  const toggleMes = (m) => onChange(seleccionados.includes(m) ? seleccionados.filter((x) => x !== m) : [...seleccionados, m]);

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        style={{ ...inputStyle, display: "inline-flex", alignItems: "center", gap: 8, width: 150, justifyContent: "space-between", cursor: "pointer" }}
      >
        <span>{label}</span>
        <span style={{ fontSize: 10, color: T.textFaint }}>{open ? "▴" : "▾"}</span>
      </button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 50,
          background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8,
          padding: 10, minWidth: 190, boxShadow: "0 8px 24px rgba(35,42,49,0.14)",
        }}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: T.text, padding: "5px 4px", cursor: "pointer", borderBottom: `1px solid ${T.borderSoft}`, marginBottom: 4, paddingBottom: 8 }}>
            <input type="checkbox" checked={seleccionados.length === 0} onChange={() => onChange([])} />
            Todos los meses
          </label>
          {mesesDisponibles.map((m) => (
            <label key={m} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: T.text, padding: "5px 4px", cursor: "pointer" }}>
              <input type="checkbox" checked={seleccionados.includes(m)} onChange={() => toggleMes(m)} />
              {m}
            </label>
          ))}
          <Button variant="ghost" onClick={() => setOpen(false)} style={{ width: "100%", marginTop: 8 }}>Cerrar</Button>
        </div>
      )}
    </div>
  );
}


function Modal({ title, subtitle, onClose, children, width = 720 }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  return (
    <div
      style={{
        position: "fixed", inset: 0, background: "rgba(6,9,10,0.72)", zIndex: 1000,
        display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "48px 20px", overflowY: "auto",
      }}
      onClick={onClose}
    >
      <div
        style={{ ...panelStyle, width: "100%", maxWidth: width, position: "relative" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{title}</div>
            {subtitle && <div style={{ fontSize: 11, color: T.textFaint, marginTop: 2 }}>{subtitle}</div>}
          </div>
          <button
            onClick={onClose}
            aria-label="Cerrar"
            style={{ background: "transparent", border: `1px solid ${T.border}`, borderRadius: 6, color: T.textDim, cursor: "pointer", width: 28, height: 28, fontSize: 14, lineHeight: 1 }}
          >
            ✕
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

const panelStyle = { background: T.panel, border: `1px solid ${T.border}`, borderRadius: 10, padding: 18 };

// Sortable table header + generic multi-type comparator (used by Partidas & Transacciones tables)
// Recuerda el ancho de cada columna (por tabla, via localStorage) y expone el
// manejador para arrastrar el borde derecho de un encabezado y redimensionarlo.
function useColumnWidths(storageKey, defaultWidth = 160) {
  const [widths, setWidths] = useState(() => {
    try { return JSON.parse(localStorage.getItem(storageKey) || "{}"); } catch { return {}; }
  });
  const getWidth = (key) => widths[key] || defaultWidth;
  const startResize = (key, e) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = getWidth(key);
    const onMove = (ev) => {
      const next = Math.max(70, startWidth + (ev.clientX - startX));
      setWidths((w) => ({ ...w, [key]: next }));
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      setWidths((w) => {
        try { localStorage.setItem(storageKey, JSON.stringify(w)); } catch {}
        return w;
      });
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  };
  return { getWidth, startResize };
}

// Recuerda el ORDEN de las columnas (por tabla, via localStorage) y reordena la
// lista de columnas visibles según esa preferencia — las que aparecen nuevas
// (por ejemplo, al cambiar el agrupamiento) se agregan al final.
function useColumnOrder(storageKey, columns) {
  const [order, setOrder] = useState(() => {
    try { return JSON.parse(localStorage.getItem(storageKey) || "[]"); } catch { return []; }
  });
  const byKey = new Map(columns.map((c) => [c.key, c]));
  const fromStored = order.filter((k) => byKey.has(k)).map((k) => byKey.get(k));
  const missing = columns.filter((c) => !order.includes(c.key));
  const ordered = [...fromStored, ...missing];

  const moveColumn = (draggedKey, targetKey) => {
    if (draggedKey === targetKey) return;
    const keys = ordered.map((c) => c.key);
    const from = keys.indexOf(draggedKey);
    const to = keys.indexOf(targetKey);
    if (from === -1 || to === -1) return;
    const next = [...keys];
    next.splice(from, 1);
    next.splice(to, 0, draggedKey);
    setOrder(next);
    try { localStorage.setItem(storageKey, JSON.stringify(next)); } catch {}
  };

  return { ordered, moveColumn };
}

function SortableTh({ label, sortKey, sort, setSort, width, onResizeStart, onDragStart, onDragOver, onDrop }) {
  const active = sort.key === sortKey;
  return (
    <th
      draggable={!!onDragStart}
      onDragStart={onDragStart}
      onDragOver={onDragOver}
      onDrop={onDrop}
      style={{ ...thStyle, position: "relative", cursor: "pointer", userSelect: "none", width, color: active ? T.accent : thStyle.color }}
      onClick={() => setSort((s) => (s.key === sortKey ? { key: sortKey, dir: s.dir === "asc" ? "desc" : "asc" } : { key: sortKey, dir: "asc" }))}
    >
      {label}{active ? (sort.dir === "asc" ? " ▲" : " ▼") : ""}
      {onResizeStart && (
        <span
          draggable={false}
          onMouseDown={(e) => onResizeStart(e)}
          onClick={(e) => e.stopPropagation()}
          style={{ position: "absolute", right: -3, top: 0, bottom: 0, width: 7, cursor: "col-resize", zIndex: 2 }}
        />
      )}
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

const SIN_DATO = "— Sin dato —";

// Groups `rows` recursively by a list of field keys (e.g. ['mes','proyecto']).
// Returns a tree: { type:'rows', rows } at the leaves, or { type:'group', key, entries:[{value,count,sum,child}] }.
// `levels` = [{ key, dir }] — dir "asc" | "desc" controla el orden alfabético del grupo.
function agruparRows(rows, levels, montoKey = "monto_estimado") {
  if (!levels.length) return { type: "rows", rows };
  const [{ field: key, dir = "asc" }, ...rest] = levels;
  const buckets = new Map();
  rows.forEach((r) => {
    const val = (r[key] ?? "").toString().trim() || SIN_DATO;
    if (!buckets.has(val)) buckets.set(val, []);
    buckets.get(val).push(r);
  });
  let entries = [...buckets.entries()];
  if (key === "mes") {
    entries.sort((a, b) => MESES.indexOf(a[0]) - MESES.indexOf(b[0]));
  } else {
    entries.sort((a, b) => a[0].localeCompare(b[0]));
  }
  if (dir === "desc") entries.reverse();
  return {
    type: "group",
    key,
    entries: entries.map(([value, groupRows]) => ({
      value,
      count: groupRows.length,
      sum: groupRows.reduce((s, r) => s + (Number(r[montoKey]) || 0), 0),
      child: agruparRows(groupRows, rest, montoKey),
    })),
  };
}

// Recorre el árbol agrupado y regresa todas las rutas de nivel-grupo (no las hojas) —
// se usa para "Contraer todo" (marcarlas todas como colapsadas de un jalón).
function collectGroupPaths(node, path = "") {
  if (node.type !== "group") return [];
  let paths = [];
  node.entries.forEach((entry) => {
    const groupPath = `${path}/${node.key}:${entry.value}`;
    paths.push(groupPath);
    paths = paths.concat(collectGroupPaths(entry.child, groupPath));
  });
  return paths;
}

// Groups rows by a fixed key path (e.g. ['proyecto','rubro','concepto']) and, at every
// level, also sums the montoKey broken out per month — used by the pivot/resumen view.
function pivotearPorMes(rows, keys, meses, montoKey = "monto_estimado") {
  const mesSums = {};
  meses.forEach((m) => { mesSums[m] = 0; });
  let total = 0;
  rows.forEach((r) => {
    const v = Number(r[montoKey]) || 0;
    if (mesSums[r.mes] !== undefined) mesSums[r.mes] += v;
    total += v;
  });
  if (!keys.length) return { type: "rows", mesSums, total };
  const [key, ...rest] = keys;
  const buckets = new Map();
  rows.forEach((r) => {
    const val = (r[key] ?? "").toString().trim() || SIN_DATO;
    if (!buckets.has(val)) buckets.set(val, []);
    buckets.get(val).push(r);
  });
  const entries = [...buckets.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  return {
    type: "group",
    key,
    mesSums,
    total,
    entries: entries.map(([value, groupRows]) => ({ value, child: pivotearPorMes(groupRows, rest, meses, montoKey) })),
  };
}

// Renders a pivotearPorMes tree as <tr> rows: bold collapsible rows for every level except
// the deepest, which renders as a plain (non-collapsible) leaf row.
function buildPivotTrs(node, path, collapsed, toggleGroup, meses, depth) {
  let out = [];
  node.entries.forEach((entry) => {
    const groupPath = `${path}/${node.key}:${entry.value}`;
    const esHoja = entry.child.type === "rows";
    const cellStyle = { ...tdStyle, fontFamily: T.fontMono, textAlign: "right" };
    if (esHoja) {
      out.push(
        <tr key={groupPath}>
          <td style={{ ...tdStyle, paddingLeft: 14 + depth * 22 }}>{entry.value}</td>
          {meses.map((m) => <td key={m} style={cellStyle}>{entry.child.mesSums[m] ? money(entry.child.mesSums[m]) : "—"}</td>)}
          <td style={{ ...cellStyle, fontWeight: 600 }}>{money(entry.child.total)}</td>
        </tr>
      );
    } else {
      const isCollapsed = collapsed.has(groupPath);
      out.push(
        <tr key={groupPath} onClick={() => toggleGroup(groupPath)} style={{ cursor: "pointer", background: T.panelAlt }}>
          <td style={{ ...tdStyle, paddingLeft: 14 + depth * 22, fontWeight: 600, display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ color: T.textFaint, fontSize: 10 }}>{isCollapsed ? "▶" : "▼"}</span>{entry.value}
          </td>
          {meses.map((m) => <td key={m} style={{ ...cellStyle, fontWeight: 600 }}>{entry.child.mesSums[m] ? money(entry.child.mesSums[m]) : "—"}</td>)}
          <td style={{ ...cellStyle, fontWeight: 600 }}>{money(entry.child.total)}</td>
        </tr>
      );
      if (!isCollapsed) out = out.concat(buildPivotTrs(entry.child, groupPath, collapsed, toggleGroup, meses, depth + 1));
    }
  });
  return out;
}

// Flattens a grouped tree into <tr> elements: a header row per group (collapsible,
// with count + sum), followed by that group's leaf rows (via renderRowTr) when expanded.
const GROUP_LEVEL_COLORS = [T.accent, T.teal, T.blue];
function buildGroupedTrs(node, path, collapsed, toggleGroup, colSpan, depth, renderRowTr, fieldLabels = {}) {
  if (node.type === "rows") return node.rows.map((r) => renderRowTr(r, depth));
  let out = [];
  const levelColor = GROUP_LEVEL_COLORS[depth % GROUP_LEVEL_COLORS.length];
  node.entries.forEach((entry) => {
    const groupPath = `${path}/${node.key}:${entry.value}`;
    const isCollapsed = collapsed.has(groupPath);
    out.push(
      <tr key={groupPath} onClick={() => toggleGroup(groupPath)} style={{ cursor: "pointer" }}>
        <td
          colSpan={colSpan}
          style={{
            ...tdStyle, background: T.panelAlt,
            paddingLeft: 14 + depth * 26,
            borderLeft: `3px solid ${levelColor}`,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ color: T.textFaint, fontSize: 10, width: 12 }}>{isCollapsed ? "▶" : "▼"}</span>
            <span style={{ fontSize: 10.5, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.04em" }}>{fieldLabels[node.key] || node.key}:</span>
            <Pill tone="accent">{entry.value}</Pill>
            <span style={{ fontSize: 11, color: T.textFaint }}>{entry.count}</span>
            <span style={{ fontSize: 11.5, fontFamily: T.fontMono, color: T.textDim, marginLeft: "auto" }}>{money(entry.sum)}</span>
          </div>
        </td>
      </tr>
    );
    if (!isCollapsed) out = out.concat(buildGroupedTrs(entry.child, groupPath, collapsed, toggleGroup, colSpan, depth + 1, renderRowTr, fieldLabels));
  });
  return out;
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
      // Marca cada fila como "actualiza" (ya existe una partida con ese folio+unidad)
      // o "nueva" — para no duplicar ni romper transacciones ya vinculadas por folio.
      const porFolio = new Map();
      partidas.forEach((p) => { if (p.folio) porFolio.set(`${p.unidad}::${p.folio}`, p); });
      const rowsConEstado = rows.map((r) => {
        const existente = r.folio ? porFolio.get(`${r.unidad}::${r.folio}`) : null;
        return { ...r, _existenteId: existente ? existente.id : null };
      });
      setPreview({ rows: rowsConEstado, sheetsFound, sheetsIgnored, fileName: file.name });
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
      const nuevas = preview.rows.filter((r) => !r._existenteId).map(({ _existenteId, ...rest }) => rest);
      const actualizaciones = preview.rows.filter((r) => r._existenteId);

      if (nuevas.length) await partidasApi.bulkInsert(nuevas);
      for (const r of actualizaciones) {
        const { id, _existenteId, ...patch } = r;
        await partidasApi.update(_existenteId, patch);
      }

      const unidadesImportadas = [...new Set(preview.rows.map((r) => r.unidad))];
      setStatus(`${nuevas.length} nuevas, ${actualizaciones.length} actualizadas (${unidadesImportadas.join(", ")}).`);
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
              <Pill key={s.sheetName} tone="accent">{s.sheetName}: {s.count} filas</Pill>
            ))}
            {preview.sheetsIgnored.map((s) => (
              <Pill key={s} tone="dim">{s} (ignorada)</Pill>
            ))}
          </div>

          <div style={{ overflowX: "auto", maxHeight: 220, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["","Unidad","Mes","Concepto","Rubro","Proyecto","Folio","Monto"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.slice(0, 12).map((r) => (
                  <tr key={r.id}>
                    <td style={tdStyle}>{r._existenteId ? <Pill tone="amber">Actualiza</Pill> : <Pill tone="teal">Nueva</Pill>}</td>
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
          <div style={{ fontSize: 11.5, color: T.textDim, marginTop: 10 }}>
            {preview.rows.filter((r) => r._existenteId).length} partidas ya existentes se van a <strong>actualizar</strong> (mismo folio, se conserva su vínculo con transacciones ya registradas) —
            {" "}{preview.rows.filter((r) => !r._existenteId).length} son <strong>nuevas</strong>.
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 16, marginTop: 14, flexWrap: "wrap" }}>
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
  const [modalOpen, setModalOpen] = useState(false);

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

  const GROUP_OPCIONES = [
    { value: "mes", label: "Mes" },
    { value: "rubro", label: "Rubro" },
    { value: "categoria", label: "Categoría" },
    { value: "proyecto", label: "Proyecto" },
  ];
  const [groupBys, setGroupBys] = useState([]);
  const [collapsedGroups, setCollapsedGroups] = useState(new Set());
  const toggleGroup = (path) => setCollapsedGroups((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });
  const groupKeys = groupBys.map((g) => g.field);
  const grouped = groupKeys.length ? agruparRows(partidasOrdenadas, groupBys) : null;

  const COLUMNAS_PARTIDA = [
    { key: "mes", label: "Mes", render: (p) => p.mes },
    { key: "anio", label: "Año", render: (p) => p.anio },
    { key: "concepto", label: "Concepto", render: (p) => p.concepto },
    { key: "rubro", label: "Rubro", render: (p) => <Pill>{p.rubro}</Pill> },
    { key: "categoria", label: "Categoría", render: (p) => <span style={{ color: T.textDim }}>{p.categoria}</span> },
    { key: "proyecto", label: "Proyecto", render: (p) => p.proyecto },
    { key: "folio", label: "Folio", render: (p) => <span style={{ fontFamily: T.fontMono, color: T.textDim }}>{p.folio || "—"}</span> },
    { key: "monto_estimado", label: "Monto", render: (p) => <span style={{ fontFamily: T.fontMono }}>{money(p.monto_estimado, p.moneda)}</span> },
  ];
  const columnasVisiblesBase = COLUMNAS_PARTIDA.filter((c) => c.key === "proyecto" || !groupKeys.includes(c.key));
  const colWidths = useColumnWidths("colw-partidas");
  const { ordered: columnasVisibles, moveColumn } = useColumnOrder("colo-partidas", columnasVisiblesBase);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };
  const renderRowTr = (p, depth = 0) => (
    <tr key={p.id}>
      {columnasVisibles.map((c, i) => (
        <td key={c.key} style={i === 0 && depth ? { ...tdStyle, paddingLeft: 14 + depth * 26 } : tdStyle}>{c.render(p)}</td>
      ))}
      <td style={tdStyle}>
        <div style={{ display: "flex", gap: 6 }}>
          <Button variant="ghost" onClick={() => startEdit(p)}>Editar</Button>
          <Button variant="danger" onClick={() => remove(p.id)}>Eliminar</Button>
        </div>
      </td>
    </tr>
  );


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
      setModalOpen(false);
    } catch (err) {
      alert("No se pudo guardar la partida: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };

  const openNew = () => { setForm({ ...blank, anio: anioDefault, proyecto: marcadores[0] || "" }); setEditId(null); setModalOpen(true); };
  const startEdit = (p) => { setForm(p); setEditId(p.id); setModalOpen(true); };
  const closeModal = () => { setModalOpen(false); setEditId(null); setForm({ ...blank, anio: anioDefault, proyecto: marcadores[0] || "" }); };
  const remove = (id) => partidasApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel
        title={`Partidas de ${unidad}`}
        subtitle={filtrosActivos ? `${partidasFiltradas.length} de ${partidasUnidad.length} registradas` : `${partidasUnidad.length} registradas`}
        right={<Button onClick={openNew}>+ Nueva partida</Button>}
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
          <Field label="Proyecto">
            <Select value={filtros.proyecto} onChange={(e) => setFiltros({ ...filtros, proyecto: e.target.value })} style={{ width: 180 }}>
              <option>Todos</option>
              {proyectosDisponiblesFiltro.map((p) => <option key={p}>{p}</option>)}
            </Select>
          </Field>
          <Field label="Rubro">
            <Select value={filtros.rubro} onChange={(e) => setFiltros({ ...filtros, rubro: e.target.value })} style={{ width: 200 }}>
              <option>Todos</option>
              {rubrosDisponiblesFiltro.map((r) => <option key={r}>{r}</option>)}
            </Select>
          </Field>
          {filtrosActivos && <Button variant="ghost" onClick={limpiarFiltros}>Limpiar filtros</Button>}
          <div style={{ width: 1, alignSelf: "stretch", background: T.borderSoft, margin: "0 4px" }} />
          <GroupByControl
            options={GROUP_OPCIONES}
            value={groupBys}
            onChange={(v) => { setGroupBys(v); setCollapsedGroups(new Set()); }}
            maxLevels={3}
            groupedTree={grouped}
            collapsed={collapsedGroups}
            setCollapsed={setCollapsedGroups}
          />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={{ ...tableStyle, tableLayout: "fixed" }}>
            <colgroup>
              {columnasVisibles.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
              <col style={{ width: 140 }} />
            </colgroup>
            <thead>
              <tr>
                {columnasVisibles.map((c) => (
                  <SortableTh
                    key={c.key} label={c.label} sortKey={c.key} sort={sort} setSort={setSort}
                    width={colWidths.getWidth(c.key)} onResizeStart={(e) => colWidths.startResize(c.key, e)}
                    onDragStart={(e) => onColDragStart(e, c.key)} onDragOver={onColDragOver} onDrop={(e) => onColDrop(e, c.key)}
                  />
                ))}
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {groupKeys.length
                ? buildGroupedTrs(grouped, "", collapsedGroups, toggleGroup, columnasVisibles.length + 1, 0, renderRowTr, Object.fromEntries(GROUP_OPCIONES.map((o) => [o.value, o.label])))
                : partidasOrdenadas.map((p) => renderRowTr(p))}
              {!partidasUnidad.length && (
                <tr><td colSpan={columnasVisibles.length + 1} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin partidas aún</td></tr>
              )}
              {partidasUnidad.length > 0 && !partidasFiltradas.length && (
                <tr><td colSpan={columnasVisibles.length + 1} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna partida coincide con estos filtros</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Panel>

      <ImportarExcelPanel partidas={partidas} partidasApi={partidasApi} />

      {modalOpen && (
        <Modal
          title={editId ? "Editar partida" : "Nueva partida presupuestal"}
          subtitle="Una fila por partida — puede recibir varias transacciones reales"
          onClose={closeModal}
        >
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
            <Field label="Concepto" style={{ gridColumn: "span 2" }}>
              <TextInput value={form.concepto} onChange={(e) => setForm({ ...form, concepto: e.target.value })} placeholder="Ej. Servicio energía eléctrica base admtva" />
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
            {form.extra && Object.keys(form.extra).length > 0 && (
              <details style={{ gridColumn: "span 4" }}>
                <summary style={{ cursor: "pointer", fontSize: 11.5, color: T.textDim, fontWeight: 600 }}>
                  Detalles adicionales del Excel ({Object.keys(form.extra).length})
                </summary>
                <div style={{ marginTop: 8, display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "4px 20px", background: T.panelAlt, borderRadius: 6, padding: 12 }}>
                  {Object.entries(form.extra).map(([k, v]) => (
                    <div key={k} style={{ fontSize: 11.5, display: "flex", justifyContent: "space-between", gap: 8, borderBottom: `1px solid ${T.borderSoft}`, padding: "4px 0" }}>
                      <span style={{ color: T.textFaint }}>{k}</span>
                      <span style={{ color: T.text, textAlign: "right" }}>{String(v)}</span>
                    </div>
                  ))}
                </div>
              </details>
            )}
            <div style={{ gridColumn: "span 4", display: "flex", gap: 10, marginTop: 4 }}>
              <Button type="submit" disabled={saving}>{saving ? "Guardando…" : editId ? "Guardar cambios" : "Agregar partida"}</Button>
              <Button type="button" variant="ghost" onClick={closeModal}>Cancelar</Button>
            </div>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------
   TABS: TRANSACCIONES
---------------------------------------------------------------------- */
function ImportarTransaccionesPanel({ partidas, proveedores, transaccionesApi }) {
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
      const { rows, sheetsFound, matched, unmatched } = parseTransaccionesWorkbook(buf, partidas, proveedores);
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
      const toInsert = preview.rows.map((r) => ({ ...r, partida_id: r.partida_id || null, proveedor_id: r.proveedor_id || null }));
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
                      <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{r.folio_original || <span style={{ color: T.amber }}>(sin folio)</span>}</td>
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

function TransaccionesTab({ unidad, unidades, partidas, transacciones, transaccionesApi, proveedoresApi, cuentasApi }) {
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadoresProyecto = marcadoresDisponibles(proyectosUnidad);
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const blank = {
    partida_id: partidasUnidad[0]?.id || "", dia: "", solicitante: "", proyecto: "", zona: "", area: "",
    proveedor: "", proveedor_id: "", cuenta_id: "", concepto_detallado: "", importe: "", moneda: "MXN", status: "",
    folio_compra_sae: "", folio_factura: "", forma_pago: "", metodo_pago: "",
  };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);
  const cuentasDelProveedorSeleccionado = form.proveedor_id ? cuentasApi.rows.filter((c) => c.proveedor_id === form.proveedor_id) : [];
  const transUnidad = transacciones.filter((t) => partidasUnidad.some((p) => p.id === t.partida_id));
  const sinVincular = transacciones.filter((t) => !t.partida_id && t.unidad_detectada === unidad);

  const [filtros, setFiltros] = useState({ texto: "", zona: "Todas", vinculo: "Todas", mes: "Todos", proyecto: "Todos" });
  const [sort, setSort] = useState({ key: "dia", dir: "desc" });

  // Filtros solo para ubicar la partida correcta dentro del selector del modal —
  // no afectan la tabla ni los datos guardados.
  const [filtroPartidaMes, setFiltroPartidaMes] = useState("Todos");
  const [filtroPartidaProyecto, setFiltroPartidaProyecto] = useState("Todos");
  const mesesPartida = MESES.filter((m) => partidasUnidad.some((p) => p.mes === m));
  const proyectosPartida = [...new Set(partidasUnidad.map((p) => p.proyecto).filter(Boolean))].sort();
  const partidasParaSelect = partidasUnidad.filter((p) => {
    if (filtroPartidaMes !== "Todos" && p.mes !== filtroPartidaMes) return false;
    if (filtroPartidaProyecto !== "Todos" && p.proyecto !== filtroPartidaProyecto) return false;
    return true;
  });
  const partidaSeleccionada = partidasUnidad.find((p) => p.id === form.partida_id);
  const opcionesPartida = partidaSeleccionada && !partidasParaSelect.some((p) => p.id === partidaSeleccionada.id)
    ? [partidaSeleccionada, ...partidasParaSelect]
    : partidasParaSelect;

  const partidaDe = (t) => partidasUnidad.find((p) => p.id === t.partida_id);
  const zonasDisponibles = [...new Set(transUnidad.map((t) => t.zona).filter(Boolean))].sort();
  const mesesFiltroTrans = MESES.filter((m) => transUnidad.some((t) => partidaDe(t)?.mes === m));
  const proyectosFiltroTrans = [...new Set(transUnidad.map((t) => partidaDe(t)?.proyecto).filter(Boolean))].sort();

  const transFiltradas = transUnidad.filter((t) => {
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const partida = partidaDe(t);
      const enTexto = [t.proveedor, t.concepto_detallado, t.solicitante, t.zona, partida?.folio, partida?.concepto]
        .some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    if (filtros.zona !== "Todas" && t.zona !== filtros.zona) return false;
    if (filtros.vinculo === "Vinculadas" && !t.partida_id) return false;
    if (filtros.vinculo === "Sin vincular" && t.partida_id) return false;
    if (filtros.mes !== "Todos" && partidaDe(t)?.mes !== filtros.mes) return false;
    if (filtros.proyecto !== "Todos" && partidaDe(t)?.proyecto !== filtros.proyecto) return false;
    return true;
  });
  const filtrosActivos = filtros.texto.trim() || filtros.zona !== "Todas" || filtros.vinculo !== "Todas" || filtros.mes !== "Todos" || filtros.proyecto !== "Todos";
  const limpiarFiltros = () => setFiltros({ texto: "", zona: "Todas", vinculo: "Todas", mes: "Todos", proyecto: "Todos" });

  const transOrdenadas = sortRows(transFiltradas, sort, {
    importe: (r) => Number(r.importe) || 0,
    proveedor: (r) => (r.proveedor || "").toLowerCase(),
    dia: (r) => r.dia || "",
    partida: (r) => (partidaDe(r)?.concepto || "").toLowerCase(),
  });

  // Enriquece cada transacción con datos de su partida vinculada, para poder
  // agrupar por Proyecto/Rubro/Mes aunque esos campos no vivan en la transacción misma.
  const transEnriquecidas = transOrdenadas.map((t) => {
    const p = partidaDe(t);
    return { ...t, _proyecto: p?.proyecto || SIN_DATO, _rubro: p?.rubro || SIN_DATO, _mes: p?.mes || SIN_DATO, _vinculo: t.partida_id ? "Vinculada" : "Sin vincular" };
  });

  const GROUP_OPCIONES_TRANS = [
    { value: "zona", label: "Zona" },
    { value: "area", label: "Área" },
    { value: "proveedor", label: "Proveedor" },
    { value: "proyecto", label: "Proyecto (transacción)" },
    { value: "status", label: "Status" },
    { value: "_proyecto", label: "Proyecto (partida)" },
    { value: "_rubro", label: "Rubro (partida)" },
    { value: "_mes", label: "Mes (partida)" },
    { value: "_vinculo", label: "Vínculo" },
  ];
  const [groupBys, setGroupBys] = useState([]);
  const [collapsedGroups, setCollapsedGroups] = useState(new Set());
  const toggleGroup = (path) => setCollapsedGroups((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });
  const groupKeys = groupBys.map((g) => g.field);
  const grouped = groupKeys.length ? agruparRows(transEnriquecidas, groupBys, "importe") : null;

  const COLUMNAS_TRANS = [
    { key: "dia", label: "Día", render: (t) => t.dia || "—" },
    {
      key: "folio", label: "Folio",
      render: (t) => {
        const p = partidaDe(t);
        return p?.folio
          ? <span style={{ fontFamily: T.fontMono, color: T.textDim }}>{p.folio}</span>
          : (t.folio_original ? <span style={{ fontFamily: T.fontMono, color: T.red }}>{t.folio_original} (sin match)</span> : "—");
      },
    },
    {
      key: "partida", label: "Partida",
      render: (t) => (
        <Select
          value={t.partida_id || ""}
          onChange={(e) => {
            const nuevoId = e.target.value || null;
            const nuevaPartida = partidasUnidad.find((p) => p.id === nuevoId);
            transaccionesApi
              .update(t.id, { partida_id: nuevoId, unidad_detectada: nuevaPartida ? nuevaPartida.unidad : t.unidad_detectada })
              .catch((err) => alert("No se pudo vincular: " + (err.message || err)));
          }}
          style={{ minWidth: 230 }}
        >
          <option value="">— Sin vincular —</option>
          {partidasUnidad.map((p) => <option key={p.id} value={p.id}>{p.mes} · {p.concepto} ({p.proyecto})</option>)}
        </Select>
      ),
    },
    { key: "proveedor", label: "Proveedor", render: (t) => t.proveedor },
    { key: "proyecto", label: "Proyecto", render: (t) => t.proyecto || "—" },
    { key: "zona", label: "Zona", render: (t) => t.zona || "—" },
    { key: "area", label: "Área", render: (t) => t.area || "—" },
    { key: "concepto_detallado", label: "Concepto", render: (t) => <span style={{ color: T.textDim }}>{t.concepto_detallado}</span> },
    { key: "importe", label: "Importe", render: (t) => <span style={{ fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</span> },
    { key: "status", label: "Status", render: (t) => t.status ? <Pill tone={/pagad/i.test(t.status) ? "teal" : "amber"}>{t.status}</Pill> : "—" },
  ];
  const columnasVisiblesBase = COLUMNAS_TRANS.filter((c) => !groupKeys.includes(c.key));
  const colWidths = useColumnWidths("colw-transacciones");
  const { ordered: columnasVisibles, moveColumn } = useColumnOrder("colo-transacciones", columnasVisiblesBase);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };
  const renderRowTr = (t, depth = 0) => (
    <tr key={t.id}>
      {columnasVisibles.map((c, i) => (
        <td key={c.key} style={i === 0 && depth ? { ...tdStyle, paddingLeft: 14 + depth * 26 } : tdStyle}>{c.render(t)}</td>
      ))}
      <td style={tdStyle}>
        <div style={{ display: "flex", gap: 6 }}>
          <Button variant="ghost" onClick={() => startEdit(t)}>Editar</Button>
          <Button variant="danger" onClick={() => remove(t.id)}>Eliminar</Button>
        </div>
      </td>
    </tr>
  );

  const submit = async (e) => {
    e.preventDefault();
    if (!form.partida_id || !form.importe) return;
    const { id, ...rest } = form;
    rest.proveedor_id = rest.proveedor_id || null;
    rest.cuenta_id = rest.cuenta_id || null;
    setSaving(true);
    try {
      if (editId) {
        await transaccionesApi.update(editId, rest);
        setEditId(null);
      } else {
        await transaccionesApi.insert({ ...rest, id: uid() });
      }
      setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" });
      setModalOpen(false);
    } catch (err) {
      alert("No se pudo guardar la transacción: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };
  const openNew = () => { setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" }); setEditId(null); setModalOpen(true); };
  const startEdit = (t) => { setForm(t); setEditId(t.id); setModalOpen(true); };
  const closeModal = () => { setModalOpen(false); setEditId(null); setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" }); };
  const remove = (id) => transaccionesApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));

  if (!partidasUnidad.length) {
    return <EmptyState title="Primero crea partidas" body={`Registra al menos una partida de ${unidad} en la pestaña Partidas antes de capturar transacciones reales.`} />;
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel
        title={`Transacciones de ${unidad}`}
        subtitle={filtrosActivos ? `${transFiltradas.length} de ${transUnidad.length} registradas` : `${transUnidad.length} registradas`}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            {(transUnidad.length + sinVincular.length) > 0 && (
              <Button
                variant="danger"
                onClick={async () => {
                  const total = transUnidad.length + sinVincular.length;
                  if (!confirm(`¿Eliminar las ${total} transacciones de ${unidad} (vinculadas y sin vincular)? Esto no se puede deshacer — útil para reimportar desde cero.`)) return;
                  for (const t of [...transUnidad, ...sinVincular]) {
                    await transaccionesApi.remove(t.id).catch(() => {});
                  }
                }}
              >
                Borrar todas ({transUnidad.length + sinVincular.length})
              </Button>
            )}
            <Button onClick={openNew}>+ Nueva transacción</Button>
          </div>
        }
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
          <Field label="Zona">
            <Select value={filtros.zona} onChange={(e) => setFiltros({ ...filtros, zona: e.target.value })} style={{ width: 170 }}>
              <option>Todas</option>
              {zonasDisponibles.map((z) => <option key={z}>{z}</option>)}
            </Select>
          </Field>
          <Field label="Vínculo">
            <Select value={filtros.vinculo} onChange={(e) => setFiltros({ ...filtros, vinculo: e.target.value })} style={{ width: 150 }}>
              <option>Todas</option>
              <option>Vinculadas</option>
              <option>Sin vincular</option>
            </Select>
          </Field>
          <Field label="Mes (de la partida)">
            <Select value={filtros.mes} onChange={(e) => setFiltros({ ...filtros, mes: e.target.value })} style={{ width: 150 }}>
              <option>Todos</option>
              {mesesFiltroTrans.map((m) => <option key={m}>{m}</option>)}
            </Select>
          </Field>
          <Field label="Proyecto (de la partida)">
            <Select value={filtros.proyecto} onChange={(e) => setFiltros({ ...filtros, proyecto: e.target.value })} style={{ width: 170 }}>
              <option>Todos</option>
              {proyectosFiltroTrans.map((p) => <option key={p}>{p}</option>)}
            </Select>
          </Field>
          {filtrosActivos && <Button variant="ghost" onClick={limpiarFiltros}>Limpiar filtros</Button>}
          <div style={{ width: 1, alignSelf: "stretch", background: T.borderSoft, margin: "0 4px" }} />
          <GroupByControl
            options={GROUP_OPCIONES_TRANS}
            value={groupBys}
            onChange={(v) => { setGroupBys(v); setCollapsedGroups(new Set()); }}
            maxLevels={3}
            groupedTree={grouped}
            collapsed={collapsedGroups}
            setCollapsed={setCollapsedGroups}
          />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={{ ...tableStyle, tableLayout: "fixed" }}>
            <colgroup>
              {columnasVisibles.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
              <col style={{ width: 140 }} />
            </colgroup>
            <thead>
              <tr>
                {columnasVisibles.map((c) => (
                  <SortableTh
                    key={c.key} label={c.label} sortKey={c.key} sort={sort} setSort={setSort}
                    width={colWidths.getWidth(c.key)} onResizeStart={(e) => colWidths.startResize(c.key, e)}
                    onDragStart={(e) => onColDragStart(e, c.key)} onDragOver={onColDragOver} onDrop={(e) => onColDrop(e, c.key)}
                  />
                ))}
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {groupKeys.length
                ? buildGroupedTrs(grouped, "", collapsedGroups, toggleGroup, columnasVisibles.length + 1, 0, renderRowTr, Object.fromEntries(GROUP_OPCIONES_TRANS.map((o) => [o.value, o.label])))
                : transEnriquecidas.map((t) => renderRowTr(t))}
              {!transUnidad.length && (
                <tr><td colSpan={columnasVisibles.length + 1} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin transacciones aún</td></tr>
              )}
              {transUnidad.length > 0 && !transFiltradas.length && (
                <tr><td colSpan={columnasVisibles.length + 1} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna transacción coincide con estos filtros</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Panel>

      {sinVincular.length > 0 && (
        <Panel
          title="Transacciones importadas sin partida vinculada"
          subtitle={`${sinVincular.length} en ${unidad} — su folio no coincidió con ninguna partida`}
          right={
            <Button
              variant="danger"
              onClick={async () => {
                if (!confirm(`¿Eliminar las ${sinVincular.length} transacciones sin vincular de ${unidad}? Esto no se puede deshacer.`)) return;
                for (const t of sinVincular) {
                  await transaccionesApi.remove(t.id).catch(() => {});
                }
              }}
            >
              Eliminar las {sinVincular.length} sin vincular
            </Button>
          }
        >
          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["Día","Folio","Proveedor","Concepto","Importe","Vincular a partida"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {sinVincular.map((t) => (
                  <tr key={t.id}>
                    <td style={tdStyle}>{t.dia || "—"}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.red }}>{t.folio_original || "(sin folio)"}</td>
                    <td style={tdStyle}>{t.proveedor}</td>
                    <td style={{ ...tdStyle, color: T.textDim }}>{t.concepto_detallado}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</td>
                    <td style={tdStyle}>
                      <Select
                        value=""
                        onChange={(e) => {
                          const nuevaPartida = partidasUnidad.find((p) => p.id === e.target.value);
                          transaccionesApi
                            .update(t.id, { partida_id: e.target.value, unidad_detectada: nuevaPartida ? nuevaPartida.unidad : t.unidad_detectada })
                            .catch((err) => alert("No se pudo vincular: " + (err.message || err)));
                        }}
                        style={{ minWidth: 230 }}
                      >
                        <option value="">Elegir partida…</option>
                        {partidasUnidad.map((p) => <option key={p.id} value={p.id}>{p.mes} · {p.concepto} ({p.proyecto})</option>)}
                      </Select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11, color: T.textFaint, marginTop: 10 }}>
            Elige una partida en el selector de cada fila para vincularla — se guarda al instante.
          </div>
        </Panel>
      )}

      <ImportarTransaccionesPanel partidas={partidas} proveedores={proveedoresApi.rows} transaccionesApi={transaccionesApi} />

      {modalOpen && (
        <Modal
          title={editId ? "Editar transacción" : "Nueva transacción real"}
          subtitle="Se vincula a una partida — una partida puede tener varias"
          onClose={closeModal}
        >
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
            <Field label="Filtrar por mes (para buscar la partida)">
              <Select value={filtroPartidaMes} onChange={(e) => setFiltroPartidaMes(e.target.value)}>
                <option>Todos</option>
                {mesesPartida.map((m) => <option key={m}>{m}</option>)}
              </Select>
            </Field>
            <Field label="Filtrar por proyecto">
              <Select value={filtroPartidaProyecto} onChange={(e) => setFiltroPartidaProyecto(e.target.value)}>
                <option>Todos</option>
                {proyectosPartida.map((p) => <option key={p}>{p}</option>)}
              </Select>
            </Field>
            <Field label="Partida" style={{ gridColumn: "span 4" }}>
              <Select value={form.partida_id} onChange={(e) => setForm({ ...form, partida_id: e.target.value })}>
                {!opcionesPartida.length && <option value="">Ningún resultado con este filtro</option>}
                {opcionesPartida.map((p) => <option key={p.id} value={p.id}>{p.mes} · {p.concepto} ({p.proyecto})</option>)}
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
            <Field label="Proyecto">
              <Select value={form.proyecto} onChange={(e) => setForm({ ...form, proyecto: e.target.value })}>
                <option value="">— Sin especificar —</option>
                {marcadoresProyecto.map((m) => <option key={m}>{m}</option>)}
              </Select>
            </Field>
            <Field label="Zona">
              <Select value={form.zona} onChange={(e) => setForm({ ...form, zona: e.target.value })}>
                <option value="">— Sin especificar —</option>
                {ZONAS.map((z) => <option key={z}>{z}</option>)}
              </Select>
            </Field>
            <Field label="Proveedor (catálogo)">
              <Select
                value={form.proveedor_id}
                onChange={(e) => {
                  const p = proveedoresUnidad.find((pr) => pr.id === e.target.value);
                  setForm({ ...form, proveedor_id: e.target.value, proveedor: p ? p.nombre : form.proveedor, cuenta_id: "" });
                }}
              >
                <option value="">— No catalogado —</option>
                {proveedoresUnidad.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
              </Select>
            </Field>
            <Field label="Cuenta bancaria">
              <Select value={form.cuenta_id} onChange={(e) => setForm({ ...form, cuenta_id: e.target.value })} disabled={!form.proveedor_id}>
                <option value="">{form.proveedor_id ? "— Elegir cuenta —" : "— Elige un proveedor primero —"}</option>
                {cuentasDelProveedorSeleccionado.map((c) => (
                  <option key={c.id} value={c.id}>{c.banco || "Banco N/D"} · {c.clabe || c.numero_cuenta || "—"} ({c.divisa})</option>
                ))}
              </Select>
            </Field>
            <Field label="Proveedor / razón social (texto)" style={{ gridColumn: "span 2" }}>
              <TextInput value={form.proveedor} onChange={(e) => setForm({ ...form, proveedor: e.target.value })} placeholder="Se llena solo al elegir del catálogo, o captúralo si no está dado de alta" />
            </Field>
            <Field label="Concepto de pago (detallado)" style={{ gridColumn: "span 2" }}>
              <TextInput value={form.concepto_detallado} onChange={(e) => setForm({ ...form, concepto_detallado: e.target.value })} />
            </Field>
            <Field label="Folio Compra SAE">
              <TextInput value={form.folio_compra_sae} onChange={(e) => setForm({ ...form, folio_compra_sae: e.target.value })} />
            </Field>
            <Field label="Folio Factura">
              <TextInput value={form.folio_factura} onChange={(e) => setForm({ ...form, folio_factura: e.target.value })} />
            </Field>
            <Field label="Forma de Pago">
              <Select value={form.forma_pago} onChange={(e) => setForm({ ...form, forma_pago: e.target.value })}>
                <option value="">— Sin especificar —</option>
                {FORMAS_PAGO.map((f) => <option key={f.value} value={f.value}>{f.label}</option>)}
              </Select>
            </Field>
            <Field label="Método de Pago">
              <Select value={form.metodo_pago} onChange={(e) => setForm({ ...form, metodo_pago: e.target.value })}>
                <option value="">— Sin especificar —</option>
                {METODOS_PAGO.map((m) => <option key={m.value} value={m.value}>{m.label}</option>)}
              </Select>
            </Field>
            <Field label="Importe">
              <TextInput type="number" step="0.01" value={form.importe} onChange={(e) => setForm({ ...form, importe: e.target.value })} placeholder="0.00" />
            </Field>
            <Field label="Moneda">
              <Select value={form.moneda} onChange={(e) => setForm({ ...form, moneda: e.target.value })}>
                {MONEDAS.map((m) => <option key={m}>{m}</option>)}
              </Select>
            </Field>
            <Field label="Status">
              <Select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
                <option value="">— Sin especificar —</option>
                <option>Pagado</option>
                <option>No Pagado</option>
              </Select>
            </Field>
            <div style={{ gridColumn: "span 4", display: "flex", gap: 10, marginTop: 4 }}>
              <Button type="submit" disabled={saving}>{saving ? "Guardando…" : editId ? "Guardar cambios" : "Registrar transacción"}</Button>
              <Button type="button" variant="ghost" onClick={closeModal}>Cancelar</Button>
            </div>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------
   TABS: CATALOGO
---------------------------------------------------------------------- */
/* ----------------------------------------------------------------------
   TAB: REPORTE DE PAGOS
---------------------------------------------------------------------- */
const COLUMNAS_REPORTE = [
  { key: "dia", label: "Día" },
  { key: "solicitante", label: "Solicitante" },
  { key: "area", label: "Área" },
  { key: "numero_solicitud", label: "No. Solicitud (SMI)" },
  { key: "no_sae", label: "No. SAE" },
  { key: "folio_compra_sae", label: "Folio Compra SAE" },
  { key: "folio_factura", label: "Folio Factura" },
  { key: "forma_pago", label: "Forma de Pago" },
  { key: "metodo_pago", label: "Método de Pago" },
  { key: "proveedor", label: "Proveedor" },
  { key: "concepto", label: "Concepto de pago" },
  { key: "banco", label: "Banco" },
  { key: "clabe", label: "Cuenta CLABE" },
  { key: "importe", label: "Importe" },
  { key: "moneda", label: "Moneda" },
];

function ReportePagosTab({ unidad, partidas, transacciones, proveedoresApi, cuentasApi }) {
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const transUnidad = transacciones.filter(
    (t) => partidasUnidad.some((p) => p.id === t.partida_id) || t.unidad_detectada === unidad
  );

  const filas = transUnidad.map((t) => {
    const partida = partidasUnidad.find((p) => p.id === t.partida_id);
    const proveedor = t.proveedor_id ? proveedoresUnidad.find((p) => p.id === t.proveedor_id) : null;
    const cuenta = t.cuenta_id ? cuentasApi.rows.find((c) => c.id === t.cuenta_id) : null;
    return {
      id: t.id,
      dia: t.dia || "",
      solicitante: t.solicitante || "",
      area: t.area || "",
      numero_solicitud: partida?.smi || "",
      no_sae: proveedor?.id_sae || "",
      folio_compra_sae: t.folio_compra_sae || "",
      folio_factura: t.folio_factura || "",
      forma_pago: FORMAS_PAGO.find((f) => f.value === t.forma_pago)?.label || t.forma_pago || "",
      metodo_pago: METODOS_PAGO.find((m) => m.value === t.metodo_pago)?.label || t.metodo_pago || "",
      proveedor: proveedor?.nombre || t.proveedor || "",
      concepto: t.concepto_detallado || "",
      banco: cuenta?.banco || "",
      clabe: cuenta?.clabe || "",
      importe: Number(t.importe) || 0,
      moneda: t.moneda || "MXN",
      _vinculadoProveedor: !!proveedor,
      _vinculadoCuenta: !!cuenta,
    };
  });

  const [buscar, setBuscar] = useState("");
  const [fechaDesde, setFechaDesde] = useState("");
  const [fechaHasta, setFechaHasta] = useState("");
  const [sort, setSort] = useState({ key: "dia", dir: "desc" });
  const filasFiltradas = filas.filter((f) => {
    if (fechaDesde && (!f.dia || f.dia < fechaDesde)) return false;
    if (fechaHasta && (!f.dia || f.dia > fechaHasta)) return false;
    if (!buscar.trim()) return true;
    const q = buscar.trim().toLowerCase();
    return [f.solicitante, f.proveedor, f.concepto, f.folio_factura, f.folio_compra_sae, f.no_sae]
      .some((v) => (v || "").toString().toLowerCase().includes(q));
  });
  const filtrosFechaActivos = fechaDesde || fechaHasta;
  const limpiarFechas = () => { setFechaDesde(""); setFechaHasta(""); };
  const filasOrdenadas = sortRows(filasFiltradas, sort, { importe: (r) => r.importe });

  const colWidths = useColumnWidths("colw-reporte");
  const { ordered: columnas, moveColumn } = useColumnOrder("colo-reporte", COLUMNAS_REPORTE);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };

  const exportarExcel = () => {
    const data = filasOrdenadas.map((f) => Object.fromEntries(columnas.map((c) => [c.label, f[c.key]])));
    const ws = XLSX.utils.json_to_sheet(data);
    const wbx = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbx, ws, "Reporte de pagos");
    XLSX.writeFile(wbx, `reporte-pagos-${unidad}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const totalImporte = filasOrdenadas.reduce((s, f) => s + f.importe, 0);
  const sinProveedorVinculado = filas.filter((f) => !f._vinculadoProveedor && f.proveedor).length;
  const sinCuentaVinculada = filas.filter((f) => f._vinculadoProveedor && !f._vinculadoCuenta).length;

  if (!transUnidad.length) {
    return (
      <EmptyState
        title="Sin transacciones para reportar"
        body={`Todavía no hay transacciones registradas para ${unidad}. Captúralas o impórtalas desde la pestaña Transacciones.`}
      />
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <KpiCard label="Transacciones" value={String(filasOrdenadas.length)} />
        <KpiCard label="Importe total (filtrado)" value={money(totalImporte)} />
      </div>

      {sinProveedorVinculado > 0 && (
        <div style={{ fontSize: 11.5, color: T.amber }}>
          {sinProveedorVinculado} transacción(es) tienen proveedor en texto pero no están ligadas al Catálogo de proveedores —
          No. SAE, Banco y CLABE saldrán vacíos para esas filas hasta que se vinculen (edítalas en Transacciones y elige el proveedor del catálogo).
        </div>
      )}
      {sinCuentaVinculada > 0 && (
        <div style={{ fontSize: 11.5, color: T.amber }}>
          {sinCuentaVinculada} transacción(es) tienen proveedor vinculado pero no una cuenta bancaria específica —
          Banco y CLABE saldrán vacíos hasta que elijas la cuenta (edítalas en Transacciones).
        </div>
      )}

      <Panel
        title={`Reporte de pagos — ${unidad}`}
        subtitle={`${filasOrdenadas.length} de ${filas.length} transacciones`}
        right={<Button onClick={exportarExcel}>Exportar a Excel</Button>}
      >
        <div style={{ marginBottom: 14, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Buscar">
            <TextInput
              value={buscar}
              onChange={(e) => setBuscar(e.target.value)}
              placeholder="Solicitante, proveedor, concepto, folio…"
              style={{ width: 280 }}
            />
          </Field>
          <Field label="Desde">
            <TextInput type="date" value={fechaDesde} onChange={(e) => setFechaDesde(e.target.value)} />
          </Field>
          <Field label="Hasta">
            <TextInput type="date" value={fechaHasta} onChange={(e) => setFechaHasta(e.target.value)} />
          </Field>
          {filtrosFechaActivos && <Button variant="ghost" onClick={limpiarFechas}>Limpiar fechas</Button>}
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ ...tableStyle, tableLayout: "fixed" }}>
            <colgroup>
              {columnas.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
            </colgroup>
            <thead>
              <tr>
                {columnas.map((c) => (
                  <SortableTh
                    key={c.key} label={c.label} sortKey={c.key} sort={sort} setSort={setSort}
                    width={colWidths.getWidth(c.key)} onResizeStart={(e) => colWidths.startResize(c.key, e)}
                    onDragStart={(e) => onColDragStart(e, c.key)} onDragOver={onColDragOver} onDrop={(e) => onColDrop(e, c.key)}
                  />
                ))}
              </tr>
            </thead>
            <tbody>
              {filasOrdenadas.map((f) => (
                <tr key={f.id}>
                  {columnas.map((c) => (
                    <td key={c.key} style={c.key === "importe" ? { ...tdStyle, fontFamily: T.fontMono } : tdStyle}>
                      {c.key === "importe" ? money(f.importe, f.moneda) : (f[c.key] || "—")}
                    </td>
                  ))}
                </tr>
              ))}
              {!filasOrdenadas.length && (
                <tr><td colSpan={columnas.length} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna transacción coincide con la búsqueda</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Panel>
    </div>
  );
}

function CatalogoTab({ unidad, unidades, proyectosApi, proveedoresApi, cuentasApi }) {
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

      <Panel title="Catálogo de zonas" subtitle="Usadas en el selector de Zona al capturar transacciones">
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {ZONAS.map((z) => <Pill key={z}>{z}</Pill>)}
        </div>
      </Panel>

      <ProveedoresPanel unidad={unidad} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />
    </div>
  );
}

function ImportarProveedoresPanel({ proveedoresApi, cuentasApi }) {
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
      const { rows, sheetsFound, nuevas, actualizaciones, sinCompania } = parseProveedoresWorkbook(buf, proveedoresApi.rows);
      if (!rows.length) {
        setError('No encontré filas con Nombre y Compañía (OSB/CTM/ISE) reconocible en este archivo.');
        return;
      }
      setPreview({ rows, sheetsFound, nuevas, actualizaciones, sinCompania, fileName: file.name });
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
      const cuentasAInsertar = []; // { proveedor_id, ...cuenta } — se llena después de saber el id real de cada proveedor

      const nuevasFilas = preview.rows.filter((r) => !r._existenteId);
      const actualizacionesFilas = preview.rows.filter((r) => r._existenteId);

      if (nuevasFilas.length) {
        const payload = nuevasFilas.map(({ _existenteId, _cuenta, ...rest }) => rest);
        const creados = await proveedoresApi.bulkInsert(payload);
        creados.forEach((c, i) => {
          const cuenta = nuevasFilas[i]._cuenta;
          if (cuenta && (cuenta.banco || cuenta.clabe || cuenta.numero_cuenta)) {
            cuentasAInsertar.push({ id: uid(), proveedor_id: c.id, ...cuenta });
          }
        });
      }
      for (const r of actualizacionesFilas) {
        const { id, _existenteId, _cuenta, ...patch } = r;
        await proveedoresApi.update(_existenteId, patch);
        if (_cuenta && (_cuenta.banco || _cuenta.clabe || _cuenta.numero_cuenta)) {
          cuentasAInsertar.push({ id: uid(), proveedor_id: _existenteId, ..._cuenta });
        }
      }
      if (cuentasAInsertar.length) await cuentasApi.bulkInsert(cuentasAInsertar);

      setStatus(`${nuevasFilas.length} nuevos, ${actualizacionesFilas.length} actualizados, ${cuentasAInsertar.length} cuenta(s) agregadas.`);
      setPreview(null);
    } catch (err) {
      setError("Ocurrió un error al importar en Supabase: " + (err.message || err));
    } finally {
      setImporting(false);
    }
  };

  return (
    <Panel title="Carga masiva de proveedores" subtitle='Una fila por proveedor, con una columna "Compañía" (OSB/CTM/ISE) que dice a cuál pertenece — o usa hojas llamadas OSB/CTM/ISE como respaldo'>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <input ref={inputRef} type="file" accept=".xlsx" onChange={onFile} style={{ fontSize: 12, color: T.textDim }} />
        {status && <Pill tone="teal">{status}</Pill>}
      </div>
      {error && <div style={{ marginTop: 10, fontSize: 12, color: T.red }}>{error}</div>}

      {preview && (
        <div style={{ marginTop: 16, borderTop: `1px solid ${T.borderSoft}`, paddingTop: 14 }}>
          <div style={{ fontSize: 12.5, color: T.text, marginBottom: 8 }}>
            <strong>{preview.fileName}</strong> — {preview.rows.length} proveedores detectados
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
            {preview.sheetsFound.map((s) => (
              <Pill key={s.sheetName} tone="accent">{s.sheetName}: {s.count} filas</Pill>
            ))}
            <Pill tone="teal">{preview.nuevas} nuevos</Pill>
            {preview.actualizaciones > 0 && <Pill tone="amber">{preview.actualizaciones} actualizan uno existente</Pill>}
            {preview.sinCompania > 0 && <Pill tone="red">{preview.sinCompania} ignoradas — sin Compañía reconocible</Pill>}
          </div>
          {preview.sinCompania > 0 && (
            <div style={{ fontSize: 11.5, color: T.amber, marginBottom: 12 }}>
              {preview.sinCompania} fila(s) no se van a importar porque la columna "Compañía" viene vacía o con un valor distinto a OSB/CTM/ISE,
              y la hoja tampoco se llama así. Corrígelo en el Excel y vuelve a subirlo.
            </div>
          )}

          <div style={{ overflowX: "auto", maxHeight: 260, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["","Unidad","Nombre","RFC","Banco","Divisa"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.slice(0, 15).map((r) => (
                  <tr key={r.id}>
                    <td style={tdStyle}>{r._existenteId ? <Pill tone="amber">Actualiza</Pill> : <Pill tone="teal">Nuevo</Pill>}</td>
                    <td style={tdStyle}>{r.unidad}</td>
                    <td style={tdStyle}>{r.nombre}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{r.rfc || "—"}</td>
                    <td style={tdStyle}>{r._cuenta.banco || "—"}</td>
                    <td style={tdStyle}>{r._cuenta.banco || r._cuenta.clabe ? r._cuenta.divisa : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {preview.rows.length > 15 && (
            <div style={{ fontSize: 11, color: T.textFaint, marginTop: 6 }}>… y {preview.rows.length - 15} más</div>
          )}

          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <Button onClick={confirmar} disabled={importing}>{importing ? "Importando…" : "Confirmar importación"}</Button>
            <Button variant="ghost" onClick={() => setPreview(null)}>Cancelar</Button>
          </div>
        </div>
      )}
    </Panel>
  );
}

function ProveedoresPanel({ unidad, proveedoresApi, cuentasApi }) {
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const blank = { nombre: "", rfc: "", id_sae: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [saving, setSaving] = useState(false);
  const [buscar, setBuscar] = useState("");

  const cuentaBlank = { banco: "", clabe: "", numero_cuenta: "", divisa: "MXN" };
  const [nuevaCuenta, setNuevaCuenta] = useState(cuentaBlank);
  const cuentasDelProveedor = editId ? cuentasApi.rows.filter((c) => c.proveedor_id === editId) : [];

  const contarCuentas = (proveedorId) => cuentasApi.rows.filter((c) => c.proveedor_id === proveedorId).length;

  const filtrados = proveedoresUnidad.filter((p) => {
    if (!buscar.trim()) return true;
    const q = buscar.trim().toLowerCase();
    return [p.nombre, p.rfc, p.id_sae].some((v) => (v || "").toLowerCase().includes(q));
  });

  const openNew = () => { setForm(blank); setEditId(null); setNuevaCuenta(cuentaBlank); setModalOpen(true); };
  const startEdit = (p) => { setForm(p); setEditId(p.id); setNuevaCuenta(cuentaBlank); setModalOpen(true); };
  const closeModal = () => { setModalOpen(false); setEditId(null); setForm(blank); setNuevaCuenta(cuentaBlank); };
  const remove = (id) => proveedoresApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));

  const submit = async (e) => {
    e.preventDefault();
    if (!form.nombre.trim()) return;
    const { id, ...rest } = form;
    setSaving(true);
    try {
      if (editId) {
        await proveedoresApi.update(editId, rest);
      } else {
        const nuevo = await proveedoresApi.insert({ ...rest, id: uid(), unidad });
        setEditId(nuevo.id); // dejamos el modal abierto para poder agregar cuentas al que se acaba de crear
        return;
      }
      closeModal();
    } catch (err) {
      alert("No se pudo guardar: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };

  const agregarCuenta = async () => {
    if (!editId) return;
    if (!nuevaCuenta.banco.trim() && !nuevaCuenta.clabe.trim() && !nuevaCuenta.numero_cuenta.trim()) return;
    try {
      await cuentasApi.insert({ id: uid(), proveedor_id: editId, ...nuevaCuenta });
      setNuevaCuenta(cuentaBlank);
    } catch (err) {
      alert("No se pudo agregar la cuenta: " + (err.message || err));
    }
  };
  const eliminarCuenta = (id) => cuentasApi.remove(id).catch((err) => alert("No se pudo eliminar la cuenta: " + (err.message || err)));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
    <Panel
      title={`Catálogo de proveedores — ${unidad}`}
      subtitle={`${proveedoresUnidad.length} registrados — cada compañía tiene el suyo. Un proveedor puede tener varias cuentas bancarias`}
      right={<Button onClick={openNew}>+ Nuevo proveedor</Button>}
    >
      <div style={{ marginBottom: 14 }}>
        <Field label="Buscar">
          <TextInput value={buscar} onChange={(e) => setBuscar(e.target.value)} placeholder="Nombre, RFC, Id SAE…" style={{ width: 280 }} />
        </Field>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={tableStyle}>
          <thead>
            <tr>{["Nombre","RFC","Id SAE","Cuentas",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
          </thead>
          <tbody>
            {filtrados.map((p) => (
              <tr key={p.id}>
                <td style={tdStyle}>{p.nombre}</td>
                <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{p.rfc || "—"}</td>
                <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{p.id_sae || "—"}</td>
                <td style={tdStyle}><Pill tone={contarCuentas(p.id) ? "teal" : "dim"}>{contarCuentas(p.id)} cuenta(s)</Pill></td>
                <td style={tdStyle}>
                  <div style={{ display: "flex", gap: 6 }}>
                    <Button variant="ghost" onClick={() => startEdit(p)}>Editar</Button>
                    <Button variant="danger" onClick={() => remove(p.id)}>Eliminar</Button>
                  </div>
                </td>
              </tr>
            ))}
            {!proveedoresUnidad.length && (
              <tr><td colSpan={5} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin proveedores aún</td></tr>
            )}
            {proveedoresUnidad.length > 0 && !filtrados.length && (
              <tr><td colSpan={5} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ningún proveedor coincide con la búsqueda</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {modalOpen && (
        <Modal title={editId ? "Editar proveedor" : "Nuevo proveedor"} subtitle={`Catálogo de ${unidad}`} onClose={closeModal} width={820}>
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 }}>
            <Field label="Nombre" style={{ gridColumn: "span 3" }}>
              <TextInput value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} />
            </Field>
            <Field label="RFC">
              <TextInput value={form.rfc} onChange={(e) => setForm({ ...form, rfc: e.target.value.toUpperCase() })} />
            </Field>
            <Field label="Id SAE">
              <TextInput value={form.id_sae} onChange={(e) => setForm({ ...form, id_sae: e.target.value })} />
            </Field>
            <div style={{ gridColumn: "span 3", display: "flex", gap: 10, marginTop: 4 }}>
              <Button type="submit" disabled={saving}>{saving ? "Guardando…" : editId ? "Guardar cambios" : "Crear y agregar cuentas"}</Button>
              <Button type="button" variant="ghost" onClick={closeModal}>Cerrar</Button>
            </div>
          </form>

          {editId ? (
            <div style={{ borderTop: `1px solid ${T.borderSoft}`, paddingTop: 16 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: T.text, marginBottom: 10 }}>
                Cuentas bancarias ({cuentasDelProveedor.length})
              </div>
              {cuentasDelProveedor.length > 0 && (
                <table style={{ ...tableStyle, marginBottom: 12 }}>
                  <thead>
                    <tr>{["Banco","CLABE","No. Cuenta","Divisa",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {cuentasDelProveedor.map((c) => (
                      <tr key={c.id}>
                        <td style={tdStyle}>{c.banco || "—"}</td>
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.clabe || "—"}</td>
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.numero_cuenta || "—"}</td>
                        <td style={tdStyle}><Pill>{c.divisa || "MXN"}</Pill></td>
                        <td style={tdStyle}><Button variant="danger" onClick={() => eliminarCuenta(c.id)}>Eliminar</Button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 8, alignItems: "end" }}>
                <Field label="Banco">
                  <TextInput value={nuevaCuenta.banco} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, banco: e.target.value })} />
                </Field>
                <Field label="CLABE">
                  <TextInput value={nuevaCuenta.clabe} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, clabe: e.target.value })} placeholder="18 dígitos" />
                </Field>
                <Field label="No. Cuenta">
                  <TextInput value={nuevaCuenta.numero_cuenta} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, numero_cuenta: e.target.value })} />
                </Field>
                <Field label="Divisa">
                  <Select value={nuevaCuenta.divisa} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, divisa: e.target.value })}>
                    {MONEDAS.map((m) => <option key={m}>{m}</option>)}
                  </Select>
                </Field>
                <Button type="button" onClick={agregarCuenta}>+ Agregar cuenta</Button>
              </div>
            </div>
          ) : (
            <div style={{ fontSize: 11.5, color: T.textFaint, borderTop: `1px solid ${T.borderSoft}`, paddingTop: 16 }}>
              Guarda el proveedor primero (botón "Crear y agregar cuentas") para poder capturarle sus cuentas bancarias.
            </div>
          )}
        </Modal>
      )}
    </Panel>

    <ImportarProveedoresPanel proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />
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
  const proveedoresApi = useCollection("proveedores");
  const cuentasApi = useCollection("proveedor_cuentas");
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
    { id: "reporte", label: "Reporte de Pagos" },
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
        select, input { color-scheme: light; }
        ::selection { background: ${T.accentBg}; }
      `}</style>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 22, flexWrap: "wrap", gap: 14 }}>
        <div>
          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>
            Control de presupuestos · prototipo
          </div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
            <div style={{ fontSize: 24, fontWeight: 700 }}>Panel de gasto</div>
            <details style={{ position: "relative" }}>
              <summary style={{ display: "inline-flex" }}>
                <Pill tone="dim">v{APP_VERSION}</Pill>
              </summary>
              <div style={{
                position: "absolute", top: "calc(100% + 6px)", left: 0, zIndex: 50,
                background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8,
                padding: 12, minWidth: 280, boxShadow: "0 8px 24px rgba(0,0,0,0.4)",
              }}>
                <div style={{ fontSize: 10.5, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>Historial de versiones</div>
                {CHANGELOG.map((c) => (
                  <div key={c.v} style={{ display: "flex", gap: 8, fontSize: 11.5, marginBottom: 5 }}>
                    <span style={{ fontFamily: T.fontMono, color: T.accent, minWidth: 44 }}>v{c.v}</span>
                    <span style={{ color: T.textDim }}>{c.desc}</span>
                  </div>
                ))}
              </div>
            </details>
          </div>
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
                  background: unidad === u ? T.accent : "transparent",
                  color: unidad === u ? "#FFFFFF" : T.textDim,
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
              borderBottom: tab === t.id ? `2px solid ${T.accent}` : "2px solid transparent",
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
          {tab === "transacciones" && <TransaccionesTab unidad={unidad} unidades={unidades} partidas={partidas} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
          {tab === "reporte" && <ReportePagosTab unidad={unidad} partidas={partidas} transacciones={transacciones} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
          {tab === "catalogo" && <CatalogoTab unidad={unidad} unidades={unidades} proyectosApi={proyectosApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
        </>
      )}
    </div>
  );
}
