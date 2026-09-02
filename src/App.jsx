import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Wallet, BarChart3, CheckCircle2, FileEdit, ArrowDownCircle, ArrowUpCircle, Info } from "lucide-react";
import { useCollection } from "./useCollection";
import { supabase } from "./supabaseClient";

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

// NO hay categoría "Diversos", y es deliberado.
//
// En la v1.73 se agregó a los 15 rubros porque 81 partidas ya la usaban:
// el catálogo no ofrecía una opción genérica y la gente tomaba la única
// que existía. Ese diagnóstico era del pasado y era correcto.
//
// Hacia adelante el cálculo cambia. Una opción cómoda que no dice nada se
// vuelve el camino de menor resistencia, y la clasificación se degrada
// sola: "Diversos" parece un dato pero no responde "en qué gastamos",
// que es para lo que existe esta aplicación.
//
// Cuando algo no se puede clasificar, el campo se deja VACÍO. Un hueco se
// ve y pide atención; una etiqueta genérica lo esconde.
const RUBROS_RESPALDO = [
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
  { rubro: "Gastos Financieros e Impuestos", categorias: ["Comisiones bancarias","Derechos","Impuestos","Gastos notariales","Seguros","Fianzas"] },
  { rubro: "Nómina y Personal", categorias: ["Sueldos y salarios","Prestaciones de ley","Prestaciones adicionales","Cuotas IMSS/Infonavit","Impuesto sobre nómina (ISN)","Finiquitos y liquidaciones","Bonos y comisiones"] },
  { rubro: "Otros", categorias: ["Gastos extraordinarios","Donativos"] },
];

const UNIDADES_BASE = ["OSB", "CTM", "ISE"];

// Vehiculos: la flotilla incluye dos companias (IZ2, JEF) que no manejan
// presupuesto propio, por eso no estan en UNIDADES_BASE — viven solo aqui.
const VEH_COMPANIAS = ["OSB", "CTM", "ISE", "IZ2", "JEF"];
const VEH_EST_FUNCIONAL = ["Funcional", "Mtto Requerido", "En Mtto", "Fuera de Servicio", "No Encontrado"];
const VEH_EST_OPERACIONAL = ["En Operación", "Resguardo", "Fuera de Operacion", "No Encontrado"];
const VEH_EST_ADMIN = ["Activa", "Vendida", "No Encontrado", "Baja"];
const VEH_TIPOS = ["Camioneta", "Camion", "Montacargas", "Camper", "Maquinaria", "Remolque"];
const VEH_GPS = ["Positrace", "WIALON"];
const MTTO_ESTATUS = ["Solicitado", "En proceso", "Concluido", "Cancelado"];
const MTTO_TIPOS = ["Preventivo", "Correctivo"];

const MONEDAS = ["MXP", "USD"];
// Respaldo de las zonas. El catálogo real vive en la tabla `zonas` de
// Supabase; esta lista solo se usa si esa tabla todavía no existe o viene
// vacía, para que subir el frontend antes de correr la migración no deje
// el selector de Zona sin opciones.
const ZONAS_RESPALDO = ["Queretaro", "Poza Rica", "Paraiso", "Altamira", "Cerro Azul", "CDMX", "Guaymas", "Torreon", "Rosarito", "Agua Dulce", "Cotaxtla"];

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


/* Opciones de agrupamiento. A nivel de módulo por dos razones: no se
   reconstruyen en cada render, y el saneamiento de las preferencias guardadas
   puede consultarlas sin caer en zona muerta temporal. */
const GROUP_OPCIONES = [
  { value: "anio", label: "Año" },
  { value: "mes", label: "Mes" },
  { value: "rubro", label: "Rubro" },
  { value: "categoria", label: "Categoría" },
  { value: "zona", label: "Zona" },
  { value: "proyecto", label: "Proyecto" },
  ];
const GROUP_OPCIONES_TRANS = [
  { value: "dia", label: "Día de Pago Programado" },
  { value: "zona", label: "Zona" },
  { value: "area", label: "Área" },
  { value: "proveedor", label: "Proveedor" },
  { value: "proyecto", label: "Proyecto (transacción)" },
  { value: "status", label: "Status" },
  { value: "moneda", label: "Moneda" },
  { value: "categoria", label: "Categoría" },
  { value: "_proyecto", label: "Proyecto (partida)" },
  { value: "_rubro", label: "Rubro (partida)" },
  { value: "_mes", label: "Mes (partida)" },
  { value: "_vinculo", label: "Vínculo" },
  ];
const GROUP_OPCIONES_SINVINC = [
  { value: "dia", label: "Día de Pago Programado" },
  { value: "zona", label: "Zona" },
  { value: "area", label: "Área" },
  { value: "proveedor", label: "Proveedor" },
  { value: "proyecto", label: "Proyecto" },
  { value: "status", label: "Status" },
  { value: "moneda", label: "Moneda" },
  ];
const GROUP_OPCIONES_PANEL = [
  { value: "proyecto", label: "Proyecto" },
  { value: "rubro", label: "Rubro" },
  { value: "categoria", label: "Categoría" },
  { value: "zona", label: "Zona" },
  { value: "concepto", label: "Concepto" },
  ];

/* ----------------------------------------------------------------------
   SUGERENCIA DE CATEGORÍA
   ---------------------------------------------------------------------- */
// Palabras que apuntan a una categoría cuando su nombre no aparece tal cual
// en el concepto. "Gasolina" no contiene la palabra "Combustible", pero
// cualquiera sabe que va ahí.
const PISTAS_CATEGORIA = {
  "Combustible": ["gasolina", "diesel", "diésel", "magna", "premium"],
  "Mantenimiento vehicular": ["mantto unidad", "mantto unidades", "mantto vehicular",
    "mantenimiento vehicular", "unidades ligeras", "unidades pesadas", "unidad vehicular"],
  "Llantas": ["llanta", "neumatico", "rin"],
  "Refacciones": ["acumulador", "bateria", "balata", "amortiguador", "clutch", "filtro"],
  "Lubricantes": ["aceite", "lubricante", "grasa"],
  "Energía eléctrica": ["cfe", "luz", "energia", "medidor", "recibo de luz"],
  "Telefonía móvil": ["celular", "telefonia celular", "radiomovil", "movil"],
  "Telefonía fija": ["telefono", "telefonia", "telmex", "linea telefonica"],
  "Internet": ["internet", "enlace", "fibra", "banda ancha"],
  "Agua": ["agua", "pipa", "hidraulico", "potable"],
  "Recolección de residuos": ["residuo", "basura", "desecho", "peligroso"],
  "Vigilancia": ["vigilancia", "seguridad privada", "guardia", "alarma", "camara"],
  "Capacitación": ["capacitacion", "curso", "entrenamiento", "certificacion"],
  "Mensajería": ["mensajeria", "paqueteria", "guia", "envio"],
  "Arrendamientos": ["arrendamiento", "renta", "alquiler", "multifuncional"],
  "Papelería": ["papeleria", "hoja", "toner", "tinta", "impresion"],
  "EPP": ["epp", "bota", "casco", "guante", "chaleco", "lente de seguridad"],
  "Consumibles industriales": ["limpieza", "consumible", "trapo", "escoba"],
  "Sueldos y Salarios": ["nomina", "sueldo", "salario", "finiquito", "aguinaldo"],
  "Impuestos": ["impuesto", "isr", "iva", "predial", "arancel"],
  "Estudios": ["estudio", "analisis", "laboratorio", "dictamen", "peritaje"],
  "Mantenimiento de instalaciones": ["mantto electrico", "instalacion", "pintura", "impermeabiliza"],
  "Herramienta especializada": ["herramienta", "extintor", "respiracion autonoma"],
  "Equipos de cómputo": ["computo", "laptop", "computadora", "impresora", "monitor"],
  "Productos químicos de operación": ["quimico", "xileno", "emulsotron", "hipoclorito", "demulsificante"],
  "Donativos": ["donativo", "donacion", "aportacion"],
};

// Índice normalizado de las pistas, construido una sola vez.
const PISTAS_NORM = {};

const normSug = (v) => String(v || "")
  .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
  .replace(/[.,;:()]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();

Object.entries(PISTAS_CATEGORIA).forEach(([cat, claves]) => {
  PISTAS_NORM[normSug(cat)] = claves;
});

/**
 * Propone una categoría a partir del concepto, DENTRO del rubro elegido.
 *
 * Puntúa cada categoría disponible: primero por sus palabras clave, luego
 * por las palabras que comparte con el concepto. Se limita al rubro para
 * no pelear con la decisión que ya tomó quien captura — el rubro se elige,
 * la categoría se sugiere.
 *
 * Devuelve null cuando ninguna alcanza: una sugerencia mala cuesta más que
 * ninguna, porque se acepta sin pensar.
 */
function sugerirCategoria(concepto, categorias) {
  const t = normSug(concepto);
  if (!t || !categorias || !categorias.length) return null;
  const relleno = new Set(["de", "del", "la", "el", "los", "las", "y", "a", "en", "para",
    "por", "con", "servicio", "servicios", "pago", "pagos", "gastos", "diversos"]);
  const palabras = new Set(t.split(" ").filter((w) => w.length > 3 && !relleno.has(w)));

  let mejor = null, mejorPts = 0;
  categorias.forEach((cat) => {
    if (normSug(cat) === "diversos") return; // el comodín nunca se sugiere
    let pts = 0;
    // La búsqueda de pistas va NORMALIZADA, no por llave exacta: la categoría
    // real es "Sueldos y salarios" y la pista decía "Sueldos y Salarios", así
    // que con búsqueda exacta la pista se perdía en silencio y ganaba
    // "Impuesto sobre nómina" por compartir una palabra.
    (PISTAS_NORM[normSug(cat)] || []).forEach((k) => { if (t.includes(normSug(k))) pts += 60; });
    normSug(cat).split(" ").forEach((w) => {
      if (w.length > 3 && !relleno.has(w) && palabras.has(w)) pts += 40;
    });
    if (pts > mejorPts) { mejorPts = pts; mejor = cat; }
  });
  return mejorPts >= 40 ? mejor : null;
}

// Todas las categorías del catálogo, sin repetir. La transacción puede usar
// cualquiera: describe QUÉ se compró, con independencia del rubro donde se
// presupuestó.
/* El catálogo vive en Supabase, pero los parsers de importación son
   funciones de módulo: no pueden leer estado de React. Por eso RUBROS es una
   variable de módulo que se refresca cuando llegan los datos, con la lista
   del código como respaldo mientras cargan o si la migración no se ha
   corrido. Enhebrar el catálogo por los 18 puntos de uso —varios fuera de
   componentes— habría sido mucho más invasivo por el mismo resultado. */
let RUBROS = RUBROS_RESPALDO;
let CATEGORIAS_TODAS = [];

function recalcularCategoriasTodas() {
  CATEGORIAS_TODAS = [...new Set(RUBROS.flatMap((r) => r.categorias))].sort();
}
recalcularCategoriasTodas();

/** Reconstruye el catálogo desde las filas de Supabase. */
function aplicarCatalogo(rubrosRows, categoriasRows) {
  if (!rubrosRows || !rubrosRows.length) {
    RUBROS = RUBROS_RESPALDO;
  } else {
    const porRubro = new Map();
    categoriasRows.forEach((c) => {
      if (c.activa === false) return;
      if (!porRubro.has(c.rubro_id)) porRubro.set(c.rubro_id, []);
      porRubro.get(c.rubro_id).push(c);
    });
    RUBROS = [...rubrosRows]
      .filter((r) => r.activo !== false)
      .sort((a, b) => (a.orden ?? 999) - (b.orden ?? 999) || String(a.nombre).localeCompare(String(b.nombre)))
      .map((r) => ({
        rubro: r.nombre,
        categorias: (porRubro.get(r.id) || [])
          .sort((a, b) => (a.orden ?? 999) - (b.orden ?? 999) || String(a.nombre).localeCompare(String(b.nombre)))
          .map((c) => c.nombre),
      }));
  }
  recalcularCategoriasTodas();
}

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const uid = () => (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2, 10));

// ----------------------------------------------------------------------
// VERSIÓN — súbela cada vez que cambies este archivo. Formato MAJOR.MINOR.PATCH:
// MINOR = feature nueva, PATCH = fix/ajuste menor. Se muestra en el header de
// la app y debe ir en el nombre del archivo que se comparte (App-v1.5.0.jsx).
// ----------------------------------------------------------------------
const APP_VERSION = "1.89.0";
const CHANGELOG = [
  { v: "1.89.0", desc: "PDF de la Solicitud de Pago: la compañía pasa a ser prefijo del folio —CTM-12— en lugar de ir suelta bajo el título, donde además lo tapaba. Se retiran las firmas y el pie de generación: el documento se firma en el sistema, no en papel, y esas líneas ocupaban un cuarto de la hoja sin usarse. El nombre del archivo sigue el mismo formato, en PDF y en Excel" },
  { v: "1.88.1", desc: "Fix: generar la Solicitud de Pago fallaba con 'catch is not a function'. El constructor de consultas de Supabase no es una promesa hasta que se le hace await, así que encadenarle .catch() rompe. Los dos pasos secundarios —adelantar el consecutivo y guardar el desglose en la transacción— van ahora en su propio try: la solicitud ya quedó registrada y no deben impedir que se descargue el documento" },
  { v: "1.88.0", desc: "Catálogo se organiza en sub-pestañas —Proyectos, Rubros y categorías, Zonas, Proveedores y Solicitudes de Pago— en vez de seis paneles apilados que obligaban a recorrer toda la página. Se prefirió esto a secciones plegables: con seis encabezados que atravesar el recorrido sigue existiendo, y el estado de plegado se olvida entre visitas. De paso se retira el panel de referencia de rubros, que solo listaba lo que el panel administrable ya muestra" },
  { v: "1.87.0", desc: "La Solicitud de Pago se genera ahora en PDF con formato de documento formal —encabezado con folio, datos en dos columnas, desglose fiscal, la cifra a pagar destacada, bloque de datos bancarios y firmas— además del Excel, que se conserva. Y en Catálogo aparece el consecutivo editable de la SPP por compañía, junto con el responsable de proyecto y el lugar de adquisición, que hasta ahora se escribían a mano en cada solicitud. El folio asignado es el MAYOR entre lo configurado y lo ya emitido: bajar el consecutivo por error no debe reutilizar folios de solicitudes que ya salieron. Requiere 20-consecutivo-spp.sql" },
  { v: "1.86.2", desc: "Fix: el botón de Generar Solicitud de Pago no hacía nada. El botón vive en Transacciones pero su estado y su modal habían quedado en Partidas, así que el clic llamaba a un setter de otro componente y el error moría en la consola sin señal visible. Es el mismo tipo de error que la vista previa del PDF en la 1.56.0; el verificador de alcance por árbol sintáctico lo detecta y ahora se corre también sobre los componentes, no solo sobre las variables" },
  { v: "1.86.1", desc: "Fix: al elegir un proveedor del catálogo maestro no pasaba nada visible. Los datos sí se cargaban, pero en el formulario de alta, que solo se dibuja cuando hay algo en edición — y eso no se activaba, así que el clic llenaba una pantalla invisible. Ahora abre el formulario con los datos puestos, cada resultado tiene un botón 'Usar' para que se vea pulsable, y el aviso sobre datos bancarios faltantes aparece DENTRO del formulario en vez de en una alerta que hay que cerrar antes de ver los datos que describe" },
  { v: "1.86.0", desc: "Solicitud de Pago a Proveedores: las transacciones marcadas como ANTICIPO en Folio SAE muestran un botón que genera la SPP en Excel, con folio autonumérico por compañía y registro en base. Todo lo que la app sabe viene precargado; el desglose fiscal —subtotal, IVA, retenciones— se DEDUCE del importe pagado y se presenta para confirmar, no como un hecho: ese camino inverso solo es exacto si el importe correspondía al esquema elegido, y las retenciones dependen del régimen del proveedor. Hay cinco esquemas predefinidos más captura manual. Los proyectos ganan Centro de Costo en el catálogo, que la solicitud necesita. Requiere 19-spp-anticipos.sql" },
  { v: "1.85.0", desc: "El selector de proveedor puede consultar el catálogo maestro de ASPEL (11,057 proveedores). Se busca contra Supabase EN EL MOMENTO y nunca se carga en memoria: traer once mil registros en cada sesión degradaría la app para todos a cambio de un catálogo que casi nunca se consulta. Al elegir uno, sus datos llenan el formulario de alta pero NO se guarda solo — el maestro trae CLABE en apenas el 30% de los casos y divisa casi en ninguno, así que darlo de alta en silencio crearía proveedores incompletos que fallan al pagar. La lista avisa si trae o no datos bancarios, y si el registro venía marcado para revisión. Requiere 18-proveedores-maestro.sql y el CSV cargado" },
  { v: "1.84.0", desc: "Proveedores: botón 'Descargar plantilla' para la carga masiva, que faltaba — había importador pero no había de dónde sacar el formato, y la única forma de conocerlo era exportar el catálogo, que no sirve cuando aún no hay proveedores. Trae las 12 columnas que el importador detecta, una hoja por compañía con su fila de ejemplo, y una hoja de instrucciones que advierte lo de la CLABE: Excel convierte 18 dígitos a notación científica y corrompe el dato antes de que nadie lo note" },
  { v: "1.83.0", desc: "Formato consistente en los nueve archivos de Excel, con una distinción deliberada: los REPORTES (Presupuestal, Pagos, Pagos Dirección) llevan título, periodo, totales y subtotales; los archivos de INTERCAMBIO (RawData de partidas, catálogo de proveedores y las dos plantillas de importación) solo reciben encabezado con fondo, panel congelado y autofiltro. No llevan título a propósito: los parsers y el preparador leen el encabezado de la PRIMERA fila, así que anteponerlo rompería la reimportación. El Reporte de Pagos gana además un subtotal por zona y moneda, que es lo que Pagos necesita para cuadrar contra el banco, e importes alineados a la derecha en todos los archivos" },
  { v: "1.82.0", desc: "El Excel del Reporte Pagos Dirección adopta el formato del Reporte Presupuestal: título y subtítulo con periodo y totales, encabezado con fondo y texto blanco, filas de total por moneda al pie, encabezado congelado y autofiltro. Los totales dejan de vivir sueltos en celdas fijas —se rompían al cambiar las columnas visibles— y pasan al subtítulo, junto al periodo que los explica. El archivo se llama ahora 'Reporte de Pagos Dirección - Compañía - Periodo', con el periodo tomado de las fechas que realmente contiene" },
  { v: "1.81.0", desc: "El Reporte Excel de Partidas se llama ahora 'Reporte Presupuestal - Compañía - Mes', y el mes sale de las partidas que realmente contiene, no de los filtros: si el filtro es amplio pero solo hay septiembre, el archivo dice septiembre. Con varios meses usa el rango en orden cronológico. El título dentro del documento y el nombre de la hoja siguen el mismo criterio" },
  { v: "1.80.1", desc: "Reporte Excel de Partidas: los subtotales de cada grupo y el encabezado etiquetan ahora las DOS monedas. El formateador general solo marca los dólares —da por hecho que sin etiqueta son pesos— y en una fila donde conviven las dos, '$211,039.96 · $25,000.00 USD' invitaba a leer la primera cifra como parte del mismo total" },
  { v: "1.80.0", desc: "Partidas: nuevo 'Reporte Excel' con su propio selector de columnas —independiente del de pantalla— que respeta el agrupamiento de la vista: cada grupo abre con su nombre, cuántas partidas contiene y su subtotal por moneda, con sangría por nivel. Se diferencia de 'Exportar', que sigue dando el listado plano para el preparador de transacciones. Los totales van separados por moneda, nunca sumadas entre sí" },
  { v: "1.79.0", desc: "Se elimina la categoría 'Diversos' de los 15 rubros y no se puede volver a crear —tampoco Varios, Otros o General. En la v1.73 se agregó porque 81 partidas ya la usaban, y ese diagnóstico del pasado era correcto; hacia adelante el cálculo se invierte: una opción cómoda que no dice nada se vuelve el camino de menor resistencia y la clasificación se degrada sola. Cuando algo no se puede clasificar el campo se deja VACÍO, no con una etiqueta genérica: un hueco se ve y pide corrección, una etiqueta lo esconde. Los rubros nuevos nacen sin categorías, para que se definan las que de verdad hacen falta" },
  { v: "1.78.0", desc: "Los rubros y categorías salen del código y pasan a ser administrables desde Catálogo: se agregan, renombran, desactivan y eliminan sin volver a desplegar. El punto delicado es renombrar — partidas.rubro, partidas.categoria y transacciones.categoria guardan el TEXTO, no un id, así que cambiar el nombre dejaría a los registros existentes apuntando a algo inexistente; por eso el cambio se propaga a los registros en uso, avisando cuántos son. Un rubro o categoría en uso no se puede eliminar, solo desactivar: sale del selector y el histórico se conserva. Todo rubro nuevo nace con la categoría Diversos. Requiere 16-catalogo-editable.sql" },
  { v: "1.77.0", desc: "Las transacciones estrenan campo Categoría: la partida dice en qué rubro se presupuestó, la transacción dice qué se compró. Está en el formulario —con las categorías del rubro de su partida arriba y el resto abajo, más la sugerencia por concepto de un clic—, en la tabla, en Agrupar por y en el buscador. La carga masiva la asigna sola cuando el archivo no la trae, deduciéndola del concepto; sin eso el campo entraría vacío en cada importación y el dato se degradaría desde el primer archivo. No se restringe al rubro de la partida a propósito: con datos reales, 39 de 228 transacciones tenían la categoría correcta para el gasto y una partida de otro rubro, y esa discrepancia es información —dice qué partidas absorben gasto que no les toca— no un error. Requiere 15-categoria-en-transacciones.sql" },
  { v: "1.76.0", desc: "Nueva categoría 'Seguros' en Gastos Financieros e Impuestos, junto a Fianzas: son el mismo tipo de gasto —transferencia de riesgo contratada— y separarlos partiría en dos algo que se lee mejor junto. La categoría 'Seguros' de Vehículos se conserva y queda reservada a pólizas vehiculares, que son parte del costo de operar la flotilla; el resto de las pólizas va a Gastos Financieros. Sin esa regla el gasto en seguros quedaría partido según quién capture" },
  { v: "1.75.0", desc: "Dos cosas sobre rubros y categorías. Al capturar una partida, la app sugiere una categoría a partir del concepto —dentro del rubro elegido— y basta un clic para usarla; se sugiere, no se impone, porque quien captura sabe cosas que el concepto no dice. Y la carga masiva ya no deja entrar combinaciones inválidas: en vez de solo avisar, las repara conservando SIEMPRE el rubro —que es el eje que se lee en el Dashboard— y ajustando la categoría, con el valor original guardado en `extra` para no perderlo. Un rubro inexistente cuya categoría solo vive en un lugar se corrige a ese lugar; si no hay forma de saberlo, cae en Otros / Diversos, que es honesto" },
  { v: "1.74.0", desc: "Las exportaciones a Excel de Partidas y del Reporte Pagos Dirección respetan ahora las columnas visibles, igual que ya lo hacían el PDF y el Reporte de Pagos. En Partidas se conservan siempre Unidad, Usado y Disponible: no existen como columna en pantalla pero sí hacen falta en la hoja. De paso, el reporte de Dirección deja de llevar encabezados, valores y anchos en tres listas paralelas —el formato de moneda se aplicaba por índice fijo, así que mover la columna Importe dejaba el signo de pesos en la de al lado" },
  { v: "1.73.0", desc: "'Diversos' pasa a ser categoría válida en los 15 rubros. El diagnóstico de clasificación encontró 81 partidas usándola bajo nueve rubros distintos: cuando el mismo error aparece en nueve rubros no son nueve equivocaciones, es que el catálogo no ofrecía una categoría genérica y se tomaba la única que existía, la de Otros. Con esto dejan de ser incoherencias sin tocar un solo dato, y el análisis por rubro se conserva. Además la carga masiva de Partidas ahora valida rubro y categoría contra el catálogo: no bloquea, pero marca las filas en el preview y dice a qué rubro sí pertenece esa categoría. Esa validación es la que faltaba — el formulario siempre encadenó rubro y categoría, pero el importador aceptaba cualquier combinación, y por ahí entraron 36 partidas con el marcador de proyecto 'Todos' copiado en ambos campos" },
  { v: "1.72.0", desc: "El Reporte de Presupuesto Mensual gana un apartado de flotilla: combustible y mantenimiento vehicular juntos, desglosados por zona y con su subtotal. Se agregó porque esos dos gastos viven en rubros distintos —Vehículos y Servicios de Mantenimiento— así que el corte por rubro los separa justo cuando interesa verlos como uno solo; entre los dos suelen ser cerca de la mitad del presupuesto. La clasificación es por palabras clave sobre concepto y categoría, y el reporte indica cuántas partidas cayeron en cada bolsa para poder detectar lo que quedó fuera" },
  { v: "1.71.0", desc: "El Reporte de Presupuesto Mensual pasa de listado a resumen ejecutivo: totales por moneda arriba, y tres cortes —por rubro, por proyecto y por zona— con importe, número de partidas y porcentaje sobre el total de su moneda. Responde cuánto en pesos, cuánto en dólares, en qué se gasta y dónde se reparte, sin obligar a quien lo recibe a sacar el resumen de ochenta renglones. El corte por zona se omite si ninguna partida tiene una asignada, y la hoja pasa a vertical porque ya no hay renglones que requieran el ancho" },
  { v: "1.70.0", desc: "Partidas: nuevo botón 'Reporte de Presupuesto Mensual' que genera un PDF con lo que está a la vista — las columnas visibles en su orden actual, las filas ya filtradas y ordenadas, y el agrupamiento vigente convertido en encabezados de sección con su subtotal" },
  { v: "1.69.0", desc: "Reporte Pagos Dirección: el PDF ahora lleva únicamente las columnas visibles, en el orden en que están en pantalla. Antes tenía sus diez columnas fijas en el código, así que ocultar o reordenar en la tabla no cambiaba nada del documento y lo que se revisaba no era lo que se enviaba. La confirmación indica cuántas columnas llevará y avisa si la tabla se pasa del ancho de la hoja" },
  { v: "1.68.0", desc: "El selector de partida gana filtros por rubro y por mes, un contador de cuántas se muestran, y el buscador ahora cubre también categoría y zona. Además explica por qué parecen faltar: desde la v1.60 solo se ofrecen las de la moneda del gasto, y eso no se decía en ninguna parte — ahora avisa cuántas se ocultan y que basta cambiar la moneda del formulario. De paso se cierra un hueco: los selectores de la tabla (vincular o cambiar la partida de una transacción existente) no filtraban por moneda, así que por ahí se podía saltar el bloqueo que el formulario sí aplicaba" },
  { v: "1.67.0", desc: "Transacciones: se retira el botón 'Borrar todas', que eliminaba la compañía entera con una sola confirmación. En su lugar, la barra de selección que ya existía gana 'Eliminar seleccionadas', así que hay que elegir explícitamente qué se borra. Sigue conservando las marcadas como Pagadas y avisando cuántas, y la confirmación indica el importe total de lo que se va a eliminar" },
  { v: "1.66.0", desc: "Una transacción marcada como Pagada ya no se puede borrar: hay que cambiarle el status a No Pagado primero, para que quede claro que se está deshaciendo un pago registrado y no solo limpiando un renglón. La regla cubre los tres caminos —el borrado de una fila y los dos masivos—; en los masivos las pagadas se conservan y se avisa cuántas, en vez de abortar toda la operación. Y se hereda a Partidas: una partida con transacciones pagadas no se elimina, porque dejarlas sin vincular haría desaparecer ese gasto del presupuesto por la puerta de atrás. Si solo tiene transacciones sin pagar, se permite pero avisando que quedarán sin partida" },
  { v: "1.65.0", desc: "Los agrupamientos y el orden de las tablas ahora se recuerdan entre sesiones, no solo mientras la pestaña sigue abierta. Los FILTROS siguen viviendo en la sesión a propósito: abrir la app mostrando el periodo que quedó de la última vez haría leer esa cifra como si fuera la actual, y eso en un tablero financiero es un error caro. Lo guardado se limpia al leerse, así que un agrupamiento u orden que apunte a un campo retirado se descarta en vez de dejar la tabla en un estado imposible" },
  { v: "1.64.1", desc: "Fix: en el panel comparativo, el último nivel del agrupamiento no se podía desplegar y sus registros quedaban inalcanzables. No se notaba mientras 'Concepto' fuera el último nivel —ahí cada hoja ya era un registro— pero al agrupar solo por Rubro, sus partidas no había forma de verlas. Ahora esa fila también abre, muestra cuántos registros contiene y los lista debajo" },
  { v: "1.64.0", desc: "Dashboard: el panel comparativo estrena control de Agrupar por, con las mismas opciones que las otras tablas (Proyecto, Rubro, Categoría, Zona, Concepto) y hasta tres niveles. La moneda queda fija como primer nivel — mezclarla dentro de un grupo haría que sus subtotales sumaran pesos con dólares. El agrupamiento gobierna las dos tablas a la vez, porque comparten el eje de columnas para leerse una contra otra y eso solo funciona si sus renglones coinciden. La primera tabla pasa a llamarse simplemente 'Presupuesto', y su encabezado ahora refleja el agrupamiento vigente en lugar de decir siempre 'Proyecto'" },
  { v: "1.63.0", desc: "Partidas: se retira la marca de Recurrente y el botón que generaba automáticamente los meses faltantes hasta diciembre. Estaba creando partidas no deseadas —duplicaba series al copiar una partida, y volvía a generar lo que ya se había ajustado o borrado a mano— y el costo superaba lo que ahorraba. Duplicar una partida sigue siendo la vía para repetir un concepto en otro mes, con control sobre cada copia. La columna es_recurrente se queda en la base sin usarse, así que ningún dato se pierde" },
  { v: "1.62.1", desc: "La plantilla de carga masiva de Partidas gana las columnas Zona y Año. El importador ya las leía desde antes, pero la plantilla no las ofrecía: nadie sabía que existían y todo entraba sin zona y con el año en curso. Las instrucciones explican que Zona es opcional —vacía significa cualquier zona— y cuándo conviene fijar el Año" },
  { v: "1.62.0", desc: "Partidas: nuevo campo Zona, alimentado por el catálogo de zonas. Es opcional — vacío significa 'cualquier zona', para que Nómina o los servicios legales no tengan que fingir una. Disponible en el formulario, la tabla, Agrupar por, el buscador, la carga masiva y la exportación; también entra en la detección de duplicados al importar, de modo que dos partidas iguales de zonas distintas dejan de contarse como la misma. Requiere correr 08-zona-en-partidas.sql" },
  { v: "1.61.1", desc: "Fix: el Dashboard se quedaba en blanco un instante después de cargar. Al unificar los filtros, el bloque que arma los controles quedó por encima de la declaración del filtro de Proyecto que usa; como el JSX se construye al evaluarse ese const, el render fallaba por zona muerta temporal en cuanto llegaban los datos" },
  { v: "1.61.0", desc: "Dashboard: un solo juego de filtros para toda la pestaña. Los de 'Presupuestado vs. pagado real' ahora gobiernan también las gráficas de abajo, y 'Presupuesto vs. ejecutado por proyecto' pierde los suyos. Antes había dos juegos en la misma pantalla y no se veía cuál mandaba sobre qué; de paso, las gráficas ganan el rango Desde-Hasta y YTD, que solo tenía el panel de arriba" },
  { v: "1.60.0", desc: "Ya no se puede vincular una transacción a una partida de otra moneda. En el formulario el selector de partida solo ofrece las de la moneda del gasto, y al cambiar la moneda se suelta la partida si deja de cuadrar; el guardado lo verifica de todos modos. La carga masiva —por donde entran casi todas las transacciones— también rechaza esos vínculos: la fila se importa sin partida y el preview dice cuántas fueron y por qué. Comparar un gasto contra un presupuesto en otra moneda no significa nada, y con prorrateo un solo movimiento mal capturado pinta un bloque entero de esa moneda en el Dashboard" },
  { v: "1.59.1", desc: "Dashboard: 'Presupuesto vs. ejecutado por proyecto' gana una tercera barra. La que decía 'Ejecutado' suma todas las transacciones vinculadas, pagadas o no, y se leía como si ya se hubieran pagado; ahora se llama 'Comprometido' y a su lado va 'Pagado', que solo cuenta las de status Pagado. Además, al guardar una transacción cuya moneda no coincide con la de su partida se pide confirmación: no se bloquea, pero un solo gasto mal capturado puede pintar un bloque entero de dólares en el Dashboard si la partida trae prorrateo" },
  { v: "1.59.0", desc: "Dashboard: se retira el panel 'Resumen financiero'. Sus filtros de Proyecto, Mes y Año se mudan al panel 'Presupuesto vs. ejecutado por proyecto', que es lo que gobiernan junto con la tendencia mensual — si se hubieran ido con el panel, las gráficas habrían quedado sin forma de filtrarse" },
  { v: "1.58.1", desc: "Fix: en el panel comparativo, las filas 'Total MXP' y 'Total USD' del final duplicaban cifras que la propia fila de cada moneda ya mostraba, y alejaban el total de los datos que resume. Se retiran: el total de cada moneda queda justo encima de su lista. Además las cifras en dólares ya se formatean como USD dentro de la tabla, en vez de imprimirse igual que los pesos" },
  { v: "1.58.0", desc: "Dashboard: el panel comparativo muestra pesos y dólares en la MISMA tabla, con la moneda como primer nivel de agrupación arriba del proyecto, y desaparece el selector que obligaba a alternar. Cada moneda tiene su propia fila de total; no hay total general único porque sumarlas exigiría fijar un tipo de cambio. Se prefirió agrupar por moneda antes que duplicar las columnas por mes, que habría llevado el panel de 9 a 17 columnas de dinero" },
  { v: "1.57.0", desc: "Catálogo de zonas editable: hasta ahora eran una lista fija en el código y solo se podían ver. Ahora se dan de alta, renombran, desactivan o eliminan desde Catálogo, y alimentan el selector de Zona al capturar. El panel avisa además de las zonas que aparecen en transacciones sin estar dadas de alta —llegan por carga masiva— y permite agregarlas de un clic. Requiere correr la migración 06-catalogo-zonas.sql" },
  { v: "1.56.0", desc: "Reporte de Pagos: botón 'Vista previa PDF' que abre el documento REAL en un visor antes de generarlo — sin descargarlo y sin marcar nada como enviado a Pagos. Avisa cuando la tabla se pasa del ancho de la hoja, que es lo que ocurre al activar muchas columnas: autoTable no falla, apreta las columnas y parte el texto sin decir nada. La exportación a Excel muestra ahora un resumen (filas, columnas incluidas y excluidas, y las primeras filas) antes de descargar" },
  { v: "1.55.2", desc: "Fix: el selector 'Columnas del Excel' mostraba las casillas sin nombre. El panel leía únicamente la propiedad `label`, pero las columnas de exportación se definen con `header`; ahora acepta cualquiera de las dos" },
  { v: "1.55.1", desc: "Reporte de Pagos: la exportación a Excel estrena su propio selector 'Columnas del Excel', independiente del de pantalla y del PDF. Antes bajaba siempre las 19 columnas. Los tres selectores comparten ahora la misma implementación, así que se comportan igual" },
  { v: "1.55.0", desc: "Reporte de Pagos: nuevo botón 'Columnas del PDF' para elegir qué lleva la solicitud que se manda a Pagos. Antes eran siete columnas fijas en el código; ahora hay dieciséis disponibles —entre ellas No. Cuenta, SWIFT, Referencia de Pago, Área y Folio Factura— y se recuerda la selección. Día, Solicitante, Proveedor, Concepto e Importe no se pueden quitar porque el área de Pagos las necesita siempre" },
  { v: "1.54.0", desc: "Carga masiva de transacciones: ahora resuelve sola la cuenta bancaria. Si el proveedor empatado tiene una única cuenta en la divisa del pago, se asigna; con varias se deja vacía a propósito, porque elegir la equivocada manda el dinero a otro lado. Antes toda transacción importada quedaba sin cuenta y había que elegirla a mano. Además, Proveedores estrena 'Exportar las 3', con una fila por cuenta bancaria (CLABE y número forzados a texto) para poder cotejarlas fuera de la app" },
  { v: "1.53.1", desc: "Partidas: se agrega 'Exportar las 3', que baja OSB, CTM e ISE en hojas separadas de un mismo archivo respetando los filtros de mes y año. El correo de pagos llega con las tres compañías mezcladas, así que un export de una sola no alcanzaba para proponer las asignaciones. Los filtros de rubro y proyecto solo se aplican a la compañía activa, porque esos valores no siempre existen en las otras dos" },
  { v: "1.53.0", desc: "Partidas: botón 'Exportar a Excel'. Descarga las partidas que estén a la vista (respeta los filtros) con Unidad, Folio, Mes, Año, Concepto, Rubro, Categoría, Proyecto, Moneda, Monto, Usado, Disponible, SMI y Recurrente. Antes lo único descargable desde esta pestaña era la plantilla vacía" },
  { v: "1.52.1", desc: "Partidas: se rediseña el bloque de transacciones vinculadas, que se confundía con las bandas de la propia tabla — su encabezado usaba exactamente el mismo estilo que los encabezados de columna. Ahora es una tarjeta blanca embutida, con un riel de color que la ata a su partida y una línea de resumen que dice cuántas transacciones son y cuánto suman por moneda" },
  { v: "1.52.0", desc: "Partidas: botón 'Duplicar' en cada fila — abre el formulario prellenado con los datos de la original para revisar y guardar como partida nueva. El folio se deja vacío para que se genere uno propio, y la marca de recurrente no se hereda (si se heredara, el duplicado generaría una segunda serie en todos los meses restantes del año). Los botones de fila pasan a íconos compactos (✎/⧉/✕) con tooltip, igual que en Transacciones" },
  { v: "1.51.2", desc: "Reporte de Pagos: se retira la columna 'Referencia (Proveedor)', que traía un dato del catálogo de proveedores y no del pago. Sale de la tabla y de la exportación a Excel. El campo sigue existiendo en el catálogo de Proveedores" },
  { v: "1.51.1", desc: "Fix: la exportación a Excel del Reporte de Pagos salía corrida — los encabezados y los anchos seguían siendo 19 columnas mientras las filas ya traían 20, y el formato de moneda caía sobre SWIFT en lugar de Importe. El layout de la exportación (encabezado, ancho y valor) queda unificado en un solo arreglo para que no pueda volver a desalinearse. Además, las columnas nuevas ahora aparecen en su posición natural dentro de la tabla en vez de irse hasta el extremo derecho, que era lo que escondía 'Referencia de Pago' si ya tenías un orden de columnas guardado" },
  { v: "1.51.0", desc: "Transacciones: nuevo campo 'Referencia de Pago' (folio SPEI, número de cheque, etc.) disponible en el alta, la edición rápida, la carga masiva y como columna del Reporte de Pagos, además de entrar en ambos buscadores. La columna 'Referencia' del reporte pasa a llamarse 'Referencia (Proveedor)' para distinguirla, ya que ese dato viene del catálogo de proveedores y no del pago. Requiere correr la migración SQL 03-referencia-pago.sql" },
  { v: "1.50.0", desc: "Dashboard: el panel comparativo estrena filtrado en cascada de tres niveles — Año (YTD o un año con datos), Periodo (todo el año o Desde-Hasta), y las listas de mes, que solo aparecen cuando eligen rango. YTD va del inicio del año en curso al mes actual. Reemplaza el selector de periodos y el botón YTD de las versiones anteriores" },
  { v: "1.49.0", desc: "Dashboard: botón YTD en el panel comparativo — fija el rango del inicio del año en curso al mes actual (a diferencia del rango por default, que corta en el mes cerrado anterior); se ilumina cuando el rango seleccionado es exactamente ese, y se deshabilita si no hay partidas del año en curso. Además corrige el componente de botón, que descartaba en silencio los estilos que recibía: los botones de acceso y el de Cerrar de los menús desplegables recuperan su ancho completo, y varios márgenes vuelven a aplicarse" },
  { v: "1.48.0", desc: "Dashboard: el panel comparativo cambia sus filtros de Mes/Año a un rango Desde/Hasta por periodo completo (ej. de 'Enero 2026' a 'Julio 2026'), lo que permite rangos que cruzan el cierre de año — imposible de expresar antes. Arranca por default en el último mes cerrado (el inmediato previo al actual). Estos filtros ahora son propios del panel y ya no se comparten con el resto del Dashboard" },
  { v: "1.47.2", desc: "Dashboard: el botón 'Contraer todo' del panel comparativo pasa a la barra de filtros (a la izquierda de Proyecto) y ahora gobierna las dos tablas a la vez en vez de una cada uno. Además alterna: si ya está todo contraído, el botón dice 'Expandir todo'" },
  { v: "1.47.1", desc: "Fix: en el panel 'Presupuestado vs. pagado real', la tabla de Pagado real solo mostraba columnas de los meses que ya tenían algún pago — si filtrabas Agosto/Septiembre/Octubre y solo había pagos en Agosto, las otras dos columnas desaparecían y las dos tablas dejaban de ser comparables. Ahora ambas comparten el mismo eje de meses (el mes filtrado aparece en las dos, en ceros si no hay pagos) y la misma regla de etiquetado de año" },
  { v: "1.47.0", desc: "Nueva pestaña Vehículos con tres vistas: Dashboard (disponibilidad de la flotilla, estatus por compañía, composición por tipo, distribución por ubicación y unidades con más SMI abiertas), Flotilla (tabla completa con búsqueda, columnas configurables y las SMI de cada unidad al expandir el renglón) y Mantenimientos (captura de folio, fecha, taller y costo). Incluye las compañías IZ2 y JEF, que solo existen en la flotilla, y avisa de posibles duplicados por VIN, placas o número de motor" },
  { v: "1.46.38", desc: "Reporte de Pagos: Forma de Pago y Método de Pago muestran solo el código (ej. '03', 'PPD') en vez del texto completo, en la tabla y en el Excel exportado — pasa el mouse encima para ver el nombre completo" },
  { v: "1.46.37", desc: "Reporte de Pagos (PDF): los bloques de Zona/Moneda ya no fuerzan una página nueva cada uno — fluyen en la misma página mientras quepan, y solo saltan de página cuando de verdad no hay espacio" },
  { v: "1.46.36", desc: "Separa 'Reportado' en dos marcas independientes: 'Reportado a Dirección' (ya se avisó que el gasto se va a hacer) y 'Enviado a Pagos' (ya se mandó a ejecutar) — cada una con su propia columna, filtro, y selección en lote en Transacciones; los botones PDF de Reporte de Pagos y Reporte Pagos Dirección ahora marcan cada uno la suya" },
  { v: "1.46.35", desc: "Reporte de Pagos: botón 'Generar reporte (PDF)' — igual que en Reporte Pagos Dirección, agrupa por Zona y Moneda (mismo criterio que el Excel) y marca las transacciones incluidas como 'Reportadas', con confirmación previa" },
  { v: "1.46.34", desc: "Etiquetas de campo en las filas de agrupamiento (ej. 'DÍA DE PAGO PROGRAMADO', 'ZONA') más grandes, en negritas y con más contraste — ya no se pierden junto al valor" },
  { v: "1.46.33", desc: "Rediseña las filas de agrupamiento: cada nivel baja de intensidad (fondo, tamaño y grosor de texto) mientras más profundo, para distinguir mejor la jerarquía de un vistazo. También corrige la fecha de 'Reportado' que se encimaba con la columna de Folio (formato más corto)" },
  { v: "1.46.32", desc: "Partidas: nuevo checkbox 'Recurrente' — con el botón 'Generar recurrentes pendientes', crea de un jalón los meses que falten hasta diciembre para cada partida marcada, con el mismo monto (🔁 en la tabla indica cuáles ya están marcadas)" },
  { v: "1.46.31", desc: "Transacciones: agrega 'Día de Pago Programado' como opción de Agrupar por (tabla principal y sin vincular)" },
  { v: "1.46.30", desc: "Agrega columna '#' con numeración dinámica en Partidas y Transacciones — cuenta las filas realmente visibles en ese momento (respeta filtros, orden, y sigue contando sin reiniciar aunque tengas Agrupar por activo)" },
  { v: "1.46.29", desc: "Fix crítico: la carga inicial de datos (partidas, transacciones, etc.) traía como máximo 1000 filas por tabla — con la base ya creciendo por las cargas masivas, esto dejaba fuera las filas más recientes en tablas grandes, sin ningún error visible. Ahora pagina hasta traer todo" },
  { v: "1.46.28", desc: "Fix crítico: Transacciones y Dashboard filtraban por unidad usando la partida vinculada en vez del campo propio 'unidad_detectada' — cualquier transacción cuya partida no perteneciera exactamente a esa compañía (partida borrada, mal elegida, etc.) desaparecía por completo de la vista y de los totales del Dashboard, aunque existiera en la base de datos" },
  { v: "1.46.27", desc: "Fix crítico: el choque de folio al crear transacciones seguía pasando porque el navegador tenía el estado desactualizado tras las cargas masivas grandes — ahora se confirma el número más alto real contra la base de datos justo antes de guardar, en vez de confiar solo en lo cargado localmente" },
  { v: "1.46.26", desc: "Fix: crear una transacción podía chocar con el folio de otra creada casi al mismo tiempo (por otra persona u otra pestaña) — ahora reintenta automáticamente con el siguiente número hasta 5 veces antes de mostrar error" },
  { v: "1.46.25", desc: "Fix crítico: al crear una transacción nueva, el ID (folio_transaccion) se generaba sin año (ej. OSB-AGO-001), así que Agosto 2025 y Agosto 2026 competían por la misma numeración y podían chocar — ahora incluye el año, igual que el folio de las partidas" },
  { v: "1.46.24", desc: "Fix: el estado de 'Contraer todo' del panel comparativo del Dashboard se perdía al cambiar de pestaña y regresar — ahora persiste en la sesión, igual que los demás filtros" },
  { v: "1.46.23", desc: "Dashboard: el selector de Moneda (MXP/USD) del panel comparativo cambia de menú desplegable a un switch deslizante estilo iOS" },
  { v: "1.46.22", desc: "Dashboard: 'Presupuestado' y 'Pagado real' se fusionan en un solo marco con filtros compartidos (Proyecto/Mes/Año) y un selector de Moneda (MXP/USD) nuevo, para dejar claro que el mismo filtro aplica a ambas tablas" },
  { v: "1.46.21", desc: "Dashboard: quita el gráfico 'Distribución por rubro'. 'Presupuesto vs. ejecutado por proyecto' y 'Tendencia mensual' ahora separan MXP y USD en bloques distintos, en vez de sumarlos juntos" },
  { v: "1.46.20", desc: "Dashboard: nuevo panel 'Pagado real por proyecto y mes' — misma estructura que el de presupuestado, pero suma solo transacciones con Status = Pagado, para que Dirección vea lo realmente pagado sin confundirlo con el plan" },
  { v: "1.46.19", desc: "Dashboard: 'Resumen presupuestado por proyecto y rubro' ahora tiene sus propios filtros de Proyecto/Mes/Año (compartidos con 'Resumen financiero') y un botón 'Contraer todo'" },
  { v: "1.46.18", desc: "Catálogo de Proyectos: cambia de edición implícita (clic+salir para guardar) a explícita — botón 'Editar' activa los campos, y 'Guardar'/'Cancelar' confirma o descarta el cambio" },
  { v: "1.46.17", desc: "Fix: 'Resumen presupuestado por proyecto y rubro' combinaba el mismo mes de distintos años en una sola columna — ahora, cuando los datos abarcan más de un año, cada columna se separa en 'Mes Año' (ej. Agosto 2025 y Agosto 2026 aparte), ordenadas cronológicamente" },
  { v: "1.46.16", desc: "Dashboard: agrega filtro de Año (junto a Proyecto/Mes) en 'Resumen financiero' — necesario ahora que hay datos de 2025 y 2026. Nota: la tabla de 'Resumen presupuestado por proyecto y rubro' de abajo aún combina el mismo mes de distintos años si dejas Año en 'Todos'" },
  { v: "1.46.15", desc: "Fix: agrupar por Mes ya considera el Año — antes 'Noviembre 2025' y 'Septiembre 2026' se ordenaban solo por nombre de mes (Noviembre después de Septiembre), ahora Noviembre 2025 sale primero; de paso, meses del mismo nombre mismo mes pero distinto año ya no se mezclan en un solo grupo" },
  { v: "1.46.14", desc: "Fix: el popup de '+ Nueva partida' (dentro del selector de Partida) siempre creaba la partida en el año actual, sin opción de cambiarlo — ahora trae un campo Año editable" },
  { v: "1.46.13", desc: "Partidas: agrega filtro de Año (multi-selección, junto a Mes) y 'Año' como opción de Agrupar por — necesario ahora que va a haber datos de 2025 y 2026 conviviendo" },
  { v: "1.46.12", desc: "Fix crítico: el popup de editar transacción desde Partidas dejaba la pantalla en blanco al abrir — Forma/Método de Pago intentaban renderizar el objeto {value,label} completo en vez de solo el texto" },
  { v: "1.46.11", desc: "Fix: al agrupar Transacciones por 'Mes (partida)' los meses salían en orden alfabético en vez de cronológico — el fix ya existía para el campo 'mes' de Partidas, pero no cubría el campo '_mes' que usa Transacciones" },
  { v: "1.46.10", desc: "Botón 'Contraer todo' visible junto a los controles de tabla (Partidas, Transacciones, sin vincular) — ya no hace falta abrir el menú de Agrupar por para encontrarlo. Partidas también tiene 'Contraer transacciones' para las filas expandidas" },
  { v: "1.46.9", desc: "Transacciones: agrega 'Moneda' como opción de Agrupar por (tanto en la tabla principal como en la de sin vincular)" },
  { v: "1.46.8", desc: "Partidas: la fila expandida de transacciones vinculadas trae un botón ✎ para editar esa transacción en un popup, sin salir de la vista de Partidas (no cambia de pestaña ni pierde tus filtros)" },
  { v: "1.46.7", desc: "Dashboard: 'Resumen general' pasa a formato de tabla ('Resumen financiero') con un ícono por columna y una insignia circular por moneda, siguiendo el diseño de referencia" },
  { v: "1.46.6", desc: "Dashboard: rediseña el panel 'Resumen general' — cada moneda queda en su propio bloque con fondo sutil y un pill de color, en vez de una etiqueta chica de texto arriba de cada fila" },
  { v: "1.46.5", desc: "Dashboard: el panel 'Resumen general' duplica sus 5 cuadros para mostrar USD además de MXP — de paso corrige un bug donde antes se sumaban MXP y USD juntos en esos mismos cuadros" },
  { v: "1.46.4", desc: "Transacciones: los botones Editar/Duplicar/Eliminar de cada fila cambian a íconos compactos (✎/⧉/✕) con tooltip, para que quepan bien los 3 en el espacio de la columna" },
  { v: "1.46.3", desc: "Transacciones: botón 'Duplicar' — abre el formulario de nueva transacción prellenado con los datos de la original (Status vuelve a 'No Pagado', y se limpian folios de compra/factura), para revisar y guardar como registro nuevo" },
  { v: "1.46.2", desc: "Agrega popup de confirmación antes de eliminar cualquier registro individual (partida, transacción, proveedor, cuenta bancaria, proyecto) — las eliminaciones en lote ya lo tenían" },
  { v: "1.46.1", desc: "Reporte Pagos Dirección: botón 'Generar reporte (PDF)' — toma las transacciones filtradas, genera el PDF, y las marca como 'Reportadas' (con confirmación previa)" },
  { v: "1.46.0", desc: "Transacciones: nuevo marcador 'Reportado' — selecciona varias con checkbox y dales 'Marcar como reportadas' al enviar tu reporte semanal; filtra por 'No reportado' para ver de un vistazo qué es nuevo desde tu último envío" },
  { v: "1.45.13", desc: "Partidas: cada fila con transacciones vinculadas trae un botón ▶ para expandirla y ver rápido Concepto/Monto/Status de esas transacciones, sin salir de la tabla" },
  { v: "1.45.12", desc: "Partidas: el filtro de Mes cambia de selector único a multi-selección con casillas (igual que en el Dashboard) — puedes filtrar por varios meses a la vez" },
  { v: "1.45.11", desc: "Dashboard: quita el panel 'Avance por proyecto' (ya no se necesita)" },
  { v: "1.45.10", desc: "Transacciones: renombra el campo 'Día' a 'Día de Pago Programado' — en la tabla, el formulario, la vista de sin vincular y la vista previa de importación" },
  { v: "1.45.9", desc: "Partidas: 'Ejercido' pasa de una segunda línea debajo del Monto a un tooltip al pasar el mouse (subrayado punteado de color) — evita que se agregue una fila extra al exportar/copiar a Excel" },
  { v: "1.45.8", desc: "Agrega 'Última actualización' a Partidas y Transacciones — se actualiza sola con un trigger de base de datos en cada edición (incluso ediciones directas o importaciones), visible como columna y en el modal de edición" },
  { v: "1.45.7", desc: "Fix crítico: el Dashboard mostraba $0.00 en todos los cuadros al cambiar de compañía si el filtro de Proyecto guardado en sesión pertenecía a otra compañía — ahora se resetea solo a 'Todos' cuando no existe en la unidad actual" },
  { v: "1.45.6", desc: "Dashboard: el panel 'Resumen general' ahora muestra 5 cifras — Presupuestado, Ocupado, Pagado, Por Pagar y Disponible — en vez de solo 3, respetando los filtros de Mes/Proyecto" },
  { v: "1.45.5", desc: "Fix: los filtros de Mes y Proyecto del Dashboard no persistían al cambiar de pestaña (se me había pasado aplicarles la persistencia de sesión que ya tienen Partidas/Transacciones/Reportes)" },
  { v: "1.45.4", desc: "Fix: 'Resumen presupuestado por proyecto y rubro' mostraba marcadores de prorrateo (Todos, X Gral) como si fueran proyectos reales — ahora se reparten entre los proyectos reales del catálogo, igual que en las gráficas y KPIs" },
  { v: "1.45.3", desc: "Tabla de Partidas: debajo del Monto de cada fila aparece \"Ejercido\" (la suma de sus transacciones vinculadas) — verde si va bien, ámbar cerca del límite, rojo si se pasó" },
  { v: "1.45.2", desc: "Dashboard: los filtros Proyecto y Mes del panel 'Resumen general' quedan alineados en la misma fila, en vez de uno arriba y otro abajo" },
  { v: "1.45.1", desc: "Fix: la suma de moneda mixta en encabezados de columna se salía del recuadro — ahora cada moneda va en su propia línea, y los encabezados en general ya no desbordan texto largo" },
  { v: "1.45.0", desc: "Las columnas de dinero (Monto/Importe) muestran la suma de lo visible/filtrado justo en el encabezado — separada por moneda si hay mezcla — en Partidas, Transacciones (incluyendo sin vincular) y ambos Reportes de Pagos" },
  { v: "1.44.9", desc: "Dashboard: los 3 cuadros de montos quedan dentro de un panel 'Resumen general' con filtro de Proyecto (usa los importes ya prorrateados); el filtro de Mes existente los sigue afectando también" },
  { v: "1.44.8", desc: "Dashboard: mueve 'Resumen presupuestado por proyecto y rubro' justo debajo de los cuadros de montos (Presupuestado/Ejecutado/Disponible)" },
  { v: "1.44.7", desc: "Botón 'Descargar plantilla' en la carga masiva de Partidas y Transacciones — genera el Excel con el formato correcto directo desde la app, sin tener que pedirlo por chat" },
  { v: "1.44.6", desc: "Agrega Fecha de Pago a Transacciones — Status nace en 'No Pagado', y marcar 'Pagado' exige capturar la fecha antes de guardar. Fix de paso: el color de Status marcaba 'No Pagado' en verde por error" },
  { v: "1.44.5", desc: "El ID de transacción cambia de formato secuencial (CTM-T-0001) a codificado por mes (CTM-AGO-001) — corre recodificar-folio-transaccion-por-mes.sql para actualizar las que ya tenían el formato viejo" },
  { v: "1.44.4", desc: "Cada transacción tiene un ID único visible (ej. CTM-T-0001) — nueva columna, buscable, y visible en el modal; se les asignó también a las transacciones que ya existían" },
  { v: "1.44.3", desc: "Filtros, fechas, orden y agrupamiento ahora persisten al cambiar de pestaña mientras la sesión siga abierta (Partidas, Transacciones y ambos Reportes) — se pierden al cerrar el navegador, a propósito" },
  { v: "1.44.2", desc: "Agrega campo SMI a Transacciones (formulario, columna, importador) — Reporte de Pagos usa el SMI de la transacción si existe, si no el de la partida" },
  { v: "1.44.1", desc: "Fix crítico: no se podía guardar una transacción nueva (violaba la política de seguridad) porque el formulario no llenaba 'unidad_detectada' al crearla manualmente" },
  { v: "1.44.0", desc: "Notas privadas por transacción — cada usuario puede dejar su propia nota (🔒), nadie más la ve ni puede consultarla, ni siquiera directo en la base de datos" },
  { v: "1.43.0", desc: "Acceso por compañía: se puede restringir a un usuario a ver solo OSB/CTM/ISE específicas — reforzado a nivel de base de datos (no solo ocultar botones), configurable desde perfiles.unidades_permitidas en Supabase" },
  { v: "1.42.2", desc: "Agrega pantalla de 'nueva contraseña' (faltaba para que los links de restablecimiento funcionaran de punta a punta) y '¿Olvidaste tu contraseña?' autoservicio en el login" },
  { v: "1.42.1", desc: "Fix crítico: tras iniciar sesión, todas las tablas aparecían vacías — la primera carga de datos salía antes de que la sesión estuviera lista; ahora se reintenta cuando el estado de sesión cambia" },
  { v: "1.42.0", desc: "Login obligatorio (Supabase Auth) y auditoría — cada partida/transacción/proveedor/cuenta guarda quién la creó y quién la editó por última vez, visible en su modal" },
  { v: "1.41.0", desc: "Los avisos de Reporte de Pagos (proveedor/cuenta sin vincular) ahora son una barra fija abajo del navegador, visible aunque hagas scroll" },
  { v: "1.40.1", desc: "Quita las comillas alrededor de los valores dinámicos (fechas, compañía, moneda, zona) en los títulos de ambos reportes exportados" },
  { v: "1.40.0", desc: "Reporte de Pagos: exportación a Excel agrupada por Zona y Moneda (bloque por cada combinación con datos, formato de la plantilla) — reemplaza la exportación plana anterior" },
  { v: "1.39.0", desc: "Reporte Pagos Dirección: exportación a Excel con el formato exacto solicitado (título con fechas/compañía, totales MXP/USD arriba, encabezados centrados, moneda formateada) — cambia de xlsx a exceljs para soportar estilos" },
  { v: "1.38.1", desc: "Reporte de Pagos: agrega columnas No. Cuenta y SWIFT (de la cuenta bancaria vinculada), visibles en tabla y exportación a Excel" },
  { v: "1.38.0", desc: "Agrega Referencia y Notas al catálogo de Proveedores — visibles en el formulario, la tabla, la carga masiva, y el Reporte de Pagos (con exportación a Excel)" },
  { v: "1.37.1", desc: "Fix de robustez: los popups de Partida/Proveedor (anidados dentro del modal de transacción) usan z-index más alto y filas como <button> real, para que el clic siempre registre" },
  { v: "1.37.0", desc: "Selector de Proveedor rediseñado: clic resalta (no cierra), botón 'Seleccionar' confirma y cierra, botón 'Editar proveedor' para el resaltado (identidad + cuentas bancarias)" },
  { v: "1.36.0", desc: "Tabla de Transacciones: la columna Proveedor muestra si está vinculado al catálogo y si tiene cuenta bancaria (Sin catálogo / Sin cuenta / Con cuenta)" },
  { v: "1.35.0", desc: "'+ Nuevo proveedor' ahora incluye un segundo paso para agregar cuenta(s) bancaria(s) antes de usarlo, sin salir de la transacción" },
  { v: "1.34.1", desc: "Fix: el popup de elegir Partida/Proveedor ahora se cierra de forma confiable al hacer clic (se cierra primero, luego procesa la selección)" },
  { v: "1.34.0", desc: "Selector de Proveedor rediseñado: popup con buscador (nombre/RFC/Id SAE) y '+ Nuevo proveedor' al vuelo, igual que el selector de Partida" },
  { v: "1.33.0", desc: "Homologa 'MXN' a 'MXP' en todo el código (defaults, catálogo de monedas, importadores, reportes) — corre homologar-mxn-a-mxp.sql para actualizar los datos ya guardados" },
  { v: "1.32.3", desc: "Fix global: texto largo sin espacios (CLABE, RFC, folios) ya no se sale de su columna y se encima con la siguiente — aplica a todas las tablas" },
  { v: "1.32.2", desc: "Reporte de Pagos y Reporte Pagos Dirección: separa el KPI de Importe total en Total MXN y Total USD (antes se sumaban las dos divisas juntas)" },
  { v: "1.32.1", desc: "Filas alternadas (cebra) en todas las tablas de la app, para facilitar la lectura" },
  { v: "1.32.0", desc: "Tabla 'sin vincular' de Transacciones: agrega filtro Desde/Hasta, Agrupar por, y control de columnas — mismo patrón que las demás tablas" },
  { v: "1.31.1", desc: "Fix: 'Proyecto' en el popup de nueva partida rápida ahora es un selector del catálogo, no texto libre" },
  { v: "1.31.0", desc: "Selector de Partida: botón '+ Nueva partida' para crear una al vuelo sin salir del formulario de transacción" },
  { v: "1.30.0", desc: "Control 'Columnas' (mostrar/ocultar con checkboxes) en Partidas, Transacciones, Reporte de Pagos y Reporte Pagos Dirección" },
  { v: "1.29.1", desc: "Selector de Partida: cada fila muestra Total y Usado (con color según % consumido)" },
  { v: "1.29.0", desc: "Selector de Partida rediseñado: popup con buscador y agrupado por mes (en modal, columna inline, y panel de sin vincular) — confirmado que ya filtraba solo por la compañía activa" },
  { v: "1.28.0", desc: "Proveedores: agrega Sucursal y SWIFT a las cuentas; la carga masiva ahora agrupa filas repetidas del mismo proveedor (varias cuentas) aunque sea nuevo en el mismo archivo" },
  { v: "1.27.2", desc: "Fix: editar una transacción agrupada fallaba porque colaban campos internos (_mes, _rubro, etc.) al guardar" },
  { v: "1.27.1", desc: "Transacciones: simplifica filtros a solo Buscar + rango de fechas (quita Zona/Vínculo/Mes/Proyecto)" },
  { v: "1.27.0", desc: "Nueva pestaña 'Reporte Pagos Dirección' — versión resumida (Día, Solicitante, Proyecto, Zona, Proveedor, Concepto, Importe, Moneda, A Partida, Status), mismo filtro de fechas" },
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

// Identificador único por transacción (no ligado al mes, a diferencia del folio
// de partida) — ej. "CTM-T-0001". Se asigna solo, no se puede editar.
function nextFolioTransaccion(unidad, mes, anio, existingTransacciones, offset = 0) {
  const abr = MES_ABR[mes] || "GEN";
  const prefix = anio ? `${unidad}-${abr}${String(anio).slice(-2)}-` : `${unidad}-${abr}-`;
  let max = 0;
  existingTransacciones.forEach((t) => {
    if (t.folio_transaccion && t.folio_transaccion.startsWith(prefix)) {
      const n = parseInt(t.folio_transaccion.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return { prefix, siguiente: max + 1 + offset };
}

// Consulta a Supabase (no al estado local, que puede estar desactualizado tras
// una carga masiva grande) el folio_transaccion más alto que ya existe para
// este prefijo — así el "siguiente número" siempre es real, no una suposición
// basada en lo que el navegador alcanzó a recibir por realtime.
async function maxFolioTransaccionReal(prefix) {
  const { data, error } = await supabase
    .from("transacciones")
    .select("folio_transaccion")
    .like("folio_transaccion", `${prefix}%`)
    .order("folio_transaccion", { ascending: false })
    .limit(1);
  if (error || !data?.length) return 0;
  const n = parseInt(data[0].folio_transaccion.slice(prefix.length), 10);
  return isNaN(n) ? 0 : n;
}

// Inserta una transacción nueva reintentando el folio si otra persona (u otra
// pestaña) alcanzó a tomar el mismo número justo antes — el cálculo del
// siguiente folio se hace en el navegador, así que dos inserciones casi
// simultáneas pueden calcular el mismo. Si el choque es por otra causa
// (no folio_transaccion), no reintenta — deja que el error normal se muestre.
async function insertTransaccionConReintento(transaccionesApi, rest, unidad, mesForm, anioForm, transUnidad) {
  const { prefix, siguiente: siguienteLocal } = nextFolioTransaccion(unidad, mesForm, anioForm, transUnidad, 0);
  // El estado local (transUnidad) puede estar desactualizado tras una carga
  // masiva grande — se confirma el número real contra la base antes de usarlo.
  const maxReal = await maxFolioTransaccionReal(prefix);
  let siguiente = Math.max(siguienteLocal, maxReal + 1);

  const maxIntentos = 8;
  for (let intento = 0; intento < maxIntentos; intento++) {
    const folio = `${prefix}${String(siguiente + intento).padStart(3, "0")}`;
    try {
      return await transaccionesApi.insert({ ...rest, id: uid(), folio_transaccion: folio });
    } catch (err) {
      const chocoPorFolio = /folio_transaccion|idx_transacciones_folio_transaccion/i.test(err?.message || "");
      if (!chocoPorFolio || intento === maxIntentos - 1) throw err;
      // reintenta con el siguiente número
    }
  }
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

const money = (n, moneda = "MXP") =>
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

// Renderiza las <option> de un selector de Partida agrupadas por Mes (<optgroup>),
// en orden cronológico — para no tener que buscar entre una lista plana larga.
function opcionesPartidaPorMes(lista) {
  const meses = MESES.filter((m) => lista.some((p) => p.mes === m));
  const sinMes = lista.filter((p) => !p.mes);
  return (
    <>
      {meses.map((mes) => (
        <optgroup key={mes} label={mes}>
          {lista.filter((p) => p.mes === mes).map((p) => (
            <option key={p.id} value={p.id}>{p.concepto} ({p.proyecto || "—"})</option>
          ))}
        </optgroup>
      ))}
      {sinMes.length > 0 && (
        <optgroup label="Sin mes">
          {sinMes.map((p) => <option key={p.id} value={p.id}>{p.concepto} ({p.proyecto || "—"})</option>)}
        </optgroup>
      )}
    </>
  );
}

// Botón que abre un popup con buscador para elegir una partida — más cómodo
// que un <select> plano cuando hay muchas. Agrupa por mes, en orden cronológico.
function PartidaPickerButton({ partidas, transacciones = [], value, onChange, placeholder = "Elegir partida…", allowClear = false, partidasApi, unidad, proyectosOpciones = [], ocultasPorMoneda = 0, moneda }) {
  const [open, setOpen] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [filtroRubro, setFiltroRubro] = useState("Todos");
  const [filtroMes, setFiltroMes] = useState("Todos");
  const seleccionada = partidas.find((p) => p.id === value);

  const rubrosDisponibles = [...new Set(partidas.map((p) => p.rubro).filter(Boolean))].sort();
  const mesesDisponibles = MESES.filter((m) => partidas.some((p) => p.mes === m));

  const filtradas = partidas.filter((p) => {
    if (filtroRubro !== "Todos" && p.rubro !== filtroRubro) return false;
    if (filtroMes !== "Todos" && p.mes !== filtroMes) return false;
    if (!busqueda.trim()) return true;
    const q = busqueda.trim().toLowerCase();
    // Se busca también en categoría y zona: con partidas agrupadas por rubro,
    // el concepto solo no basta para distinguirlas.
    return [p.concepto, p.folio, p.proyecto, p.rubro, p.categoria, p.zona]
      .some((v) => (v || "").toLowerCase().includes(q));
  });
  const meses = MESES.filter((m) => filtradas.some((p) => p.mes === m));
  const sinMes = filtradas.filter((p) => !p.mes);

  const elegir = (id) => {
    setOpen(false); setBusqueda(""); setFiltroRubro("Todos"); setFiltroMes("Todos"); setCreando(false);
    try { onChange(id); } catch (err) { console.error("Error al elegir partida:", err); }
  };

  const usadoDe = (p) => transacciones.filter((t) => t.partida_id === p.id).reduce((s, t) => s + (Number(t.importe) || 0), 0);

  const nuevaPartidaBlank = { mes: MESES[0], anio: new Date().getFullYear(), concepto: "", rubro: RUBROS[0]?.rubro || "", proyecto: proyectosOpciones[0] || "", monto_estimado: "", moneda: "MXP" };
  const [creando, setCreando] = useState(false);
  const [nuevaPartida, setNuevaPartida] = useState(nuevaPartidaBlank);
  const [guardandoPartida, setGuardandoPartida] = useState(false);

  const crearPartida = async () => {
    if (!nuevaPartida.concepto.trim() || !nuevaPartida.monto_estimado) return;
    setGuardandoPartida(true);
    try {
      const anio = Number(nuevaPartida.anio) || new Date().getFullYear();
      const existingFolios = partidas.filter((p) => p.unidad === unidad).map((p) => p.folio);
      const folio = autoFolio(unidad, nuevaPartida.mes, anio, existingFolios);
      const categoriaDefault = RUBROS.find((r) => r.rubro === nuevaPartida.rubro)?.categorias?.[0] || "";
      const creada = await partidasApi.insert({
        id: uid(), unidad, mes: nuevaPartida.mes, anio, smi: "", concepto: nuevaPartida.concepto.trim(),
        rubro: nuevaPartida.rubro, categoria: categoriaDefault, proyecto: nuevaPartida.proyecto.trim(),
        monto_estimado: Number(nuevaPartida.monto_estimado) || 0, moneda: nuevaPartida.moneda, folio,
      });
      setNuevaPartida(nuevaPartidaBlank);
      elegir(creada.id);
    } catch (err) {
      alert("No se pudo crear la partida: " + (err.message || err));
    } finally {
      setGuardandoPartida(false);
    }
  };

  const FilaPartida = (p) => {
    const total = Number(p.monto_estimado) || 0;
    const usado = usadoDe(p);
    const pct = total ? (usado / total) * 100 : 0;
    const tone = pct > 100 ? T.red : pct > 85 ? T.amber : T.teal;
    return (
      <button
        type="button"
        key={p.id}
        onClick={() => elegir(p.id)}
        style={{ display: "block", width: "100%", textAlign: "left", border: "none", padding: "9px 12px", cursor: "pointer", borderBottom: `1px solid ${T.borderSoft}`, background: p.id === value ? T.accentBg : "transparent" }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 12.5, color: T.text }}>{p.concepto}</div>
            <div style={{ fontSize: 11, color: T.textFaint, marginTop: 2 }}>{p.proyecto || "—"}{p.folio ? ` · ${p.folio}` : ""}</div>
          </div>
          <div style={{ textAlign: "right", fontSize: 10.5, fontFamily: T.fontMono, flexShrink: 0, whiteSpace: "nowrap" }}>
            <div style={{ color: T.textDim }}>Total {money(total, p.moneda)}</div>
            <div style={{ color: tone }}>Usado {money(usado, p.moneda)}</div>
          </div>
        </div>
      </button>
    );
  };

  return (
    <>
      <button
        type="button"
        onClick={() => setOpen(true)}
        style={{
          ...inputStyle, width: "100%", textAlign: "left", cursor: "pointer",
          display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8,
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {seleccionada ? `${seleccionada.mes} · ${seleccionada.concepto}` : placeholder}
        </span>
        <span style={{ color: T.textFaint, fontSize: 10, flexShrink: 0 }}>▾</span>
      </button>

      {open && (
        <Modal title="Elegir partida" subtitle="Busca por concepto, folio, rubro o proyecto — Total y Usado por partida" onClose={() => { setOpen(false); setCreando(false); }} width={620} zIndex={1100}>
          {partidasApi && !creando && (
            <Button type="button" variant="ghost" onClick={() => setCreando(true)} style={{ marginBottom: 10 }}>
              + Nueva partida
            </Button>
          )}

          {creando ? (
            <div style={{ border: `1px solid ${T.borderSoft}`, borderRadius: 6, padding: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: T.text, marginBottom: 10 }}>Nueva partida — {unidad}</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10 }}>
                <Field label="Concepto" style={{ gridColumn: "span 2" }}>
                  <TextInput autoFocus value={nuevaPartida.concepto} onChange={(e) => setNuevaPartida({ ...nuevaPartida, concepto: e.target.value })} />
                </Field>
                <Field label="Mes">
                  <Select value={nuevaPartida.mes} onChange={(e) => setNuevaPartida({ ...nuevaPartida, mes: e.target.value })}>
                    {MESES.map((m) => <option key={m}>{m}</option>)}
                  </Select>
                </Field>
                <Field label="Año">
                  <TextInput type="number" value={nuevaPartida.anio} onChange={(e) => setNuevaPartida({ ...nuevaPartida, anio: e.target.value })} />
                </Field>
                <Field label="Rubro">
                  <Select value={nuevaPartida.rubro} onChange={(e) => setNuevaPartida({ ...nuevaPartida, rubro: e.target.value })}>
                    {RUBROS.map((r) => <option key={r.rubro}>{r.rubro}</option>)}
                  </Select>
                </Field>
                <Field label="Proyecto">
                  <Select value={nuevaPartida.proyecto} onChange={(e) => setNuevaPartida({ ...nuevaPartida, proyecto: e.target.value })}>
                    {proyectosOpciones.length === 0 && <option value="">Sin proyectos configurados — ve a Catálogo</option>}
                    {proyectosOpciones.map((m) => <option key={m}>{m}</option>)}
                  </Select>
                </Field>
                <Field label="Monto estimado">
                  <TextInput type="number" step="0.01" value={nuevaPartida.monto_estimado} onChange={(e) => setNuevaPartida({ ...nuevaPartida, monto_estimado: e.target.value })} placeholder="0.00" />
                </Field>
                <Field label="Moneda">
                  <Select value={nuevaPartida.moneda} onChange={(e) => setNuevaPartida({ ...nuevaPartida, moneda: e.target.value })}>
                    {MONEDAS.map((m) => <option key={m}>{m}</option>)}
                  </Select>
                </Field>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <Button type="button" onClick={crearPartida} disabled={guardandoPartida}>{guardandoPartida ? "Creando…" : "Crear y usar esta partida"}</Button>
                <Button type="button" variant="ghost" onClick={() => { setCreando(false); setNuevaPartida(nuevaPartidaBlank); }}>Cancelar</Button>
              </div>
            </div>
          ) : (
            <>
              <TextInput
                autoFocus
                value={busqueda}
                onChange={(e) => setBusqueda(e.target.value)}
                placeholder="Buscar por concepto, folio, rubro, categoría, proyecto o zona…"
                style={{ width: "100%", marginBottom: 8 }}
              />
              <div style={{ display: "flex", gap: 8, marginBottom: 10, alignItems: "center", flexWrap: "wrap" }}>
                <Select value={filtroRubro} onChange={(e) => setFiltroRubro(e.target.value)} style={{ flex: 1, minWidth: 150 }}>
                  <option>Todos</option>
                  {rubrosDisponibles.map((r) => <option key={r}>{r}</option>)}
                </Select>
                <Select value={filtroMes} onChange={(e) => setFiltroMes(e.target.value)} style={{ width: 140 }}>
                  <option>Todos</option>
                  {mesesDisponibles.map((m) => <option key={m}>{m}</option>)}
                </Select>
                <span style={{ fontSize: 11.5, color: T.textFaint, fontFamily: T.fontMono, whiteSpace: "nowrap" }}>
                  {filtradas.length} de {partidas.length}
                </span>
              </div>
              {/* El formulario solo ofrece partidas de la moneda del gasto. Sin
                  decirlo, parece que faltan partidas — que es exactamente como
                  se reportó este comportamiento. */}
              {ocultasPorMoneda > 0 && (
                <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "8px 11px", borderRadius: "0 6px 6px 0", fontSize: 11.5, marginBottom: 10 }}>
                  Se ocultan <b>{ocultasPorMoneda}</b> partida(s) de otra moneda. Esta transacción está en <b>{moneda || "MXP"}</b>;
                  cambia la moneda del formulario si el gasto es en la otra.
                </div>
              )}
              <div style={{ maxHeight: 420, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
                {allowClear && (
                  <button
                    type="button"
                    onClick={() => elegir("")}
                    style={{ display: "block", width: "100%", textAlign: "left", border: "none", padding: "9px 12px", cursor: "pointer", fontSize: 12.5, color: T.textFaint, borderBottom: `1px solid ${T.borderSoft}` }}
                  >
                    — Sin vincular —
                  </button>
                )}
                {meses.map((mes) => (
                  <div key={mes}>
                    <div style={{ padding: "6px 12px", fontSize: 10.5, fontWeight: 700, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.05em", background: T.panelAlt }}>
                      {mes}
                    </div>
                    {filtradas.filter((p) => p.mes === mes).map((p) => FilaPartida(p))}
                  </div>
                ))}
                {sinMes.map((p) => FilaPartida(p))}
                {!filtradas.length && (
                  <div style={{ padding: 16, textAlign: "center", fontSize: 12, color: T.textFaint }}>Sin resultados</div>
                )}
              </div>
            </>
          )}
        </Modal>
      )}
    </>
  );
}

// Botón que abre un popup con buscador para elegir un proveedor del catálogo —
// mismo patrón que PartidaPickerButton. Incluye "+ Nuevo proveedor" al vuelo.
function ProveedorPickerButton({ proveedores, value, onChange, placeholder = "Elegir proveedor…", proveedoresApi, cuentasApi, unidad }) {
  const [open, setOpen] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [resaltadoId, setResaltadoId] = useState(value || "");
  const seleccionado = proveedores.find((p) => p.id === value);

  const filtrados = proveedores.filter((p) => {
    if (!busqueda.trim()) return true;
    const q = busqueda.trim().toLowerCase();
    return [p.nombre, p.rfc, p.id_sae, p.referencia, p.notas].some((v) => (v || "").toLowerCase().includes(q));
  });

  const cuentaBlank = { banco: "", sucursal: "", swift: "", clabe: "", numero_cuenta: "", divisa: "MXP" };
  const [editando, setEditando] = useState(null); // null = lista; {id:null,...} = creando; {id:"x",...} = editando existente
  const [guardando, setGuardando] = useState(false);
  const [nuevaCuenta, setNuevaCuenta] = useState(cuentaBlank);
  const cuentasDeEditando = (editando?.id && cuentasApi) ? cuentasApi.rows.filter((c) => c.proveedor_id === editando.id) : [];

  /* Búsqueda en el catálogo maestro (11 mil proveedores de ASPEL).
     Se consulta a Supabase EN EL MOMENTO, no se carga en memoria: traer once
     mil registros en cada sesión degradaría la app para todos a cambio de un
     catálogo que casi nunca se consulta. */
  const [buscandoMaestro, setBuscandoMaestro] = useState(false);
  const [resMaestro, setResMaestro] = useState(null); // null = no se ha buscado
  const [avisoMaestro, setAvisoMaestro] = useState("");

  const buscarEnMaestro = async () => {
    const q = busqueda.trim();
    if (q.length < 3) { alert("Escribe al menos 3 letras del nombre o el RFC para buscar en el maestro."); return; }
    setBuscandoMaestro(true);
    try {
      const { data, error } = await supabase
        .from("proveedores_maestro")
        .select("*")
        .eq("compania", unidad)
        .or(`nombre.ilike.%${q}%,rfc.ilike.%${q}%,id_sae.eq.${/^\d+$/.test(q) ? q : -1}`)
        .order("nombre")
        .limit(25);
      if (error) throw error;
      setResMaestro(data || []);
    } catch (err) {
      alert("No se pudo buscar en el maestro: " + (err.message || err));
      setResMaestro([]);
    } finally {
      setBuscandoMaestro(false);
    }
  };

  /* Trae un proveedor del maestro al formulario de alta. No lo guarda solo:
     el maestro trae CLABE en apenas el 30% de los casos y la divisa casi
     nunca, así que darlo de alta en silencio crearía proveedores incompletos
     que fallan al momento de pagar. */
  const usarDelMaestro = (m) => {
    // Llenar el formulario no basta: sólo se dibuja cuando `editando` tiene
    // valor, así que sin esto el clic llenaba una pantalla invisible y
    // parecía que no hacía nada.
    setEditando({ id: null, nombre: m.nombre || "", rfc: m.rfc || "", id_sae: m.id_sae || "" });
    setNuevaCuenta({
      banco: m.banco || "", sucursal: m.sucursal || "", swift: m.swift || "",
      clabe: m.clabe || "", numero_cuenta: m.numero_cuenta || "",
      divisa: m.divisa || "MXP",
    });
    setResMaestro(null);
    setBusqueda("");
    // El aviso se muestra DENTRO del formulario, no en un alert que hay que
    // cerrar antes de ver los datos que describe.
    setAvisoMaestro(
      m.revisar ? `Este proveedor viene marcado en el maestro: ${m.revisar}. Verifica los datos bancarios contra la factura antes de guardarlo.`
      : (!m.clabe && !m.numero_cuenta)
        ? "El maestro no tiene datos bancarios de este proveedor. Se llenaron nombre, RFC e Id SAE; captura la cuenta antes de usarlo para pagar."
        : `Datos traídos del catálogo maestro (SAE ${m.id_sae}). Revísalos y guarda para agregarlo a ${unidad}.`
    );
  };

  const abrir = () => { setResaltadoId(value || ""); setEditando(null); setBusqueda(""); setResMaestro(null); setAvisoMaestro(""); setOpen(true); };
  const cerrar = () => { setOpen(false); setBusqueda(""); setEditando(null); setResMaestro(null); setAvisoMaestro(""); };

  const confirmar = () => {
    const p = proveedores.find((pr) => pr.id === resaltadoId) || (editando?.id === resaltadoId ? editando : null);
    cerrar();
    try { onChange(resaltadoId, p); } catch (err) { console.error("Error al elegir proveedor:", err); }
  };

  const iniciarCrear = () => setEditando({ id: null, nombre: "", rfc: "", id_sae: "" });
  const iniciarEditar = () => {
    const p = proveedores.find((pr) => pr.id === resaltadoId);
    if (p) setEditando({ id: p.id, nombre: p.nombre || "", rfc: p.rfc || "", id_sae: p.id_sae || "" });
  };

  const guardarIdentidad = async () => {
    if (!editando.nombre.trim()) return;
    setGuardando(true);
    try {
      if (editando.id) {
        await proveedoresApi.update(editando.id, { nombre: editando.nombre.trim(), rfc: editando.rfc.trim().toUpperCase(), id_sae: editando.id_sae.trim() });
      } else {
        const creado = await proveedoresApi.insert({ id: uid(), unidad, nombre: editando.nombre.trim(), rfc: editando.rfc.trim().toUpperCase(), id_sae: editando.id_sae.trim() });
        setEditando({ ...editando, id: creado.id });
        setResaltadoId(creado.id);
      }
    } catch (err) {
      alert("No se pudo guardar: " + (err.message || err));
    } finally {
      setGuardando(false);
    }
  };

  const agregarCuentaEditando = async () => {
    if (!editando?.id) return;
    if (!nuevaCuenta.banco.trim() && !nuevaCuenta.clabe.trim() && !nuevaCuenta.numero_cuenta.trim()) return;
    try {
      await cuentasApi.insert({ id: uid(), proveedor_id: editando.id, ...nuevaCuenta });
      setNuevaCuenta(cuentaBlank);
    } catch (err) {
      alert("No se pudo agregar la cuenta: " + (err.message || err));
    }
  };
  const eliminarCuentaEditando = (id) => {
    if (!confirm("¿Eliminar esta cuenta bancaria? Esto no se puede deshacer.")) return;
    cuentasApi.remove(id).catch((err) => alert("No se pudo eliminar la cuenta: " + (err.message || err)));
  };

  return (
    <>
      <button
        type="button"
        onClick={abrir}
        style={{
          ...inputStyle, width: "100%", textAlign: "left", cursor: "pointer",
          display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8,
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {seleccionado ? seleccionado.nombre : placeholder}
        </span>
        <span style={{ color: T.textFaint, fontSize: 10, flexShrink: 0 }}>▾</span>
      </button>

      {open && (
        <Modal title="Elegir proveedor" subtitle="Busca por nombre, RFC o Id SAE" onClose={cerrar} width={560} zIndex={1100}>
          {editando ? (
            <div style={{ border: `1px solid ${T.borderSoft}`, borderRadius: 6, padding: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: T.text, marginBottom: 10 }}>
                {editando.id ? "Editar proveedor" : "Nuevo proveedor"} — {unidad}
              </div>
              {avisoMaestro && (
                <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "8px 11px",
                              borderRadius: "0 6px 6px 0", fontSize: 11.5, marginBottom: 12 }}>
                  {avisoMaestro}
                </div>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10 }}>
                <Field label="Nombre" style={{ gridColumn: "span 2" }}>
                  <TextInput autoFocus value={editando.nombre} onChange={(e) => setEditando({ ...editando, nombre: e.target.value })} />
                </Field>
                <Field label="RFC">
                  <TextInput value={editando.rfc} onChange={(e) => setEditando({ ...editando, rfc: e.target.value.toUpperCase() })} />
                </Field>
                <Field label="Id SAE">
                  <TextInput value={editando.id_sae} onChange={(e) => setEditando({ ...editando, id_sae: e.target.value })} />
                </Field>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <Button type="button" onClick={guardarIdentidad} disabled={guardando}>
                  {guardando ? "Guardando…" : editando.id ? "Guardar cambios" : "Crear proveedor"}
                </Button>
                <Button type="button" variant="ghost" onClick={() => { setEditando(null); setAvisoMaestro(""); }}>Volver a la lista</Button>
              </div>

              {editando.id && (
                <div style={{ marginTop: 16, borderTop: `1px solid ${T.borderSoft}`, paddingTop: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: T.text, marginBottom: 10 }}>
                    Cuentas bancarias ({cuentasDeEditando.length})
                  </div>
                  {cuentasDeEditando.length > 0 && (
                    <table style={{ ...tableStyle, marginBottom: 10 }}>
                      <thead>
                        <tr>{["Banco","CLABE","No. Cuenta","Divisa",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
                      </thead>
                      <tbody>
                        {cuentasDeEditando.map((c) => (
                          <tr key={c.id}>
                            <td style={tdStyle}>{c.banco || "—"}</td>
                            <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.clabe || "—"}</td>
                            <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.numero_cuenta || "—"}</td>
                            <td style={tdStyle}><Pill>{c.divisa || "MXP"}</Pill></td>
                            <td style={tdStyle}><Button type="button" variant="danger" onClick={() => eliminarCuentaEditando(c.id)}>Eliminar</Button></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 8 }}>
                    <Field label="Banco">
                      <TextInput value={nuevaCuenta.banco} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, banco: e.target.value })} />
                    </Field>
                    <Field label="Sucursal">
                      <TextInput value={nuevaCuenta.sucursal} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, sucursal: e.target.value })} />
                    </Field>
                    <Field label="SWIFT">
                      <TextInput value={nuevaCuenta.swift} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, swift: e.target.value })} />
                    </Field>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, alignItems: "end" }}>
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
                    <Button type="button" variant="ghost" onClick={agregarCuentaEditando}>+ Agregar cuenta</Button>
                  </div>
                  <Button type="button" onClick={confirmar} style={{ marginTop: 12 }}>Seleccionar este proveedor</Button>
                </div>
              )}
            </div>
          ) : (
            <>
              <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                {proveedoresApi && (
                  <Button type="button" variant="ghost" onClick={iniciarCrear}>+ Nuevo proveedor</Button>
                )}
                {proveedoresApi && resaltadoId && (
                  <Button type="button" variant="ghost" onClick={iniciarEditar}>Editar proveedor</Button>
                )}
              </div>
              <TextInput
                autoFocus
                value={busqueda}
                onChange={(e) => setBusqueda(e.target.value)}
                placeholder="Buscar…"
                style={{ width: "100%", marginBottom: 12 }}
              />
              <div style={{ maxHeight: 380, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6, marginBottom: 12 }}>
                <button
                  type="button"
                  onClick={() => setResaltadoId("")}
                  style={{ display: "block", width: "100%", textAlign: "left", border: "none", padding: "9px 12px", cursor: "pointer", fontSize: 12.5, color: T.textFaint, borderBottom: `1px solid ${T.borderSoft}`, background: resaltadoId === "" ? T.accentBg : "transparent" }}
                >
                  — No catalogado —
                </button>
                {filtrados.map((p) => (
                  <button
                    type="button"
                    key={p.id}
                    onClick={() => setResaltadoId(p.id)}
                    onDoubleClick={confirmar}
                    style={{ display: "block", width: "100%", textAlign: "left", border: "none", padding: "9px 12px", cursor: "pointer", borderBottom: `1px solid ${T.borderSoft}`, background: p.id === resaltadoId ? T.accentBg : "transparent" }}
                  >
                    <div style={{ fontSize: 12.5, color: T.text }}>{p.nombre}</div>
                    <div style={{ fontSize: 11, color: T.textFaint, marginTop: 2 }}>{p.rfc || "—"}{p.id_sae ? ` · SAE ${p.id_sae}` : ""}</div>
                  </button>
                ))}
                {!filtrados.length && (
                  <div style={{ padding: 16, textAlign: "center", fontSize: 12, color: T.textFaint }}>
                    Sin resultados en {unidad}
                  </div>
                )}
              </div>

              {/* Puente al catálogo maestro. Aparece siempre, no solo cuando la
                  búsqueda local falla: un proveedor puede existir aquí con
                  datos incompletos y estar completo allá. */}
              <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${T.borderSoft}` }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <Button type="button" variant="ghost" onClick={buscarEnMaestro} disabled={buscandoMaestro}>
                    {buscandoMaestro ? "Buscando…" : "Buscar en el catálogo maestro"}
                  </Button>
                  <span style={{ fontSize: 11.5, color: T.textFaint }}>
                    11 mil proveedores de ASPEL — se consulta al momento, no está cargado en la app
                  </span>
                </div>

                {resMaestro !== null && (
                  <div style={{ marginTop: 10 }}>
                    {!resMaestro.length ? (
                      <div style={{ fontSize: 12, color: T.textFaint, padding: "8px 2px" }}>
                        Tampoco está en el maestro de {unidad}. Usa "+ Nuevo proveedor" para darlo de alta.
                      </div>
                    ) : (
                      <div style={{ maxHeight: 220, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
                        {resMaestro.map((m) => (
                          <div key={m.id}
                            onClick={() => usarDelMaestro(m)}
                            title="Usar este proveedor"
                            style={{ padding: "8px 11px", borderBottom: `1px solid ${T.borderSoft}`, cursor: "pointer", fontSize: 12.5,
                                     display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                            <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
                              <span style={{ fontWeight: 600 }}>{m.nombre}</span>
                              <span style={{ fontFamily: T.fontMono, color: T.textFaint, fontSize: 11, whiteSpace: "nowrap" }}>
                                SAE {m.id_sae}
                              </span>
                            </div>
                            <div style={{ color: T.textDim, fontSize: 11.5, marginTop: 2 }}>
                              {m.rfc || "sin RFC"}
                              {m.banco ? ` · ${m.banco}` : ""}
                              {/* Que tenga o no datos bancarios cambia lo que falta capturar
                                  después, así que se dice desde la lista. */}
                              {m.clabe || m.numero_cuenta
                                ? <span style={{ color: T.teal }}> · con datos bancarios</span>
                                : <span style={{ color: T.amber }}> · sin cuenta</span>}
                              {m.revisar ? <span style={{ color: T.red }}> · revisar: {m.revisar}</span> : null}
                            </div>
                            </div>
                            {/* Un botón explícito: la fila entera es pulsable, pero
                                sin una señal visible nadie lo descubre. */}
                            <Button type="button" variant="ghost" style={{ padding: "4px 10px", fontSize: 11.5, whiteSpace: "nowrap" }}
                              onClick={(e) => { e.stopPropagation(); usarDelMaestro(m); }}>
                              Usar
                            </Button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <Button type="button" onClick={confirmar}>Seleccionar</Button>
                <Button type="button" variant="ghost" onClick={cerrar}>Cancelar</Button>
              </div>
            </>
          )}
        </Modal>
      )}
    </>
  );
}

/* ----------------------------------------------------------------------
   SOLICITUD DE PAGO A PROVEEDORES (SPP)
   ---------------------------------------------------------------------- */

/**
 * Esquemas fiscales. Las retenciones dependen del RÉGIMEN del proveedor y
 * del tipo de servicio, no del importe, así que la app no puede deducirlas
 * sola: propone el más común y quien genera la solicitud confirma o ajusta.
 *
 * Calcularlas mal en un documento que va a Pagos es peor que preguntarlas.
 */
const ESQUEMAS_FISCALES = [
  { id: "pf_servicios", label: "Persona física — servicios profesionales",
    iva: 0.16, retIsr: 0.0125, retIvaSobreIva: 2 / 3,
    nota: "IVA 16%, ISR 1.25%, retención de 2/3 del IVA" },
  { id: "pf_arrendamiento", label: "Persona física — arrendamiento",
    iva: 0.16, retIsr: 0.10, retIvaSobreIva: 2 / 3,
    nota: "IVA 16%, ISR 10%, retención de 2/3 del IVA" },
  { id: "pm_sin_ret", label: "Persona moral — sin retenciones",
    iva: 0.16, retIsr: 0, retIvaSobreIva: 0,
    nota: "IVA 16%, sin retenciones — el caso más común entre empresas" },
  { id: "pm_flete", label: "Persona moral — autotransporte de carga",
    iva: 0.16, retIsr: 0, retIvaSobreIva: 0.04 / 0.16,
    nota: "IVA 16%, retención de 4% de IVA" },
  { id: "sin_iva", label: "Exento / sin IVA",
    iva: 0, retIsr: 0, retIvaSobreIva: 0,
    nota: "Sin IVA ni retenciones" },
  { id: "manual", label: "Capturar a mano",
    iva: null, retIsr: null, retIvaSobreIva: null,
    nota: "Los montos se escriben uno por uno" },
];

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

/**
 * Calcula el desglose a partir del subtotal. Las retenciones se devuelven
 * en NEGATIVO, como aparecen en el formato que recibe Pagos.
 */
function calcularDesglose(subtotal, esquemaId) {
  const e = ESQUEMAS_FISCALES.find((x) => x.id === esquemaId) || ESQUEMAS_FISCALES[0];
  if (e.iva === null) return null; // captura manual
  const sub = Number(subtotal) || 0;
  const iva = r2(sub * e.iva);
  return {
    subtotal: r2(sub),
    iva,
    ret_isr: r2(-sub * e.retIsr),
    ret_iva: r2(-iva * e.retIvaSobreIva),
  };
}

const totalDesglose = (d) =>
  r2((Number(d.subtotal) || 0) + (Number(d.iva) || 0) + (Number(d.ret_isr) || 0) +
     (Number(d.ret_iva) || 0) - (Number(d.descuento) || 0));

/**
 * Deduce el subtotal a partir de un importe que ya trae impuestos.
 *
 * Las transacciones guardan el TOTAL, que es lo que se paga. Para llenar la
 * solicitud hace falta el camino inverso, y por eso el resultado se propone
 * en un diálogo en vez de darse por bueno: si el importe capturado no
 * correspondía exactamente a este esquema, el subtotal deducido no cuadra.
 */
function subtotalDesdeTotal(total, esquemaId) {
  const e = ESQUEMAS_FISCALES.find((x) => x.id === esquemaId) || ESQUEMAS_FISCALES[0];
  if (e.iva === null) return 0;
  const factor = 1 + e.iva - e.retIsr - e.iva * e.retIvaSobreIva;
  return factor ? r2((Number(total) || 0) / factor) : 0;
}

/** ¿La transacción está marcada como anticipo? El marcador va en Folio SAE. */
function esAnticipo(t) {
  return /ANTICIP/i.test(String(t?.folio_compra_sae || ""));
}

/* ----------------------------------------------------------------------
   IMPORTACION DESDE EXCEL (hojas RawData-*)
---------------------------------------------------------------------- */
const UNIDAD_KEYS = ["OSB", "CTM", "ISE"];

/**
 * Da formato al encabezado de un archivo de INTERCAMBIO (RawData, catálogos,
 * plantillas). No agrega filas de título a propósito: los parsers de
 * importación y el preparador leen el encabezado de la PRIMERA fila, así que
 * anteponer un título rompería la reimportación. Lo que sí se puede: fondo,
 * panel congelado y autofiltro, que no mueven nada de lugar.
 */
function formatearHojaDatos(ws, headerRow, nCols) {
  headerRow.eachCell((cell) => {
    cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
    cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
  });
  headerRow.height = 22;
  ws.views = [{ state: "frozen", ySplit: headerRow.number }];
  ws.autoFilter = { from: { row: headerRow.number, column: 1 }, to: { row: headerRow.number, column: nCols } };
}

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
      zona: findCol(headers, ["zona"]),
      folio: findExactCol(headers, ["id", "folio", "a partida", "partida", "no. partida"]),
    };

    let count = 0;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const concepto = col.concepto !== -1 ? row[col.concepto] : null;
      if (concepto === null || concepto === undefined || String(concepto).trim() === "") continue;

      const moneda = col.moneda !== -1 && row[col.moneda] ? String(row[col.moneda]).trim().toUpperCase() : "MXP";
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

      let rubroFila = (col.rubro !== -1 && row[col.rubro]) ? String(row[col.rubro]).trim() : "Otros";
      let categoriaFila = (col.categoria !== -1 && row[col.categoria]) ? String(row[col.categoria]).trim() : "";

      /* Validación contra el catálogo. Es la que faltaba: el formulario
         encadena rubro -> categoría desde siempre, pero la carga masiva
         aceptaba cualquier combinación, y por ahí entraron 117 partidas con
         rubros o categorías que no existen —entre ellas 36 con "Todos" en
         ambos campos, que es el marcador de proyecto corrido de columna.
         No se bloquea la importación: se marca la fila para que se vea en el
         preview y se decida. Rechazarla obligaría a rehacer el archivo por
         un dato que casi siempre es corregible después. */
      /* Se REPARA, no solo se avisa. Una combinación inválida entra a la base
         y luego queda fuera de todos los cortes del Dashboard sin que nada lo
         señale; ese fue el origen de las 117 partidas del diagnóstico.

         El criterio: se conserva el RUBRO, que es el eje que se lee, y se
         ajusta la categoría. Nunca al revés — inventar un rubro cambiaría de
         lugar el gasto en el reporte de Dirección.

         El valor original se guarda en `extra` para no perderlo. */
      let defRubro = RUBROS.find((r) => normHeader(r.rubro) === normHeader(rubroFila));
      const avisosClasif = [];
      const reparado = {};

      if (!defRubro) {
        // Si la categoría sí existe y pertenece a un solo rubro, ese rubro es
        // la corrección evidente. Si no, se manda a Otros/Diversos, que es
        // honesto: "no sabemos", en vez de un rubro inventado.
        const porCategoria = RUBROS.filter((r) => r.categorias.some((c) => normHeader(c) === normHeader(categoriaFila)));
        reparado.rubro_original = rubroFila;
        if (porCategoria.length === 1) {
          rubroFila = porCategoria[0].rubro;
          avisosClasif.push(`Rubro "${reparado.rubro_original}" no existe — se usó ${rubroFila}, que es donde vive "${categoriaFila}"`);
        } else {
          // Sin categoría, no con una etiqueta genérica: el hueco se ve en el
          // preview y pide corrección; "Diversos" parecería un dato resuelto.
          reparado.categoria_original = categoriaFila;
          rubroFila = "Otros";
          categoriaFila = "";
          avisosClasif.push(`Rubro "${reparado.rubro_original}" no existe en el catálogo — se usó Otros, sin categoría`);
        }
        defRubro = RUBROS.find((r) => r.rubro === rubroFila);
      }

      if (defRubro && !defRubro.categorias.some((c) => normHeader(c) === normHeader(categoriaFila))) {
        const dondeSiVa = RUBROS.filter((r) => r.categorias.some((c) => normHeader(c) === normHeader(categoriaFila)))
          .map((r) => r.rubro);
        reparado.categoria_original = categoriaFila;
        // Se sugiere por el concepto antes de caer en Diversos.
        const sugerida = sugerirCategoria(concepto, defRubro.categorias);
        categoriaFila = sugerida || "";
        const destino = categoriaFila || "sin categoría";
        avisosClasif.push(
          dondeSiVa.length
            ? `"${reparado.categoria_original}" no pertenece a ${rubroFila} (vive en ${dondeSiVa.join(" / ")}) — se usó ${destino}`
            : `La categoría "${reparado.categoria_original}" no existe — se usó ${destino}`
        );
      }

      rows.push({
        id: uid(),
        unidad,
        mes,
        anio,
        smi: (col.smi !== -1 && row[col.smi]) ? String(row[col.smi]).trim() : "",
        concepto: String(concepto).trim(),
        rubro: rubroFila,
        categoria: categoriaFila,
        _avisoClasificacion: avisosClasif.length ? avisosClasif.join(" · ") : null,
        zona: (col.zona !== -1 && row[col.zona]) ? String(row[col.zona]).trim() : "",
        proyecto: (col.proyecto !== -1 && row[col.proyecto]) ? String(row[col.proyecto]).trim() : "",
        monto_estimado: monto,
        moneda,
        folio,
        extra: { ...buildExtra(row, headers, promotedIdx), ...reparado },
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
  const porClave = new Map(); // "unidad|rfcOnombre" -> fila agregada (para agrupar varias cuentas del mismo proveedor en un solo archivo)
  const sheetsFound = [];
  let sinCompania = 0;

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
      referencia: findExactCol(headers, ["referencia"]),
      notas: findExactCol(headers, ["notas"]),
      banco: findExactCol(headers, ["banco"]),
      sucursal: findExactCol(headers, ["sucursal"]),
      swift: findExactCol(headers, ["swift"]),
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
      const cuenta = {
        banco: (col.banco !== -1 && row[col.banco]) ? String(row[col.banco]).trim() : "",
        sucursal: (col.sucursal !== -1 && row[col.sucursal]) ? String(row[col.sucursal]).trim() : "",
        swift: (col.swift !== -1 && row[col.swift]) ? String(row[col.swift]).trim() : "",
        clabe: (col.clabe !== -1 && row[col.clabe]) ? String(row[col.clabe]).trim() : "",
        numero_cuenta: (col.cuenta !== -1 && row[col.cuenta]) ? String(row[col.cuenta]).trim() : "",
        divisa: (col.divisa !== -1 && row[col.divisa]) ? String(row[col.divisa]).trim().toUpperCase() : "MXP",
      };
      const tieneCuenta = cuenta.banco || cuenta.clabe || cuenta.numero_cuenta || cuenta.sucursal || cuenta.swift;

      // Clave para agrupar varias filas del MISMO archivo que representan al
      // mismo proveedor (varias cuentas) — por RFC si lo trae, si no por Nombre.
      const clave = `${unidad}|${rfc || nombre.toUpperCase()}`;

      if (porClave.has(clave)) {
        if (tieneCuenta) porClave.get(clave)._cuentas.push(cuenta);
      } else {
        const existente = existingProveedores.find((p) => {
          if (p.unidad !== unidad) return false;
          if (rfc && p.rfc) return p.rfc.trim().toUpperCase() === rfc;
          if (!rfc && !p.rfc) return (p.nombre || "").trim().toUpperCase() === nombre.toUpperCase();
          return false;
        });
        porClave.set(clave, {
          id: uid(),
          unidad,
          nombre,
          rfc,
          id_sae: (col.idSae !== -1 && row[col.idSae]) ? String(row[col.idSae]).trim() : "",
          referencia: (col.referencia !== -1 && row[col.referencia]) ? String(row[col.referencia]).trim() : "",
          notas: (col.notas !== -1 && row[col.notas]) ? String(row[col.notas]).trim() : "",
          _existenteId: existente ? existente.id : null,
          _cuentas: tieneCuenta ? [cuenta] : [],
        });
      }
      count++;
    }
    if (count) sheetsFound.push({ sheetName, count });
  });

  const rows = [...porClave.values()];
  const nuevas = rows.filter((r) => !r._existenteId).length;
  const actualizaciones = rows.filter((r) => r._existenteId).length;

  const conAvisoClasif = rows.filter((r) => r._avisoClasificacion).length;
  return { rows, sheetsFound, nuevas, actualizaciones, sinCompania, conAvisoClasif };
}

function parseTransaccionesWorkbook(arrayBuffer, partidas, proveedores = [], cuentas = []) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const rows = [];
  const sheetsFound = [];
  let matched = 0, unmatched = 0, conCuenta = 0, sinCuenta = 0, desajustes = 0;

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
      smi: findExactCol(headers, ["smi"]),
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
      referenciaPago: findCol(headers, ["referencia", "pago"], ["referencia", "transferencia"], ["referencia bancaria"]),
      categoria: findCol(headers, ["categoria"]),
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

      const partidaPorFolio = folio ? partidas.find(
        (p) => p.folio && p.folio.trim().toUpperCase() === folio.toUpperCase() &&
               (!unidad_detectada || p.unidad === unidad_detectada)
      ) : null;

      // La carga masiva es por donde entran casi todas las transacciones, así
      // que el vínculo entre monedas distintas se rechaza aquí también. Sin
      // esto, bloquear solo el formulario sería inútil: el problema seguiría
      // entrando por la puerta grande. La fila se importa, pero SIN partida,
      // para que quede visible y se corrija en vez de perderse.
      const monNorm = (v) => ((v || "MXP") === "USD" ? "USD" : "MXP");
      const monedaDeLaFila = (col.moneda !== -1 && row[col.moneda]) ? String(row[col.moneda]).trim().toUpperCase() : "MXP";
      const desajusteMoneda = !!partidaPorFolio && monNorm(partidaPorFolio.moneda) !== monNorm(monedaDeLaFila);
      if (desajusteMoneda) desajustes++;
      const partida = desajusteMoneda ? null : partidaPorFolio;
      if (partida) matched++; else unmatched++;

      const proveedorNombre = (col.proveedor !== -1 && row[col.proveedor]) ? String(row[col.proveedor]).trim() : "";
      const proveedorMatch = proveedorNombre ? proveedores.find(
        (pv) => pv.nombre && pv.nombre.trim().toUpperCase() === proveedorNombre.toUpperCase() &&
                (!unidad_detectada || pv.unidad === unidad_detectada)
      ) : null;

      // La cuenta bancaria vive en el catálogo del proveedor, no en la
      // transacción. Si el proveedor tiene UNA sola cuenta en la divisa del
      // pago, se asigna sola; con varias se deja vacía a propósito, porque
      // elegir la equivocada manda el dinero a otro lado.
      const monedaFila = (col.moneda !== -1 && row[col.moneda]) ? String(row[col.moneda]).trim().toUpperCase() : "MXP";
      const cuentasDelProv = proveedorMatch
        ? cuentas.filter((c) => c.proveedor_id === proveedorMatch.id)
        : [];
      const cuentasDivisa = cuentasDelProv.filter((c) => (c.divisa || "MXP").toUpperCase() === monedaFila);
      const cuentaElegida = cuentasDivisa.length === 1 ? cuentasDivisa[0]
                          : (cuentasDelProv.length === 1 && !cuentasDivisa.length ? null : null);
      if (cuentaElegida) conCuenta++; else if (proveedorMatch) sinCuenta++;

      const conceptoDetallado = (col.concepto !== -1 && row[col.concepto]) ? String(row[col.concepto]).trim() : "";
      const categoriaDelArchivo = (col.categoria !== -1 && row[col.categoria]) ? String(row[col.categoria]).trim() : "";
      // Las opciones se acotan al rubro de la partida cuando la hay; si no,
      // se busca en todo el catálogo.
      const catsPosibles = partida
        ? (RUBROS.find((r) => r.rubro === partida.rubro)?.categorias || CATEGORIAS_TODAS)
        : CATEGORIAS_TODAS;
      const categoriaImportada = categoriaDelArchivo
        || sugerirCategoria(conceptoDetallado, catsPosibles)
        || "";

      rows.push({
        id: uid(),
        _desajusteMoneda: desajusteMoneda
          ? `${monNorm(monedaDeLaFila)} contra partida ${partidaPorFolio.folio} en ${monNorm(partidaPorFolio.moneda)}`
          : null,
        cuenta_id: cuentaElegida ? cuentaElegida.id : "",
        _cuentasDisponibles: cuentasDelProv.length,
        partida_id: partida ? partida.id : "",
        folio_original: folio,
        unidad_detectada: partida ? partida.unidad : unidad_detectada,
        dia: col.dia !== -1 ? toISODate(row[col.dia]) : "",
        solicitante: (col.solicitante !== -1 && row[col.solicitante]) ? String(row[col.solicitante]).trim() : "",
        smi: (col.smi !== -1 && row[col.smi]) ? String(row[col.smi]).trim() : "",
        proyecto: (col.proyecto !== -1 && row[col.proyecto]) ? String(row[col.proyecto]).trim() : "",
        zona: (col.zona !== -1 && row[col.zona]) ? String(row[col.zona]).trim() : "",
        area: (col.area !== -1 && row[col.area]) ? String(row[col.area]).trim() : "",
        proveedor: proveedorNombre,
        proveedor_id: proveedorMatch ? proveedorMatch.id : "",
        concepto_detallado: conceptoDetallado,
        importe,
        moneda: (col.moneda !== -1 && row[col.moneda]) ? String(row[col.moneda]).trim().toUpperCase() : "MXP",
        status: (col.status !== -1 && row[col.status]) ? String(row[col.status]).trim() : "",
        folio_compra_sae: (col.folioCompraSae !== -1 && row[col.folioCompraSae]) ? String(row[col.folioCompraSae]).trim() : "",
        folio_factura: (col.folioFactura !== -1 && row[col.folioFactura]) ? String(row[col.folioFactura]).trim() : "",
        forma_pago: (col.formaPago !== -1 && row[col.formaPago]) ? String(row[col.formaPago]).trim() : "",
        metodo_pago: (col.metodoPago !== -1 && row[col.metodoPago]) ? String(row[col.metodoPago]).trim() : "",
        referencia_pago: (col.referenciaPago !== -1 && row[col.referenciaPago]) ? String(row[col.referenciaPago]).trim() : "",
        // Si el archivo trae categoría se respeta; si no, se deduce del
        // concepto. Sin esto el campo entraría vacío en cada importación y el
        // dato que acabamos de poblar se degradaría desde el primer archivo.
        categoria: categoriaImportada,
        _categoriaSugerida: categoriaImportada && !categoriaDelArchivo,
      });
      count++;
    }
    if (count) sheetsFound.push({ sheetName, count });
  });

  return { rows, sheetsFound, matched, unmatched, conCuenta, sinCuenta, desajustes };
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

function Button({ children, variant = "primary", style, ...rest }) {
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
        ...style,
      }}
    >
      {children}
    </button>
  );
}

// Botón compacto de solo ícono, para filas de tabla con varias acciones donde
// el texto completo ("Editar"/"Duplicar"/"Eliminar") ya no cabe bien.
function IconButton({ icon, label, tone = T.textDim, ...rest }) {
  return (
    <button
      {...rest}
      title={label}
      aria-label={label}
      style={{
        width: 28, height: 28, borderRadius: 6, cursor: "pointer",
        background: "transparent", border: `1px solid ${tone}55`, color: tone,
        fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", padding: 0,
        flexShrink: 0,
      }}
    >
      {icon}
    </button>
  );
}

// Switch deslizante estilo iOS, para elegir entre dos opciones (ej. MXP/USD).
function SlidingToggle({ opciones, value, onChange }) {
  const idx = Math.max(0, opciones.indexOf(value));
  const anchoBoton = 68;
  return (
    <div style={{
      position: "relative", width: anchoBoton * opciones.length + 4, height: 30,
      background: T.panelAlt, borderRadius: 15, border: `1px solid ${T.border}`, display: "flex",
    }}>
      <div style={{
        position: "absolute", top: 2, left: 2 + idx * anchoBoton, width: anchoBoton - 2, height: 24,
        background: T.accent, borderRadius: 12, transition: "left 0.2s ease",
        boxShadow: "0 1px 3px rgba(35,42,49,0.2)",
      }} />
      {opciones.map((op) => (
        <button
          key={op}
          type="button"
          onClick={() => onChange(op)}
          style={{
            position: "relative", zIndex: 1, width: anchoBoton, border: "none", background: "transparent",
            fontSize: 11.5, fontWeight: 700, cursor: "pointer", fontFamily: T.fontUI, letterSpacing: "0.02em",
            color: value === op ? "#FFFFFF" : T.textDim, transition: "color 0.2s ease",
          }}
        >
          {op}
        </button>
      ))}
    </div>
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
  const transUnidad = transacciones.filter((t) => idsPartidas.has(t.partida_id) || t.unidad_detectada === unidad);

  /* -----------------------------------------------------------------
     FILTRO ÚNICO DEL DASHBOARD
     Vivía dentro del panel comparativo, pero gobierna también las gráficas
     de abajo, así que sube aquí. Antes había dos juegos de filtros en la
     misma pantalla —uno arriba y otro a media página— y no se veía cuál
     mandaba sobre qué.
     Cascada de tres niveles:
       1. Año     -> "YTD" o un año concreto con datos
       2. Periodo -> todo el año, o un rango de meses (no aplica en YTD)
       3. Desde/Hasta -> meses sueltos, porque el año ya quedó fijo arriba
     ----------------------------------------------------------------- */
  // Filtro de Proyecto. Va ANTES de `controlesFiltro`: ese const arma su JSX
  // al evaluarse, así que si la declaración quedara más abajo, el propio
  // render fallaría por zona muerta temporal y la pantalla saldría en blanco.
  const [proyectoKpi, setProyectoKpi] = useSessionState("ss-dashboard-proyecto", "Todos");

  const hoy = new Date();
  const anioActual = hoy.getFullYear();
  const mesActualIdx = hoy.getMonth();

  const aniosConDatos = [...new Set(partidasUnidad.map((p) => Number(p.anio)).filter(Boolean))].sort((a, b) => b - a);
  const hayYTD = partidasUnidad.some((p) => Number(p.anio) === anioActual && MESES.indexOf(p.mes) <= mesActualIdx);
  const opcionesAnio = [...(hayYTD ? ["YTD"] : []), ...aniosConDatos.map(String)];

  const [anioGuardado, setAnioSel] = useSessionState("ss-dashboard-f1-anio", null);
  const anioSel = opcionesAnio.includes(anioGuardado) ? anioGuardado : (opcionesAnio[0] ?? null);
  const esYTD = anioSel === "YTD";
  const anioEfectivo = esYTD ? anioActual : Number(anioSel);

  const [modoGuardado, setModo] = useSessionState("ss-dashboard-f2-periodo", "todo");
  const modo = esYTD ? "todo" : (modoGuardado === "rango" ? "rango" : "todo");

  // Solo se listan los meses que existen en datos para ese año, para que no se
  // pueda elegir un rango que garantiza tablas vacías.
  const mesesDelAnio = MESES.filter((m) => partidasUnidad.some((p) => Number(p.anio) === anioEfectivo && p.mes === m));
  const [desdeGuardado, setMesDesde] = useSessionState("ss-dashboard-f3-desde", null);
  const [hastaGuardado, setMesHasta] = useSessionState("ss-dashboard-f3-hasta", null);
  const mesDesde = mesesDelAnio.includes(desdeGuardado) ? desdeGuardado : (mesesDelAnio[0] ?? MESES[0]);
  const mesHasta = mesesDelAnio.includes(hastaGuardado) ? hastaGuardado : (mesesDelAnio[mesesDelAnio.length - 1] ?? MESES[11]);
  const cambiarDesde = (m) => { setMesDesde(m); if (MESES.indexOf(m) > MESES.indexOf(mesHasta)) setMesHasta(m); };
  const cambiarHasta = (m) => { setMesHasta(m); if (MESES.indexOf(m) < MESES.indexOf(mesDesde)) setMesDesde(m); };

  // Ventana efectiva de meses (índices inclusivos) que resulta de los 3 filtros.
  const [iDesde, iHasta] = esYTD
    ? [0, mesActualIdx]
    : modo === "rango"
      ? [MESES.indexOf(mesDesde), MESES.indexOf(mesHasta)]
      : [0, 11];

  const partidasRango = partidasUnidad.filter((p) => {
    if (Number(p.anio) !== anioEfectivo) return false;
    const i = MESES.indexOf(p.mes);
    return i >= iDesde && i <= iHasta;
  });
  const idsRango = new Set(partidasRango.map((p) => p.id));

  // Las gráficas de abajo ya consumían `partidasFiltradasMes`; ahora ese
  // nombre apunta al rango del filtro único, así que obedecen sin tocarlas.
  const partidasFiltradasMes = partidasRango;

  // Los controles se arman aquí y se le entregan al panel para que los pinte
  // en su cabecera. Se dibujan UNA sola vez aunque gobiernen todo el tablero.
  const controlesFiltro = (
    <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
      <Field label="Proyecto">
        <Select value={proyectoKpi} onChange={(e) => setProyectoKpi(e.target.value)} style={{ width: 170 }}>
          <option>Todos</option>
          {proyectosUnidad.map((p) => <option key={p.nombre}>{p.nombre}</option>)}
        </Select>
      </Field>
      <Field label="Año">
        <Select value={anioSel ?? ""} onChange={(e) => setAnioSel(e.target.value)} style={{ width: 110 }} disabled={!opcionesAnio.length}>
          {opcionesAnio.map((o) => <option key={o} value={o}>{o}</option>)}
        </Select>
      </Field>
      {!esYTD && (
        <Field label="Periodo">
          <Select value={modo} onChange={(e) => setModo(e.target.value)} style={{ width: 150 }}>
            <option value="todo">Todo</option>
            <option value="rango">Desde - Hasta</option>
          </Select>
        </Field>
      )}
      {!esYTD && modo === "rango" && (
        <>
          <Field label="Desde">
            <Select value={mesDesde} onChange={(e) => cambiarDesde(e.target.value)} style={{ width: 135 }}>
              {mesesDelAnio.map((m) => <option key={m} value={m}>{m}</option>)}
            </Select>
          </Field>
          <Field label="Hasta">
            <Select value={mesHasta} onChange={(e) => cambiarHasta(e.target.value)} style={{ width: 135 }}>
              {mesesDelAnio.map((m) => <option key={m} value={m}>{m}</option>)}
            </Select>
          </Field>
        </>
      )}
    </div>
  );

  const mesLabel = esYTD
    ? ` · YTD ${anioActual}`
    : modo === "rango"
      ? ` · ${mesDesde} a ${mesHasta} ${anioEfectivo}`
      : ` · ${anioEfectivo}`;


  const idsPartidasFiltradas = new Set(partidasFiltradasMes.map((p) => p.id));
  const transFiltradasMes = transUnidad.filter((t) => idsPartidasFiltradas.has(t.partida_id));

  // Igual que porProyecto, pero separado por moneda — nunca se suma MXP con USD.
  const porProyectoPorMoneda = useMemo(() => {
    const nuevoMapa = () => {
      const m = {};
      proyectosUnidad.forEach((p) => { m[p.nombre] = { proyecto: p.nombre, presupuestado: 0, ejecutado: 0, pagado: 0 }; });
      return m;
    };
    const mapas = { MXP: nuevoMapa(), USD: nuevoMapa() };

    partidasFiltradasMes.forEach((partida) => {
      const m = (partida.moneda || "MXP") === "USD" ? "USD" : "MXP";
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!mapas[m][proyecto]) mapas[m][proyecto] = { proyecto, presupuestado: 0, ejecutado: 0, pagado: 0 };
        mapas[m][proyecto].presupuestado += (Number(partida.monto_estimado) || 0) * fraccion;
      });
    });

    transFiltradasMes.forEach((t) => {
      const partida = partidasFiltradasMes.find((p) => p.id === t.partida_id);
      if (!partida) return;
      const m = (t.moneda || "MXP") === "USD" ? "USD" : "MXP";
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!mapas[m][proyecto]) mapas[m][proyecto] = { proyecto, presupuestado: 0, ejecutado: 0, pagado: 0 };
        mapas[m][proyecto].ejecutado += (Number(t.importe) || 0) * fraccion;
        if (t.status === "Pagado") mapas[m][proyecto].pagado += (Number(t.importe) || 0) * fraccion;
      });
    });

    return { MXP: Object.values(mapas.MXP), USD: Object.values(mapas.USD) };
  }, [proyectosUnidad, partidasFiltradasMes, transFiltradasMes]);
  const porProyecto = porProyectoPorMoneda.MXP; // usado por el pivot/gráficas existentes (solo MXP)
  const porProyectoUSD = porProyectoPorMoneda.USD;

  // La tendencia mensual siempre usa TODOS los meses (sin filtrar), es la única
  // vista pensada justo para comparar entre meses. Separada por moneda, para
  // nunca sumar MXP con USD en la misma línea.
  const porMesPorMoneda = useMemo(() => {
    const nuevoMapa = () => {
      const m = {};
      MESES.forEach((mes) => { m[mes] = { mes, presupuestado: 0, ejecutado: 0 }; });
      return m;
    };
    const mapas = { MXP: nuevoMapa(), USD: nuevoMapa() };
    partidasUnidad.forEach((p) => {
      const m = (p.moneda || "MXP") === "USD" ? "USD" : "MXP";
      if (mapas[m][p.mes]) mapas[m][p.mes].presupuestado += Number(p.monto_estimado) || 0;
    });
    transUnidad.forEach((t) => {
      const partida = partidasUnidad.find((p) => p.id === t.partida_id);
      if (!partida) return;
      const m = (t.moneda || "MXP") === "USD" ? "USD" : "MXP";
      if (mapas[m][partida.mes]) mapas[m][partida.mes].ejecutado += Number(t.importe) || 0;
    });
    return {
      MXP: Object.values(mapas.MXP).filter((m) => m.presupuestado > 0 || m.ejecutado > 0),
      USD: Object.values(mapas.USD).filter((m) => m.presupuestado > 0 || m.ejecutado > 0),
    };
  }, [partidasUnidad, transUnidad]);


  // Si el proyecto guardado no existe en ESTA compañía (ej. veníamos de otra
  // unidad), se regresa solo a "Todos" en vez de quedarse "huérfano" en $0.00.
  useEffect(() => {
    if (proyectoKpi !== "Todos" && !proyectosUnidad.some((p) => p.nombre === proyectoKpi)) {
      setProyectoKpi("Todos");
    }
  }, [unidad, proyectosUnidad.map((p) => p.nombre).join(","), proyectoKpi]);

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
      <ResumenComparativoPanel
        partidasRango={partidasRango}
        idsRango={idsRango}
        controlesFiltro={controlesFiltro}
        transacciones={transUnidad}
        proyectosUnidad={proyectosUnidad}
        proyectoKpi={proyectoKpi}
        setProyectoKpi={setProyectoKpi}
      />

      <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
        <Panel
          title="Presupuesto vs. ejecutado por proyecto"
          subtitle={`Gastos compartidos ya prorrateados según su marcador — obedece los filtros de arriba${mesLabel}`}
        >
          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.08em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>MXP</div>
          <ResponsiveContainer width="100%" height={Math.max(180, porProyecto.length * 34)}>
            <BarChart data={porProyecto} layout="vertical" margin={{ left: 8, right: 16 }}>
              <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} horizontal={false} />
              <XAxis type="number" tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
              <YAxis type="category" dataKey="proyecto" width={110} tick={{ fill: T.textDim, fontSize: 11 }} stroke={T.border} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
              {/* Tres barras, no dos. "Ejecutado" suma TODAS las transacciones
                  vinculadas, pagadas o no, y sin la tercera barra se lee como
                  si ya se hubiera pagado. La diferencia entre ellas es
                  justamente lo comprometido y pendiente. */}
              <Bar dataKey="presupuestado" name="Presupuestado" fill={T.border} radius={[0,3,3,0]} />
              <Bar dataKey="ejecutado" name="Comprometido" fill={T.accent} radius={[0,3,3,0]} />
              <Bar dataKey="pagado" name="Pagado" fill={T.teal} radius={[0,3,3,0]} />
            </BarChart>
          </ResponsiveContainer>

          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.08em", textTransform: "uppercase", fontFamily: T.fontMono, margin: "18px 0 4px" }}>USD</div>
          {porProyectoUSD.some((p) => p.presupuestado || p.ejecutado) ? (
            <ResponsiveContainer width="100%" height={Math.max(180, porProyectoUSD.length * 34)}>
              <BarChart data={porProyectoUSD} layout="vertical" margin={{ left: 8, right: 16 }}>
                <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} horizontal={false} />
                <XAxis type="number" tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
                <YAxis type="category" dataKey="proyecto" width={110} tick={{ fill: T.textDim, fontSize: 11 }} stroke={T.border} />
                <Tooltip content={<ChartTooltip />} />
                <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
                <Bar dataKey="presupuestado" name="Presupuestado" fill={T.border} radius={[0,3,3,0]} />
                <Bar dataKey="ejecutado" name="Comprometido" fill={T.accent} radius={[0,3,3,0]} />
                <Bar dataKey="pagado" name="Pagado" fill={T.teal} radius={[0,3,3,0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div style={{ textAlign: "center", color: T.textFaint, fontSize: 12, padding: 16 }}>Sin datos en USD para estos filtros</div>
          )}
        </Panel>
      </div>

      {(porMesPorMoneda.MXP.length > 1 || porMesPorMoneda.USD.length > 1) && (
        <Panel title="Tendencia mensual" subtitle="Presupuestado vs. ejecutado">
          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.08em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>MXP</div>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={porMesPorMoneda.MXP}>
              <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} />
              <XAxis dataKey="mes" tick={{ fill: T.textFaint, fontSize: 10 }} stroke={T.border} />
              <YAxis tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
              <Line type="monotone" dataKey="presupuestado" name="Presupuestado" stroke={T.textFaint} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="ejecutado" name="Ejecutado" stroke={T.accent} strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>

          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.08em", textTransform: "uppercase", fontFamily: T.fontMono, margin: "18px 0 4px" }}>USD</div>
          {porMesPorMoneda.USD.length > 0 ? (
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={porMesPorMoneda.USD}>
                <CartesianGrid strokeDasharray="2 4" stroke={T.borderSoft} />
                <XAxis dataKey="mes" tick={{ fill: T.textFaint, fontSize: 10 }} stroke={T.border} />
                <YAxis tick={{ fill: T.textFaint, fontSize: 10 }} tickFormatter={(v) => `$${(v/1000).toFixed(0)}k`} stroke={T.border} />
                <Tooltip content={<ChartTooltip />} />
                <Legend wrapperStyle={{ fontSize: 11, color: T.textDim }} />
                <Line type="monotone" dataKey="presupuestado" name="Presupuestado" stroke={T.textFaint} strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="ejecutado" name="Ejecutado" stroke={T.teal} strokeWidth={2} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          ) : (
            <div style={{ textAlign: "center", color: T.textFaint, fontSize: 12, padding: 16 }}>Sin datos en USD para estos filtros</div>
          )}
        </Panel>
      )}

    </div>
  );
}

// Combina "Presupuestado" y "Pagado real" en un solo marco, con una barra de
// filtros PROPIA (Proyecto/Desde/Hasta/Moneda). El rango se expresa en periodos
// completos (mes + año), no en meses sueltos: con datos de 2025 y 2026 en la
// misma base, "de Agosto a Octubre" sería ambiguo, y un rango que cruza
// diciembre es imposible de expresar con una lista de meses.
/**
 * Recibe el rango ya resuelto en lugar de calcularlo. Los filtros de este
 * panel gobiernan TODO el Dashboard —también las gráficas de abajo—, así que
 * el estado vive arriba y aquí solo se dibujan los controles.
 */
function ResumenComparativoPanel({
  partidasRango, idsRango, transacciones, proyectosUnidad,
  proyectoKpi, setProyectoKpi, controlesFiltro,
}) {
  const partidaDe = (t) => partidasRango.find((p) => p.id === t.partida_id);
  const monedaDe = (x) => (x.moneda || "MXP") === "USD" ? "USD" : "MXP";

  // Ya NO se filtra por moneda: las dos conviven en la misma tabla. La
  // moneda entra como primer nivel de agrupación, arriba del proyecto, para
  // que cada rama tenga su propio subtotal y ninguna suma cruce monedas.
  // Duplicar las columnas por mes era la otra vía, pero llevaba el panel de
  // 9 a 17 columnas de dinero — ilegible de un vistazo, que es justo para lo
  // que se usa.
  const partidasMoneda = partidasRango;
  const pagadasMoneda = transacciones.filter((t) => t.status === "Pagado" && idsRango.has(t.partida_id));

  const aniosPresentes = [...new Set([
    ...partidasMoneda.map((p) => p.anio),
    ...pagadasMoneda.map((t) => partidaDe(t)?.anio),
  ].filter(Boolean))];
  const multiAnio = aniosPresentes.length > 1;

  const etiquetaColumna = (mes, anio) => (multiAnio && anio) ? `${mes} ${anio}` : mes;
  const columnaDePptado = (p) => etiquetaColumna(p.mes, p.anio);
  const columnaDePagado = (t) => { const p = partidaDe(t); return p ? etiquetaColumna(p.mes, p.anio) : null; };

  const columnas = [...new Set([
    ...partidasMoneda.map(columnaDePptado),
    ...pagadasMoneda.map(columnaDePagado),
  ])].filter(Boolean).sort((a, b) => {
    const [mesA, anioA] = a.split(" "); const [mesB, anioB] = b.split(" ");
    if ((anioA || "") !== (anioB || "")) return (anioA || "").localeCompare(anioB || "");
    return MESES.indexOf(mesA) - MESES.indexOf(mesB);
  });

  const [collapsedPptado, setCollapsedPptado] = useSessionSetState("ss-dashboard-collapsed-pptado");
  const [collapsedPagado, setCollapsedPagado] = useSessionSetState("ss-dashboard-collapsed-pagado");
  const togglePptado = (path) => setCollapsedPptado((prev) => { const n = new Set(prev); n.has(path) ? n.delete(path) : n.add(path); return n; });
  const togglePagado = (path) => setCollapsedPagado((prev) => { const n = new Set(prev); n.has(path) ? n.delete(path) : n.add(path); return n; });

  let partidasResueltas = partidasMoneda.flatMap((p) =>
    resolverProrrateo(p.proyecto, proyectosUnidad).map(({ proyecto, fraccion }) => ({
      ...p, proyecto, moneda: monedaDe(p),
      monto_estimado: (Number(p.monto_estimado) || 0) * fraccion, _columna: columnaDePptado(p),
    }))
  );
  if (proyectoKpi !== "Todos") partidasResueltas = partidasResueltas.filter((p) => p.proyecto === proyectoKpi);

  let filasPagadoResueltas = pagadasMoneda.flatMap((t) => {
    const p = partidaDe(t);
    return resolverProrrateo(p.proyecto, proyectosUnidad).map(({ proyecto, fraccion }) => ({
      proyecto, rubro: p.rubro, concepto: t.concepto_detallado || t.proveedor || "—",
      // Se copian de la partida para que las dos tablas puedan agruparse por
      // los mismos campos; la transacción no los trae por su cuenta.
      categoria: p.categoria, zona: p.zona || t.zona || "",
      moneda: monedaDe(t),
      importe: (Number(t.importe) || 0) * fraccion, _columna: columnaDePagado(t),
    }));
  });
  if (proyectoKpi !== "Todos") filasPagadoResueltas = filasPagadoResueltas.filter((f) => f.proyecto === proyectoKpi);

  /* El agrupamiento es configurable, pero la MONEDA va fija como primer
     nivel: si no, los pesos y los dólares se mezclarían dentro de un mismo
     grupo y sus subtotales sumarían monedas distintas.

     El mismo agrupamiento gobierna las DOS tablas a propósito. Comparten el
     eje de columnas para poder leerse una contra otra, y eso solo funciona
     si sus renglones también coinciden. */
  const [groupBysPanel, setGroupBysPanel] = usePrefState(
    "pref-dashboard-panel-groupbys",
    [{ field: "proyecto" }, { field: "rubro" }, { field: "concepto" }],
    sanearGroupBys(GROUP_OPCIONES_PANEL)
  );
  const AGRUPACION = ["moneda", ...groupBysPanel.map((g) => g.field)];
  // La moneda es el primer nivel, así que la ruta siempre empieza por ella:
  // "/moneda:USD/proyecto:Desh Gral/..."
  const monedaDeRuta = (ruta) => (String(ruta).match(/^\/moneda:(MXP|USD)\b/) || [])[1] || "MXP";
  const pivotPptado = pivotearPorMes(partidasResueltas, AGRUPACION, columnas, "monto_estimado", "_columna");
  const pivotPagado = pivotearPorMes(filasPagadoResueltas, AGRUPACION, columnas, "importe", "_columna");

  const rutasPptado = collectGroupPaths(pivotPptado);
  const rutasPagado = collectGroupPaths(pivotPagado);
  const todoContraido =
    (rutasPptado.length + rutasPagado.length) > 0 &&
    rutasPptado.every((r) => collapsedPptado.has(r)) &&
    rutasPagado.every((r) => collapsedPagado.has(r));
  const toggleTodo = () => {
    if (todoContraido) {
      setCollapsedPptado(new Set());
      setCollapsedPagado(new Set());
    } else {
      setCollapsedPptado(new Set(rutasPptado));
      setCollapsedPagado(new Set(rutasPagado));
    }
  };

  return (
    <Panel
      title="Presupuestado vs. pagado real, por proyecto"
      subtitle="Pesos y dólares en la misma tabla, agrupados por moneda — no se suman entre sí"
      right={
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 8 }}>
            <GroupByControl
              options={GROUP_OPCIONES_PANEL}
              value={groupBysPanel}
              onChange={setGroupBysPanel}
              maxLevels={3}
            />
            <Button variant="ghost" onClick={toggleTodo} style={{ height: 34 }}>
              {todoContraido ? "Expandir todo" : "Contraer todo"}
            </Button>
          </div>
          {controlesFiltro}
        </div>
      }
    >
      {[
        { titulo: "Presupuesto", pivot: pivotPptado, colapsadas: collapsedPptado,
          toggle: togglePptado, vacio: !partidasResueltas.length },
        { titulo: "Pagado real (Status = Pagado)", pivot: pivotPagado, colapsadas: collapsedPagado,
          toggle: togglePagado, vacio: !filasPagadoResueltas.length },
      ].map((t, i) => (
        <div key={t.titulo} style={{ marginBottom: i === 0 ? 28 : 0 }}>
          <div style={{ marginBottom: 8 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{t.titulo}</div>
          </div>
          {t.vacio ? (
            <div style={{ textAlign: "center", color: T.textFaint, fontSize: 12.5, padding: 24 }}>
              Sin movimientos para estos filtros
            </div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table style={tableStyle}>
                <thead>
                  <tr>
                    <th style={thStyle}>
                      {/* Refleja el agrupamiento vigente en vez de un texto fijo:
                          si se agrupa por Zona, decir "Proyecto" sería mentira. */}
                      {["Moneda", ...groupBysPanel.map((gb) =>
                        (GROUP_OPCIONES_PANEL.find((o) => o.value === gb.field) || {}).label || gb.field
                      )].join(" / ")}
                    </th>
                    {columnas.map((m) => <th key={m} style={{ ...thStyle, textAlign: "right" }}>{m}</th>)}
                    <th style={{ ...thStyle, textAlign: "right" }}>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {/* No hay filas de "Total" al final: la fila de cada moneda
                      YA es su total —buildPivotTrs suma cada grupo— y queda
                      justo encima de su propia lista, que es donde sirve.
                      Ponerlas otra vez abajo duplicaba las cifras y alejaba
                      el total de los datos que resume. */}
                  {buildPivotTrs(t.pivot, "", t.colapsadas, t.toggle, columnas, 0, monedaDeRuta)}
                </tbody>
              </table>
            </div>
          )}
        </div>
      ))}
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

// Barra de avisos fija en la parte inferior del navegador — se queda visible
// aunque se haga scroll. `avisos` = [{ tone, texto }].
function AvisosFlotantes({ avisos }) {
  const activos = avisos.filter((a) => a.texto);
  if (!activos.length) return null;
  const toneColor = { amber: T.amberDim, red: T.red, accent: T.accent, teal: T.teal };
  return (
    <div style={{
      position: "fixed", bottom: 0, left: 0, right: 0, zIndex: 900,
      background: T.panel, borderTop: `3px solid ${toneColor[activos[0].tone] || T.amberDim}`,
      boxShadow: "0 -6px 18px rgba(35,42,49,0.14)", padding: "10px 20px",
      display: "flex", flexDirection: "column", gap: 4, maxHeight: 140, overflowY: "auto",
    }}>
      {activos.map((a, i) => (
        <div key={i} style={{ fontSize: 11.5, color: toneColor[a.tone] || T.amberDim }}>
          {a.texto}
        </div>
      ))}
    </div>
  );
}

// Leyenda de auditoría ("Creado por X · Editado por Y") para modales de edición.
// No muestra nada si el registro es nuevo, o si no hay datos de autoría aún
// (registros capturados antes de activar usuarios).
// Formatea un timestamp de Supabase (ISO) a algo legible en es-MX.
function formatFechaHora(iso) {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString("es-MX", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
  } catch {
    return "";
  }
}

// Versión corta (sin año, sin "p.m." completo) para celdas angostas donde el
// formato completo se encima con la columna de al lado.
function formatFechaCorta(iso) {
  if (!iso) return "";
  try {
    const d = new Date(iso);
    const fecha = d.toLocaleDateString("es-MX", { day: "2-digit", month: "short" });
    const hora = d.toLocaleTimeString("es-MX", { hour: "numeric", minute: "2-digit", hour12: true }).replace(/\s?[ap]\.?\s?m\.?/i, (m) => m.trim()[0].toLowerCase());
    return `${fecha}, ${hora}`;
  } catch {
    return "";
  }
}

function AutoriaCaption({ record, perfilesApi }) {
  if (!record?.id) return null;
  const nombreDe = (uid) => perfilesApi.rows.find((p) => p.id === uid)?.nombre || null;
  const creador = record.created_by ? nombreDe(record.created_by) : null;
  const editor = record.updated_by ? nombreDe(record.updated_by) : null;
  const actualizado = formatFechaHora(record.updated_at);
  if (!creador && !editor && !actualizado) return null;
  return (
    <div style={{ fontSize: 10.5, color: T.textFaint, marginBottom: 12 }}>
      {creador && `Creado por ${creador}`}
      {creador && editor && editor !== creador ? ` · Última edición por ${editor}` : ""}
      {actualizado && ` · Última actualización: ${actualizado}`}
    </div>
  );
}

// Como useState, pero respaldado en sessionStorage — sobrevive cambios de
// pestaña/vista mientras la sesión del navegador siga abierta (se pierde al
// cerrar la pestaña o el navegador, a propósito).
function useSessionState(key, defaultValue) {
  const [value, setValue] = useState(() => {
    try {
      const raw = sessionStorage.getItem(key);
      return raw !== null ? JSON.parse(raw) : defaultValue;
    } catch {
      return defaultValue;
    }
  });
  useEffect(() => {
    try { sessionStorage.setItem(key, JSON.stringify(value)); } catch {}
  }, [key, value]);
  return [value, setValue];
}

/**
 * Como useSessionState, pero en localStorage: sobrevive a cerrar la pestaña.
 *
 * Se reserva para PREFERENCIAS DE TRABAJO —agrupamientos, orden de columnas—
 * no para filtros. Un filtro persistido es peligroso: la app abriría mostrando
 * el periodo que quedó de la última vez y esa cifra se leería como si fuera la
 * actual. El acomodo visual, en cambio, no cambia lo que los números dicen.
 *
 * `sanear` limpia lo guardado antes de usarlo. Hace falta porque una
 * preferencia apunta a campos POR NOMBRE, y los campos cambian: al retirar
 * `es_recurrente` cualquier orden guardado que lo usara habría quedado
 * apuntando al vacío.
 */
function usePrefState(key, defaultValue, sanear) {
  const [value, setValue] = useState(() => {
    try {
      const raw = localStorage.getItem(key);
      if (raw === null) return defaultValue;
      const guardado = JSON.parse(raw);
      return sanear ? sanear(guardado, defaultValue) : guardado;
    } catch {
      return defaultValue;
    }
  });
  useEffect(() => {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch {}
  }, [key, value]);
  return [value, setValue];
}

/** Descarta niveles de agrupamiento cuyo campo ya no existe en el catálogo. */
const sanearGroupBys = (opciones) => (guardado, porDefecto) => {
  if (!Array.isArray(guardado)) return porDefecto;
  const validos = new Set(opciones.map((o) => o.value));
  return guardado.filter((g) => g && validos.has(g.field));
};

/** Si la columna de ordenamiento desapareció, se regresa al orden por defecto. */
const sanearSort = (claves) => (guardado, porDefecto) => {
  if (!guardado || typeof guardado !== "object") return porDefecto;
  if (guardado.key && !claves.includes(guardado.key)) return porDefecto;
  return guardado;
};

// Igual que useSessionState, pero para un Set (ej. filas colapsadas) — un Set
// no se guarda directo como JSON, así que se convierte a/desde arreglo por dentro.
function useSessionSetState(key, defaultArray = []) {
  const [arr, setArr] = useSessionState(key, defaultArray);
  const set = new Set(arr);
  const setSet = (next) => {
    setArr((prevArr) => {
      const prevSet = new Set(prevArr);
      const nextSet = typeof next === "function" ? next(prevSet) : next;
      return [...nextSet];
    });
  };
  return [set, setSet];
}

function EmptyState({ title, body }) {
  return (
    <div style={{ ...panelStyle, textAlign: "center", padding: "48px 24px" }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: T.text, marginBottom: 6 }}>{title}</div>
      <div style={{ fontSize: 12.5, color: T.textDim, maxWidth: 420, margin: "0 auto" }}>{body}</div>
    </div>
  );
}

// Recuerda qué columnas están ocultas (por tabla, via localStorage).
function useColumnVisibility(storageKey, columns) {
  const [hidden, setHidden] = useState(() => {
    try { return new Set(JSON.parse(localStorage.getItem(storageKey) || "[]")); } catch { return new Set(); }
  });
  const toggle = (key) => {
    setHidden((h) => {
      const next = new Set(h);
      next.has(key) ? next.delete(key) : next.add(key);
      try { localStorage.setItem(storageKey, JSON.stringify([...next])); } catch {}
      return next;
    });
  };
  const showAll = () => {
    setHidden(new Set());
    try { localStorage.setItem(storageKey, JSON.stringify([])); } catch {}
  };
  const visible = columns.filter((c) => !hidden.has(c.key));
  return { visible, hidden, toggle, showAll };
}

// Botón "Columnas" con un panel de checkboxes para mostrar/ocultar cada una.
function ColumnVisibilityControl({ columns, hidden, onToggle, onShowAll, etiqueta = "Columnas" }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const onClickFuera = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClickFuera);
    return () => document.removeEventListener("mousedown", onClickFuera);
  }, []);
  const ocultas = columns.filter((c) => hidden.has(c.key)).length;

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <Button variant={ocultas ? "primary" : "ghost"} onClick={() => setOpen((o) => !o)}>
        {etiqueta}{ocultas ? ` (${columns.length - ocultas}/${columns.length})` : ""}
      </Button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 50,
          background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8,
          padding: 14, minWidth: 220, maxHeight: 360, overflowY: "auto",
          boxShadow: "0 8px 24px rgba(35,42,49,0.14)",
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: T.text }}>{etiqueta === "Columnas" ? "Columnas visibles" : etiqueta}</span>
            {ocultas > 0 && (
              <button
                type="button"
                onClick={onShowAll}
                style={{ background: "none", border: "none", color: T.accent, fontSize: 11, cursor: "pointer", padding: 0 }}
              >
                Mostrar todas
              </button>
            )}
          </div>
          {columns.map((c) => (
            <label key={c.key} style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", fontSize: 12.5, color: T.text, cursor: "pointer" }}>
              {/* Las columnas fijas se muestran marcadas y deshabilitadas: se ve
                  que están, y se ve que no son opcionales. */}
              <input type="checkbox" checked={!hidden.has(c.key)} disabled={c.fija} onChange={() => onToggle(c.key)} />
              {/* Las columnas de la tabla usan `label` y las de exportación `header`.
                  Se aceptan ambos: al leer solo uno, el otro juego de columnas
                  salía con las casillas en blanco, sin nombre. */}
              <span style={{ color: c.fija ? T.textFaint : T.text }}>
                {c.label || c.header || c.key}{c.fija ? " (siempre)" : ""}
              </span>
            </label>
          ))}
        </div>
      )}
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


// Selector de casillas múltiples genérico — se usa tanto para Mes como para Año.
// `opciones` = lista de valores disponibles; `todosLabel` = texto cuando no hay
// nada seleccionado (equivale a "todos").
function MultiSelect({ opciones, seleccionados, onChange, todosLabel, unidadLabel = "" }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const onClickFuera = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClickFuera);
    return () => document.removeEventListener("mousedown", onClickFuera);
  }, []);

  const label = seleccionados.length === 0 ? todosLabel : seleccionados.length === 1 ? String(seleccionados[0]) : `${seleccionados.length}${unidadLabel}`;
  const toggleValor = (v) => onChange(seleccionados.includes(v) ? seleccionados.filter((x) => x !== v) : [...seleccionados, v]);

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
            {todosLabel}
          </label>
          {opciones.map((v) => (
            <label key={v} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: T.text, padding: "5px 4px", cursor: "pointer" }}>
              <input type="checkbox" checked={seleccionados.includes(v)} onChange={() => toggleValor(v)} />
              {v}
            </label>
          ))}
          <Button variant="ghost" onClick={() => setOpen(false)} style={{ width: "100%", marginTop: 8 }}>Cerrar</Button>
        </div>
      )}
    </div>
  );
}

function MesMultiSelect({ mesesDisponibles, seleccionados, onChange }) {
  return <MultiSelect opciones={mesesDisponibles} seleccionados={seleccionados} onChange={onChange} todosLabel="Todos los meses" unidadLabel=" meses" />;
}

function AnioMultiSelect({ aniosDisponibles, seleccionados, onChange }) {
  return <MultiSelect opciones={aniosDisponibles} seleccionados={seleccionados} onChange={onChange} todosLabel="Todos los años" unidadLabel=" años" />;
}



function Modal({ title, subtitle, onClose, children, width = 720, zIndex = 1000 }) {
  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  return (
    <div
      style={{
        position: "fixed", inset: 0, background: "rgba(6,9,10,0.72)", zIndex,
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
// lista de columnas visibles según esa preferencia. Las columnas nuevas — las
// que aún no existían cuando se guardó la preferencia — se insertan en su lugar
// natural (justo después de la que las precede en la definición), no al final:
// si se agregaran al final, una columna nueva aparecería escondida al extremo
// derecho de la tabla y parecería que no se agregó.
function useColumnOrder(storageKey, columns) {
  const [order, setOrder] = useState(() => {
    try { return JSON.parse(localStorage.getItem(storageKey) || "[]"); } catch { return []; }
  });
  const byKey = new Map(columns.map((c) => [c.key, c]));
  const ordered = order.filter((k) => byKey.has(k)).map((k) => byKey.get(k));
  columns.forEach((c, i) => {
    if (order.includes(c.key)) return;
    const previa = columns.slice(0, i).reverse().find((x) => ordered.some((o) => o.key === x.key));
    const at = previa ? ordered.findIndex((o) => o.key === previa.key) + 1 : ordered.length;
    ordered.splice(at, 0, c);
  });

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

// Suma un campo de dinero sobre una lista de filas, separado por moneda (para
// no mezclar pesos con dólares) — para mostrarlo en el encabezado de la columna.
function sumaPorMoneda(rows, montoKey, monedaKey = "moneda") {
  const totals = {};
  rows.forEach((r) => {
    const m = r[monedaKey] || "MXP";
    const v = Number(r[montoKey]) || 0;
    if (!v) return;
    totals[m] = (totals[m] || 0) + v;
  });
  return Object.entries(totals).map(([m, v]) => money(v, m)); // una línea por moneda
}

function SortableTh({ label, sortKey, sort, setSort, width, onResizeStart, onDragStart, onDragOver, onDrop, sumLabel }) {
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
      <div>{label}{active ? (sort.dir === "asc" ? " ▲" : " ▼") : ""}</div>
      {sumLabel && sumLabel.length > 0 && (
        <div style={{ fontSize: 9.5, fontWeight: 400, textTransform: "none", letterSpacing: "normal", color: T.teal, fontFamily: T.fontMono, marginTop: 2 }}>
          {sumLabel.map((linea, i) => <div key={i} style={{ whiteSpace: "nowrap" }}>{i === 0 ? "Σ " : ""}{linea}</div>)}
        </div>
      )}
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
  const esCampoMes = key === "mes" || key === "_mes";
  const anioDe = (r) => Number(key === "mes" ? r.anio : r._anio) || null;

  const buckets = new Map();
  rows.forEach((r) => {
    let val;
    if (esCampoMes) {
      const mes = (r[key] ?? "").toString().trim();
      const anio = anioDe(r);
      // "Noviembre 2025" y "Noviembre 2026" son grupos distintos, no el mismo —
      // si no se conoce el año, se queda como antes (solo el nombre del mes).
      val = mes ? (anio ? `${mes} ${anio}` : mes) : SIN_DATO;
    } else {
      val = (r[key] ?? "").toString().trim() || SIN_DATO;
    }
    if (!buckets.has(val)) buckets.set(val, []);
    buckets.get(val).push(r);
  });
  let entries = [...buckets.entries()];
  if (esCampoMes) {
    // Ordena por año real y luego por posición del mes en el calendario —
    // no alfabéticamente ni solo por nombre de mes.
    entries.sort((a, b) => {
      const anioA = anioDe(a[1][0]) || 0;
      const anioB = anioDe(b[1][0]) || 0;
      if (anioA !== anioB) return anioA - anioB;
      const mesA = a[1][0][key] ?? "";
      const mesB = b[1][0][key] ?? "";
      return MESES.indexOf(mesA) - MESES.indexOf(mesB);
    });
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
function pivotearPorMes(rows, keys, meses, montoKey = "monto_estimado", colKey = "mes") {
  const mesSums = {};
  meses.forEach((m) => { mesSums[m] = 0; });
  let total = 0;
  rows.forEach((r) => {
    const v = Number(r[montoKey]) || 0;
    if (mesSums[r[colKey]] !== undefined) mesSums[r[colKey]] += v;
    total += v;
  });
  // Se conservan las filas: sin ellas, el último nivel del agrupamiento no
  // tiene qué desplegar y sus registros quedan inalcanzables.
  if (!keys.length) return { type: "rows", rows, mesSums, total };
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
    entries: entries.map(([value, groupRows]) => ({ value, child: pivotearPorMes(groupRows, rest, meses, montoKey, colKey) })),
  };
}

// Renders a pivotearPorMes tree as <tr> rows: bold collapsible rows for every level except
// the deepest, which renders as a plain (non-collapsible) leaf row.
/**
 * `resolverMoneda` recibe la ruta del grupo y devuelve la moneda con la que
 * se formatean sus cifras. Se usa cuando la moneda es un nivel de
 * agrupación: sin esto, los dólares se imprimirían con formato de pesos y
 * las dos ramas se verían idénticas.
 */
function buildPivotTrs(node, path, collapsed, toggleGroup, meses, depth, resolverMoneda) {
  let out = [];
  node.entries.forEach((entry) => {
    const groupPath = `${path}/${node.key}:${entry.value}`;
    const esHoja = entry.child.type === "rows";
    const mon = resolverMoneda ? resolverMoneda(groupPath) : undefined;
    const cellStyle = { ...tdStyle, fontFamily: T.fontMono, textAlign: "right" };
    if (esHoja) {
      // El último nivel del agrupamiento también se puede desplegar, para ver
      // los registros que contiene. Antes se dibujaba como hoja muerta: con
      // "Concepto" al final no se notaba —cada hoja era un registro— pero al
      // agrupar solo por Rubro, sus partidas quedaban inalcanzables.
      const filas = entry.child.rows || [];
      const puedeAbrir = filas.length > 1 || (filas.length === 1 && String(filas[0].concepto || "") !== String(entry.value));
      const isCollapsed = collapsed.has(groupPath);
      out.push(
        <tr
          key={groupPath}
          onClick={puedeAbrir ? () => toggleGroup(groupPath) : undefined}
          style={puedeAbrir ? { cursor: "pointer" } : undefined}
        >
          <td style={{ ...tdStyle, paddingLeft: 14 + depth * 22, display: "flex", alignItems: "center", gap: 8 }}>
            {puedeAbrir
              ? <span style={{ color: T.textFaint, fontSize: 10 }}>{isCollapsed ? "▶" : "▼"}</span>
              : <span style={{ width: 10 }} />}
            {entry.value}
            {puedeAbrir && <span style={{ color: T.textFaint, fontSize: 10.5 }}>({filas.length})</span>}
          </td>
          {meses.map((m) => <td key={m} style={cellStyle}>{entry.child.mesSums[m] ? money(entry.child.mesSums[m], mon) : "—"}</td>)}
          <td style={{ ...cellStyle, fontWeight: 600 }}>{money(entry.child.total, mon)}</td>
        </tr>
      );
      if (puedeAbrir && !isCollapsed) {
        filas.forEach((r, i) => {
          const etiqueta = r.concepto || r.proveedor || "—";
          out.push(
            <tr key={`${groupPath}#${i}`}>
              <td style={{ ...tdStyle, paddingLeft: 14 + (depth + 1) * 22, color: T.textDim, fontSize: 12 }}>
                {etiqueta}
              </td>
              {meses.map((m) => (
                <td key={m} style={{ ...cellStyle, color: T.textDim }}>
                  {r._columna === m ? money(r.monto_estimado ?? r.importe, mon) : "—"}
                </td>
              ))}
              <td style={{ ...cellStyle, color: T.textDim }}>{money(r.monto_estimado ?? r.importe, mon)}</td>
            </tr>
          );
        });
      }
    } else {
      const isCollapsed = collapsed.has(groupPath);
      out.push(
        <tr key={groupPath} onClick={() => toggleGroup(groupPath)} style={{ cursor: "pointer", background: T.panelAlt }}>
          <td style={{ ...tdStyle, paddingLeft: 14 + depth * 22, fontWeight: 600, display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ color: T.textFaint, fontSize: 10 }}>{isCollapsed ? "▶" : "▼"}</span>{entry.value}
          </td>
          {meses.map((m) => <td key={m} style={{ ...cellStyle, fontWeight: 600 }}>{entry.child.mesSums[m] ? money(entry.child.mesSums[m], mon) : "—"}</td>)}
          <td style={{ ...cellStyle, fontWeight: 600 }}>{money(entry.child.total, mon)}</td>
        </tr>
      );
      if (!isCollapsed) out = out.concat(buildPivotTrs(entry.child, groupPath, collapsed, toggleGroup, meses, depth + 1, resolverMoneda));
    }
  });
  return out;
}

// Flattens a grouped tree into <tr> elements: a header row per group (collapsible,
// with count + sum), followed by that group's leaf rows (via renderRowTr) when expanded.
const GROUP_LEVEL_COLORS = [T.accent, T.teal, T.blue];
function buildGroupedTrs(node, path, collapsed, toggleGroup, colSpan, depth, renderRowTr, fieldLabels = {}, counter = { n: 0 }) {
  if (node.type === "rows") return node.rows.map((r) => renderRowTr(r, depth, ++counter.n));
  let out = [];
  const levelColor = GROUP_LEVEL_COLORS[depth % GROUP_LEVEL_COLORS.length];
  // La jerarquía visual baja de intensidad mientras más profundo el nivel:
  // fondo más oscuro y texto más grande/grueso arriba, más sutil abajo.
  const bgShade = ["#EEF1F4", "#F4F6F8", "#F8F9FB"][Math.min(depth, 2)];
  const valueFontSize = [13, 12, 11.5][Math.min(depth, 2)];
  const valueFontWeight = [700, 600, 600][Math.min(depth, 2)];
  const labelFontSize = [12, 11, 10.5][Math.min(depth, 2)];
  const labelFontWeight = [800, 700, 600][Math.min(depth, 2)];
  const labelColor = [T.text, T.textDim, T.textFaint][Math.min(depth, 2)];
  const rowPadding = depth === 0 ? "12px 10px" : "8px 10px";
  const borderWidth = depth === 0 ? 4 : depth === 1 ? 3 : 2;
  const pillTone = depth === 0 ? "accent" : depth === 1 ? "teal" : "dim";

  node.entries.forEach((entry) => {
    const groupPath = `${path}/${node.key}:${entry.value}`;
    const isCollapsed = collapsed.has(groupPath);
    out.push(
      <tr key={groupPath} onClick={() => toggleGroup(groupPath)} style={{ cursor: "pointer" }}>
        <td
          colSpan={colSpan}
          style={{
            ...tdStyle, background: bgShade,
            padding: rowPadding,
            paddingLeft: 14 + depth * 24,
            borderLeft: `${borderWidth}px solid ${levelColor}`,
            borderBottom: `1px solid ${T.borderSoft}`,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <span style={{ color: T.textFaint, fontSize: depth === 0 ? 11 : 9.5, width: 12, flexShrink: 0 }}>{isCollapsed ? "▶" : "▼"}</span>
            <span style={{ fontSize: labelFontSize, fontWeight: labelFontWeight, color: labelColor, textTransform: "uppercase", letterSpacing: "0.04em", flexShrink: 0 }}>{fieldLabels[node.key] || node.key}</span>
            <Pill tone={pillTone}>
              <span style={{ fontSize: valueFontSize, fontWeight: valueFontWeight, fontFamily: T.fontUI, letterSpacing: 0 }}>{entry.value}</span>
            </Pill>
            <span style={{ fontSize: 10.5, color: T.textFaint, background: T.panel, borderRadius: 999, padding: "1px 7px", flexShrink: 0 }}>{entry.count}</span>
            <span style={{ fontSize: depth === 0 ? 13 : 11.5, fontWeight: depth === 0 ? 700 : 600, fontFamily: T.fontMono, color: T.text, marginLeft: "auto", flexShrink: 0 }}>{money(entry.sum)}</span>
          </div>
        </td>
      </tr>
    );
    if (!isCollapsed) out = out.concat(buildGroupedTrs(entry.child, groupPath, collapsed, toggleGroup, colSpan, depth + 1, renderRowTr, fieldLabels, counter));
  });
  return out;
}

const tableStyle = { width: "100%", borderCollapse: "collapse", fontSize: 12.5 };
const thStyle = { textAlign: "left", padding: "8px 10px", color: T.textFaint, fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.05em", borderBottom: `1px solid ${T.border}`, overflowWrap: "break-word", wordBreak: "break-word" };
const tdStyle = { padding: "9px 10px", borderBottom: `1px solid ${T.borderSoft}`, color: T.text, overflowWrap: "break-word", wordBreak: "break-word" };

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

  const descargarPlantilla = async () => {
    const wbx = new ExcelJS.Workbook();
    // El importador ya leía Año y Zona, pero la plantilla no los ofrecía, así
    // que nadie sabía que podía usarlos y todo caía al año por defecto sin zona.
    const headers = ["Mes", "Año", "Concepto", "Rubro", "Categoria", "Proyecto", "Zona", "Sub Total MXN", "Sub Total USD", "Moneda", "SMI", "ID"];
    const ejemplos = {
      OSB: ["Agosto", 2026, "Ejemplo: Servicio de energía eléctrica", "Servicios Básicos", "Energía eléctrica", "Todos", "Queretaro", 12000, 0, "MXP", "", ""],
      CTM: ["Agosto", 2026, "Ejemplo: Renta de oficina base Poza Rica", "Servicios Operativos", "Arrendamientos", "Todos", "Poza Rica", 27000, 0, "MXP", "", ""],
      ISE: ["Agosto", 2026, "Ejemplo: Producto químico deshidratación", "Productos Químicos", "Productos químicos de operación", "Desh Gral", "", 0, 45000, "USD", "", ""],
    };
    UNIDAD_KEYS.forEach((u) => {
      const ws = wbx.addWorksheet(`RawData-${u}`);
      const headerRow = ws.addRow(headers);
      formatearHojaDatos(ws, headerRow, headers.length);
      const ejemploRow = ws.addRow(ejemplos[u]);
      ejemploRow.eachCell((cell) => { cell.font = { italic: true, color: { argb: "FF8B99A6" } }; });
      ws.columns = [{ width: 11 }, { width: 7 }, { width: 45 }, { width: 24 }, { width: 30 }, { width: 16 }, { width: 15 }, { width: 14 }, { width: 14 }, { width: 9 }, { width: 8 }, { width: 15 }];
    });
    const notas = wbx.addWorksheet("Instrucciones");
    [
      "Cómo usar esta plantilla",
      "",
      "1. Cada hoja \"RawData-OSB/CTM/ISE\" es el catálogo de esa compañía — no las mezcles.",
      "2. BORRA la fila de ejemplo (fila 2, en cursiva gris) de cada hoja antes de subir.",
      "3. ID/Folio: déjalo vacío para que la app le asigne uno automático, o pon uno",
      "   existente (ej. CTM-AGO26-045) para ACTUALIZAR esa partida en vez de crear otra.",
      "4. Sub Total MXN / Sub Total USD: llena solo el que aplique, deja el otro en 0.",
      "5. Proyecto: usa un proyecto real del Catálogo, o los marcadores Desh Gral /",
      "   Prod Gral / Todos.",
      "6. Zona: OPCIONAL. Vacía significa \"cualquier zona\", que es lo correcto para",
      "   partidas como Nómina o servicios legales. Usa una zona del Catálogo.",
      "7. Año: si lo dejas vacío se usa el año en curso. Conviene ponerlo cuando",
      "   cargues presupuesto de otro ejercicio.",
    ].forEach((linea, i) => { notas.getCell(`A${i + 1}`).value = linea; });
    notas.getCell("A1").font = { bold: true };
    notas.getColumn(1).width = 90;

    const buffer = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "plantilla_partidas.xlsx";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

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
      // _avisoClasificacion es un dato de pantalla, no de la tabla: se
      // descarta antes de insertar o Supabase rechaza la columna desconocida.
      const nuevas = preview.rows.filter((r) => !r._existenteId)
        .map(({ _existenteId, _avisoClasificacion, ...rest }) => rest);
      const actualizaciones = preview.rows.filter((r) => r._existenteId);

      if (nuevas.length) await partidasApi.bulkInsert(nuevas);
      for (const r of actualizaciones) {
        const { id, _existenteId, _avisoClasificacion, ...patch } = r;
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
    <Panel
      title="Carga masiva desde Excel"
      subtitle='Sube el libro con hojas "RawData-OSB", "RawData-CTM" y/o "RawData-ISE" — se detectan automáticamente'
      right={<Button variant="ghost" onClick={descargarPlantilla}>Descargar plantilla</Button>}
    >
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
            {preview.rows.some((r) => r._avisoClasificacion) && (
              <Pill tone="amber">
                {preview.rows.filter((r) => r._avisoClasificacion).length} con rubro o categoría fuera del catálogo
              </Pill>
            )}
          </div>

          {/* No bloquea la importación: avisa. Rechazar el archivo entero por
              una clasificación obligaría a rehacerlo, y casi siempre es más
              fácil corregirla después desde la tabla. */}
          {preview.rows.some((r) => r._avisoClasificacion) && (
            <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "10px 13px", borderRadius: "0 6px 6px 0", fontSize: 12, marginBottom: 12 }}>
              <b>Hay filas cuyo rubro o categoría no existe en el catálogo</b>
              Se van a importar de todas formas, pero conviene revisarlas: el Dashboard agrupa por rubro,
              y un rubro inventado queda fuera de todos los cortes.
              <ul style={{ margin: "8px 0 0", paddingLeft: 18, color: T.textDim }}>
                {[...new Set(preview.rows.filter((r) => r._avisoClasificacion).map((r) => r._avisoClasificacion))]
                  .slice(0, 6).map((a) => <li key={a}>{a}</li>)}
              </ul>
            </div>
          )}

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

// Modal ligero para editar una transacción sin salir de la vista en la que
// estás (ej. desde la fila expandida de una partida) — no cambia de pestaña.
// No incluye el selector de Partida (aquí ya se sabe a cuál pertenece).
/**
 * Una transacción PAGADA no se borra.
 *
 * Un pago ejecutado es un hecho contable: ya salió el dinero. Borrarlo hace
 * que el gasto desaparezca del presupuesto sin que nada lo explique, y lo
 * reportado a Dirección deja de cuadrar con el banco. Para eliminarla hay que
 * marcarla primero como No Pagada, que obliga a decidir explícitamente que
 * ese pago no ocurrió.
 *
 * Devuelve true si se puede borrar.
 */
function puedeBorrarTransaccion(t) {
  if (!t || t.status !== "Pagado") return true;
  alert(
    `No se puede eliminar: la transacción "${t.concepto_detallado || t.proveedor || t.id}" ` +
    `(${money(t.importe, t.moneda)}) está marcada como Pagada.\n\n` +
    `Si de verdad hay que borrarla, primero cámbiale el status a "No Pagado". ` +
    `Así queda claro que se está deshaciendo un pago registrado, no solo limpiando un renglón.`
  );
  return false;
}

/** Normaliza la moneda: vacío o cualquier cosa distinta de USD cuenta como MXP. */
const monedaNorm = (v) => ((v || "MXP") === "USD" ? "USD" : "MXP");
const mismaMoneda = (a, b) => monedaNorm(a) === monedaNorm(b);

/**
 * Impide vincular un gasto a una partida de OTRA moneda.
 *
 * Es un bloqueo, no un aviso: la comparación "presupuestado vs. pagado" no
 * significa nada si el presupuesto está en pesos y el gasto en dólares. Y el
 * daño no se queda en esa fila — el Dashboard dibuja el bloque de moneda
 * según la transacción, así que si la partida trae marcador de prorrateo, un
 * solo movimiento mal capturado pinta una sección entera de dólares repartida
 * entre todos los proyectos.
 *
 * Devuelve true si se puede guardar.
 */
function validarMonedaContraPartida(form, partidas) {
  const partida = partidas.find((p) => p.id === form.partida_id);
  if (!partida) return true;
  if (mismaMoneda(form.moneda, partida.moneda)) return true;
  alert(
    `No se puede guardar: la transacción está en ${monedaNorm(form.moneda)} y la partida ` +
    `"${partida.folio || partida.concepto}" es en ${monedaNorm(partida.moneda)}.\n\n` +
    `Elige una partida en ${monedaNorm(form.moneda)}, o corrige la moneda de la transacción ` +
    `si el importe está en ${monedaNorm(partida.moneda)}.`
  );
  return false;
}

/**
 * Genera el archivo de la Solicitud de Pago.
 *
 * No replica el formato original —eso no era el objetivo— pero conserva
 * TODOS los datos y su orden, para que quien lo recibe en Pagos reconozca
 * el documento y no tenga que buscar dónde quedó cada cosa.
 */
async function generarExcelSPP(r) {
  const wbx = new ExcelJS.Workbook();
  const ws = wbx.addWorksheet(`SPP ${r.folio}`);
  ws.columns = [{ width: 30 }, { width: 46 }, { width: 14 }, { width: 15 },
                { width: 15 }, { width: 14 }, { width: 14 }, { width: 15 }];

  const num = '"$"#,##0.00';
  const gris = { type: "pattern", pattern: "solid", fgColor: { argb: "FFECEEF1" } };
  const azul = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };

  const tit = ws.addRow([`Solicitud de Pago a Proveedores ${r.compania}`]);
  tit.font = { bold: true, size: 15, name: "Calibri" };
  ws.mergeCells(1, 1, 1, 8);
  ws.addRow([]);

  const dato = (etq, val) => {
    const row = ws.addRow([etq, val]);
    row.getCell(1).font = { bold: true, name: "Calibri", size: 11 };
    row.getCell(2).font = { name: "Calibri", size: 11 };
    return row;
  };
  dato("Folio", r.folio);
  dato("Fecha Elaboracion", r.fecha_elaboracion);
  dato("Fecha de Pago", r.fecha_pago || "");
  dato("Zona", r.zona);
  dato("Proyecto", r.centro_costo ? `${r.centro_costo} — ${r.proyecto}` : r.proyecto);
  dato("Responsable de Proyecto", r.responsable);
  dato("Solicitante", r.solicitante);
  dato("Lugar de Adquisicion", r.lugar_adquisicion);
  ws.addRow([]);

  const intro = ws.addRow(["Por medio de la presente solicito se realice el pago al Proveedor como sigue:"]);
  intro.font = { name: "Calibri", size: 11, italic: true };
  ws.mergeCells(intro.number, 1, intro.number, 8);
  ws.addRow([]);

  const hdr = ws.addRow(["PARTIDA", "DESCRIPCION DE PRODUCTOS O SERVICIOS", "CANTIDAD",
                         "P.U", "SUBTOTAL", "IVA", "RET ISR", "RET IVA"]);
  hdr.eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Calibri", size: 10.5 };
    c.fill = azul;
    c.alignment = { horizontal: "center", vertical: "center", wrapText: true };
  });
  hdr.height = 24;

  const fila = ws.addRow([1, r.descripcion, r.cantidad, r.precio_unitario,
                          r.subtotal, r.iva, r.ret_isr, r.ret_iva]);
  [4, 5, 6, 7, 8].forEach((i) => { fila.getCell(i).numFmt = num; fila.getCell(i).alignment = { horizontal: "right" }; });
  fila.getCell(2).alignment = { horizontal: "left", wrapText: true, vertical: "top" };
  ws.addRow([]);

  const totRow = (etq, val, fuerte) => {
    const row = ws.addRow(["", "", "", "", "", etq, "", val]);
    row.getCell(6).font = { bold: true, name: "Calibri", size: 11 };
    row.getCell(6).alignment = { horizontal: "right" };
    row.getCell(8).numFmt = num;
    row.getCell(8).font = { bold: true, name: "Calibri", size: fuerte ? 12 : 11 };
    row.getCell(8).alignment = { horizontal: "right" };
    if (fuerte) [6, 7, 8].forEach((i) => { row.getCell(i).fill = gris; });
    return row;
  };
  totRow("TOTAL", r.total);
  totRow("Descuento", r.descuento || 0);
  totRow(`TOTAL A PAGAR ${r.pct_pago}%`, r.total_a_pagar, true);
  ws.addRow([]);

  dato("Tipo de Moneda", r.moneda);
  dato("Condiciones de Pago", `${r.condicion_pago}${r.tipo_pago ? ` — ${r.tipo_pago}` : ""}`);
  dato("Numero de Proveedor ASPEL-SAE", r.proveedor_sae || "");
  ws.addRow([]);
  dato("Concepto de Pago", r.concepto);
  ws.addRow([]);
  dato("Observaciones:", r.observaciones || "");
  ws.addRow([]);

  const ban = ws.addRow(["Datos Bancarios del Proveedor"]);
  ban.font = { bold: true, name: "Calibri", size: 12 };
  ban.getCell(1).fill = gris;
  ws.mergeCells(ban.number, 1, ban.number, 8);
  dato("Nombre o Razon Social", r.proveedor);
  dato("Banco", r.banco || "");
  dato("Sucursal Bancaria", r.sucursal || "");
  // Cuenta y CLABE como TEXTO: Excel convierte 18 dígitos a notación
  // científica y el número llegaría corrupto al área de Pagos.
  const fc = dato("Cuenta Bancaria", String(r.cuenta || ""));
  fc.getCell(2).numFmt = "@";
  dato("Referencia Bancaria", r.referencia_bancaria || "");
  const fcl = dato("Cuenta CLABE", String(r.clabe || ""));
  fcl.getCell(2).numFmt = "@";

  const buf = await wbx.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `SPP ${r.compania}-${r.folio} - ${String(r.proveedor || "").slice(0, 30)}.xlsx`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * PDF de la Solicitud de Pago. Es el documento que llega al área de Pagos,
 * así que se cuida la presentación: bloques delimitados, la cifra a pagar
 * destacada y los datos bancarios agrupados aparte, que es lo que se
 * consulta al ejecutar la transferencia.
 */
function generarPdfSPP(r) {
  const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "letter" });
  const M = 40;                    // margen
  const A = 612 - M * 2;           // ancho útil
  const AZUL = [62, 92, 118];
  const GRIS = [107, 119, 133];
  const num = (v) => "$" + (Number(v) || 0).toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  let y = M;

  // --- Encabezado ---
  doc.setFillColor(...AZUL);
  doc.rect(M, y, A, 46, "F");
  doc.setTextColor(255).setFontSize(14).setFont(undefined, "bold");
  doc.text("SOLICITUD DE PAGO A PROVEEDORES", M + 14, y + 29);
  // La compañía va como prefijo del folio: "CTM-12" identifica el documento
  // de un vistazo, y el título deja de competir con ella por el mismo renglón.
  doc.setFontSize(9).setFont(undefined, "normal");
  doc.text("FOLIO", M + A - 14, y + 17, { align: "right" });
  doc.setFontSize(19).setFont(undefined, "bold");
  doc.text(`${r.compania}-${r.folio}`, M + A - 14, y + 36, { align: "right" });
  y += 62;

  /* Los datos generales en dos columnas: en una sola, el documento se
     alarga y obliga a recorrerlo para encontrar un dato. */
  const filaDoble = (e1, v1, e2, v2) => {
    doc.setFontSize(8).setTextColor(...GRIS).setFont(undefined, "normal");
    doc.text(String(e1).toUpperCase(), M, y);
    if (e2) doc.text(String(e2).toUpperCase(), M + A / 2, y);
    doc.setFontSize(10).setTextColor(35, 42, 49);
    doc.text(String(v1 || "—"), M, y + 13);
    if (e2) doc.text(String(v2 || "—"), M + A / 2, y + 13);
    y += 28;
  };
  filaDoble("Fecha de elaboración", r.fecha_elaboracion, "Fecha de pago", r.fecha_pago);
  filaDoble("Zona", r.zona, "Proyecto", r.centro_costo ? `${r.centro_costo} — ${r.proyecto}` : r.proyecto);
  filaDoble("Responsable de proyecto", r.responsable, "Solicitante", r.solicitante);
  filaDoble("Lugar de adquisición", r.lugar_adquisicion, "Proveedor ASPEL-SAE", r.proveedor_sae);

  y += 4;
  doc.setDrawColor(220).line(M, y, M + A, y);
  y += 16;
  doc.setFontSize(9.5).setTextColor(...GRIS).setFont(undefined, "italic");
  doc.text("Por medio de la presente solicito se realice el pago al Proveedor como sigue:", M, y);
  doc.setFont(undefined, "normal");
  y += 16;

  // --- Concepto y desglose ---
  autoTable(doc, {
    startY: y,
    head: [["PARTIDA", "DESCRIPCIÓN DE PRODUCTOS O SERVICIOS", "CANT.", "P.U.", "SUBTOTAL"]],
    body: [[1, r.descripcion || "", r.cantidad, num(r.precio_unitario), num(r.subtotal)]],
    styles: { fontSize: 8.5, cellPadding: 6 },
    headStyles: { fillColor: AZUL, textColor: 255, halign: "center" },
    columnStyles: { 0: { halign: "center", cellWidth: 50 }, 1: { halign: "left" },
                    2: { halign: "center", cellWidth: 48 }, 3: { halign: "right", cellWidth: 70 },
                    4: { halign: "right", cellWidth: 80 } },
    margin: { left: M, right: M },
  });
  y = doc.lastAutoTable.finalY + 14;

  // Los importes alineados a la derecha, junto al total, para poder
  // seguir la aritmética de arriba abajo.
  const linea = (etq, val, fuerte) => {
    doc.setFontSize(fuerte ? 11 : 9.5).setFont(undefined, fuerte ? "bold" : "normal");
    doc.setTextColor(fuerte ? 35 : 90, fuerte ? 42 : 100, fuerte ? 49 : 110);
    doc.text(etq, M + A - 150, y, { align: "right" });
    doc.text(val, M + A, y, { align: "right" });
    y += fuerte ? 18 : 14;
  };
  linea("Subtotal", num(r.subtotal));
  linea("IVA", num(r.iva));
  if (Number(r.ret_isr)) linea("Retención ISR", num(r.ret_isr));
  if (Number(r.ret_iva)) linea("Retención IVA", num(r.ret_iva));
  if (Number(r.descuento)) linea("Descuento", num(r.descuento));
  linea("TOTAL", num(r.total), true);

  y += 6;
  // La cifra que se va a transferir, destacada: es el dato que Pagos busca.
  doc.setFillColor(...AZUL);
  doc.rect(M + A - 260, y, 260, 34, "F");
  doc.setTextColor(255).setFontSize(9).setFont(undefined, "normal");
  doc.text(`TOTAL A PAGAR ${r.pct_pago}%`, M + A - 248, y + 14);
  doc.setFontSize(15).setFont(undefined, "bold");
  doc.text(num(r.total_a_pagar), M + A - 12, y + 26, { align: "right" });
  y += 50;

  doc.setTextColor(35, 42, 49);
  filaDoble("Tipo de moneda", r.moneda, "Condiciones de pago",
    `${r.condicion_pago}${r.tipo_pago ? ` — ${r.tipo_pago}` : ""}`);

  const bloqueTexto = (etq, val) => {
    doc.setFontSize(8).setTextColor(...GRIS);
    doc.text(String(etq).toUpperCase(), M, y);
    doc.setFontSize(10).setTextColor(35, 42, 49);
    const lineas = doc.splitTextToSize(String(val || "—"), A);
    doc.text(lineas, M, y + 13);
    y += 13 + lineas.length * 12 + 8;
  };
  bloqueTexto("Concepto de pago", r.concepto);
  if (r.observaciones) bloqueTexto("Observaciones", r.observaciones);

  // --- Datos bancarios, en su propio bloque ---
  y += 6;
  doc.setFillColor(236, 238, 241);
  doc.rect(M, y, A, 22, "F");
  doc.setFontSize(10).setFont(undefined, "bold").setTextColor(35, 42, 49);
  doc.text("DATOS BANCARIOS DEL PROVEEDOR", M + 10, y + 15);
  doc.setFont(undefined, "normal");
  y += 34;
  filaDoble("Nombre o razón social", r.proveedor, "Banco", r.banco);
  filaDoble("Sucursal bancaria", r.sucursal, "Referencia bancaria", r.referencia_bancaria);
  filaDoble("Cuenta bancaria", r.cuenta, "Cuenta CLABE", r.clabe);

  doc.save(`SPP ${r.compania}-${r.folio} - ${String(r.proveedor || "").slice(0, 30)}.pdf`);
}

/**
 * Diálogo de la Solicitud de Pago a Proveedores.
 *
 * Todo lo que la app sabe viene precargado; todo lo que la app DEDUCE se
 * muestra para que se confirme. La diferencia importa: el desglose fiscal
 * se calcula hacia atrás desde el importe pagado, y ese camino inverso solo
 * es exacto si el importe correspondía justo a ese esquema. Presentarlo como
 * un hecho invitaría a firmar un documento con cifras que nadie revisó.
 */
function SolicitudPagoModal({ transaccion, onClose, unidad, partidas, proyectosUnidad, proveedoresApi, cuentasApi, transaccionesApi, session }) {
  const t = transaccion;
  const partida = partidas.find((p) => p.id === t.partida_id);
  const proveedor = proveedoresApi.rows.find((p) => p.id === t.proveedor_id);
  const cuenta = cuentasApi.rows.find((c) => c.id === t.cuenta_id)
    || (t.proveedor_id ? cuentasApi.rows.find((c) => c.proveedor_id === t.proveedor_id) : null);
  const proyNombre = t.proyecto || partida?.proyecto || "";
  const proy = proyectosUnidad.find((p) => p.nombre === proyNombre);

  /* Los datos fijos de la compañía se leen al abrir: son los mismos en cada
     solicitud y hasta ahora se escribían a mano una y otra vez. */
  const [cfgCia, setCfgCia] = useState(null);
  useEffect(() => {
    let vivo = true;
    supabase.from("config_companias").select("*").eq("compania", unidad).maybeSingle()
      .then(({ data }) => { if (vivo && data) { setCfgCia(data);
        setF((prev) => ({ ...prev,
          responsable: prev.responsable || data.spp_responsable || "",
          lugar_adquisicion: data.spp_lugar_adquisicion || prev.lugar_adquisicion })); } });
    return () => { vivo = false; };
  }, [unidad]);

  const [esquema, setEsquema] = useState("pf_servicios");
  const [pctPago, setPctPago] = useState(50);
  const [guardando, setGuardando] = useState(false);
  const [formato, setFormato] = useState("pdf");

  // El subtotal se deduce del importe pagado, que es lo único que la
  // transacción guarda hoy.
  const [d, setD] = useState(() => {
    const sub = subtotalDesdeTotal(t.importe, "pf_servicios");
    return { ...calcularDesglose(sub, "pf_servicios"), descuento: 0 };
  });
  const recalcular = (esq, sub) => {
    const nuevo = calcularDesglose(sub, esq);
    if (nuevo) setD((prev) => ({ ...nuevo, descuento: prev.descuento || 0 }));
  };

  const [f, setF] = useState({
    fecha_elaboracion: new Date().toISOString().slice(0, 10),
    fecha_pago: t.fecha_pago || t.dia || "",
    zona: t.zona || partida?.zona || "",
    responsable: cfgCia?.spp_responsable || "",
    solicitante: t.solicitante || session?.user?.email?.split("@")[0] || "",
    lugar_adquisicion: cfgCia?.spp_lugar_adquisicion || "Queretaro",
    tipo_pago: "Anticipo",
    condicion_pago: "TRANSFERENCIA",
    concepto: t.concepto_detallado || "",
    descripcion: t.concepto_detallado || "",
    cantidad: 1,
    observaciones: "",
    moneda: t.moneda === "USD" ? "Dólares" : "Pesos",
  });

  const total = totalDesglose(d);
  const aPagar = r2(total * (Number(pctPago) || 0) / 100);
  const num = (v) => (Number(v) || 0).toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const faltantes = [];
  if (!proveedor) faltantes.push("proveedor no vinculado");
  if (!cuenta) faltantes.push("sin cuenta bancaria");
  if (!proy?.centro_costo) faltantes.push("el proyecto no tiene centro de costo");
  if (!f.zona) faltantes.push("sin zona");

  const generar = async () => {
    setGuardando(true);
    try {
      // Folio autonumérico por compañía, leído de la base y no del estado
      // local: dos personas generando a la vez calcularían el mismo.
      const { data: ultimos, error: errF } = await supabase
        .from("solicitudes_pago").select("folio")
        .eq("compania", unidad).order("folio", { ascending: false }).limit(1);
      if (errF) throw errF;
      const { data: cfg } = await supabase.from("config_companias")
        .select("spp_ultimo").eq("compania", unidad).maybeSingle();
      // El MAYOR entre lo configurado y lo ya emitido: bajar el consecutivo
      // por error no debe reutilizar folios de solicitudes que ya salieron.
      const folio = Math.max(Number(cfg?.spp_ultimo) || 0, (ultimos && ultimos[0]?.folio) || 0) + 1;

      const reg = {
        id: uid(), compania: unidad, folio, transaccion_id: t.id,
        fecha_elaboracion: f.fecha_elaboracion, fecha_pago: f.fecha_pago || null,
        zona: f.zona, proyecto: proyNombre, centro_costo: proy?.centro_costo || "",
        responsable: f.responsable, solicitante: f.solicitante,
        lugar_adquisicion: f.lugar_adquisicion,
        proveedor: proveedor?.nombre || t.proveedor || "",
        proveedor_sae: proveedor?.id_sae || "",
        banco: cuenta?.banco || "", sucursal: cuenta?.sucursal || "",
        cuenta: cuenta?.numero_cuenta || "", clabe: cuenta?.clabe || "",
        referencia_bancaria: proveedor?.referencia || "",
        moneda: f.moneda, tipo_pago: f.tipo_pago, condicion_pago: f.condicion_pago,
        concepto: f.concepto, descripcion: f.descripcion,
        cantidad: Number(f.cantidad) || 1,
        precio_unitario: r2((Number(d.subtotal) || 0) / (Number(f.cantidad) || 1)),
        subtotal: d.subtotal, iva: d.iva, ret_isr: d.ret_isr, ret_iva: d.ret_iva,
        descuento: Number(d.descuento) || 0, total,
        pct_pago: Number(pctPago) || 0, total_a_pagar: aPagar,
        observaciones: f.observaciones,
      };
      const { error } = await supabase.from("solicitudes_pago").insert(reg);
      if (error) throw error;
      /* El constructor de consultas de Supabase no es una promesa hasta que
         se le hace await, así que encadenarle .catch() lanza "catch is not a
         function". Estos dos pasos son secundarios —la solicitud ya quedó
         guardada— y no deben impedir que se descargue el documento, así que
         cada uno va en su propio try. */
      try {
        await supabase.from("config_companias")
          .upsert({ compania: unidad, spp_ultimo: folio, updated_at: new Date().toISOString() });
      } catch { /* el folio ya quedó en solicitudes_pago; el consecutivo se recalcula solo */ }

      try {
        // El desglose confirmado se guarda también en la transacción: es el
        // dato que faltaba para saber de qué se compone el importe.
        await transaccionesApi.update(t.id, {
          subtotal: d.subtotal, iva: d.iva, ret_isr: d.ret_isr, ret_iva: d.ret_iva,
          descuento: Number(d.descuento) || 0, pct_pago: Number(pctPago) || 0,
          tipo_pago: f.tipo_pago, condicion_pago: f.condicion_pago,
          observaciones_spp: f.observaciones,
        });
      } catch { /* el documento es lo que importa; el desglose se puede recapturar */ }

      if (formato === "pdf") generarPdfSPP(reg);
      else await generarExcelSPP(reg);
      onClose();
    } catch (err) {
      alert("No se pudo generar la solicitud: " + (err.message || err));
    } finally {
      setGuardando(false);
    }
  };

  const campo = (etq, val, set, ancho = "100%") => (
    <Field label={etq}>
      <TextInput value={val} onChange={(e) => set(e.target.value)} style={{ width: ancho }} />
    </Field>
  );

  return (
    <Modal title={`Solicitud de Pago a Proveedores — ${unidad}`}
      subtitle="Revisa cada cifra antes de generar: el desglose fiscal se deduce del importe pagado y solo es exacto si corresponde al esquema elegido"
      onClose={onClose} width={900}>
      {faltantes.length > 0 && (
        <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "10px 13px", borderRadius: "0 6px 6px 0", fontSize: 12, marginBottom: 14 }}>
          <b>La solicitud va a salir incompleta</b>
          Falta: {faltantes.join(" · ")}. Puedes generarla igual y completar a mano,
          pero conviene corregirlo en la app para que la próxima salga sola.
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 6 }}>
        {campo("Fecha de elaboración", f.fecha_elaboracion, (v) => setF({ ...f, fecha_elaboracion: v }))}
        {campo("Fecha de pago", f.fecha_pago, (v) => setF({ ...f, fecha_pago: v }))}
        {campo("Zona", f.zona, (v) => setF({ ...f, zona: v }))}
        <Field label="Proyecto / Centro de costo">
          <TextInput value={`${proyNombre}${proy?.centro_costo ? ` · ${proy.centro_costo}` : " · sin CC"}`} disabled />
        </Field>
        {campo("Responsable de proyecto", f.responsable, (v) => setF({ ...f, responsable: v }))}
        {campo("Solicitante", f.solicitante, (v) => setF({ ...f, solicitante: v }))}
      </div>

      <div style={{ borderTop: `1px solid ${T.borderSoft}`, margin: "14px 0", paddingTop: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>Desglose fiscal</div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr", gap: 12, marginBottom: 10 }}>
          <Field label="Esquema fiscal del proveedor">
            <Select value={esquema} onChange={(e) => {
              const esq = e.target.value;
              setEsquema(esq);
              if (esq !== "manual") recalcular(esq, subtotalDesdeTotal(t.importe, esq));
            }}>
              {ESQUEMAS_FISCALES.map((x) => <option key={x.id} value={x.id}>{x.label}</option>)}
            </Select>
            <span style={{ fontSize: 11, color: T.textFaint, marginTop: 4, display: "block" }}>
              {(ESQUEMAS_FISCALES.find((x) => x.id === esquema) || {}).nota}
            </span>
          </Field>
          <Field label="Subtotal">
            <TextInput type="number" value={d.subtotal}
              onChange={(e) => { const v = e.target.value; setD({ ...d, subtotal: v });
                if (esquema !== "manual") recalcular(esquema, v); }} />
          </Field>
          <Field label="% que se paga">
            <TextInput type="number" value={pctPago} onChange={(e) => setPctPago(e.target.value)} />
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
          <Field label="IVA">
            <TextInput type="number" value={d.iva} onChange={(e) => setD({ ...d, iva: e.target.value })} />
          </Field>
          <Field label="Retención ISR">
            <TextInput type="number" value={d.ret_isr} onChange={(e) => setD({ ...d, ret_isr: e.target.value })} />
          </Field>
          <Field label="Retención IVA">
            <TextInput type="number" value={d.ret_iva} onChange={(e) => setD({ ...d, ret_iva: e.target.value })} />
          </Field>
          <Field label="Descuento">
            <TextInput type="number" value={d.descuento} onChange={(e) => setD({ ...d, descuento: e.target.value })} />
          </Field>
        </div>

        <div style={{ display: "flex", gap: 24, marginTop: 12, padding: "10px 14px", background: T.panelAlt, borderRadius: 6, flexWrap: "wrap" }}>
          <div><span style={{ fontSize: 11, color: T.textFaint }}>TOTAL</span>
            <div style={{ fontFamily: T.fontMono, fontSize: 15, fontWeight: 700 }}>${num(total)}</div></div>
          <div><span style={{ fontSize: 11, color: T.textFaint }}>A PAGAR {pctPago}%</span>
            <div style={{ fontFamily: T.fontMono, fontSize: 15, fontWeight: 700, color: T.teal }}>${num(aPagar)}</div></div>
          {/* El contraste con lo capturado revela si el desglose cuadra. */}
          <div><span style={{ fontSize: 11, color: T.textFaint }}>IMPORTE DE LA TRANSACCIÓN</span>
            <div style={{ fontFamily: T.fontMono, fontSize: 15,
              color: Math.abs(total - (Number(t.importe) || 0)) < 1 ? T.textDim : T.amber }}>
              ${num(t.importe)}
              {Math.abs(total - (Number(t.importe) || 0)) >= 1 &&
                <span style={{ fontSize: 11 }}> · no coincide</span>}
            </div></div>
        </div>
      </div>

      <div style={{ borderTop: `1px solid ${T.borderSoft}`, margin: "14px 0", paddingTop: 14,
                    display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
        <Field label="Tipo de pago">
          <Select value={f.tipo_pago} onChange={(e) => setF({ ...f, tipo_pago: e.target.value })}>
            {["Anticipo", "Pago Total", "Finiquito"].map((x) => <option key={x}>{x}</option>)}
          </Select>
        </Field>
        <Field label="Condición de pago">
          <Select value={f.condicion_pago} onChange={(e) => setF({ ...f, condicion_pago: e.target.value })}>
            {["TRANSFERENCIA", "CHEQUE", "EFECTIVO"].map((x) => <option key={x}>{x}</option>)}
          </Select>
        </Field>
        <Field label="Moneda">
          <Select value={f.moneda} onChange={(e) => setF({ ...f, moneda: e.target.value })}>
            {["Pesos", "Dólares"].map((x) => <option key={x}>{x}</option>)}
          </Select>
        </Field>
        <Field label="Concepto de pago" style={{ gridColumn: "span 3" }}>
          <TextInput value={f.concepto} onChange={(e) => setF({ ...f, concepto: e.target.value })} />
        </Field>
        <Field label="Descripción del producto o servicio" style={{ gridColumn: "span 3" }}>
          <TextInput value={f.descripcion} onChange={(e) => setF({ ...f, descripcion: e.target.value })} />
        </Field>
        <Field label="Observaciones" style={{ gridColumn: "span 3" }}>
          <TextInput value={f.observaciones} onChange={(e) => setF({ ...f, observaciones: e.target.value })}
            placeholder="Correos, nombres comerciales, datos informativos…" />
        </Field>
      </div>

      <div style={{ fontSize: 11.5, color: T.textFaint, marginBottom: 12 }}>
        Datos bancarios: {proveedor?.nombre || t.proveedor || "—"}
        {cuenta ? ` · ${cuenta.banco || "sin banco"} · CLABE ${cuenta.clabe || "—"}` : " · sin cuenta registrada"}
        {proveedor?.id_sae ? ` · SAE ${proveedor.id_sae}` : ""}
      </div>

      <div style={{ display: "flex", gap: 8 }}>
        <Button onClick={() => { setFormato("pdf"); generar(); }} disabled={guardando}>
          {guardando ? "Generando…" : "Generar PDF"}
        </Button>
        <Button variant="ghost" onClick={() => { setFormato("xlsx"); generar(); }} disabled={guardando}>
          Generar Excel
        </Button>
        <Button variant="ghost" onClick={onClose}>Cancelar</Button>
      </div>
    </Modal>
  );
}

function TransaccionQuickEditModal({ transaccion, onClose, transaccionesApi, proveedoresApi, cuentasApi, unidad, partidasUnidad = [] }) {
  const [form, setForm] = useState({ ...transaccion });
  const [saving, setSaving] = useState(false);
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const cuentasDelProveedorSeleccionado = form.proveedor_id ? cuentasApi.rows.filter((c) => c.proveedor_id === form.proveedor_id) : [];

  const submit = async (e) => {
    e.preventDefault();
    if (form.status === "Pagado" && !form.fecha_pago) {
      alert("Para marcar esta transacción como Pagada, primero indica la Fecha de Pago.");
      return;
    }
    if (!validarMonedaContraPartida(form, partidasUnidad)) return;
    setSaving(true);
    try {
      const { id, ...restRaw } = form;
      const rest = Object.fromEntries(Object.entries(restRaw).filter(([k]) => !k.startsWith("_")));
      rest.proveedor_id = rest.proveedor_id || null;
      rest.cuenta_id = rest.cuenta_id || null;
      rest.fecha_pago = rest.fecha_pago || null;
      await transaccionesApi.update(transaccion.id, rest);
      onClose();
    } catch (err) {
      alert("No se pudo guardar: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <Modal
      title="Editar transacción"
      subtitle={form.folio_transaccion ? `ID: ${form.folio_transaccion}` : undefined}
      onClose={onClose}
      width={760}
      zIndex={1100}
    >
      <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        <Field label="Día de Pago Programado">
          <TextInput type="date" value={form.dia || ""} onChange={(e) => setForm({ ...form, dia: e.target.value })} />
        </Field>
        <Field label="Solicitante">
          <TextInput value={form.solicitante || ""} onChange={(e) => setForm({ ...form, solicitante: e.target.value })} />
        </Field>
        <Field label="SMI">
          <TextInput value={form.smi || ""} onChange={(e) => setForm({ ...form, smi: e.target.value })} />
        </Field>
        <Field label="Área">
          <TextInput value={form.area || ""} onChange={(e) => setForm({ ...form, area: e.target.value })} />
        </Field>
        <Field label="Proyecto">
          <TextInput value={form.proyecto || ""} onChange={(e) => setForm({ ...form, proyecto: e.target.value })} />
        </Field>
        <Field label="Zona">
          <TextInput value={form.zona || ""} onChange={(e) => setForm({ ...form, zona: e.target.value })} />
        </Field>
        <Field label="Proveedor (catálogo)" style={{ gridColumn: "span 2" }}>
          <ProveedorPickerButton
            proveedores={proveedoresUnidad}
            proveedoresApi={proveedoresApi}
            cuentasApi={cuentasApi}
            unidad={unidad}
            value={form.proveedor_id}
            onChange={(id, p) => {
              const proveedor = p !== undefined ? p : (proveedoresUnidad.find((pr) => pr.id === id) || null);
              setForm({ ...form, proveedor_id: id, proveedor: proveedor ? proveedor.nombre : form.proveedor, cuenta_id: "" });
            }}
          />
        </Field>
        <Field label="Proveedor / Razón social (texto)" style={{ gridColumn: "span 2" }}>
          <TextInput value={form.proveedor || ""} onChange={(e) => setForm({ ...form, proveedor: e.target.value })} />
        </Field>
        {form.proveedor_id && (
          <Field label="Cuenta bancaria" style={{ gridColumn: "span 2" }}>
            <Select value={form.cuenta_id || ""} onChange={(e) => setForm({ ...form, cuenta_id: e.target.value })}>
              <option value="">— Elige una cuenta —</option>
              {cuentasDelProveedorSeleccionado.map((c) => <option key={c.id} value={c.id}>{c.banco} · {c.clabe} ({c.divisa})</option>)}
            </Select>
          </Field>
        )}
        <Field label="Concepto de pago (detallado)" style={{ gridColumn: "span 4" }}>
          <TextInput value={form.concepto_detallado || ""} onChange={(e) => setForm({ ...form, concepto_detallado: e.target.value })} />
        </Field>
        <Field label="Folio Compra SAE">
          <TextInput value={form.folio_compra_sae || ""} onChange={(e) => setForm({ ...form, folio_compra_sae: e.target.value })} />
        </Field>
        <Field label="Folio Factura">
          <TextInput value={form.folio_factura || ""} onChange={(e) => setForm({ ...form, folio_factura: e.target.value })} />
        </Field>
        <Field label="Forma de Pago">
          <Select value={form.forma_pago || ""} onChange={(e) => setForm({ ...form, forma_pago: e.target.value })}>
            <option value="">— Sin especificar —</option>
            {FORMAS_PAGO.map((f) => <option key={f.value} value={f.value}>{f.label}</option>)}
          </Select>
        </Field>
        <Field label="Método de Pago">
          <Select value={form.metodo_pago || ""} onChange={(e) => setForm({ ...form, metodo_pago: e.target.value })}>
            <option value="">— Sin especificar —</option>
            {METODOS_PAGO.map((m) => <option key={m.value} value={m.value}>{m.label}</option>)}
          </Select>
        </Field>
        <Field label="Referencia de Pago">
          <TextInput value={form.referencia_pago || ""} onChange={(e) => setForm({ ...form, referencia_pago: e.target.value })} placeholder="Folio SPEI, cheque, etc." />
        </Field>
        <Field label="Importe">
          <TextInput type="number" step="0.01" value={form.importe ?? ""} onChange={(e) => setForm({ ...form, importe: e.target.value })} />
        </Field>
        <Field label="Moneda">
          <Select value={form.moneda || "MXP"} onChange={(e) => setForm({ ...form, moneda: e.target.value })}>
            {MONEDAS.map((m) => <option key={m}>{m}</option>)}
          </Select>
        </Field>
        <Field label="Status">
          <Select value={form.status || ""} onChange={(e) => setForm({ ...form, status: e.target.value })}>
            <option value="">— Sin especificar —</option>
            <option>Pagado</option>
            <option>No Pagado</option>
          </Select>
        </Field>
        {form.status === "Pagado" && (
          <Field label="Fecha de Pago">
            <TextInput type="date" value={form.fecha_pago || ""} onChange={(e) => setForm({ ...form, fecha_pago: e.target.value })} required />
          </Field>
        )}
        <div style={{ gridColumn: "span 4", display: "flex", gap: 10, marginTop: 4 }}>
          <Button type="submit" disabled={saving}>{saving ? "Guardando…" : "Guardar cambios"}</Button>
          <Button type="button" variant="ghost" onClick={onClose}>Cancelar</Button>
        </div>
      </form>
    </Modal>
  );
}

function PartidasTab({ unidad, unidades, partidas, partidasApi, perfilesApi, transacciones, transaccionesApi, proveedoresApi, cuentasApi, zonas = ZONAS_RESPALDO }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadores = marcadoresDisponibles(proyectosUnidad);
  const anioDefault = (() => {
    const anios = partidas.filter((p) => p.unidad === unidad).map((p) => p.anio).filter(Boolean);
    return anios.length ? Math.max(...anios) : new Date().getFullYear();
  })();
  const blank = { unidad, mes: "Agosto", anio: anioDefault, smi: "", concepto: "", rubro: RUBROS[0].rubro, categoria: RUBROS[0].categorias[0], proyecto: marcadores[0] || "", zona: "", monto_estimado: "", moneda: "MXP", folio: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);

  const categoriasDisponibles = RUBROS.find((r) => r.rubro === form.rubro)?.categorias || [];
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);

  const [filtros, setFiltros] = useSessionState("ss-partidas-filtros", { texto: "", mes: [], anio: [], rubro: "Todos", proyecto: "Todos" });
  const filtrosMes = Array.isArray(filtros.mes) ? filtros.mes : [];
  const filtrosAnio = Array.isArray(filtros.anio) ? filtros.anio : [];
  const rubrosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.rubro).filter(Boolean))].sort();
  const proyectosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.proyecto).filter(Boolean))].sort();
  const mesesDisponiblesFiltro = MESES.filter((m) => partidasUnidad.some((p) => p.mes === m));
  const aniosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.anio).filter(Boolean))].sort((a, b) => a - b);

  // Los filtros de la tabla, extraídos como función para poder aplicarlos
  // también a otras compañías al exportar las tres de un jalón.
  // Nota: el filtro de rubro y el de proyecto se saltan cuando el valor
  // elegido no existe en la otra compañía, en vez de devolver cero filas.
  const aplicarFiltrosPartidas = (lista, estricto = true) => lista.filter((p) => {
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const enTexto = [p.concepto, p.folio, p.smi, p.categoria, p.zona].some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    if (filtrosMes.length && !filtrosMes.includes(p.mes)) return false;
    if (filtrosAnio.length && !filtrosAnio.includes(p.anio)) return false;
    if (estricto) {
      if (filtros.rubro !== "Todos" && p.rubro !== filtros.rubro) return false;
      if (filtros.proyecto !== "Todos" && p.proyecto !== filtros.proyecto) return false;
    }
    return true;
  });
  const partidasFiltradas = aplicarFiltrosPartidas(partidasUnidad);
  const filtrosActivos = filtros.texto.trim() || filtrosMes.length > 0 || filtrosAnio.length > 0 || filtros.rubro !== "Todos" || filtros.proyecto !== "Todos";
  const limpiarFiltros = () => setFiltros({ texto: "", mes: [], anio: [], rubro: "Todos", proyecto: "Todos" });

  const [sort, setSort] = usePrefState("pref-partidas-sort", { key: null, dir: "asc" }, sanearSort(["mes","anio","concepto","rubro","categoria","proyecto","zona","folio","monto_estimado","updated_at"]));
  const partidasOrdenadas = sortRows(partidasFiltradas, sort, {
    mes: (r) => MESES.indexOf(r.mes),
    monto_estimado: (r) => Number(r.monto_estimado) || 0,
    anio: (r) => Number(r.anio) || 0,
  });

  const [groupBys, setGroupBys] = usePrefState("pref-partidas-groupbys", [], sanearGroupBys(GROUP_OPCIONES));
  const [collapsedGroups, setCollapsedGroups] = useState(new Set());
  const toggleGroup = (path) => setCollapsedGroups((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });
  const groupKeys = groupBys.map((g) => g.field);
  const grouped = groupKeys.length ? agruparRows(partidasOrdenadas, groupBys) : null;

  const usadoDe = (p) => transacciones.filter((t) => t.partida_id === p.id).reduce((s, t) => s + (Number(t.importe) || 0), 0);

  // Exporta las partidas que están a la vista (con los filtros aplicados).
  // Encabezado, ancho y valor viven juntos en un solo arreglo para que agregar
  // una columna no pueda desalinear el archivo.
  const exportarExcel = async (todasLasCompanias = false) => {
    /* Catálogo completo de lo exportable. `col` amarra cada columna con la de
       la tabla: las que la tienen se incluyen solo si están visibles, y las
       que no —Unidad, Usado, Disponible— van siempre, porque no existen en
       pantalla pero sí hacen falta en la hoja. */
    const TODAS = [
      { header: "Unidad",    width: 9,  get: (p) => p.unidad },
      { header: "Folio",     width: 16, get: (p) => p.folio || "",        col: "folio" },
      { header: "Mes",       width: 11, get: (p) => p.mes,                col: "mes" },
      { header: "Año",       width: 7,  get: (p) => p.anio,               col: "anio" },
      { header: "Concepto",  width: 50, get: (p) => p.concepto,           col: "concepto" },
      { header: "Rubro",     width: 26, get: (p) => p.rubro,              col: "rubro" },
      { header: "Categoria", width: 30, get: (p) => p.categoria,          col: "categoria" },
      { header: "Zona",      width: 16, get: (p) => p.zona || "",         col: "zona" },
      { header: "Proyecto",  width: 17, get: (p) => p.proyecto || "",     col: "proyecto" },
      { header: "Moneda",    width: 9,  get: (p) => p.moneda || "MXP" },
      { header: "Monto",     width: 15, get: (p) => Number(p.monto_estimado) || 0, money: true, col: "monto_estimado" },
      { header: "Usado",     width: 15, get: (p) => usadoDe(p), money: true },
      { header: "Disponible",width: 15, get: (p) => (Number(p.monto_estimado) || 0) - usadoDe(p), money: true },
      { header: "SMI",       width: 10, get: (p) => p.smi || "" },
    ];
    const visibles = new Set(columnasVisibles.map((c) => c.key));
    const COLS = TODAS.filter((c) => !c.col || visibles.has(c.col));
    // Con todasLasCompanias, se ignora el selector global de compañía y se
    // exportan las tres en hojas separadas, aplicando los MISMOS filtros de mes,
    // año, rubro y proyecto que estén puestos. Es lo que necesita el preparador
    // de transacciones, porque el correo de pagos llega mezclado.
    const hojas = todasLasCompanias
      ? UNIDAD_KEYS.map((u) => ({
          unidad: u,
          filas: aplicarFiltrosPartidas(partidas.filter((p) => p.unidad === u), u === unidad),
        })).filter((h) => h.filas.length)
      : [{ unidad, filas: partidasOrdenadas }];

    if (!hojas.length) { alert("No hay partidas que exportar con los filtros actuales."); return; }

    const wbx = new ExcelJS.Workbook();
    hojas.forEach(({ unidad: u, filas }) => {
      const ws = wbx.addWorksheet(`RawData-${u}`);
      ws.columns = COLS.map((c) => ({ width: c.width }));
      const hr = ws.addRow(COLS.map((c) => c.header));
      formatearHojaDatos(ws, hr, COLS.length);
      filas.forEach((p) => {
        const row = ws.addRow(COLS.map((c) => c.get(p)));
        COLS.forEach((c, ci) => {
          const cell = row.getCell(ci + 1);
          cell.font = { name: "Calibri", size: 11 };
          if (c.money) { cell.numFmt = '"$"#,##0.00'; cell.alignment = { horizontal: "right" }; }
          // El folio se fuerza a texto: si Excel lo interpreta como número,
          // se pierden ceros y el archivo deja de servir para reimportar.
          if (c.header === "Folio") cell.alignment = { horizontal: "left" };
          if (c.header === "Concepto") cell.alignment = { horizontal: "left", wrapText: true, vertical: "top" };
        });
      });
    });
    const buf = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Partidas_${todasLasCompanias ? "OSB-CTM-ISE" : unidad}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };


  /* -----------------------------------------------------------------
     REPORTE EXCEL DE PARTIDAS
     Se diferencia de "Exportar" en dos cosas: tiene su propio selector de
     columnas —independiente de lo que esté visible en pantalla— y respeta
     el AGRUPAMIENTO de la vista, emitiendo una fila de encabezado con
     subtotal por cada grupo. Un listado plano de ochenta renglones no se
     parece a lo que se está viendo y obliga a rehacer el resumen a mano.
     ----------------------------------------------------------------- */
  const COLUMNAS_REPORTE_PARTIDAS = [
    { key: "folio",       header: "Folio",       width: 16, fija: true, get: (p) => p.folio || "" },
    { key: "mes",         header: "Mes",         width: 12, get: (p) => p.mes },
    { key: "anio",        header: "Año",         width: 8,  get: (p) => p.anio },
    { key: "concepto",    header: "Concepto",    width: 52, fija: true, get: (p) => p.concepto, izq: true },
    { key: "rubro",       header: "Rubro",       width: 26, get: (p) => p.rubro, izq: true },
    { key: "categoria",   header: "Categoría",   width: 30, get: (p) => p.categoria || "", izq: true },
    { key: "zona",        header: "Zona",        width: 16, get: (p) => p.zona || "Cualquiera" },
    { key: "proyecto",    header: "Proyecto",    width: 18, get: (p) => p.proyecto || "" },
    { key: "smi",         header: "SMI",         width: 11, get: (p) => p.smi || "" },
    { key: "moneda",      header: "Moneda",      width: 9,  get: (p) => p.moneda || "MXP" },
    { key: "monto",       header: "Presupuesto", width: 15, fija: true, money: true, get: (p) => Number(p.monto_estimado) || 0 },
    { key: "usado",       header: "Usado",       width: 15, money: true, get: (p) => usadoDe(p) },
    { key: "disponible",  header: "Disponible",  width: 15, money: true, get: (p) => (Number(p.monto_estimado) || 0) - usadoDe(p) },
  ];
  const repVis = useVisibilidadColumnas("colv-partidas-reporte", COLUMNAS_REPORTE_PARTIDAS, ["smi", "anio"]);
  const [generandoReporteXls, setGenerandoReporteXls] = useState(false);

  const generarReporteExcel = async () => {
    if (!partidasOrdenadas.length) {
      alert("No hay partidas en el filtro actual para el reporte.");
      return;
    }
    const COLS = repVis.visibles;
    setGenerandoReporteXls(true);
    try {
      /* El periodo sale de las partidas que REALMENTE salen en el reporte, no
         de los filtros: si el filtro es amplio pero solo hay septiembre, el
         archivo debe decir septiembre. Con varios meses se usa el rango, en
         orden cronológico y no alfabético. */
      const periodos = [...new Set(partidasOrdenadas.map((p) => `${p.mes}|${p.anio}`))]
        .map((k) => { const [mes, anio] = k.split("|"); return { mes, anio: Number(anio) }; })
        .sort((a, b) => a.anio - b.anio || MESES.indexOf(a.mes) - MESES.indexOf(b.mes));
      const etiquetaPeriodo = !periodos.length ? "Sin periodo"
        : periodos.length === 1 ? `${periodos[0].mes} ${periodos[0].anio}`
        : `${periodos[0].mes} ${periodos[0].anio} a ${periodos[periodos.length - 1].mes} ${periodos[periodos.length - 1].anio}`;

      const wbx = new ExcelJS.Workbook();
      // El nombre de hoja de Excel no admite : \\ / ? * [ ] y topa en 31.
      const nombreHoja = `${unidad} ${etiquetaPeriodo}`.replace(/[:\\/?*[\]]/g, "-").slice(0, 31);
      const ws = wbx.addWorksheet(nombreHoja);
      ws.columns = COLS.map((c) => ({ width: c.width }));

      const totalPorMoneda = (lista) => {
        const t = {};
        lista.forEach((p) => {
          const m = (p.moneda || "MXP") === "USD" ? "USD" : "MXP";
          t[m] = (t[m] || 0) + (Number(p.monto_estimado) || 0);
        });
        return t;
      };
      /* En los subtotales se etiquetan AMBAS monedas. `money()` solo marca los
         dólares —da por hecho que sin etiqueta son pesos— y en una fila donde
         conviven las dos, "$211,039.96 · $25,000.00 USD" invita a leer la
         primera cifra como parte del mismo total. */
      const fmtTot = (t) => Object.entries(t)
        .map(([m, v]) => `$${(Number(v) || 0).toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${m}`)
        .join("   ·   ");

      // Encabezado del documento
      const tit = ws.addRow([`Reporte Presupuestal — ${unidad} — ${etiquetaPeriodo}`]);
      tit.font = { bold: true, size: 14 };
      ws.mergeCells(1, 1, 1, Math.max(COLS.length, 2));
      const sub = ws.addRow([
        `${partidasOrdenadas.length} partidas   ·   ${fmtTot(totalPorMoneda(partidasOrdenadas))}` +
        (groupKeys.length ? `   ·   agrupado por ${groupBys.map((g) => (GROUP_OPCIONES.find((o) => o.value === g.field) || {}).label || g.field).join(" > ")}` : ""),
      ]);
      sub.font = { size: 10, color: { argb: "FF6B7785" } };
      ws.mergeCells(2, 1, 2, Math.max(COLS.length, 2));
      ws.addRow([]);

      const hr = ws.addRow(COLS.map((c) => c.header));
      hr.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
        cell.alignment = { horizontal: "center", vertical: "center" };
      });

      const filaDe = (p, sangria) => {
        const row = ws.addRow(COLS.map((c) => c.get(p)));
        COLS.forEach((c, i) => {
          const cell = row.getCell(i + 1);
          if (c.money) cell.numFmt = '"$"#,##0.00';
          if (c.izq) cell.alignment = { horizontal: "left", indent: sangria };
        });
        return row;
      };

      // Recorre el árbol agrupado. Cada grupo abre con su nombre, cuántas
      // partidas contiene y su subtotal por moneda.
      const hojas = (nodo) => nodo.type === "rows" ? nodo.rows : nodo.entries.flatMap((e) => hojas(e.child));
      const etiquetaCampo = (campo) =>
        (GROUP_OPCIONES.find((o) => o.value === campo) || {}).label || campo;

      if (groupKeys.length && grouped) {
        (function recorrer(nodo, nivel) {
          if (nodo.type === "rows") { nodo.rows.forEach((p) => filaDe(p, nivel)); return; }
          nodo.entries.forEach((e) => {
            const lista = hojas(e.child);
            const row = ws.addRow([`${etiquetaCampo(nodo.key)}: ${e.value}   (${lista.length})   ${fmtTot(totalPorMoneda(lista))}`]);
            ws.mergeCells(row.number, 1, row.number, Math.max(COLS.length, 2));
            row.getCell(1).font = { bold: true, color: { argb: nivel === 0 ? "FF232A31" : "FF4A5560" } };
            row.getCell(1).fill = { type: "pattern", pattern: "solid",
              fgColor: { argb: nivel === 0 ? "FFECEEF1" : "FFF6F7F9" } };
            row.getCell(1).alignment = { indent: nivel };
            recorrer(e.child, nivel + 1);
          });
        })(grouped, 0);
      } else {
        partidasOrdenadas.forEach((p) => filaDe(p, 0));
      }

      // Totales por moneda: nunca sumadas entre sí.
      ws.addRow([]);
      Object.entries(totalPorMoneda(partidasOrdenadas)).forEach(([m, v]) => {
        const row = ws.addRow([`TOTAL ${m}`, ...Array(Math.max(COLS.length - 2, 0)).fill(""), v]);
        row.font = { bold: true };
        row.getCell(COLS.length).numFmt = '"$"#,##0.00';
      });

      ws.views = [{ state: "frozen", ySplit: 4 }];

      const buf = await wbx.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Reporte Presupuestal - ${unidad} - ${etiquetaPeriodo}.xlsx`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      alert("No se pudo generar el reporte: " + (err.message || err));
    } finally {
      setGenerandoReporteXls(false);
    }
  };

  const COLUMNAS_PARTIDA = [
    { key: "mes", label: "Mes", render: (p) => p.mes },
    { key: "anio", label: "Año", render: (p) => p.anio },
    {
      key: "concepto", label: "Concepto",
      render: (p) => (
        <span>
          {p.concepto}
        </span>
      ),
    },
    { key: "rubro", label: "Rubro", render: (p) => <Pill>{p.rubro}</Pill> },
    { key: "categoria", label: "Categoría", render: (p) => <span style={{ color: T.textDim }}>{p.categoria}</span> },
    { key: "zona", label: "Zona", render: (p) => p.zona ? <Pill>{p.zona}</Pill> : <span style={{ color: T.textFaint }}>Cualquiera</span> },
    { key: "proyecto", label: "Proyecto", render: (p) => p.proyecto },
    { key: "folio", label: "Folio", render: (p) => <span style={{ fontFamily: T.fontMono, color: T.textDim }}>{p.folio || "—"}</span> },
    {
      key: "monto_estimado", label: "Monto",
      render: (p) => {
        const usado = usadoDe(p);
        const pct = p.monto_estimado ? (usado / p.monto_estimado) * 100 : 0;
        const tone = pct > 100 ? T.red : pct > 85 ? T.amber : T.teal;
        return (
          <span
            style={{ fontFamily: T.fontMono, borderBottom: usado > 0 ? `1px dashed ${tone}` : "none", cursor: usado > 0 ? "help" : "default" }}
            title={usado > 0 ? `Ejercido: ${money(usado, p.moneda)} (${pct.toFixed(0)}%)` : undefined}
          >
            {money(p.monto_estimado, p.moneda)}
          </span>
        );
      },
    },
    { key: "updated_at", label: "Última actualización", render: (p) => <span style={{ fontSize: 11, color: T.textFaint }}>{formatFechaHora(p.updated_at) || "—"}</span> },
  ];
  const colVisibility = useColumnVisibility("colv-partidas", COLUMNAS_PARTIDA);
  const columnasVisiblesBase = COLUMNAS_PARTIDA.filter((c) => (c.key === "proyecto" || !groupKeys.includes(c.key)) && !colVisibility.hidden.has(c.key));
  const colWidths = useColumnWidths("colw-partidas");
  const { ordered: columnasVisibles, moveColumn } = useColumnOrder("colo-partidas", columnasVisiblesBase);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };
  const [expandedIds, setExpandedIds] = useState(new Set());
  const [transaccionEditando, setTransaccionEditando] = useState(null);
  const toggleExpand = (id) => setExpandedIds((prev) => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });

  const [generandoPDF, setGenerandoPDF] = useState(false);

  /**
   * Reporte de Presupuesto Mensual — versión ejecutiva.
   *
   * No lista partidas: responde cuatro preguntas. Cuánto en pesos, cuánto en
   * dólares, en dónde se reparte (proyecto y zona) y en qué (rubro). Un
   * listado de ochenta renglones obliga a que quien lo recibe haga el
   * resumen; esto se lo da hecho.
   *
   * Cada corte lleva su porcentaje sobre el total de SU moneda. El porcentaje
   * es lo que hace comparable un rubro contra otro sin tener que dividir
   * mentalmente, y calcularlo por moneda evita que un renglón en dólares se
   * vea diminuto junto a los pesos.
   */
  const generarReportePDF = async () => {
    if (!partidasOrdenadas.length) {
      alert("No hay partidas en el filtro actual para generar el reporte.");
      return;
    }
    if (!confirm(`Generar el Reporte de Presupuesto Mensual con ${partidasOrdenadas.length} partida(s) resumidas por rubro, proyecto y zona.`)) return;

    setGenerandoPDF(true);
    try {
      const monedaDe = (p) => ((p.moneda || "MXP") === "USD" ? "USD" : "MXP");
      const totales = {};
      partidasOrdenadas.forEach((p) => {
        const m = monedaDe(p);
        totales[m] = (totales[m] || 0) + (Number(p.monto_estimado) || 0);
      });
      const monedas = Object.keys(totales).sort();

      /* Agrupa por un campo y devuelve filas ordenadas de mayor a menor. Lo
         vacío se rotula en vez de omitirse: una partida sin zona sigue siendo
         presupuesto, y esconderla haría que los cortes no sumaran el total. */
      const cortePor = (campo, vacio) => {
        const mapa = {};
        partidasOrdenadas.forEach((p) => {
          const clave = String(p[campo] || "").trim() || vacio;
          const m = monedaDe(p);
          if (!mapa[clave]) mapa[clave] = { clave, MXP: 0, USD: 0, n: 0 };
          mapa[clave][m] += Number(p.monto_estimado) || 0;
          mapa[clave].n++;
        });
        return Object.values(mapa).sort((a, b) => (b.MXP + b.USD) - (a.MXP + a.USD));
      };

      const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "letter" });
      const anchoUtil = 552; // carta vertical menos márgenes
      const periodo = [...new Set(partidasOrdenadas.map((p) => `${p.mes} ${p.anio}`))].join(" · ");

      doc.setFontSize(15);
      doc.setTextColor(35, 42, 49);
      doc.text(`Presupuesto Mensual — ${unidad}`, 30, 36);
      doc.setFontSize(9.5);
      doc.setTextColor(120);
      doc.text(`${periodo || "Sin periodo"}   ·   ${partidasOrdenadas.length} partidas`, 30, 52);

      // Los totales van arriba y en grande: son la respuesta a la primera
      // pregunta, y quien lo lee no debería tener que buscarlos.
      let y = 76;
      monedas.forEach((m, i) => {
        const x = 30 + i * 270;
        doc.setFillColor(m === "USD" ? 30 : 62, m === "USD" ? 143 : 92, m === "USD" ? 115 : 118);
        doc.rect(x, y, 250, 46, "F");
        doc.setTextColor(255);
        doc.setFontSize(9);
        doc.text(`TOTAL ${m}`, x + 12, y + 17);
        doc.setFontSize(16);
        doc.text(money(totales[m], m), x + 12, y + 37);
      });
      y += 68;

      const seccion = (titulo, filas) => {
        if (!filas.length) return;
        const cuerpo = filas.map((f) => {
          const cel = [f.clave, String(f.n)];
          monedas.forEach((m) => {
            const pct = totales[m] ? (f[m] / totales[m]) * 100 : 0;
            cel.push(f[m] ? `${money(f[m], m)}   ${pct.toFixed(1)}%` : "—");
          });
          return cel;
        });
        autoTable(doc, {
          startY: y,
          head: [[titulo, "Partidas", ...monedas.map((m) => `Importe ${m}`)]],
          body: cuerpo,
          styles: { fontSize: 8.5, cellPadding: 5 },
          headStyles: { fillColor: [62, 92, 118], textColor: 255, halign: "left" },
          bodyStyles: { halign: "left" },
          columnStyles: {
            0: { cellWidth: monedas.length > 1 ? 170 : 240 },
            1: { halign: "center", cellWidth: 55 },
            ...Object.fromEntries(monedas.map((_, i) => [i + 2, { halign: "right" }])),
          },
          margin: { left: 30, right: 30 },
          tableWidth: anchoUtil,
        });
        y = doc.lastAutoTable.finalY + 22;
      };

      // El orden responde a la pregunta: primero EN QUÉ se gasta, luego DÓNDE.
      seccion("Rubro", cortePor("rubro", "Sin rubro"));

      /* Apartado de flotilla.
         Combustible y mantenimiento vehicular viven en RUBROS DISTINTOS
         —Vehículos y Servicios de Mantenimiento—, así que el corte por rubro
         los separa justo cuando interesa verlos juntos: entre los dos suelen
         ser la mitad del presupuesto.
         Se clasifican por palabras clave sobre concepto y categoría, así que
         el reporte dice cuántas partidas cayó en cada bolsa: si el número no
         cuadra con lo esperado, hay algo mal escrito o mal capturado. */
      const CLAVES_FLOTILLA = [
        { bolsa: "Combustible", claves: ["combustible", "gasolina", "diesel", "diésel"] },
        { bolsa: "Mantenimiento vehicular", claves: [
          "mantto unidad", "mantto. unidad", "mantenimiento unidad", "mantto vehic",
          "mantenimiento vehic", "unidades ligeras", "unidades pesadas",
          "unidad vehicular", "unidades vehiculares", "llanta", "neumatico", "neumático",
          "acumulador", "refaccion", "refacción", "verificacion", "verificación",
        ] },
      ];
      const normTexto = (v) => String(v || "")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[.,;:()]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
      const bolsaDe = (p) => {
        const t = normTexto(p.concepto) + " " + normTexto(p.categoria);
        for (const { bolsa, claves } of CLAVES_FLOTILLA) {
          if (claves.some((k) => t.includes(normTexto(k)))) return bolsa;
        }
        return null;
      };
      const deFlotilla = partidasOrdenadas.filter((p) => bolsaDe(p));
      if (deFlotilla.length) {
        const mapa = {};
        deFlotilla.forEach((p) => {
          // El desglose es por ZONA: en flotilla la pregunta operativa es
          // dónde se está gastando, no bajo qué rubro quedó clasificado.
          const clave = `${bolsaDe(p)} · ${String(p.zona || "").trim() || "Sin zona"}`;
          const m = monedaDe(p);
          if (!mapa[clave]) mapa[clave] = { clave, MXP: 0, USD: 0, n: 0 };
          mapa[clave][m] += Number(p.monto_estimado) || 0;
          mapa[clave].n++;
        });
        const filas = Object.values(mapa).sort((a, b) => a.clave.localeCompare(b.clave));

        // Subtotal de la flotilla completa, que es el número que se busca.
        const sub = { clave: "TOTAL FLOTILLA", MXP: 0, USD: 0, n: deFlotilla.length };
        deFlotilla.forEach((p) => { sub[monedaDe(p)] += Number(p.monto_estimado) || 0; });
        filas.push(sub);

        seccion("Flotilla: combustible y mantenimiento vehicular", filas);
      }
      seccion("Proyecto", cortePor("proyecto", "Sin proyecto"));
      const porZona = cortePor("zona", "Cualquier zona");
      // Si ninguna partida tiene zona, el corte no aporta nada y se omite.
      if (!(porZona.length === 1 && porZona[0].clave === "Cualquier zona")) {
        seccion("Zona", porZona);
      }

      doc.save(`presupuesto-mensual-${unidad}-${new Date().toISOString().slice(0, 10)}.pdf`);
    } catch (err) {
      alert("No se pudo generar el reporte: " + (err.message || err));
    } finally {
      setGenerandoPDF(false);
    }
  };

  const renderRowTr = (p, depth = 0, n) => {
    const expandido = expandedIds.has(p.id);
    const transDeEsta = transacciones.filter((t) => t.partida_id === p.id);
    return (
      <React.Fragment key={p.id}>
        <tr>
          <td style={{ ...tdStyle, width: 36, textAlign: "right", color: T.textFaint, fontFamily: T.fontMono, fontSize: 11 }}>{n}</td>
          <td style={{ ...tdStyle, width: 30, textAlign: "center" }}>
            {transDeEsta.length > 0 && (
              <button
                type="button"
                onClick={() => toggleExpand(p.id)}
                title={expandido ? "Ocultar transacciones" : `Ver ${transDeEsta.length} transacción(es)`}
                style={{ background: "none", border: "none", cursor: "pointer", color: T.textFaint, fontSize: 11, padding: 2 }}
              >
                {expandido ? "▼" : "▶"}
              </button>
            )}
          </td>
          {columnasVisibles.map((c, i) => (
            <td key={c.key} style={i === 0 && depth ? { ...tdStyle, paddingLeft: 14 + depth * 26 } : tdStyle}>{c.render(p)}</td>
          ))}
          <td style={tdStyle}>
            <div style={{ display: "flex", gap: 4 }}>
              <IconButton icon="✎" label="Editar" tone={T.accent} onClick={() => startEdit(p)} />
              <IconButton icon="⧉" label="Duplicar" tone={T.textDim} onClick={() => duplicar(p)} />
              <IconButton icon="✕" label="Eliminar" tone={T.red} onClick={() => remove(p.id)} />
            </div>
          </td>
        </tr>
        {expandido && (
          <tr>
            {/* Cajón embutido. Tres señales de anidamiento, cada una haciendo un
                trabajo distinto: el riel de acento ata el bloque a su fila padre,
                la tarjeta blanca sobre el gris lo separa como otra superficie, y
                el resumen en minúsculas evita el choque con los encabezados de
                columna y las píldoras de agrupamiento, que ya usan mayúsculas. */}
            <td colSpan={columnasVisibles.length + 3} style={{ padding: 0, background: T.panelAlt, borderBottom: `1px solid ${T.border}` }}>
              <div style={{ padding: "10px 14px 12px 40px" }}>
                <div style={{ borderLeft: `3px solid ${T.accent}`, borderRadius: "0 6px 6px 0", background: T.panel, boxShadow: "0 1px 2px rgba(35,42,49,0.06)", overflow: "hidden" }}>
                  <div style={{ display: "flex", alignItems: "baseline", gap: 8, padding: "8px 12px 6px", borderBottom: `1px solid ${T.borderSoft}` }}>
                    <span style={{ fontSize: 12, fontWeight: 600, color: T.text }}>
                      {transDeEsta.length === 1 ? "1 transacción vinculada" : `${transDeEsta.length} transacciones vinculadas`}
                    </span>
                    <span style={{ fontSize: 11.5, color: T.textFaint, fontFamily: T.fontMono }}>
                      {[...new Set(transDeEsta.map((t) => t.moneda || "MXP"))]
                        .map((m) => money(transDeEsta.filter((t) => (t.moneda || "MXP") === m).reduce((a, t) => a + (Number(t.importe) || 0), 0), m))
                        .join("  ·  ")}
                    </span>
                  </div>
                  <table style={{ ...tableStyle, fontSize: 11.5 }}>
                    <tbody>
                      {transDeEsta.map((t, idx) => (
                        <tr key={t.id}>
                          <td style={{ ...tdStyle, padding: "7px 12px", color: T.textDim, borderBottom: idx === transDeEsta.length - 1 ? "none" : `1px solid ${T.borderSoft}` }}>
                            {t.concepto_detallado || "—"}
                          </td>
                          <td style={{ ...tdStyle, padding: "7px 12px", fontFamily: T.fontMono, textAlign: "right", whiteSpace: "nowrap", borderBottom: idx === transDeEsta.length - 1 ? "none" : `1px solid ${T.borderSoft}` }}>
                            {money(t.importe, t.moneda)}
                          </td>
                          <td style={{ ...tdStyle, padding: "7px 12px", width: 110, borderBottom: idx === transDeEsta.length - 1 ? "none" : `1px solid ${T.borderSoft}` }}>
                            {t.status ? <Pill tone={t.status === "Pagado" ? "teal" : "amber"}>{t.status}</Pill> : "—"}
                          </td>
                          <td style={{ ...tdStyle, padding: "7px 12px", width: 44, borderBottom: idx === transDeEsta.length - 1 ? "none" : `1px solid ${T.borderSoft}` }}>
                            <IconButton icon="✎" label="Editar transacción" tone={T.accent} onClick={() => setTransaccionEditando(t)} />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </td>
          </tr>
        )}
      </React.Fragment>
    );
  };


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
  const duplicar = (p) => {
    // Copia la partida como registro NUEVO. El folio se limpia a propósito:
    // submit() respeta el que venga en el formulario, así que heredarlo haría
    // chocar el duplicado con la original. Vacío, se genera uno nuevo.
    const { id, folio, created_at, updated_at, created_by, updated_by, ...resto } = p;
    const limpio = Object.fromEntries(Object.entries(resto).filter(([k]) => !k.startsWith("_")));
    setForm({ ...limpio, folio: "" });
    setEditId(null);
    setModalOpen(true);
  };
  const startEdit = (p) => { setForm(p); setEditId(p.id); setModalOpen(true); };

  const closeModal = () => { setModalOpen(false); setEditId(null); setForm({ ...blank, anio: anioDefault, proyecto: marcadores[0] || "" }); };
  const remove = (id) => {
    const p = partidasUnidad.find((x) => x.id === id);
    const vinculadas = transacciones.filter((t) => t.partida_id === id);
    const pagadas = vinculadas.filter((t) => t.status === "Pagado");

    // La regla de las transacciones pagadas se hereda: borrar la partida las
    // dejaría apuntando a nada, con lo que el gasto desaparecería del
    // presupuesto sin dejar rastro. Es el mismo daño, por la puerta de atrás.
    if (pagadas.length) {
      const total = pagadas.reduce((sum, t) => sum + (Number(t.importe) || 0), 0);
      alert(
        `No se puede eliminar la partida "${p?.concepto || id}" (folio ${p?.folio || "—"}).\n\n` +
        `Tiene ${pagadas.length} transacción(es) marcadas como Pagadas por ${money(total, p?.moneda)}. ` +
        `Borrarla las dejaría sin partida y ese gasto desaparecería del presupuesto.\n\n` +
        `Primero reasígnalas a otra partida, o cámbiales el status a "No Pagado".`
      );
      return;
    }

    // Las no pagadas sí dejan borrar, pero se avisa: quedarán sin vincular.
    const aviso = vinculadas.length
      ? `\n\nOJO: ${vinculadas.length} transacción(es) sin pagar quedarán sin partida vinculada. Van a aparecer en "Transacciones importadas sin partida vinculada".`
      : "";
    if (!confirm(`¿Eliminar la partida "${p?.concepto || id}" (folio ${p?.folio || "—"})? Esto no se puede deshacer.${aviso}`)) return;
    partidasApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel
        title={`Partidas de ${unidad}`}
        subtitle={filtrosActivos ? `${partidasFiltradas.length} de ${partidasUnidad.length} registradas` : `${partidasUnidad.length} registradas`}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <ColumnVisibilityControl
              columns={COLUMNAS_REPORTE_PARTIDAS}
              hidden={repVis.hidden}
              onToggle={repVis.toggle}
              onShowAll={repVis.showAll}
              etiqueta="Columnas del reporte"
            />
            <Button variant="ghost" onClick={generarReporteExcel} disabled={generandoReporteXls}
              title="Excel con las columnas elegidas, respetando el agrupamiento de la vista">
              {generandoReporteXls ? "Generando…" : "Reporte Excel"}
            </Button>
            <Button variant="ghost" onClick={generarReportePDF} disabled={generandoPDF}>
              {generandoPDF ? "Generando…" : "Reporte de Presupuesto Mensual"}
            </Button>
            <Button variant="ghost" onClick={() => exportarExcel(false)}>Exportar {unidad}</Button>
            <Button variant="ghost" onClick={() => exportarExcel(true)} title="Las tres compañías en hojas separadas, con los mismos filtros — para el preparador de transacciones">Exportar las 3</Button>
            <Button onClick={openNew}>+ Nueva partida</Button>
          </div>
        }
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
            <MesMultiSelect mesesDisponibles={mesesDisponiblesFiltro} seleccionados={filtrosMes} onChange={(nuevo) => setFiltros({ ...filtros, mes: nuevo })} />
          </Field>
          <Field label="Año">
            <AnioMultiSelect aniosDisponibles={aniosDisponiblesFiltro} seleccionados={filtrosAnio} onChange={(nuevo) => setFiltros({ ...filtros, anio: nuevo })} />
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
          {groupKeys.length > 0 && (
            <Button variant="ghost" onClick={() => setCollapsedGroups(new Set(collectGroupPaths(grouped)))}>Contraer todo</Button>
          )}
          {expandedIds.size > 0 && (
            <Button variant="ghost" onClick={() => setExpandedIds(new Set())}>Contraer transacciones</Button>
          )}
          <ColumnVisibilityControl
            columns={COLUMNAS_PARTIDA}
            hidden={colVisibility.hidden}
            onToggle={colVisibility.toggle}
            onShowAll={colVisibility.showAll}
          />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table style={{ ...tableStyle, tableLayout: "fixed" }}>
            <colgroup>
              <col style={{ width: 36 }} />
              <col style={{ width: 30 }} />
              {columnasVisibles.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
              <col style={{ width: 140 }} />
            </colgroup>
            <thead>
              <tr>
                <th style={thStyle}>#</th>
                <th style={thStyle}></th>
                {columnasVisibles.map((c) => (
                  <SortableTh
                    key={c.key} label={c.label} sortKey={c.key} sort={sort} setSort={setSort}
                    width={colWidths.getWidth(c.key)} onResizeStart={(e) => colWidths.startResize(c.key, e)}
                    onDragStart={(e) => onColDragStart(e, c.key)} onDragOver={onColDragOver} onDrop={(e) => onColDrop(e, c.key)}
                    sumLabel={c.key === "monto_estimado" ? sumaPorMoneda(partidasOrdenadas, "monto_estimado") : undefined}
                  />
                ))}
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {groupKeys.length
                ? buildGroupedTrs(grouped, "", collapsedGroups, toggleGroup, columnasVisibles.length + 3, 0, renderRowTr, Object.fromEntries(GROUP_OPCIONES.map((o) => [o.value, o.label])))
                : partidasOrdenadas.map((p, i) => renderRowTr(p, 0, i + 1))}
              {!partidasUnidad.length && (
                <tr><td colSpan={columnasVisibles.length + 3} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin partidas aún</td></tr>
              )}
              {partidasUnidad.length > 0 && !partidasFiltradas.length && (
                <tr><td colSpan={columnasVisibles.length + 3} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna partida coincide con estos filtros</td></tr>
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
          <AutoriaCaption record={form} perfilesApi={perfilesApi} />
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
              {/* Se sugiere, no se impone: quien captura sabe cosas que el
                  concepto no dice. Solo aparece cuando difiere de lo elegido. */}
              {(() => {
                const sug = sugerirCategoria(form.concepto, categoriasDisponibles);
                if (!sug || sug === form.categoria) return null;
                return (
                  <button
                    type="button"
                    onClick={() => setForm({ ...form, categoria: sug })}
                    style={{
                      marginTop: 5, background: "transparent", border: "none", padding: 0,
                      color: T.accent, fontSize: 11.5, cursor: "pointer", textAlign: "left",
                      fontFamily: T.fontUI,
                    }}
                  >
                    Por el concepto, quizá sea <b>{sug}</b> — usar
                  </button>
                );
              })()}
            </Field>
            <Field label="Concepto" style={{ gridColumn: "span 2" }}>
              <TextInput value={form.concepto} onChange={(e) => setForm({ ...form, concepto: e.target.value })} placeholder="Ej. Servicio energía eléctrica base admtva" />
            </Field>
            <Field label="Zona">
              {/* Opcional a propósito: vacío significa que el presupuesto no
                  está atado a una zona, como Nómina o los servicios legales. */}
              <Select value={form.zona || ""} onChange={(e) => setForm({ ...form, zona: e.target.value })}>
                <option value="">— Cualquier zona —</option>
                {zonas.map((z) => <option key={z}>{z}</option>)}
                {/* Si la partida trae una zona que salió del catálogo, se
                    conserva como opción para no perderla al editar. */}
                {form.zona && !zonas.includes(form.zona) && <option key={form.zona}>{form.zona}</option>}
              </Select>
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
      {transaccionEditando && (
        <TransaccionQuickEditModal
          partidasUnidad={partidasUnidad}
          transaccion={transaccionEditando}
          onClose={() => setTransaccionEditando(null)}
          transaccionesApi={transaccionesApi}
          proveedoresApi={proveedoresApi}
          cuentasApi={cuentasApi}
          unidad={unidad}
        />
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------
   TABS: TRANSACCIONES
---------------------------------------------------------------------- */
function ImportarTransaccionesPanel({ partidas, proveedores, cuentas = [], transaccionesApi }) {
  const inputRef = useRef(null);
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [importing, setImporting] = useState(false);

  const descargarPlantilla = async () => {
    const wbx = new ExcelJS.Workbook();
    const headers = [
      "Dia", "SMI", "SOLICITANTE", "Proyecto", "Zona", "ÁREA",
      "NOMBRE/DENOMINACIÓN O RAZON SOCIAL", "CONCEPTO DE PAGO(DETALLADO",
      "FOLIO COMPRA SAE", "FOLIO FACTURA", "FORMA DE PAGO", "MÉTODO DE PAGO",
      "IMPORTE", "Moneda", "A Partida", "Status",
    ];
    const ejemplo = [
      "11/08/2026", "", "Mariel Diaz", "Todos", "Queretaro", "Administracion",
      "Comision Federal De Electricidad", "Servicio energía base operativa",
      "", "", "03 - Transferencia electrónica de fondos", "PUE - Pago en una sola exhibición",
      5137, "MXP", "CTM-AGO26-003", "No Pagado",
    ];
    UNIDAD_KEYS.forEach((u) => {
      const ws = wbx.addWorksheet(u);
      const headerRow = ws.addRow(headers);
      formatearHojaDatos(ws, headerRow, headers.length);
      const ejemploRow = ws.addRow(ejemplo);
      ejemploRow.eachCell((cell) => { cell.font = { italic: true, color: { argb: "FF8B99A6" } }; });
      ws.columns = headers.map(() => ({ width: 16 }));
      ws.getColumn(7).width = 32; ws.getColumn(8).width = 40;
      ws.getColumn(11).width = 22; ws.getColumn(12).width = 22;
    });
    const notas = wbx.addWorksheet("Instrucciones");
    [
      "Cómo usar esta plantilla",
      "",
      "1. El nombre de cada hoja (OSB, CTM, ISE) le dice a la app de qué compañía",
      "   son esas transacciones — no cambies esos nombres.",
      "2. BORRA la fila de ejemplo (fila 2, en cursiva gris) de cada hoja antes de subir.",
      "3. A Partida: el folio EXACTO de la partida a la que corresponde ese pago",
      "   (ej. CTM-AGO26-003). Si no lo tienes a la mano, déjalo vacío — la",
      "   transacción se importa igual, solo queda \"sin vincular\" para asignarla después.",
      "4. Proyecto: un proyecto real del Catálogo, o los marcadores Desh Gral /",
      "   Prod Gral / Todos.",
      "5. Zona: una de las zonas del Catálogo (Queretaro, Poza Rica, Paraiso, etc.).",
      "6. Status: Pagado o No Pagado.",
      "7. Moneda: MXP o USD.",
    ].forEach((linea, i) => { notas.getCell(`A${i + 1}`).value = linea; });
    notas.getCell("A1").font = { bold: true };
    notas.getColumn(1).width = 90;

    const buffer = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "plantilla_transacciones.xlsx";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(""); setStatus(""); setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const { rows, sheetsFound, matched, unmatched, conCuenta, sinCuenta, desajustes } = parseTransaccionesWorkbook(buf, partidas, proveedores, cuentas);
      if (!rows.length) {
        setError('No encontré una hoja con columnas "Día", "Importe" y "A Partida" en este archivo.');
        return;
      }
      setPreview({ rows, sheetsFound, matched, unmatched, conCuenta, sinCuenta, desajustes, fileName: file.name });
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
      // _cuentasDisponibles es un dato de pantalla, no de la tabla: se descarta
      // antes de insertar o Supabase rechaza la columna desconocida.
      const toInsert = preview.rows.map(({ _cuentasDisponibles, _desajusteMoneda, _categoriaSugerida, ...r }) => ({
        ...r,
        partida_id: r.partida_id || null,
        proveedor_id: r.proveedor_id || null,
        cuenta_id: r.cuenta_id || null,
      }));
      await transaccionesApi.bulkInsert(toInsert);
      const vinculadas = preview.rows.filter((r) => r.partida_id).length;
      const conCta = preview.rows.filter((r) => r.cuenta_id).length;
      setStatus(`Importadas ${preview.rows.length} transacciones (${vinculadas} con partida, ${conCta} con cuenta bancaria asignada).`);
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
    <Panel
      title="Carga masiva de transacciones reales"
      subtitle='Sube el registro de pagos (columnas Día, Solicitante, Área, Proveedor, Importe, A Partida) — se vincula por el folio de la partida'
      right={<Button variant="ghost" onClick={descargarPlantilla}>Descargar plantilla</Button>}
    >
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
            {preview.desajustes > 0 && <Pill tone="amber">{preview.desajustes} con moneda distinta</Pill>}
            {preview.rows.filter((r) => r._categoriaSugerida).length > 0 && (
              <Pill tone="accent">
                {preview.rows.filter((r) => r._categoriaSugerida).length} con categoría deducida del concepto
              </Pill>
            )}
          </div>
          {preview.desajustes > 0 && (
            <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "10px 13px", borderRadius: "0 6px 6px 0", fontSize: 12, marginBottom: 12 }}>
              <b>{preview.desajustes} fila(s) traen un folio de partida en otra moneda</b>
              No se vincularon a propósito: comparar un gasto en una moneda contra un presupuesto
              en otra no significa nada, y si la partida tiene prorrateo el importe se reparte entre
              todos los proyectos del Dashboard. Corrige la moneda en el archivo, o apunta esas
              filas a una partida de la misma moneda con el selector de cada renglón.
            </div>
          )}
          {sinVincularActuales > 0 && (
            <div style={{ fontSize: 11.5, color: T.amber, marginBottom: 12 }}>
              Para las filas sin folio coincidente, elige la partida correcta en el selector de esa fila antes de confirmar — o déjalas sin vincular e impórtalas igual, podrás asignarlas después.
            </div>
          )}
          <div style={{ overflowX: "auto", maxHeight: 340, overflowY: "auto", border: `1px solid ${T.borderSoft}`, borderRadius: 6 }}>
            <table style={tableStyle}>
              <thead>
                <tr>{["Día de Pago Programado","Folio","Proveedor","Importe","Partida"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
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

function TransaccionesTab({ unidad, unidades, partidas, partidasApi, transacciones, transaccionesApi, proveedoresApi, cuentasApi, perfilesApi, notasApi, session, zonas = ZONAS_RESPALDO }) {
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadoresProyecto = marcadoresDisponibles(proyectosUnidad);
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const blank = {
    partida_id: partidasUnidad[0]?.id || "", unidad_detectada: unidad, dia: "", solicitante: "", smi: "", proyecto: "", zona: "", area: "",
    proveedor: "", proveedor_id: "", cuenta_id: "", concepto_detallado: "", importe: "", moneda: "MXP", status: "No Pagado", fecha_pago: "",
    folio_compra_sae: "", folio_factura: "", forma_pago: "", metodo_pago: "", referencia_pago: "", categoria: "",
  };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);
  // La Solicitud de Pago se genera desde una transacción, así que su estado
  // vive aquí. Estaba en PartidasTab y el botón llamaba a un setter de otro
  // componente: fallaba en la consola sin que nada visible ocurriera.
  const [sppDe, setSppDe] = useState(null);
  const cuentasDelProveedorSeleccionado = form.proveedor_id ? cuentasApi.rows.filter((c) => c.proveedor_id === form.proveedor_id) : [];
  const transUnidad = transacciones.filter((t) => t.unidad_detectada === unidad);
  const sinVincular = transacciones.filter((t) => !t.partida_id && t.unidad_detectada === unidad);

  const [filtros, setFiltros] = useSessionState("ss-transacciones-filtros", { texto: "", fechaDesde: "", fechaHasta: "", reportado: "Todos", enviadoPagos: "Todos" });
  const [sort, setSort] = usePrefState("pref-transacciones-sort", { key: "dia", dir: "desc" }, sanearSort(["dia","folio_transaccion","partida","proveedor","proyecto","zona","area","concepto_detallado","importe","status","fecha_pago","updated_at"]));
  const [seleccionadas, setSeleccionadas] = useState(new Set());
  const [marcandoReportado, setMarcandoReportado] = useState(false);
  const [marcandoEnviado, setMarcandoEnviado] = useState(false);

  const partidaDe = (t) => partidasUnidad.find((p) => p.id === t.partida_id);

  const transFiltradas = transUnidad.filter((t) => {
    if (filtros.fechaDesde && (!t.dia || t.dia < filtros.fechaDesde)) return false;
    if (filtros.fechaHasta && (!t.dia || t.dia > filtros.fechaHasta)) return false;
    if (filtros.reportado === "Reportado" && !t.reportado_at) return false;
    if (filtros.reportado === "No reportado" && t.reportado_at) return false;
    if (filtros.enviadoPagos === "Enviado" && !t.enviado_pagos_at) return false;
    if (filtros.enviadoPagos === "No enviado" && t.enviado_pagos_at) return false;
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const partida = partidaDe(t);
      const enTexto = [t.proveedor, t.concepto_detallado, t.solicitante, t.zona, t.folio_transaccion, partida?.folio, partida?.concepto]
        .some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    return true;
  });
  const filtrosActivos = filtros.texto.trim() || filtros.fechaDesde || filtros.fechaHasta || filtros.reportado !== "Todos" || filtros.enviadoPagos !== "Todos";
  const limpiarFiltros = () => setFiltros({ texto: "", fechaDesde: "", fechaHasta: "", reportado: "Todos", enviadoPagos: "Todos" });

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
    return { ...t, _proyecto: p?.proyecto || SIN_DATO, _rubro: p?.rubro || SIN_DATO, _mes: p?.mes || SIN_DATO, _anio: p?.anio || null, _vinculo: t.partida_id ? "Vinculada" : "Sin vincular" };
  });

  const [groupBys, setGroupBys] = usePrefState("pref-transacciones-groupbys", [], sanearGroupBys(GROUP_OPCIONES_TRANS));
  const [collapsedGroups, setCollapsedGroups] = useState(new Set());
  const toggleGroup = (path) => setCollapsedGroups((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });
  const groupKeys = groupBys.map((g) => g.field);
  const grouped = groupKeys.length ? agruparRows(transEnriquecidas, groupBys, "importe") : null;

  // --- Tabla "sin vincular": mismos controles que la tabla principal (fechas, agrupar, columnas) ---
  // Indicador de vínculo del proveedor: sin catálogo / con catálogo sin cuenta / con cuenta bancaria.
  const renderProveedorConIndicador = (t) => {
    if (!t.proveedor && !t.proveedor_id) return "—";
    let tone = "red", texto = "Sin catálogo";
    if (t.proveedor_id) {
      const tieneCuenta = cuentasApi.rows.some((c) => c.proveedor_id === t.proveedor_id);
      tone = tieneCuenta ? "teal" : "amber";
      texto = tieneCuenta ? "Con cuenta" : "Sin cuenta";
    }
    return (
      <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
        <span>{t.proveedor || "—"}</span>
        <Pill tone={tone}>{texto}</Pill>
      </div>
    );
  };

  const COLUMNAS_SINVINC = [
    { key: "dia", label: "Día de Pago Programado", render: (t) => t.dia || "—" },
    { key: "folio_original", label: "Folio", render: (t) => <span style={{ fontFamily: T.fontMono, color: T.red }}>{t.folio_original || "(sin folio)"}</span> },
    { key: "proveedor", label: "Proveedor", render: renderProveedorConIndicador },
    { key: "zona", label: "Zona", render: (t) => t.zona || "—" },
    { key: "area", label: "Área", render: (t) => t.area || "—" },
    { key: "proyecto", label: "Proyecto", render: (t) => t.proyecto || "—" },
    { key: "concepto_detallado", label: "Concepto", render: (t) => <span style={{ color: T.textDim }}>{t.concepto_detallado}</span> },
    { key: "importe", label: "Importe", render: (t) => <span style={{ fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</span> },
    { key: "status", label: "Status", render: (t) => t.status || "—" },
  ];
  const [filtrosSV, setFiltrosSV] = useSessionState("ss-transacciones-sv-filtros", { fechaDesde: "", fechaHasta: "" });
  const [groupBysSV, setGroupBysSV] = usePrefState("pref-transacciones-sv-groupbys", [], sanearGroupBys(GROUP_OPCIONES_SINVINC));
  const [collapsedGroupsSV, setCollapsedGroupsSV] = useState(new Set());
  const toggleGroupSV = (path) => setCollapsedGroupsSV((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });
  const sinVincularFiltrado = sinVincular.filter((t) => {
    if (filtrosSV.fechaDesde && (!t.dia || t.dia < filtrosSV.fechaDesde)) return false;
    if (filtrosSV.fechaHasta && (!t.dia || t.dia > filtrosSV.fechaHasta)) return false;
    return true;
  });
  const filtrosSVActivos = filtrosSV.fechaDesde || filtrosSV.fechaHasta;
  const limpiarFiltrosSV = () => setFiltrosSV({ fechaDesde: "", fechaHasta: "" });
  const groupKeysSV = groupBysSV.map((g) => g.field);
  const groupedSV = groupKeysSV.length ? agruparRows(sinVincularFiltrado, groupBysSV, "importe") : null;
  const colVisibilitySV = useColumnVisibility("colv-sinvinc", COLUMNAS_SINVINC);
  const columnasVisiblesSVBase = COLUMNAS_SINVINC.filter((c) => !groupKeysSV.includes(c.key) && !colVisibilitySV.hidden.has(c.key));
  const colWidthsSV = useColumnWidths("colw-sinvinc");
  const { ordered: columnasSV, moveColumn: moveColumnSV } = useColumnOrder("colo-sinvinc", columnasVisiblesSVBase);
  const dragKeyRefSV = useRef(null);
  const onColDragStartSV = (e, key) => { dragKeyRefSV.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOverSV = (e) => e.preventDefault();
  const onColDropSV = (e, targetKey) => { e.preventDefault(); if (dragKeyRefSV.current) { moveColumnSV(dragKeyRefSV.current, targetKey); dragKeyRefSV.current = null; } };
  const renderRowSinVinc = (t, depth = 0) => (
    <tr key={t.id}>
      {columnasSV.map((c) => <td key={c.key} style={{ ...tdStyle, paddingLeft: depth ? 14 + depth * 26 : undefined }}>{c.render(t)}</td>)}
      <td style={tdStyle}>
        {/* Mismo criterio que el formulario: solo partidas de la moneda de ESTA
            transacción. Sin esto, vincular desde la tabla se saltaba el bloqueo
            entre monedas que el formulario sí aplica. */}
        <PartidaPickerButton
          partidas={partidasUnidad.filter((p) => mismaMoneda(p.moneda, t.moneda))}
          ocultasPorMoneda={partidasUnidad.filter((p) => !mismaMoneda(p.moneda, t.moneda)).length}
          moneda={t.moneda}
          transacciones={transUnidad}
          partidasApi={partidasApi}
          unidad={unidad}
          proyectosOpciones={marcadoresProyecto}
          value=""
          placeholder="Elegir partida…"
          onChange={(nuevoId) => {
            const nuevaPartida = partidasUnidad.find((p) => p.id === nuevoId);
            transaccionesApi
              .update(t.id, { partida_id: nuevoId, unidad_detectada: nuevaPartida ? nuevaPartida.unidad : t.unidad_detectada })
              .catch((err) => alert("No se pudo vincular: " + (err.message || err)));
          }}
        />
      </td>
    </tr>
  );

  const COLUMNAS_TRANS = [
    {
      key: "folio_transaccion", label: "ID",
      render: (t) => <span style={{ fontFamily: T.fontMono, color: T.accent, fontSize: 11 }}>{t.folio_transaccion || "—"}</span>,
    },
    { key: "dia", label: "Día de Pago Programado", render: (t) => t.dia || "—" },
    { key: "smi", label: "SMI", render: (t) => t.smi || "—" },
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
        <PartidaPickerButton
          partidas={partidasUnidad.filter((p) => mismaMoneda(p.moneda, t.moneda))}
          ocultasPorMoneda={partidasUnidad.filter((p) => !mismaMoneda(p.moneda, t.moneda)).length}
          moneda={t.moneda}
          transacciones={transUnidad}
          partidasApi={partidasApi}
          unidad={unidad}
          proyectosOpciones={marcadoresProyecto}
          value={t.partida_id || ""}
          allowClear
          onChange={(nuevoId) => {
            const nuevaPartida = partidasUnidad.find((p) => p.id === nuevoId);
            transaccionesApi
              .update(t.id, { partida_id: nuevoId || null, unidad_detectada: nuevaPartida ? nuevaPartida.unidad : t.unidad_detectada })
              .catch((err) => alert("No se pudo vincular: " + (err.message || err)));
          }}
        />
      ),
    },
    { key: "proveedor", label: "Proveedor", render: renderProveedorConIndicador },
    { key: "proyecto", label: "Proyecto", render: (t) => t.proyecto || "—" },
    { key: "zona", label: "Zona", render: (t) => t.zona || "—" },
    { key: "area", label: "Área", render: (t) => t.area || "—" },
    {
      key: "concepto_detallado", label: "Concepto",
      render: (t) => (
        <span style={{ color: T.textDim }}>
          {notaActual(t.id) && <span title="Tienes una nota privada aquí" style={{ marginRight: 5 }}>🔒</span>}
          {t.concepto_detallado}
        </span>
      ),
    },
    { key: "categoria", label: "Categoría", render: (t) => t.categoria ? <Pill>{t.categoria}</Pill> : <span style={{ color: T.textFaint }}>—</span> },
    { key: "importe", label: "Importe", render: (t) => <span style={{ fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</span> },
    { key: "status", label: "Status", render: (t) => t.status ? <Pill tone={t.status === "Pagado" ? "teal" : "amber"}>{t.status}</Pill> : "—" },
    { key: "fecha_pago", label: "Fecha de Pago", render: (t) => t.fecha_pago || "—" },
    {
      key: "reportado_at", label: "Reportado a Dirección",
      render: (t) => t.reportado_at
        ? <Pill tone="accent">{formatFechaCorta(t.reportado_at)}</Pill>
        : <Pill tone="amber">No reportado</Pill>,
    },
    {
      key: "enviado_pagos_at", label: "Enviado a Pagos",
      render: (t) => t.enviado_pagos_at
        ? <Pill tone="teal">{formatFechaCorta(t.enviado_pagos_at)}</Pill>
        : <Pill tone="amber">No enviado</Pill>,
    },
    { key: "updated_at", label: "Última actualización", render: (t) => <span style={{ fontSize: 11, color: T.textFaint }}>{formatFechaHora(t.updated_at) || "—"}</span> },
  ];
  const colVisibility = useColumnVisibility("colv-transacciones", COLUMNAS_TRANS);
  const columnasVisiblesBase = COLUMNAS_TRANS.filter((c) => !groupKeys.includes(c.key) && !colVisibility.hidden.has(c.key));
  const colWidths = useColumnWidths("colw-transacciones");
  const { ordered: columnasVisibles, moveColumn } = useColumnOrder("colo-transacciones", columnasVisiblesBase);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };
  const [eliminando, setEliminando] = useState(false);
  /**
   * Borra solo lo seleccionado. Reemplaza al botón "Borrar todas", que
   * eliminaba TODA la compañía con una sola confirmación — demasiado poder
   * para un clic, y sin forma de acotar el alcance.
   *
   * Las pagadas se conservan y se avisa cuántas: la regla es la misma que en
   * el borrado de una fila.
   */
  const eliminarSeleccionadas = async () => {
    const todas = [...transUnidad, ...sinVincular].filter((t) => seleccionadas.has(t.id));
    const pagadas = todas.filter((t) => t.status === "Pagado");
    const borrables = todas.filter((t) => t.status !== "Pagado");
    if (!borrables.length) {
      alert(`Las ${pagadas.length} transacciones seleccionadas están marcadas como Pagadas y no se pueden borrar. Cámbiales el status a "No Pagado" si de verdad hay que eliminarlas.`);
      return;
    }
    const total = borrables.reduce((sum, t) => sum + (Number(t.importe) || 0), 0);
    const aviso = pagadas.length ? `\n\nSe van a CONSERVAR ${pagadas.length} marcadas como Pagadas.` : "";
    if (!confirm(`¿Eliminar ${borrables.length} transacción(es) por ${money(total)}? Esto no se puede deshacer.${aviso}`)) return;
    setEliminando(true);
    try {
      for (const t of borrables) {
        await transaccionesApi.remove(t.id).catch(() => {});
      }
      setSeleccionadas(new Set());
    } finally {
      setEliminando(false);
    }
  };

  const renderRowTr = (t, depth = 0, n) => (
    <tr key={t.id}>
      <td style={{ ...tdStyle, width: 36, textAlign: "right", color: T.textFaint, fontFamily: T.fontMono, fontSize: 11 }}>{n}</td>
      <td style={{ ...tdStyle, textAlign: "center" }}>
        <input type="checkbox" checked={seleccionadas.has(t.id)} onChange={() => toggleSeleccion(t.id)} />
      </td>
      {columnasVisibles.map((c, i) => (
        <td key={c.key} style={i === 0 && depth ? { ...tdStyle, paddingLeft: 14 + depth * 26 } : tdStyle}>{c.render(t)}</td>
      ))}
      <td style={tdStyle}>
        <div style={{ display: "flex", gap: 4 }}>
          {/* La SPP solo aplica a anticipos, y el marcador vive en Folio SAE.
              Mostrar el botón en todas invitaría a generar solicitudes de
              pagos que no las necesitan. */}
          {esAnticipo(t) && (
            <IconButton icon="§" label="Generar Solicitud de Pago" tone={T.teal} onClick={() => setSppDe(t)} />
          )}
          <IconButton icon="✎" label="Editar" tone={T.accent} onClick={() => startEdit(t)} />
          <IconButton icon="⧉" label="Duplicar" tone={T.textDim} onClick={() => duplicar(t)} />
          <IconButton icon="✕" label="Eliminar" tone={T.red} onClick={() => remove(t.id)} />
        </div>
      </td>
    </tr>
  );

  const [notaPrivada, setNotaPrivada] = useState("");
  const notaActual = (transaccionId) => notasApi.rows.find((n) => n.transaccion_id === transaccionId && n.usuario_id === session?.user?.id);

  const guardarNotaPrivada = async (transaccionId) => {
    if (!session?.user?.id) return;
    const existente = notaActual(transaccionId);
    try {
      if (!notaPrivada.trim()) {
        if (existente) await notasApi.remove(existente.id);
        return;
      }
      if (existente) {
        await notasApi.update(existente.id, { nota: notaPrivada, updated_at: new Date().toISOString() });
      } else {
        await notasApi.insert({ id: uid(), transaccion_id: transaccionId, usuario_id: session.user.id, nota: notaPrivada });
      }
    } catch (err) {
      alert("No se pudo guardar tu nota privada: " + (err.message || err));
    }
  };

  const submit = async (e) => {
    e.preventDefault();
    if (!form.partida_id || !form.importe) return;
    if (form.status === "Pagado" && !form.fecha_pago) {
      alert("Para marcar esta transacción como Pagada, primero indica la Fecha de Pago.");
      return;
    }
    if (!validarMonedaContraPartida(form, partidasUnidad)) return;
    const { id, ...restRaw } = form;
    const rest = Object.fromEntries(Object.entries(restRaw).filter(([k]) => !k.startsWith("_")));
    rest.proveedor_id = rest.proveedor_id || null;
    rest.cuenta_id = rest.cuenta_id || null;
    rest.fecha_pago = rest.fecha_pago || null;
    setSaving(true);
    try {
      let transaccionId = editId;
      if (editId) {
        await transaccionesApi.update(editId, rest);
        setEditId(null);
      } else {
        const fechaForm = form.dia ? new Date(`${form.dia}T00:00:00`) : new Date();
        const mesForm = MESES[fechaForm.getMonth()];
        const anioForm = fechaForm.getFullYear();
        const creada = await insertTransaccionConReintento(transaccionesApi, rest, unidad, mesForm, anioForm, transUnidad);
        transaccionId = creada.id;
      }
      await guardarNotaPrivada(transaccionId);
      setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" });
      setNotaPrivada("");
      setModalOpen(false);
    } catch (err) {
      alert("No se pudo guardar la transacción: " + (err.message || err));
    } finally {
      setSaving(false);
    }
  };
  const openNew = () => { setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" }); setEditId(null); setNotaPrivada(""); setModalOpen(true); };
  const startEdit = (t) => {
    // Quita campos internos (_proyecto, _rubro, _mes, _vinculo) que se agregan
    // solo para el agrupamiento — no existen como columnas reales en Supabase.
    const limpio = Object.fromEntries(Object.entries(t).filter(([k]) => !k.startsWith("_")));
    setForm(limpio);
    setEditId(t.id);
    setNotaPrivada(notaActual(t.id)?.nota || "");
    setModalOpen(true);
  };
  const closeModal = () => { setModalOpen(false); setEditId(null); setForm({ ...blank, partida_id: partidasUnidad[0]?.id || "" }); setNotaPrivada(""); };
  const duplicar = (t) => {
    // Copia los datos de la transacción, pero como registro NUEVO: sin folio propio,
    // sin fecha de pago/status heredado (nace "No Pagado"), y sin folios de compra/
    // factura (son específicos de cada pago, no algo que tenga sentido clonar).
    const { id, folio_transaccion, created_by, updated_by, updated_at, created_at, reportado_at, ...resto } = t;
    const limpio = Object.fromEntries(Object.entries(resto).filter(([k]) => !k.startsWith("_")));
    setForm({ ...limpio, status: "No Pagado", fecha_pago: "", folio_compra_sae: "", folio_factura: "", referencia_pago: "" });
    setEditId(null);
    setNotaPrivada("");
    setModalOpen(true);
  };
  const remove = (id) => {
    const t = transUnidad.find((x) => x.id === id) || sinVincular.find((x) => x.id === id);
    if (!puedeBorrarTransaccion(t)) return;
    if (!confirm(`¿Eliminar la transacción "${t?.concepto_detallado || id}" (${money(t?.importe, t?.moneda)})? Esto no se puede deshacer.`)) return;
    transaccionesApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  const toggleSeleccion = (id) => setSeleccionadas((prev) => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });
  const toggleSeleccionTodas = () => {
    const idsVisibles = transOrdenadas.map((t) => t.id);
    const todasSeleccionadas = idsVisibles.length > 0 && idsVisibles.every((id) => seleccionadas.has(id));
    setSeleccionadas(todasSeleccionadas ? new Set() : new Set(idsVisibles));
  };
  const marcarReportadas = async (reportar) => {
    setMarcandoReportado(true);
    try {
      for (const id of seleccionadas) {
        await transaccionesApi.update(id, { reportado_at: reportar ? new Date().toISOString() : null });
      }
      setSeleccionadas(new Set());
    } catch (err) {
      alert("No se pudo actualizar: " + (err.message || err));
    } finally {
      setMarcandoReportado(false);
    }
  };
  const marcarEnviadasPagos = async (enviar) => {
    setMarcandoEnviado(true);
    try {
      for (const id of seleccionadas) {
        await transaccionesApi.update(id, { enviado_pagos_at: enviar ? new Date().toISOString() : null });
      }
      setSeleccionadas(new Set());
    } catch (err) {
      alert("No se pudo actualizar: " + (err.message || err));
    } finally {
      setMarcandoEnviado(false);
    }
  };

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
          <Field label="Desde">
            <TextInput type="date" value={filtros.fechaDesde} onChange={(e) => setFiltros({ ...filtros, fechaDesde: e.target.value })} />
          </Field>
          <Field label="Hasta">
            <TextInput type="date" value={filtros.fechaHasta} onChange={(e) => setFiltros({ ...filtros, fechaHasta: e.target.value })} />
          </Field>
          <Field label="Reportado a Dirección">
            <Select value={filtros.reportado} onChange={(e) => setFiltros({ ...filtros, reportado: e.target.value })} style={{ width: 170 }}>
              <option>Todos</option>
              <option>Reportado</option>
              <option>No reportado</option>
            </Select>
          </Field>
          <Field label="Enviado a Pagos">
            <Select value={filtros.enviadoPagos} onChange={(e) => setFiltros({ ...filtros, enviadoPagos: e.target.value })} style={{ width: 160 }}>
              <option>Todos</option>
              <option>Enviado</option>
              <option>No enviado</option>
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
          {groupKeys.length > 0 && (
            <Button variant="ghost" onClick={() => setCollapsedGroups(new Set(collectGroupPaths(grouped)))}>Contraer todo</Button>
          )}
          <ColumnVisibilityControl
            columns={COLUMNAS_TRANS}
            hidden={colVisibility.hidden}
            onToggle={colVisibility.toggle}
            onShowAll={colVisibility.showAll}
          />
        </div>

        {seleccionadas.size > 0 && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, padding: "10px 14px", background: T.accentBg, border: `1px solid ${T.accent}55`, borderRadius: 6, flexWrap: "wrap" }}>
            <span style={{ fontSize: 12.5, color: T.text }}>{seleccionadas.size} seleccionada(s)</span>
            <Button onClick={() => marcarReportadas(true)} disabled={marcandoReportado}>
              {marcandoReportado ? "Marcando…" : "Marcar como reportadas a Dirección"}
            </Button>
            <Button variant="ghost" onClick={() => marcarReportadas(false)} disabled={marcandoReportado}>Quitar marca</Button>
            <div style={{ width: 1, alignSelf: "stretch", background: T.accent, opacity: 0.3, margin: "0 2px" }} />
            <Button onClick={() => marcarEnviadasPagos(true)} disabled={marcandoEnviado}>
              {marcandoEnviado ? "Marcando…" : "Marcar como enviadas a Pagos"}
            </Button>
            <Button variant="ghost" onClick={() => marcarEnviadasPagos(false)} disabled={marcandoEnviado}>Quitar marca</Button>
            <Button variant="ghost" onClick={() => setSeleccionadas(new Set())}>Cancelar selección</Button>
            <div style={{ width: 1, alignSelf: "stretch", background: T.accent, opacity: 0.3, margin: "0 2px" }} />
            <Button variant="danger" onClick={eliminarSeleccionadas} disabled={eliminando}>
              {eliminando ? "Eliminando…" : "Eliminar seleccionadas"}
            </Button>
          </div>
        )}

        <div style={{ overflowX: "auto" }}>
          <table style={{ ...tableStyle, tableLayout: "fixed" }}>
            <colgroup>
              <col style={{ width: 36 }} />
              <col style={{ width: 28 }} />
              {columnasVisibles.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
              <col style={{ width: 140 }} />
            </colgroup>
            <thead>
              <tr>
                <th style={thStyle}>#</th>
                <th style={{ ...thStyle, textAlign: "center" }}>
                  <input type="checkbox" checked={transOrdenadas.length > 0 && transOrdenadas.every((t) => seleccionadas.has(t.id))} onChange={toggleSeleccionTodas} />
                </th>
                {columnasVisibles.map((c) => (
                  <SortableTh
                    key={c.key} label={c.label} sortKey={c.key} sort={sort} setSort={setSort}
                    width={colWidths.getWidth(c.key)} onResizeStart={(e) => colWidths.startResize(c.key, e)}
                    onDragStart={(e) => onColDragStart(e, c.key)} onDragOver={onColDragOver} onDrop={(e) => onColDrop(e, c.key)}
                    sumLabel={c.key === "importe" ? sumaPorMoneda(transOrdenadas, "importe") : undefined}
                  />
                ))}
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {groupKeys.length
                ? buildGroupedTrs(grouped, "", collapsedGroups, toggleGroup, columnasVisibles.length + 3, 0, renderRowTr, Object.fromEntries(GROUP_OPCIONES_TRANS.map((o) => [o.value, o.label])))
                : transEnriquecidas.map((t, i) => renderRowTr(t, 0, i + 1))}
              {!transUnidad.length && (
                <tr><td colSpan={columnasVisibles.length + 3} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin transacciones aún</td></tr>
              )}
              {transUnidad.length > 0 && !transFiltradas.length && (
                <tr><td colSpan={columnasVisibles.length + 3} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna transacción coincide con estos filtros</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Panel>

      {sinVincular.length > 0 && (
        <Panel
          title="Transacciones importadas sin partida vinculada"
          subtitle={`${sinVincularFiltrado.length} de ${sinVincular.length} en ${unidad} — su folio no coincidió con ninguna partida`}
          right={
            <Button
              variant="danger"
              onClick={async () => {
                const pagadas = sinVincular.filter((t) => t.status === "Pagado");
                const borrables = sinVincular.filter((t) => t.status !== "Pagado");
                if (!borrables.length) {
                  alert(`Las ${pagadas.length} transacciones sin vincular están marcadas como Pagadas y no se pueden borrar.`);
                  return;
                }
                const aviso = pagadas.length ? `\n\nSe van a CONSERVAR ${pagadas.length} marcadas como Pagadas.` : "";
                if (!confirm(`¿Eliminar ${borrables.length} transacciones sin vincular de ${unidad}? Esto no se puede deshacer.${aviso}`)) return;
                for (const t of borrables) {
                  await transaccionesApi.remove(t.id).catch(() => {});
                }
              }}
            >
              Eliminar las {sinVincular.length} sin vincular
            </Button>
          }
        >
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${T.borderSoft}` }}>
            <Field label="Desde">
              <TextInput type="date" value={filtrosSV.fechaDesde} onChange={(e) => setFiltrosSV({ ...filtrosSV, fechaDesde: e.target.value })} />
            </Field>
            <Field label="Hasta">
              <TextInput type="date" value={filtrosSV.fechaHasta} onChange={(e) => setFiltrosSV({ ...filtrosSV, fechaHasta: e.target.value })} />
            </Field>
            {filtrosSVActivos && <Button variant="ghost" onClick={limpiarFiltrosSV}>Limpiar filtros</Button>}
            <div style={{ width: 1, alignSelf: "stretch", background: T.borderSoft, margin: "0 4px" }} />
            <GroupByControl
              options={GROUP_OPCIONES_SINVINC}
              value={groupBysSV}
              onChange={(v) => { setGroupBysSV(v); setCollapsedGroupsSV(new Set()); }}
              maxLevels={3}
              groupedTree={groupedSV}
              collapsed={collapsedGroupsSV}
              setCollapsed={setCollapsedGroupsSV}
            />
            {groupKeysSV.length > 0 && (
              <Button variant="ghost" onClick={() => setCollapsedGroupsSV(new Set(collectGroupPaths(groupedSV)))}>Contraer todo</Button>
            )}
            <ColumnVisibilityControl
              columns={COLUMNAS_SINVINC}
              hidden={colVisibilitySV.hidden}
              onToggle={colVisibilitySV.toggle}
              onShowAll={colVisibilitySV.showAll}
            />
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ ...tableStyle, tableLayout: "fixed" }}>
              <colgroup>
                {columnasSV.map((c) => <col key={c.key} style={{ width: colWidthsSV.getWidth(c.key) }} />)}
                <col style={{ width: 230 }} />
              </colgroup>
              <thead>
                <tr>
                  {columnasSV.map((c) => (
                    <SortableTh
                      key={c.key} label={c.label} sortKey={c.key} sort={{ key: "", dir: "asc" }} setSort={() => {}}
                      width={colWidthsSV.getWidth(c.key)} onResizeStart={(e) => colWidthsSV.startResize(c.key, e)}
                      onDragStart={(e) => onColDragStartSV(e, c.key)} onDragOver={onColDragOverSV} onDrop={(e) => onColDropSV(e, c.key)}
                      sumLabel={c.key === "importe" ? sumaPorMoneda(sinVincularFiltrado, "importe") : undefined}
                    />
                  ))}
                  <th style={thStyle}>Vincular a partida</th>
                </tr>
              </thead>
              <tbody>
                {groupKeysSV.length
                  ? buildGroupedTrs(groupedSV, "", collapsedGroupsSV, toggleGroupSV, columnasSV.length + 1, 0, renderRowSinVinc, Object.fromEntries(GROUP_OPCIONES_SINVINC.map((o) => [o.value, o.label])))
                  : sinVincularFiltrado.map((t) => renderRowSinVinc(t))}
                {!sinVincularFiltrado.length && (
                  <tr><td colSpan={columnasSV.length + 1} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna coincide con estos filtros</td></tr>
                )}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11, color: T.textFaint, marginTop: 10 }}>
            Elige una partida en el selector de cada fila para vincularla — se guarda al instante.
          </div>
        </Panel>
      )}

      <ImportarTransaccionesPanel partidas={partidas} proveedores={proveedoresApi.rows} cuentas={cuentasApi.rows} transaccionesApi={transaccionesApi} />

      {sppDe && (
        <SolicitudPagoModal
          transaccion={sppDe}
          onClose={() => setSppDe(null)}
          unidad={unidad}
          partidas={partidasUnidad}
          proyectosUnidad={proyectosUnidad}
          proveedoresApi={proveedoresApi}
          cuentasApi={cuentasApi}
          transaccionesApi={transaccionesApi}
          session={session}
        />
      )}
      {modalOpen && (
        <Modal
          title={editId ? "Editar transacción" : "Nueva transacción real"}
          subtitle={form.folio_transaccion ? `ID: ${form.folio_transaccion} — se vincula a una partida` : "Se vincula a una partida — una partida puede tener varias"}
          onClose={closeModal}
        >
          <AutoriaCaption record={form} perfilesApi={perfilesApi} />
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
            <Field label="Partida" style={{ gridColumn: "span 4" }}>
              {/* Solo las partidas de la MISMA moneda. Prevenir es mejor que
                  rechazar al guardar: la opción equivocada ni se ofrece. */}
              <PartidaPickerButton
                partidas={partidasUnidad.filter((p) => mismaMoneda(p.moneda, form.moneda))}
                ocultasPorMoneda={partidasUnidad.filter((p) => !mismaMoneda(p.moneda, form.moneda)).length}
                moneda={form.moneda}
                transacciones={transUnidad}
                partidasApi={partidasApi}
                unidad={unidad}
                proyectosOpciones={marcadoresProyecto}
                value={form.partida_id}
                onChange={(id) => setForm({ ...form, partida_id: id })}
              />
            </Field>
            <Field label="Día de Pago Programado">
              <TextInput type="date" value={form.dia} onChange={(e) => setForm({ ...form, dia: e.target.value })} />
            </Field>
            <Field label="Solicitante">
              <TextInput value={form.solicitante} onChange={(e) => setForm({ ...form, solicitante: e.target.value })} />
            </Field>
            <Field label="SMI (No. de Solicitud)">
              <TextInput value={form.smi} onChange={(e) => setForm({ ...form, smi: e.target.value })} />
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
                {zonas.map((z) => <option key={z}>{z}</option>)}
                {/* Si la transacción trae una zona que ya no está en el
                    catálogo, se agrega como opción para no perderla al editar. */}
                {form.zona && !zonas.includes(form.zona) && <option key={form.zona}>{form.zona}</option>}
              </Select>
            </Field>
            <Field label="Proveedor (catálogo)">
              <ProveedorPickerButton
                proveedores={proveedoresUnidad}
                proveedoresApi={proveedoresApi}
                cuentasApi={cuentasApi}
                unidad={unidad}
                value={form.proveedor_id}
                onChange={(id, p) => {
                  const proveedor = p !== undefined ? p : (proveedoresUnidad.find((pr) => pr.id === id) || null);
                  setForm({ ...form, proveedor_id: id, proveedor: proveedor ? proveedor.nombre : form.proveedor, cuenta_id: "" });
                }}
              />
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
            <Field label="Referencia de Pago">
              <TextInput value={form.referencia_pago} onChange={(e) => setForm({ ...form, referencia_pago: e.target.value })} placeholder="Folio SPEI, cheque, etc." />
            </Field>
            <Field label="Importe">
              <TextInput type="number" step="0.01" value={form.importe} onChange={(e) => setForm({ ...form, importe: e.target.value })} placeholder="0.00" />
            </Field>
            <Field label="Categoría del gasto">
              {/* Las opciones salen del rubro de la partida vinculada, pero NO
                  se restringe a él: con datos reales, 39 de 228 transacciones
                  tenían la categoría correcta para el gasto y una partida de
                  otro rubro. La discrepancia es información, no error. */}
              <Select value={form.categoria || ""} onChange={(e) => setForm({ ...form, categoria: e.target.value })}>
                <option value="">— Sin categoría —</option>
                {(() => {
                  const p = partidasUnidad.find((x) => x.id === form.partida_id);
                  const delRubro = p ? (RUBROS.find((r) => r.rubro === p.rubro)?.categorias || []) : [];
                  const resto = CATEGORIAS_TODAS.filter((c) => !delRubro.includes(c));
                  return (
                    <>
                      {delRubro.length > 0 && (
                        <optgroup label={`Del rubro ${p.rubro}`}>
                          {delRubro.map((c) => <option key={c}>{c}</option>)}
                        </optgroup>
                      )}
                      <optgroup label="Otras categorías">
                        {resto.map((c) => <option key={c}>{c}</option>)}
                      </optgroup>
                    </>
                  );
                })()}
              </Select>
              {(() => {
                const p = partidasUnidad.find((x) => x.id === form.partida_id);
                const cats = p ? (RUBROS.find((r) => r.rubro === p.rubro)?.categorias || []) : CATEGORIAS_TODAS;
                const sug = sugerirCategoria(form.concepto_detallado, cats);
                if (!sug || sug === form.categoria) return null;
                return (
                  <button type="button" onClick={() => setForm({ ...form, categoria: sug })}
                    style={{ marginTop: 5, background: "transparent", border: "none", padding: 0,
                             color: T.accent, fontSize: 11.5, cursor: "pointer", textAlign: "left", fontFamily: T.fontUI }}>
                    Por el concepto, quizá sea <b>{sug}</b> — usar
                  </button>
                );
              })()}
            </Field>
            <Field label="Moneda">
              <Select
                value={form.moneda}
                onChange={(e) => {
                  // Al cambiar de moneda, la partida elegida puede dejar de ser
                  // válida. Se suelta en lugar de quedarse como una selección
                  // que el guardado va a rechazar después.
                  const nueva = e.target.value;
                  const actual = partidasUnidad.find((p) => p.id === form.partida_id);
                  const sigueValida = !actual || mismaMoneda(actual.moneda, nueva);
                  setForm({ ...form, moneda: nueva, partida_id: sigueValida ? form.partida_id : "" });
                }}
              >
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
            {form.status === "Pagado" && (
              <Field label="Fecha de Pago">
                <TextInput type="date" value={form.fecha_pago} onChange={(e) => setForm({ ...form, fecha_pago: e.target.value })} required />
              </Field>
            )}
            <Field label="🔒 Tu nota privada (solo tú la ves)" style={{ gridColumn: "span 4" }}>
              <TextInput value={notaPrivada} onChange={(e) => setNotaPrivada(e.target.value)} placeholder="Recordatorios, pendientes, contexto — nadie más puede ver esto" />
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
/**
 * Recuerda qué columnas están ocultas para un destino concreto (pantalla,
 * PDF, Excel). Cada destino guarda su propia preferencia bajo su llave, de
 * modo que ocultar algo en el PDF no afecta al Excel.
 *
 * Las columnas marcadas `fija` no se pueden ocultar: son las que el
 * documento necesita para servir de algo.
 */
function useVisibilidadColumnas(storageKey, columnas, ocultasIniciales = []) {
  const [ocultas, setOcultas] = useState(() => {
    try {
      const raw = localStorage.getItem(storageKey);
      return new Set(raw ? JSON.parse(raw) : ocultasIniciales);
    } catch { return new Set(ocultasIniciales); }
  });
  const guardar = (next) => {
    setOcultas(next);
    try { localStorage.setItem(storageKey, JSON.stringify([...next])); } catch {}
  };
  return {
    hidden: ocultas,
    toggle: (k) => {
      const col = columnas.find((c) => c.key === k);
      if (col && col.fija) return;
      const next = new Set(ocultas);
      next.has(k) ? next.delete(k) : next.add(k);
      guardar(next);
    },
    showAll: () => guardar(new Set()),
    visibles: columnas.filter((c) => !ocultas.has(c.key)),
  };
}

// Columnas que puede llevar el PDF de solicitud de pagos. `fija: true` son
// las que el área de Pagos necesita siempre y no se pueden quitar; el resto
// se elige desde el botón "Columnas del PDF".
const COLUMNAS_PDF = [
  { key: "dia",            label: "Día",              fija: true,  get: (f) => f.dia || "" },
  { key: "solicitante",    label: "Solicitante",      fija: true,  get: (f) => f.solicitante || "" },
  { key: "area",           label: "Área",             get: (f) => f.area || "" },
  { key: "numero_solicitud", label: "No. Solicitud",  get: (f) => f.numero_solicitud || "" },
  { key: "folio_factura",  label: "Folio Factura",    get: (f) => f.folio_factura || "" },
  { key: "forma_pago",     label: "Forma de Pago",    get: (f) => f.forma_pago || "" },
  { key: "metodo_pago",    label: "Método de Pago",   get: (f) => f.metodo_pago || "" },
  { key: "proveedor",      label: "Proveedor",        fija: true,  ancho: "left", get: (f) => f.proveedor || "" },
  { key: "concepto",       label: "Concepto de pago", fija: true,  ancho: "left", get: (f) => f.concepto || "" },
  { key: "banco",          label: "Banco",            get: (f) => f.banco || "" },
  { key: "clabe",          label: "Cuenta CLABE",     get: (f) => f.clabe || "" },
  { key: "numero_cuenta",  label: "No. Cuenta",       get: (f) => f.numero_cuenta || "" },
  { key: "swift",          label: "SWIFT",            get: (f) => f.swift || "" },
  { key: "referencia_pago", label: "Referencia de Pago", get: (f) => f.referencia_pago || "" },
  { key: "importe",        label: "Importe",          fija: true,  get: (f) => money(f.importe, f.moneda) },
  { key: "notas",          label: "Notas",            ancho: "left", get: (f) => f.notas || "" },
];
// Las que trae marcadas la primera vez: el formato que se venía usando.
const PDF_OCULTAS_INICIAL = ["area", "numero_solicitud", "folio_factura", "forma_pago",
  "metodo_pago", "numero_cuenta", "swift", "referencia_pago", "notas"];

// Columnas del Excel del Reporte de Pagos. Encabezado, ancho y valor viven
// juntos: cuando eran arreglos paralelos, agregar una columna a uno y
// olvidar los otros corría todo el archivo en silencio.
const COLUMNAS_EXCEL = [
  { key: "dia",              header: "Día",                 width: 8.875, fija: true, get: (f) => f.dia },
  { key: "solicitante",      header: "Solicitante",         width: 10,    fija: true, get: (f) => f.solicitante },
  { key: "area",             header: "Área",                width: 8.125, get: (f) => f.area },
  { key: "numero_solicitud", header: "No. Solicitud (SMI)", width: 10.125, get: (f) => f.numero_solicitud },
  { key: "no_sae",           header: "No. SAE",             width: 10.125, get: (f) => f.no_sae },
  { key: "folio_compra_sae", header: "Folio Compra SAE",    width: 8.625, get: (f) => f.folio_compra_sae },
  { key: "folio_factura",    header: "Folio Factura",       width: 8.25,  get: (f) => f.folio_factura },
  { key: "forma_pago",       header: "Forma de Pago",       width: 15,    get: (f) => f.forma_pago },
  { key: "metodo_pago",      header: "Método de Pago",      width: 12,    get: (f) => f.metodo_pago },
  { key: "proveedor",        header: "Proveedor",           width: 17.625, fija: true, get: (f) => f.proveedor },
  { key: "referencia_pago",  header: "Referencia de Pago",  width: 14,    get: (f) => f.referencia_pago },
  { key: "concepto",         header: "Concepto de pago",    width: 13.625, fija: true, get: (f) => f.concepto },
  { key: "banco",            header: "Banco",               width: 15.5,  get: (f) => f.banco },
  { key: "clabe",            header: "Cuenta CLABE",        width: 18,    get: (f) => f.clabe },
  { key: "numero_cuenta",    header: "No. Cuenta",          width: 12.5,  get: (f) => f.numero_cuenta },
  { key: "swift",            header: "SWIFT",               width: 9,     get: (f) => f.swift },
  { key: "importe",          header: "Importe",             width: 10.5,  fija: true, money: true, get: (f) => f.importe },
  { key: "moneda",           header: "Moneda",              width: 10.5,  get: (f) => f.moneda },
  { key: "notas",            header: "Notas",               width: 8.875, get: (f) => f.notas },
];

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
  { key: "referencia_pago", label: "Referencia de Pago" },
  { key: "concepto", label: "Concepto de pago" },
  { key: "banco", label: "Banco" },
  { key: "clabe", label: "Cuenta CLABE" },
  { key: "numero_cuenta", label: "No. Cuenta" },
  { key: "swift", label: "SWIFT" },
  { key: "importe", label: "Importe" },
  { key: "moneda", label: "Moneda" },
  { key: "notas", label: "Notas" },
];

function ReportePagosTab({ unidad, partidas, transacciones, transaccionesApi, proveedoresApi, cuentasApi }) {
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
      zona: t.zona || "",
      numero_solicitud: t.smi || partida?.smi || "",
      no_sae: proveedor?.id_sae || "",
      folio_compra_sae: t.folio_compra_sae || "",
      folio_factura: t.folio_factura || "",
      forma_pago: t.forma_pago || "",
      forma_pago_label: FORMAS_PAGO.find((f) => f.value === t.forma_pago)?.label || t.forma_pago || "",
      metodo_pago: t.metodo_pago || "",
      metodo_pago_label: METODOS_PAGO.find((m) => m.value === t.metodo_pago)?.label || t.metodo_pago || "",
      proveedor: proveedor?.nombre || t.proveedor || "",
      referencia_pago: t.referencia_pago || "",
      notas: proveedor?.notas || "",
      concepto: t.concepto_detallado || "",
      banco: cuenta?.banco || "",
      clabe: cuenta?.clabe || "",
      numero_cuenta: cuenta?.numero_cuenta || "",
      swift: cuenta?.swift || "",
      importe: Number(t.importe) || 0,
      moneda: t.moneda || "MXP",
      _vinculadoProveedor: !!proveedor,
      _vinculadoCuenta: !!cuenta,
    };
  });

  const [buscar, setBuscar] = useSessionState("ss-reporte-buscar", "");
  const [fechaDesde, setFechaDesde] = useSessionState("ss-reporte-desde", "");
  const [fechaHasta, setFechaHasta] = useSessionState("ss-reporte-hasta", "");
  const [sort, setSort] = usePrefState("pref-reporte-sort", { key: "dia", dir: "desc" }, sanearSort(["dia","solicitante","area","numero_solicitud","no_sae","folio_compra_sae","folio_factura","forma_pago","metodo_pago","proveedor","referencia_pago","concepto","banco","clabe","numero_cuenta","swift","importe","moneda","notas"]));
  const filasFiltradas = filas.filter((f) => {
    if (fechaDesde && (!f.dia || f.dia < fechaDesde)) return false;
    if (fechaHasta && (!f.dia || f.dia > fechaHasta)) return false;
    if (!buscar.trim()) return true;
    const q = buscar.trim().toLowerCase();
    return [f.solicitante, f.proveedor, f.concepto, f.folio_factura, f.folio_compra_sae, f.no_sae, f.referencia_pago]
      .some((v) => (v || "").toString().toLowerCase().includes(q));
  });
  const filtrosFechaActivos = fechaDesde || fechaHasta;
  const limpiarFechas = () => { setFechaDesde(""); setFechaHasta(""); };
  const filasOrdenadas = sortRows(filasFiltradas, sort, { importe: (r) => r.importe });

  const colVisibility = useColumnVisibility("colv-reporte", COLUMNAS_REPORTE);
  const pdfVis = useVisibilidadColumnas("colv-reporte-pdf", COLUMNAS_PDF, PDF_OCULTAS_INICIAL);
  const columnasPDF = pdfVis.visibles;
  const excelVis = useVisibilidadColumnas("colv-reporte-excel", COLUMNAS_EXCEL);
  const columnasExcel = excelVis.visibles;
  const [previewPDF, setPreviewPDF] = useState(null);
  const colWidths = useColumnWidths("colw-reporte");
  const { ordered: columnas, moveColumn } = useColumnOrder("colo-reporte", colVisibility.visible);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };

  const exportarExcel = async () => {
    // Resumen en vez de vista previa: renderizar una tabla que IMITE la hoja
    // sería una segunda implementación del layout, y en cuanto se desincronice
    // el preview estaría mintiendo. Esto responde la misma pregunta —¿va lo
    // que necesito?— con datos del archivo real.
    const ocultas = COLUMNAS_EXCEL.filter((c) => excelVis.hidden.has(c.key));
    const muestra = filasOrdenadas.slice(0, 3)
      .map((f, i) => `  ${i + 1}. ` + columnasExcel.slice(0, 5).map((c) => String(c.get(f) ?? "—")).join(" · "))
      .join("\n");
    const okExcel = confirm(
      `Se va a descargar un Excel con:\n\n` +
      `  ${filasOrdenadas.length} transacción(es)\n` +
      `  ${columnasExcel.length} de ${COLUMNAS_EXCEL.length} columnas\n` +
      (ocultas.length ? `  Fuera: ${ocultas.map((c) => c.header).join(", ")}\n` : "") +
      `\nPrimeras filas (5 columnas):\n${muestra}\n\n¿Descargar?`
    );
    if (!okExcel) return;

    const wbx = new ExcelJS.Workbook();
    const ws = wbx.addWorksheet("Reporte de pagos");
    ws.columns = columnasExcel.map((c) => ({ width: c.width }));

    const diasOrdenados = filasOrdenadas.map((f) => f.dia).filter(Boolean).sort();
    const inicio = fechaDesde || diasOrdenados[0] || "";
    const fin = fechaHasta || diasOrdenados[diasOrdenados.length - 1] || "";

    const zonas = [...new Set(filasOrdenadas.map((f) => f.zona).filter(Boolean))].sort();
    const ordenMoneda = (m) => (m === "MXP" ? 0 : m === "USD" ? 1 : 2);
    let fila = 1;
    zonas.forEach((zona) => {
      const monedasEnZona = [...new Set(filasOrdenadas.filter((f) => f.zona === zona).map((f) => f.moneda))]
        .sort((a, b) => ordenMoneda(a) - ordenMoneda(b));

      monedasEnZona.forEach((moneda) => {
        const filasGrupo = filasOrdenadas.filter((f) => f.zona === zona && f.moneda === moneda);
        if (!filasGrupo.length) return;

        const tituloCell = ws.getCell(`B${fila}`);
        tituloCell.value = `Solicitud de Pagos del dia ${inicio} al dia ${fin} Compañía ${unidad} - ${moneda}`;
        tituloCell.font = { bold: true, size: 14, name: "Calibri" };
        fila += 1;

        const zonaCell = ws.getCell(`D${fila}`);
        zonaCell.value = `Zona: ${zona}`;
        zonaCell.font = { bold: true, size: 12, name: "Calibri" };
        fila += 2; // una fila en blanco, como en la plantilla

        const headerRow = ws.getRow(fila);
        columnasExcel.forEach((c, i) => {
          const cell = headerRow.getCell(i + 1);
          cell.value = c.header;
          cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
          cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
        });
        headerRow.height = 22;
        fila += 1;

        filasGrupo.forEach((f) => {
          const row = ws.getRow(fila);
          columnasExcel.forEach((c, ci) => {
            const cell = row.getCell(ci + 1);
            cell.value = c.get(f);
            // Las cifras a la derecha, los textos largos a la izquierda con
            // ajuste: centrarlo todo obliga a leer cada celda para comparar.
            cell.alignment = {
              horizontal: c.money ? "right" : (["proveedor", "concepto", "notas"].includes(c.key) ? "left" : "center"),
              vertical: "top",
              wrapText: ["proveedor", "concepto", "notas"].includes(c.key),
            };
            cell.font = { name: "Calibri", size: 11 };
            if (c.money) cell.numFmt = '"$"#,##0.00';
          });
          fila += 1;
        });

        // Subtotal del bloque: cada zona y moneda cierra con su suma, que es
        // lo que Pagos necesita para cuadrar contra el banco.
        const idxImporte = columnasExcel.findIndex((c) => c.money);
        if (idxImporte !== -1) {
          const totRow = ws.getRow(fila);
          const etq = totRow.getCell(Math.max(idxImporte, 1));
          etq.value = `Total ${zona} ${moneda}`;
          etq.font = { name: "Calibri", size: 11, bold: true };
          etq.alignment = { horizontal: "right" };
          const val = totRow.getCell(idxImporte + 1);
          val.value = filasGrupo.reduce((sum, f) => sum + (Number(f.importe) || 0), 0);
          val.numFmt = '"$"#,##0.00';
          val.font = { name: "Calibri", size: 11, bold: true };
          val.alignment = { horizontal: "right" };
          for (let i = 1; i <= columnasExcel.length; i++) {
            totRow.getCell(i).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFECEEF1" } };
          }
          fila += 1;
        }

        fila += 2; // espacio antes de la siguiente sección
      });
    });

    const buffer = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Reporte de Pagos - ${unidad} - ${(!inicio ? "Sin periodo" : (inicio === fin ? inicio : `${inicio} a ${fin}`))}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const [generandoReporte, setGenerandoReporte] = useState(false);
  /**
   * Estima si la tabla cabe a lo ancho de la hoja. autoTable no falla cuando
   * se pasa: reparte el sobrante partiendo palabras, y el resultado se vuelve
   * ilegible sin previo aviso. Vale más advertirlo antes de mandarlo a Pagos.
   */
  const anchoEstimadoPDF = () => {
    // Carta horizontal son 792 pt; quedan ~712 descontando márgenes.
    const disponible = 712;
    const usado = columnasPDF.reduce((suma, c) => {
      const largos = filasOrdenadas.map((f) => String(c.get(f) ?? "").length);
      const max = Math.max(c.label.length, ...(largos.length ? largos : [0]));
      // A 8 pt, cada carácter mide ~4.4 pt; más 8 pt de relleno por celda.
      return suma + Math.min(max * 4.4, 140) + 8;
    }, 0);
    return { usado: Math.round(usado), disponible, cabe: usado <= disponible };
  };

  /**
   * Construye el documento y lo devuelve SIN guardarlo. La vista previa y el
   * archivo final salen los dos de aquí: con dos generadores distintos, el
   * preview acabaría mostrando algo que no es lo que se descarga.
   */
  const construirPDF = () => {
      const diasOrdenados = filasOrdenadas.map((f) => f.dia).filter(Boolean).sort();
      const inicio = fechaDesde || diasOrdenados[0] || "";
      const fin = fechaHasta || diasOrdenados[diasOrdenados.length - 1] || "";
      const zonas = [...new Set(filasOrdenadas.map((f) => f.zona).filter(Boolean))].sort();
      const ordenMoneda = (m) => (m === "MXP" ? 0 : m === "USD" ? 1 : 2);

      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "letter" });
      const alturaPagina = doc.internal.pageSize.getHeight();
      const margenInferior = 40;
      let cursorY = 30; // dónde va el próximo título — solo salta de página si de verdad no cabe

      zonas.forEach((zona) => {
        const monedasEnZona = [...new Set(filasOrdenadas.filter((f) => f.zona === zona).map((f) => f.moneda))]
          .sort((a, b) => ordenMoneda(a) - ordenMoneda(b));

        monedasEnZona.forEach((moneda) => {
          const filasGrupo = filasOrdenadas.filter((f) => f.zona === zona && f.moneda === moneda);
          if (!filasGrupo.length) return;

          // Espacio mínimo para el título + al menos una fila de la tabla —
          // si no cabe en lo que resta de la página, ahí sí se salta.
          const alturaMinima = 58 + 24 + 20;
          if (cursorY > 30 && cursorY + alturaMinima > alturaPagina - margenInferior) {
            doc.addPage();
            cursorY = 30;
          } else if (cursorY > 30) {
            cursorY += 26; // espacio entre un bloque y el siguiente en la misma página
          }

          doc.setFontSize(13);
          doc.setTextColor(0);
          doc.text(`Solicitud de Pagos del dia ${inicio} al dia ${fin} Compañía ${unidad} - ${moneda}`, 30, cursorY);
          doc.setFontSize(10);
          doc.setTextColor(120);
          doc.text(`Zona: ${zona}`, 30, cursorY + 16);

          autoTable(doc, {
            startY: cursorY + 28,
            head: [columnasPDF.map((c) => c.label)],
            body: filasGrupo.map((f) => columnasPDF.map((c) => c.get(f))),
            styles: { fontSize: 8, cellPadding: 4 },
            headStyles: { fillColor: [62, 92, 118], textColor: 255, halign: "center" },
            bodyStyles: { halign: "center" },
            // Los textos largos se alinean a la izquierda; lo demás centrado.
            columnStyles: Object.fromEntries(
              columnasPDF.map((c, i) => [i, { halign: c.ancho === "left" ? "left" : "center" }])
            ),
            margin: { bottom: margenInferior },
          });

          cursorY = doc.lastAutoTable.finalY;
        });
      });

      return doc;
  };

  /**
   * Muestra el PDF real en un visor, sin guardarlo y —lo importante— sin
   * marcar nada como enviado a Pagos. Es para revisar antes de comprometerse.
   */
  const vistaPreviaPDF = () => {
    if (!filasOrdenadas.length) {
      alert("No hay transacciones en el filtro actual para previsualizar.");
      return;
    }
    try {
      const doc = construirPDF();
      const url = URL.createObjectURL(doc.output("blob"));
      setPreviewPDF({ url, paginas: doc.internal.getNumberOfPages(), ancho: anchoEstimadoPDF() });
    } catch (err) {
      alert("No se pudo generar la vista previa: " + (err.message || err));
    }
  };

  const cerrarPreviewPDF = () => {
    if (previewPDF) URL.revokeObjectURL(previewPDF.url);
    setPreviewPDF(null);
  };

  const generarReportePDF = async () => {
    if (!filasOrdenadas.length) {
      alert("No hay transacciones en el filtro actual para generar el reporte.");
      return;
    }
    const ancho = anchoEstimadoPDF();
    const aviso = ancho.cabe ? "" :
      `\n\nOJO: con ${columnasPDF.length} columnas la tabla se pasa del ancho de la hoja ` +
      `(~${ancho.usado} pt contra ${ancho.disponible} disponibles). Las columnas se van a apretar ` +
      `y el texto se va a partir. Considera quitar algunas en "Columnas del PDF".`;
    const confirmado = confirm(
      `Esto va a generar un PDF con las ${filasOrdenadas.length} transacción(es) que tienes filtradas ahora, y las va a marcar como "Enviadas a Pagos". ¿Continuar?${aviso}`
    );
    if (!confirmado) return;

    setGenerandoReporte(true);
    try {
      const doc = construirPDF();
      doc.save(`reporte-pagos-${unidad}-${new Date().toISOString().slice(0, 10)}.pdf`);
      cerrarPreviewPDF();

      for (const f of filasOrdenadas) {
        await transaccionesApi.update(f.id, { enviado_pagos_at: new Date().toISOString() });
      }
    } catch (err) {
      alert("No se pudo generar el reporte: " + (err.message || err));
    } finally {
      setGenerandoReporte(false);
    }
  };
  const totalMXN = filasOrdenadas.filter((f) => f.moneda === "MXP").reduce((s, f) => s + f.importe, 0);
  const totalUSD = filasOrdenadas.filter((f) => f.moneda === "USD").reduce((s, f) => s + f.importe, 0);
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

  const hayAvisos = sinProveedorVinculado > 0 || sinCuentaVinculada > 0;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20, paddingBottom: hayAvisos ? 70 : 0 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <KpiCard label="Transacciones" value={String(filasOrdenadas.length)} />
        <KpiCard label="Total MXP (filtrado)" value={money(totalMXN, "MXP")} />
        <KpiCard label="Total USD (filtrado)" value={money(totalUSD, "USD")} />
      </div>

      <AvisosFlotantes
        avisos={[
          {
            tone: "amber",
            texto: sinProveedorVinculado > 0
              ? `${sinProveedorVinculado} transacción(es) tienen proveedor en texto pero no están ligadas al Catálogo de proveedores — No. SAE, Banco y CLABE saldrán vacíos para esas filas hasta que se vinculen (edítalas en Transacciones y elige el proveedor del catálogo).`
              : "",
          },
          {
            tone: "amber",
            texto: sinCuentaVinculada > 0
              ? `${sinCuentaVinculada} transacción(es) tienen proveedor vinculado pero no una cuenta bancaria específica — Banco y CLABE saldrán vacíos hasta que elijas la cuenta (edítalas en Transacciones).`
              : "",
          },
        ]}
      />

      <Panel
        title={`Reporte de pagos — ${unidad}`}
        subtitle={`${filasOrdenadas.length} de ${filas.length} transacciones`}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <ColumnVisibilityControl
              columns={COLUMNAS_REPORTE}
              hidden={colVisibility.hidden}
              onToggle={colVisibility.toggle}
              onShowAll={colVisibility.showAll}
            />
            <ColumnVisibilityControl
              columns={COLUMNAS_PDF}
              hidden={pdfVis.hidden}
              onToggle={pdfVis.toggle}
              onShowAll={pdfVis.showAll}
              etiqueta="Columnas del PDF"
            />
            <ColumnVisibilityControl
              columns={COLUMNAS_EXCEL}
              hidden={excelVis.hidden}
              onToggle={excelVis.toggle}
              onShowAll={excelVis.showAll}
              etiqueta="Columnas del Excel"
            />
            <Button onClick={exportarExcel}>Exportar a Excel</Button>
            <Button variant="ghost" onClick={vistaPreviaPDF} disabled={generandoReporte}>
              Vista previa PDF
            </Button>
            <Button onClick={generarReportePDF} disabled={generandoReporte}>
              {generandoReporte ? "Generando…" : "Generar reporte (PDF)"}
            </Button>
          </div>
        }
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
                    sumLabel={c.key === "importe" ? sumaPorMoneda(filasOrdenadas, "importe") : undefined}
                  />
                ))}
              </tr>
            </thead>
            <tbody>
              {filasOrdenadas.map((f) => (
                <tr key={f.id}>
                  {columnas.map((c) => (
                    <td
                      key={c.key}
                      style={c.key === "importe" ? { ...tdStyle, fontFamily: T.fontMono } : tdStyle}
                      title={c.key === "forma_pago" ? f.forma_pago_label : c.key === "metodo_pago" ? f.metodo_pago_label : undefined}
                    >
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
      {previewPDF && (
        <Modal onClose={cerrarPreviewPDF} width={1100}>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, height: "80vh" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 14, fontWeight: 700 }}>Vista previa del reporte</div>
                <div style={{ fontSize: 11.5, color: T.textFaint }}>
                  {filasOrdenadas.length} transacción(es) · {columnasPDF.length} columnas · {previewPDF.paginas} página(s)
                  {" · "}Nada se ha marcado como enviado a Pagos todavía
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <Button variant="ghost" onClick={cerrarPreviewPDF}>Cerrar</Button>
                <Button onClick={() => { cerrarPreviewPDF(); generarReportePDF(); }}>
                  Generar y marcar como enviado
                </Button>
              </div>
            </div>
            {!previewPDF.ancho.cabe && (
              <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "8px 12px", borderRadius: "0 6px 6px 0", fontSize: 12 }}>
                <b>La tabla se pasa del ancho de la hoja</b> (~{previewPDF.ancho.usado} pt contra {previewPDF.ancho.disponible} disponibles).
                Las columnas se aprietan y el texto se parte. Quita algunas en "Columnas del PDF".
              </div>
            )}
            <iframe
              title="Vista previa del reporte"
              src={previewPDF.url}
              style={{ flex: 1, width: "100%", border: `1px solid ${T.border}`, borderRadius: 6 }}
            />
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------
   TAB: REPORTE DE PAGOS DIRECCIÓN
---------------------------------------------------------------------- */
// Cada columna lleva su forma de imprimirse en el PDF (`pdf`) junto a la
// etiqueta, para que el documento salga de la MISMA definición que la tabla.
// Antes el PDF tenía sus diez columnas fijas en el código y no había manera de
// que respetara lo elegido en "Columnas".
// Cada columna lleva junto todo lo que necesita: cómo se ve en pantalla
// (`label`), cómo se imprime en el PDF (`pdf`) y cómo se exporta a Excel
// (`xls` + `width`). Antes eran listas paralelas y el formato de moneda se
// aplicaba por índice fijo, así que al mover una columna el signo de pesos
// se quedaba en la anterior.
const COLUMNAS_REPORTE_DIRECCION = [
  { key: "dia", label: "Día", pdf: (f) => f.dia || "", xls: (f) => f.dia, width: 12 },
  { key: "solicitante", label: "Solicitante", pdf: (f) => f.solicitante || "", xls: (f) => f.solicitante, width: 18 },
  { key: "proyecto", label: "Proyecto", pdf: (f) => f.proyecto || "", xls: (f) => f.proyecto, width: 16 },
  { key: "zona", label: "Zona", pdf: (f) => f.zona || "", xls: (f) => f.zona, width: 14 },
  { key: "proveedor", label: "Proveedor", pdf: (f) => f.proveedor || "", izq: true, xls: (f) => f.proveedor, width: 30 },
  { key: "concepto", label: "Concepto", pdf: (f) => f.concepto || "", izq: true, xls: (f) => f.concepto, width: 38 },
  { key: "importe", label: "Importe", pdf: (f) => money(f.importe, f.moneda), xls: (f) => f.importe, width: 14, money: true },
  { key: "moneda", label: "Moneda", pdf: (f) => f.moneda || "", xls: (f) => f.moneda, width: 9 },
  { key: "a_partida", label: "A Partida", pdf: (f) => f.a_partida || "", xls: (f) => f.a_partida, width: 16 },
  { key: "status", label: "Status", pdf: (f) => f.status || "", xls: (f) => f.status, width: 12 },
];

function ReportePagosDireccionTab({ unidad, partidas, transacciones, transaccionesApi, proveedoresApi }) {
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const transUnidad = transacciones.filter(
    (t) => partidasUnidad.some((p) => p.id === t.partida_id) || t.unidad_detectada === unidad
  );

  const filas = transUnidad.map((t) => {
    const partida = partidasUnidad.find((p) => p.id === t.partida_id);
    const proveedor = t.proveedor_id ? proveedoresUnidad.find((p) => p.id === t.proveedor_id) : null;
    return {
      id: t.id,
      dia: t.dia || "",
      solicitante: t.solicitante || "",
      proyecto: t.proyecto || "",
      zona: t.zona || "",
      proveedor: proveedor?.nombre || t.proveedor || "",
      concepto: t.concepto_detallado || "",
      importe: Number(t.importe) || 0,
      moneda: t.moneda || "MXP",
      a_partida: partida?.folio || "",
      status: t.status || "",
    };
  });

  const [buscar, setBuscar] = useSessionState("ss-reporte-direccion-buscar", "");
  const [fechaDesde, setFechaDesde] = useSessionState("ss-reporte-direccion-desde", "");
  const [fechaHasta, setFechaHasta] = useSessionState("ss-reporte-direccion-hasta", "");
  const [sort, setSort] = usePrefState("pref-reporte-direccion-sort", { key: "dia", dir: "desc" }, sanearSort(["dia","solicitante","proyecto","zona","proveedor","concepto","importe","moneda","a_partida","status"]));
  const filasFiltradas = filas.filter((f) => {
    if (fechaDesde && (!f.dia || f.dia < fechaDesde)) return false;
    if (fechaHasta && (!f.dia || f.dia > fechaHasta)) return false;
    if (!buscar.trim()) return true;
    const q = buscar.trim().toLowerCase();
    return [f.solicitante, f.proveedor, f.concepto, f.a_partida, f.proyecto, f.zona]
      .some((v) => (v || "").toString().toLowerCase().includes(q));
  });
  const filtrosFechaActivos = fechaDesde || fechaHasta;
  const limpiarFechas = () => { setFechaDesde(""); setFechaHasta(""); };
  const filasOrdenadas = sortRows(filasFiltradas, sort, { importe: (r) => r.importe });

  const colVisibility = useColumnVisibility("colv-reporte-direccion", COLUMNAS_REPORTE_DIRECCION);
  const colWidths = useColumnWidths("colw-reporte-direccion");
  const { ordered: columnas, moveColumn } = useColumnOrder("colo-reporte-direccion", colVisibility.visible);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };

  const exportarExcel = async () => {
    if (!filasOrdenadas.length) {
      alert("No hay transacciones en el filtro actual para exportar.");
      return;
    }
    const colsXls = columnas.filter((c) => c.xls);

    /* El periodo sale de las fechas que REALMENTE contiene el reporte, no de
       los filtros: si el rango es amplio pero solo hay pagos de un día, el
       archivo debe decir ese día. */
    const diasOrdenados = filasOrdenadas.map((f) => f.dia).filter(Boolean).sort();
    const inicio = fechaDesde || diasOrdenados[0] || "";
    const fin = fechaHasta || diasOrdenados[diasOrdenados.length - 1] || "";
    const etiquetaPeriodo = !inicio ? "Sin periodo" : (inicio === fin ? inicio : `${inicio} a ${fin}`);

    const fmtTot = (m, v) =>
      `$${(Number(v) || 0).toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${m}`;

    const wbx = new ExcelJS.Workbook();
    const nombreHoja = `${unidad} ${etiquetaPeriodo}`.replace(/[:\\/?*[\]]/g, "-").slice(0, 31);
    const ws = wbx.addWorksheet(nombreHoja);
    ws.columns = colsXls.map((c) => ({ width: c.width || 14 }));

    // --- Encabezado, con el mismo tratamiento que el Reporte Presupuestal ---
    const tit = ws.addRow([`Reporte de Pagos a Dirección — ${unidad} — ${etiquetaPeriodo}`]);
    tit.font = { bold: true, size: 14, name: "Calibri" };
    ws.mergeCells(1, 1, 1, Math.max(colsXls.length, 2));

    // Los totales van en el subtítulo, no sueltos en celdas fijas: ahí se leen
    // junto al periodo que los explica, y dejan de romperse si cambian las
    // columnas visibles.
    const totales = [];
    if (totalMXN) totales.push(fmtTot("MXP", totalMXN));
    if (totalUSD) totales.push(fmtTot("USD", totalUSD));
    const sub = ws.addRow([`${filasOrdenadas.length} transacciones   ·   ${totales.join("   ·   ") || "sin importes"}`]);
    sub.font = { size: 10, color: { argb: "FF6B7785" }, name: "Calibri" };
    ws.mergeCells(2, 1, 2, Math.max(colsXls.length, 2));
    ws.addRow([]);

    const hr = ws.addRow(colsXls.map((c) => c.label));
    hr.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Calibri" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
      cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
    });
    ws.getRow(hr.number).height = 22;

    filasOrdenadas.forEach((f) => {
      const row = ws.addRow(colsXls.map((c) => c.xls(f)));
      colsXls.forEach((c, ci) => {
        const cell = row.getCell(ci + 1);
        cell.font = { name: "Calibri", size: 11 };
        // Las cifras a la derecha: alineadas por el punto decimal se comparan
        // de un vistazo, que es para lo que sirve una columna de importes.
        cell.alignment = { horizontal: c.money ? "right" : (c.izq ? "left" : "center"), vertical: "top", wrapText: !!c.izq };
        if (c.money) cell.numFmt = '"$"#,##0.00';
      });
    });

    // --- Totales al pie, separados por moneda ---
    ws.addRow([]);
    [["MXP", totalMXN], ["USD", totalUSD]].forEach(([m, v]) => {
      if (!v) return;
      const row = ws.addRow([`TOTAL ${m}`, ...Array(Math.max(colsXls.length - 2, 0)).fill(""), v]);
      row.font = { bold: true, name: "Calibri" };
      const ult = row.getCell(colsXls.length);
      ult.numFmt = '"$"#,##0.00';
      row.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFECEEF1" } };
      });
    });

    // Encabezado congelado: con cien renglones, saber en qué columna se está
    // parado deja de ser evidente al hacer scroll.
    ws.views = [{ state: "frozen", ySplit: hr.number }];
    ws.autoFilter = { from: { row: hr.number, column: 1 }, to: { row: hr.number, column: colsXls.length } };

    const buffer = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Reporte de Pagos Dirección - ${unidad} - ${etiquetaPeriodo}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const [generandoReporte, setGenerandoReporte] = useState(false);
  /**
   * Estima si la tabla cabe a lo ancho de la hoja. autoTable no falla al
   * desbordarse: aprieta las columnas y parte las palabras sin avisar, y el
   * resultado llega ilegible a Dirección.
   */
  const anchoEstimadoPDF = () => {
    const disponible = 712; // carta horizontal menos márgenes
    const usado = columnas.reduce((suma, c) => {
      const largos = filasOrdenadas.map((f) => String(c.pdf ? c.pdf(f) : "").length);
      const max = Math.max(c.label.length, ...(largos.length ? largos : [0]));
      return suma + Math.min(max * 4.2, 140) + 8;
    }, 0);
    return { usado: Math.round(usado), disponible, cabe: usado <= disponible };
  };

  const generarReportePDF = async () => {
    if (!filasOrdenadas.length) {
      alert("No hay transacciones en el filtro actual para generar el reporte.");
      return;
    }
    const ancho = anchoEstimadoPDF();
    const avisoAncho = ancho.cabe ? "" :
      `\n\nOJO: con ${columnas.length} columnas la tabla se pasa del ancho de la hoja ` +
      `(~${ancho.usado} pt contra ${ancho.disponible} disponibles). Se van a apretar y el texto se va a partir. ` +
      `Considera ocultar algunas en "Columnas".`;
    const confirmado = confirm(
      `Esto va a generar un PDF con las ${filasOrdenadas.length} transacción(es) que tienes filtradas ahora, ` +
      `con las ${columnas.length} columnas visibles, y las va a marcar como "Reportadas". ¿Continuar?${avisoAncho}`
    );
    if (!confirmado) return;

    setGenerandoReporte(true);
    try {
      const diasOrdenados = filasOrdenadas.map((f) => f.dia).filter(Boolean).sort();
      const inicio = fechaDesde || diasOrdenados[0] || "";
      const fin = fechaHasta || diasOrdenados[diasOrdenados.length - 1] || "";

      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "letter" });
      doc.setFontSize(13);
      doc.text(`Reporte de pagos a realizar del dia ${inicio} al dia ${fin} Compañía ${unidad}`, 30, 30);
      doc.setFontSize(9);
      doc.setTextColor(120);
      doc.text(`Total a pagar MXP: ${money(totalMXN, "MXP")}   ·   Total a pagar USD: ${money(totalUSD, "USD")}`, 30, 46);

      // Las mismas columnas que están a la vista, y en el mismo orden: si el
      // PDF mostrara otras, la tabla dejaría de servir para revisar lo que se
      // va a enviar.
      autoTable(doc, {
        startY: 58,
        head: [columnas.map((c) => c.label)],
        body: filasOrdenadas.map((f) => columnas.map((c) => (c.pdf ? c.pdf(f) : ""))),
        styles: { fontSize: 7.5, cellPadding: 4 },
        headStyles: { fillColor: [62, 92, 118], textColor: 255, halign: "center" },
        bodyStyles: { halign: "center" },
        columnStyles: Object.fromEntries(
          columnas.map((c, i) => [i, { halign: c.izq ? "left" : "center" }])
        ),
      });

      doc.save(`reporte-pagos-direccion-${unidad}-${new Date().toISOString().slice(0, 10)}.pdf`);

      for (const f of filasOrdenadas) {
        await transaccionesApi.update(f.id, { reportado_at: new Date().toISOString() });
      }
    } catch (err) {
      alert("No se pudo generar el reporte: " + (err.message || err));
    } finally {
      setGenerandoReporte(false);
    }
  };
  const totalMXN = filasOrdenadas.filter((f) => f.moneda === "MXP").reduce((s, f) => s + f.importe, 0);
  const totalUSD = filasOrdenadas.filter((f) => f.moneda === "USD").reduce((s, f) => s + f.importe, 0);

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
        <KpiCard label="Total MXP (filtrado)" value={money(totalMXN, "MXP")} />
        <KpiCard label="Total USD (filtrado)" value={money(totalUSD, "USD")} />
      </div>

      <Panel
        title={`Reporte de Pagos Dirección — ${unidad}`}
        subtitle={`${filasOrdenadas.length} de ${filas.length} transacciones`}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <ColumnVisibilityControl
              columns={COLUMNAS_REPORTE_DIRECCION}
              hidden={colVisibility.hidden}
              onToggle={colVisibility.toggle}
              onShowAll={colVisibility.showAll}
            />
            <Button onClick={exportarExcel}>Exportar a Excel</Button>
            <Button onClick={generarReportePDF} disabled={generandoReporte}>
              {generandoReporte ? "Generando…" : "Generar reporte (PDF)"}
            </Button>
          </div>
        }
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
                    sumLabel={c.key === "importe" ? sumaPorMoneda(filasOrdenadas, "importe") : undefined}
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

/**
 * Catálogo de zonas. Se pueden dar de alta, renombrar, desactivar y borrar.
 *
 * Desactivar y borrar no son lo mismo, y la diferencia importa:
 *   · Desactivar la saca del selector pero la deja como referencia.
 *   · Borrar la quita del catálogo — las transacciones que ya la usaban
 *     conservan el texto, porque `transacciones.zona` no es llave foránea.
 *     Por eso se avisa cuántas quedarían con una zona fuera de catálogo.
 */
function ZonasPanel({ zonasApi, transacciones = [] }) {
  const [nueva, setNueva] = useState("");
  const [editandoId, setEditandoId] = useState(null);
  const [borrador, setBorrador] = useState("");

  const zonas = [...zonasApi.rows].sort(
    (a, b) => (a.orden ?? 999) - (b.orden ?? 999) || String(a.nombre).localeCompare(String(b.nombre))
  );
  const usos = (nombre) => transacciones.filter(
    (t) => String(t.zona || "").trim().toLowerCase() === String(nombre).trim().toLowerCase()
  ).length;

  // Zonas que aparecen en transacciones pero no están dadas de alta: llegan
  // por carga masiva o desde los correos, y sin esto pasarían inadvertidas.
  const enCatalogo = new Set(zonas.map((z) => String(z.nombre).trim().toLowerCase()));
  const huerfanas = [...new Set(
    transacciones.map((t) => String(t.zona || "").trim()).filter(Boolean)
      .filter((z) => !enCatalogo.has(z.toLowerCase()))
  )].sort();

  const agregar = async () => {
    const nombre = nueva.trim();
    if (!nombre) return;
    if (enCatalogo.has(nombre.toLowerCase())) { alert(`"${nombre}" ya está en el catálogo.`); return; }
    try {
      await zonasApi.insert({ id: uid(), nombre, activa: true, orden: (zonas[zonas.length - 1]?.orden ?? zonas.length) + 1 });
      setNueva("");
    } catch (err) { alert("No se pudo agregar: " + (err.message || err)); }
  };

  const guardar = async (id) => {
    const nombre = borrador.trim();
    if (!nombre) return;
    try {
      await zonasApi.update(id, { nombre });
      setEditandoId(null);
    } catch (err) { alert("No se pudo actualizar: " + (err.message || err)); }
  };

  const alternarActiva = (z) =>
    zonasApi.update(z.id, { activa: z.activa === false })
      .catch((err) => alert("No se pudo actualizar: " + (err.message || err)));

  const borrar = (z) => {
    const n = usos(z.nombre);
    const aviso = n
      ? `\n\nOJO: ${n} transacción(es) usan esta zona. No se van a modificar —conservan el texto— pero quedarán con una zona que ya no existe en el catálogo. Si solo quieres sacarla del selector, mejor desactívala.`
      : "";
    if (!confirm(`¿Eliminar la zona "${z.nombre}"?${aviso}`)) return;
    zonasApi.remove(z.id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  return (
    <Panel
      title="Catálogo de zonas"
      subtitle={`${zonas.filter((z) => z.activa !== false).length} activas — alimentan el selector de Zona al capturar transacciones`}
    >
      <div style={{ display: "flex", gap: 8, alignItems: "flex-end", marginBottom: 14, flexWrap: "wrap" }}>
        <Field label="Nueva zona">
          <TextInput
            value={nueva}
            onChange={(e) => setNueva(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") agregar(); }}
            placeholder="Nombre de la zona"
            style={{ width: 240 }}
          />
        </Field>
        <Button onClick={agregar} disabled={!nueva.trim()}>+ Agregar</Button>
      </div>

      {!zonasApi.rows.length && (
        <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "10px 13px", borderRadius: "0 6px 6px 0", fontSize: 12, marginBottom: 12 }}>
          <b>La tabla de zonas está vacía</b>
          Mientras tanto el selector usa la lista de respaldo. Corre la migración
          <code style={{ fontFamily: T.fontMono, margin: "0 4px" }}>06-catalogo-zonas.sql</code>
          para poblarla con las 11 originales.
        </div>
      )}

      {huerfanas.length > 0 && (
        <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "10px 13px", borderRadius: "0 6px 6px 0", fontSize: 12, marginBottom: 12 }}>
          <b>Zonas en uso que no están en el catálogo</b>
          Aparecen en transacciones pero nadie las dio de alta, así que no salen en el selector:
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8 }}>
            {huerfanas.map((h) => (
              <Button key={h} variant="ghost" style={{ padding: "3px 9px", fontSize: 11 }}
                onClick={() => zonasApi.insert({ id: uid(), nombre: h, activa: true, orden: 900 })
                  .catch((err) => alert("No se pudo agregar: " + (err.message || err)))}>
                + {h} <span style={{ color: T.textFaint }}>({usos(h)})</span>
              </Button>
            ))}
          </div>
        </div>
      )}

      <table style={tableStyle}>
        <thead>
          <tr>
            <th style={thStyle}>Zona</th>
            <th style={thStyle}>Transacciones</th>
            <th style={thStyle}>Estado</th>
            <th style={thStyle}></th>
          </tr>
        </thead>
        <tbody>
          {zonas.map((z) => (
            <tr key={z.id} style={{ opacity: z.activa === false ? 0.55 : 1 }}>
              <td style={tdStyle}>
                {editandoId === z.id ? (
                  <TextInput value={borrador} onChange={(e) => setBorrador(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") guardar(z.id); }} style={{ width: 220 }} />
                ) : z.nombre}
              </td>
              <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{usos(z.nombre) || "—"}</td>
              <td style={tdStyle}>
                <Pill tone={z.activa === false ? "amber" : "teal"}>{z.activa === false ? "Inactiva" : "Activa"}</Pill>
              </td>
              <td style={tdStyle}>
                <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                  {editandoId === z.id ? (
                    <>
                      <Button onClick={() => guardar(z.id)} style={{ padding: "4px 10px", fontSize: 11.5 }}>Guardar</Button>
                      <Button variant="ghost" onClick={() => setEditandoId(null)} style={{ padding: "4px 10px", fontSize: 11.5 }}>Cancelar</Button>
                    </>
                  ) : (
                    <>
                      <IconButton icon="✎" label="Renombrar" tone={T.accent}
                        onClick={() => { setEditandoId(z.id); setBorrador(z.nombre); }} />
                      <IconButton icon={z.activa === false ? "◻" : "◼"}
                        label={z.activa === false ? "Activar" : "Desactivar"}
                        onClick={() => alternarActiva(z)} />
                      <IconButton icon="✕" label="Eliminar" tone={T.red} onClick={() => borrar(z)} />
                    </>
                  )}
                </div>
              </td>
            </tr>
          ))}
          {!zonas.length && (
            <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>
              Todavía no hay zonas en el catálogo
            </td></tr>
          )}
        </tbody>
      </table>
    </Panel>
  );
}

/**
 * Catálogo de rubros y categorías.
 *
 * El punto delicado es RENOMBRAR. `partidas.rubro`, `partidas.categoria` y
 * `transacciones.categoria` guardan el TEXTO, no un id, así que cambiar el
 * nombre aquí dejaría a los registros existentes apuntando a algo que ya no
 * existe en el catálogo. Por eso al renombrar se propaga el cambio a los
 * registros en uso, en el mismo paso y avisando cuántos son.
 *
 * Se guarda el texto y no un id a propósito: así el histórico conserva cómo
 * se llamaba el rubro cuando se capturó, y borrar una categoría no deja
 * registros rotos.
 */
function RubrosPanel({ rubrosApi, categoriasApi, partidas = [], transacciones = [] }) {
  const [nuevoRubro, setNuevoRubro] = useState("");
  const [rubroSel, setRubroSel] = useState(null);
  const [nuevaCat, setNuevaCat] = useState("");

  const rubros = [...rubrosApi.rows].sort(
    (a, b) => (a.orden ?? 999) - (b.orden ?? 999) || String(a.nombre).localeCompare(String(b.nombre))
  );
  const catsDe = (rid) => categoriasApi.rows
    .filter((c) => c.rubro_id === rid)
    .sort((a, b) => (a.orden ?? 999) - (b.orden ?? 999) || String(a.nombre).localeCompare(String(b.nombre)));

  const igual = (a, b) => String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
  const usosRubro = (nom) => partidas.filter((p) => igual(p.rubro, nom)).length;
  const usosCat = (nom) =>
    partidas.filter((p) => igual(p.categoria, nom)).length +
    transacciones.filter((t) => igual(t.categoria, nom)).length;

  const agregarRubro = async () => {
    const nombre = nuevoRubro.trim();
    if (!nombre) return;
    if (rubros.some((r) => igual(r.nombre, nombre))) { alert(`El rubro "${nombre}" ya existe.`); return; }
    try {
      const orden = (rubros[rubros.length - 1]?.orden ?? rubros.length) + 1;
      await rubrosApi.insert({ id: uid(), nombre, orden, activo: true });
      // Nace SIN categorías, a propósito. Crearle una genérica invitaría a
      // usarla en vez de definir las que el rubro realmente necesita.
      setNuevoRubro("");
      alert(`Rubro "${nombre}" creado. Agrégale sus categorías: un rubro sin categorías no sirve para clasificar.`);
    } catch (err) { alert("No se pudo agregar: " + (err.message || err)); }
  };

  const renombrarRubro = async (r) => {
    const nombre = (prompt(`Nuevo nombre para el rubro "${r.nombre}":`, r.nombre) || "").trim();
    if (!nombre || igual(nombre, r.nombre)) return;
    if (rubros.some((x) => x.id !== r.id && igual(x.nombre, nombre))) { alert(`Ya existe un rubro "${nombre}".`); return; }
    const n = usosRubro(r.nombre);
    if (n && !confirm(`${n} partida(s) usan el rubro "${r.nombre}". Se les va a cambiar el nombre también, para que no queden apuntando a un rubro inexistente.\n\n¿Continuar?`)) return;
    try {
      await rubrosApi.update(r.id, { nombre });
      for (const p of partidas.filter((p) => igual(p.rubro, r.nombre))) {
        await supabase.from("partidas").update({ rubro: nombre }).eq("id", p.id);
      }
    } catch (err) { alert("No se pudo renombrar: " + (err.message || err)); }
  };

  const borrarRubro = (r) => {
    const n = usosRubro(r.nombre);
    if (n) { alert(`No se puede eliminar "${r.nombre}": ${n} partida(s) lo usan.\n\nSi ya no se debe usar, desactívalo — sale del selector pero el histórico se conserva.`); return; }
    if (!confirm(`¿Eliminar el rubro "${r.nombre}" y sus ${catsDe(r.id).length} categoría(s)?`)) return;
    rubrosApi.remove(r.id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  const agregarCat = async () => {
    const nombre = nuevaCat.trim();
    if (!nombre || !rubroSel) return;
    if (catsDe(rubroSel.id).some((c) => igual(c.nombre, nombre))) { alert(`"${nombre}" ya existe en ${rubroSel.nombre}.`); return; }
    // Se bloquea recrear el comodín: si vuelve a existir, vuelve a usarse.
    if (["diversos", "varios", "otros", "general", "generales"].includes(nombre.toLowerCase())) {
      alert(`No se permite "${nombre}" como categoría.\n\nUna opción genérica se vuelve el camino de menor resistencia y la clasificación se degrada: acaba respondiendo "en cosas" a la pregunta de en qué gastamos. Si un gasto no encaja, déjalo sin categoría — el hueco se ve y se corrige.`);
      return;
    }
    try {
      const cs = catsDe(rubroSel.id);
      await categoriasApi.insert({ id: uid(), rubro_id: rubroSel.id, nombre,
        orden: (cs[cs.length - 1]?.orden ?? cs.length) + 1, activa: true });
      setNuevaCat("");
    } catch (err) { alert("No se pudo agregar: " + (err.message || err)); }
  };

  const renombrarCat = async (c) => {
    const nombre = (prompt(`Nuevo nombre para la categoría "${c.nombre}":`, c.nombre) || "").trim();
    if (!nombre || igual(nombre, c.nombre)) return;
    const n = usosCat(c.nombre);
    if (n && !confirm(`${n} registro(s) usan la categoría "${c.nombre}". Se les va a cambiar el nombre también.\n\n¿Continuar?`)) return;
    try {
      await categoriasApi.update(c.id, { nombre });
      for (const p of partidas.filter((p) => igual(p.categoria, c.nombre))) {
        await supabase.from("partidas").update({ categoria: nombre }).eq("id", p.id);
      }
      for (const t of transacciones.filter((t) => igual(t.categoria, c.nombre))) {
        await supabase.from("transacciones").update({ categoria: nombre }).eq("id", t.id);
      }
    } catch (err) { alert("No se pudo renombrar: " + (err.message || err)); }
  };

  const borrarCat = (c) => {
    const n = usosCat(c.nombre);
    if (n) { alert(`No se puede eliminar "${c.nombre}": ${n} registro(s) la usan.\n\nDesactívala si ya no se debe usar.`); return; }
    if (!confirm(`¿Eliminar la categoría "${c.nombre}"?`)) return;
    categoriasApi.remove(c.id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  const activo = rubroSel && rubros.find((r) => r.id === rubroSel.id);

  return (
    <Panel
      title="Rubros y categorías"
      subtitle={`${rubros.length} rubros · ${categoriasApi.rows.length} categorías — alimentan Partidas y Transacciones`}
    >
      {!rubrosApi.rows.length && (
        <div style={{ borderLeft: `3px solid ${T.amber}`, background: "#FDF8EF", padding: "10px 13px", borderRadius: "0 6px 6px 0", fontSize: 12, marginBottom: 14 }}>
          <b>El catálogo de la base está vacío</b>
          Mientras tanto se usa la lista del código. Corre
          <code style={{ fontFamily: T.fontMono, margin: "0 4px" }}>16-catalogo-editable.sql</code>
          para poblarla con los 15 rubros y 122 categorías actuales.
        </div>
      )}

      <div style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
        {/* Rubros */}
        <div style={{ flex: "1 1 340px", minWidth: 300 }}>
          <div style={{ display: "flex", gap: 8, alignItems: "flex-end", marginBottom: 10 }}>
            <Field label="Nuevo rubro">
              <TextInput value={nuevoRubro} onChange={(e) => setNuevoRubro(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") agregarRubro(); }}
                placeholder="Nombre del rubro" style={{ width: 210 }} />
            </Field>
            <Button onClick={agregarRubro} disabled={!nuevoRubro.trim()}>+ Agregar</Button>
          </div>
          <table style={tableStyle}>
            <thead><tr><th style={thStyle}>Rubro</th><th style={thStyle}>Cats.</th><th style={thStyle}>Partidas</th><th style={thStyle}></th></tr></thead>
            <tbody>
              {rubros.map((r) => (
                <tr key={r.id}
                  onClick={() => setRubroSel(r)}
                  style={{ cursor: "pointer", opacity: r.activo === false ? 0.55 : 1,
                           background: rubroSel?.id === r.id ? T.panelAlt : "transparent" }}>
                  <td style={tdStyle}>{r.nombre}</td>
                  <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{catsDe(r.id).length}</td>
                  <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{usosRubro(r.nombre) || "—"}</td>
                  <td style={tdStyle}>
                    <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                      <IconButton icon="✎" label="Renombrar" tone={T.accent}
                        onClick={(e) => { e.stopPropagation(); renombrarRubro(r); }} />
                      <IconButton icon={r.activo === false ? "◻" : "◼"}
                        label={r.activo === false ? "Activar" : "Desactivar"}
                        onClick={(e) => { e.stopPropagation();
                          rubrosApi.update(r.id, { activo: r.activo === false })
                            .catch((err) => alert("No se pudo: " + (err.message || err))); }} />
                      <IconButton icon="✕" label="Eliminar" tone={T.red}
                        onClick={(e) => { e.stopPropagation(); borrarRubro(r); }} />
                    </div>
                  </td>
                </tr>
              ))}
              {!rubros.length && (
                <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin rubros en la base</td></tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Categorías del rubro elegido */}
        <div style={{ flex: "1 1 340px", minWidth: 300 }}>
          {!activo ? (
            <div style={{ color: T.textFaint, fontSize: 12.5, padding: "40px 0", textAlign: "center" }}>
              Elige un rubro de la izquierda para ver y editar sus categorías
            </div>
          ) : (
            <>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10 }}>
                Categorías de {activo.nombre}
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "flex-end", marginBottom: 10 }}>
                <Field label="Nueva categoría">
                  <TextInput value={nuevaCat} onChange={(e) => setNuevaCat(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") agregarCat(); }}
                    placeholder="Nombre de la categoría" style={{ width: 210 }} />
                </Field>
                <Button onClick={agregarCat} disabled={!nuevaCat.trim()}>+ Agregar</Button>
              </div>
              <table style={tableStyle}>
                <thead><tr><th style={thStyle}>Categoría</th><th style={thStyle}>En uso</th><th style={thStyle}></th></tr></thead>
                <tbody>
                  {catsDe(activo.id).map((c) => (
                    <tr key={c.id} style={{ opacity: c.activa === false ? 0.55 : 1 }}>
                      <td style={tdStyle}>{c.nombre}</td>
                      <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{usosCat(c.nombre) || "—"}</td>
                      <td style={tdStyle}>
                        <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                          <IconButton icon="✎" label="Renombrar" tone={T.accent} onClick={() => renombrarCat(c)} />
                          <IconButton icon={c.activa === false ? "◻" : "◼"}
                            label={c.activa === false ? "Activar" : "Desactivar"}
                            onClick={() => categoriasApi.update(c.id, { activa: c.activa === false })
                              .catch((err) => alert("No se pudo: " + (err.message || err)))} />
                          <IconButton icon="✕" label="Eliminar" tone={T.red} onClick={() => borrarCat(c)} />
                        </div>
                      </td>
                    </tr>
                  ))}
                  {!catsDe(activo.id).length && (
                    <tr><td colSpan={3} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin categorías</td></tr>
                  )}
                </tbody>
              </table>
            </>
          )}
        </div>
      </div>
    </Panel>
  );
}

/**
 * Ajustes por compañía. Hoy solo la numeración de las Solicitudes de Pago y
 * los datos que se repiten en todas ellas, escritos hasta ahora a mano.
 */
function ConfigCompaniaPanel({ unidad }) {
  const [cfg, setCfg] = useState(null);
  const [ultimoEmitido, setUltimoEmitido] = useState(0);
  const [guardando, setGuardando] = useState(false);

  useEffect(() => {
    let vivo = true;
    (async () => {
      const { data } = await supabase.from("config_companias").select("*").eq("compania", unidad).maybeSingle();
      const { data: max } = await supabase.from("solicitudes_pago").select("folio")
        .eq("compania", unidad).order("folio", { ascending: false }).limit(1);
      if (!vivo) return;
      setCfg(data || { compania: unidad, spp_ultimo: 0, spp_responsable: "", spp_lugar_adquisicion: "" });
      setUltimoEmitido((max && max[0]?.folio) || 0);
    })();
    return () => { vivo = false; };
  }, [unidad]);

  if (!cfg) return null;
  const siguiente = Math.max(Number(cfg.spp_ultimo) || 0, ultimoEmitido) + 1;

  const guardar = async () => {
    setGuardando(true);
    try {
      const { error } = await supabase.from("config_companias").upsert({
        compania: unidad,
        spp_ultimo: Number(cfg.spp_ultimo) || 0,
        spp_responsable: cfg.spp_responsable || "",
        spp_lugar_adquisicion: cfg.spp_lugar_adquisicion || "",
        updated_at: new Date().toISOString(),
      });
      if (error) throw error;
    } catch (err) {
      alert("No se pudo guardar: " + (err.message || err));
    } finally { setGuardando(false); }
  };

  return (
    <Panel title={`Solicitudes de Pago — ${unidad}`}
      subtitle="Numeración y datos que se repiten en todas las solicitudes de esta compañía">
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 12 }}>
        <Field label="Último folio usado">
          <TextInput type="number" value={cfg.spp_ultimo ?? 0}
            onChange={(e) => setCfg({ ...cfg, spp_ultimo: e.target.value })} />
        </Field>
        <Field label="Responsable de proyecto">
          <TextInput value={cfg.spp_responsable || ""}
            onChange={(e) => setCfg({ ...cfg, spp_responsable: e.target.value })}
            placeholder="Se precarga en cada solicitud" />
        </Field>
        <Field label="Lugar de adquisición">
          <TextInput value={cfg.spp_lugar_adquisicion || ""}
            onChange={(e) => setCfg({ ...cfg, spp_lugar_adquisicion: e.target.value })}
            placeholder="Queretaro" />
        </Field>
      </div>

      {/* El folio real es el mayor entre lo configurado y lo ya emitido: si
          se baja el número por error, no se reutilizan folios de solicitudes
          que ya salieron. */}
      <div style={{ fontSize: 12, color: T.textDim, marginBottom: 12 }}>
        La siguiente solicitud llevará el folio <b style={{ fontFamily: T.fontMono }}>{siguiente}</b>.
        {ultimoEmitido > (Number(cfg.spp_ultimo) || 0) && (
          <span style={{ color: T.amber }}>
            {" "}Ya se emitió el folio {ultimoEmitido}, así que la numeración continúa desde ahí
            aunque el valor configurado sea menor.
          </span>
        )}
      </div>

      <Button onClick={guardar} disabled={guardando}>
        {guardando ? "Guardando…" : "Guardar"}
      </Button>
    </Panel>
  );
}

const SUBS_CATALOGO = [
  { id: "proyectos", label: "Proyectos" },
  { id: "rubros", label: "Rubros y categorías" },
  { id: "zonas", label: "Zonas" },
  { id: "proveedores", label: "Proveedores" },
  { id: "solicitudes", label: "Solicitudes de Pago" },
];

function CatalogoTab({ unidad, unidades, proyectosApi, zonasApi, rubrosApi, categoriasApi, partidas = [], transacciones = [], proveedoresApi, cuentasApi, perfilesApi }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const [sub, setSub] = useSessionState("ss-catalogo-sub", "proyectos");
  const [nuevo, setNuevo] = useState({ nombre: "", grupo: "", pct: "", centro_costo: "" });
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({ nombre: "", grupo: "", pct: "", centro_costo: "" });
  const [guardandoId, setGuardandoId] = useState(null);

  const empezarEditar = (p) => {
    setEditingId(p.id);
    setDraft({ nombre: p.nombre || "", grupo: p.grupo || "", pct: p.pct ?? "", centro_costo: p.centro_costo || "" });
  };
  const cancelarEditar = () => { setEditingId(null); setDraft({ nombre: "", grupo: "", pct: "" }); };
  const guardarEditar = async (id) => {
    setGuardandoId(id);
    try {
      await proyectosApi.update(id, { nombre: draft.nombre, grupo: draft.grupo, pct: Number(draft.pct) || 0, centro_costo: draft.centro_costo.trim() });
      cancelarEditar();
    } catch (err) {
      alert("No se pudo actualizar: " + (err.message || err));
    } finally {
      setGuardandoId(null);
    }
  };
  const removeProyecto = (id) => {
    const p = proyectosApi.rows.find((x) => x.id === id);
    if (!confirm(`¿Eliminar el proyecto "${p?.nombre || id}" del catálogo? Esto no se puede deshacer.`)) return;
    proyectosApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };
  const addProyecto = () => {
    if (!nuevo.nombre) return;
    proyectosApi.insert({ id: uid(), unidad, nombre: nuevo.nombre, grupo: nuevo.grupo || "General", pct: Number(nuevo.pct) || 0, centro_costo: nuevo.centro_costo.trim() })
      .catch((err) => alert("No se pudo agregar: " + (err.message || err)));
    setNuevo({ nombre: "", grupo: "", pct: "" });
  };

  const totalPct = proyectosUnidad.reduce((s, p) => s + Number(p.pct || 0), 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* Sub-pestañas en vez de seis paneles apilados. Se prefirió esto a
          secciones plegables: con seis encabezados que atravesar el recorrido
          sigue existiendo, y el estado de plegado se olvida entre visitas. */}
      <div style={{ display: "flex", background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8, padding: 3, alignSelf: "flex-start", flexWrap: "wrap" }}>
        {SUBS_CATALOGO.map((sc) => (
          <button
            key={sc.id}
            onClick={() => setSub(sc.id)}
            style={{
              padding: "7px 16px", borderRadius: 6, border: "none", cursor: "pointer",
              background: sub === sc.id ? T.accent : "transparent",
              color: sub === sc.id ? "#FFFFFF" : T.textDim,
              fontWeight: 600, fontSize: 12.5, fontFamily: T.fontUI,
            }}
          >
            {sc.label}
          </button>
        ))}
      </div>

      {sub === "proyectos" && (
      <Panel
        title={`Proyectos y % de prorrateo — ${unidad}`}
        subtitle="El % se usa para repartir gastos compartidos marcados como “Todos” o “<Grupo> Gral”"
        right={<Pill tone={Math.abs(totalPct - 100) < 0.01 ? "teal" : "red"}>{totalPct.toFixed(1)}% asignado</Pill>}
      >
        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>{["Proyecto","Grupo","% Administrativos","Centro de costo",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {proyectosUnidad.map((p) => {
                const editando = editingId === p.id;
                return (
                  <tr key={p.id}>
                    {editando ? (
                      <>
                        <td style={tdStyle}><TextInput autoFocus value={draft.nombre} onChange={(e) => setDraft({ ...draft, nombre: e.target.value })} /></td>
                        <td style={tdStyle}><TextInput value={draft.grupo} onChange={(e) => setDraft({ ...draft, grupo: e.target.value })} placeholder="Desh / Prod / IMP" /></td>
                        <td style={tdStyle}><TextInput type="number" value={draft.pct} onChange={(e) => setDraft({ ...draft, pct: e.target.value })} style={{ width: 90 }} /></td>
                        <td style={tdStyle}><TextInput value={draft.centro_costo} onChange={(e) => setDraft({ ...draft, centro_costo: e.target.value.toUpperCase() })} placeholder="CC-015" style={{ width: 100 }} /></td>
                        <td style={tdStyle}>
                          <div style={{ display: "flex", gap: 6 }}>
                            <Button onClick={() => guardarEditar(p.id)} disabled={guardandoId === p.id}>{guardandoId === p.id ? "Guardando…" : "Guardar"}</Button>
                            <Button variant="ghost" onClick={cancelarEditar}>Cancelar</Button>
                          </div>
                        </td>
                      </>
                    ) : (
                      <>
                        <td style={tdStyle}>{p.nombre}</td>
                        <td style={{ ...tdStyle, color: T.textDim }}>{p.grupo || "—"}</td>
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{p.pct ?? 0}%</td>
                        {/* Sin centro de costo la Solicitud de Pago sale incompleta, así que
                            el hueco se marca en vez de mostrarse como un guion cualquiera. */}
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>
                          {p.centro_costo || <span style={{ color: T.amber }}>falta</span>}
                        </td>
                        <td style={tdStyle}>
                          <div style={{ display: "flex", gap: 6 }}>
                            <Button variant="ghost" onClick={() => empezarEditar(p)}>Editar</Button>
                            <Button variant="danger" onClick={() => removeProyecto(p.id)}>Eliminar</Button>
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                );
              })}
              <tr>
                <td style={tdStyle}><TextInput placeholder="Nuevo proyecto" value={nuevo.nombre} onChange={(e) => setNuevo({ ...nuevo, nombre: e.target.value })} /></td>
                <td style={tdStyle}><TextInput placeholder="Grupo" value={nuevo.grupo} onChange={(e) => setNuevo({ ...nuevo, grupo: e.target.value })} /></td>
                <td style={tdStyle}><TextInput type="number" placeholder="%" value={nuevo.pct} onChange={(e) => setNuevo({ ...nuevo, pct: e.target.value })} style={{ width: 90 }} /></td>
                <td style={tdStyle}><TextInput placeholder="CC-015" value={nuevo.centro_costo} onChange={(e) => setNuevo({ ...nuevo, centro_costo: e.target.value.toUpperCase() })} style={{ width: 100 }} /></td>
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
      )}

      {sub === "rubros" && (
        <RubrosPanel rubrosApi={rubrosApi} categoriasApi={categoriasApi} partidas={partidas} transacciones={transacciones} />
      )}
      {sub === "zonas" && <ZonasPanel zonasApi={zonasApi} transacciones={transacciones} />}
      {sub === "proveedores" && (
        <ProveedoresPanel unidad={unidad} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} />
      )}
      {sub === "solicitudes" && <ConfigCompaniaPanel unidad={unidad} />}
    </div>
  );
}

function ImportarProveedoresPanel({ proveedoresApi, cuentasApi }) {
  const inputRef = useRef(null);
  /* Plantilla de carga masiva. Faltaba: había importador pero no había de
     dónde sacar el formato, así que la única forma de saber qué columnas
     espera era exportar el catálogo existente — que no sirve si aún no hay
     proveedores capturados. */
  const descargarPlantilla = async () => {
    const headers = ["Compañía", "Nombre", "RFC", "ID SAE", "Referencia", "Notas",
      "Banco", "Sucursal", "SWIFT", "CLABE", "Numero de Cuenta", "Divisa"];
    const anchos = [12, 40, 16, 12, 18, 30, 20, 12, 14, 24, 20, 9];
    const ejemplos = {
      OSB: ["OSB", "Ejemplo: Comisión Federal de Electricidad", "CFE370814QI0", "", "", "", "BANORTE", "", "", "072680013482913231", "", "MXP"],
      CTM: ["CTM", "Ejemplo: Servicio de Gasolinería Lepacaba SA de CV", "SGL010203AB1", "", "", "", "BBVA", "", "", "012680001234567890", "", "MXP"],
      ISE: ["ISE", "Ejemplo: Champion X", "", "", "", "Proveedor extranjero", "CITIBANK", "", "CITIUS33", "", "1234567890", "USD"],
    };

    const wbx = new ExcelJS.Workbook();
    UNIDAD_KEYS.forEach((u) => {
      const ws = wbx.addWorksheet(u);
      ws.columns = anchos.map((w) => ({ width: w }));
      const headerRow = ws.addRow(headers);
      formatearHojaDatos(ws, headerRow, headers.length);
      const ejemploRow = ws.addRow(ejemplos[u]);
      ejemploRow.eachCell((cell) => { cell.font = { italic: true, color: { argb: "FF8B99A6" } }; });
    });

    const notas = wbx.addWorksheet("Instrucciones");
    notas.columns = [{ width: 100 }];
    [
      "CÓMO LLENAR ESTA PLANTILLA",
      "",
      "1. Borra la fila de ejemplo (va en gris cursiva) antes de importar.",
      "2. Compañía: OSB, CTM o ISE. Si la dejas vacía se usa el nombre de la hoja.",
      "3. Nombre: obligatorio. Es lo que empareja con los proveedores que ya existen,",
      "   junto con el RFC. Si el RFC coincide, se ACTUALIZA el proveedor en vez de",
      "   duplicarlo.",
      "4. Un proveedor con VARIAS cuentas bancarias va en varias filas: repite el",
      "   nombre y el RFC, y cambia solo los datos de la cuenta.",
      "5. CLABE y Número de Cuenta: escríbelos como TEXTO, no como número.",
      "   Excel convierte una CLABE de 18 dígitos a notación científica y se pierden",
      "   los últimos dígitos y el cero inicial. Si la celda muestra 7.26801E+17,",
      "   el dato ya se corrompió: da formato de texto a la columna ANTES de pegar.",
      "6. Divisa: MXP o USD. Determina con qué cuenta se paga cada transacción,",
      "   así que una cuenta en dólares mal marcada manda el pago a la cuenta",
      "   equivocada.",
      "7. SWIFT: solo para transferencias internacionales.",
      "8. ID SAE: el identificador del proveedor en ASPEL. Sirve para empatar por id",
      "   en vez de por nombre, que es mucho más confiable.",
    ].forEach((t) => notas.addRow([t]));
    notas.getRow(1).font = { bold: true, size: 13 };

    const buf = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Plantilla-Proveedores.xlsx";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };


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
        const payload = nuevasFilas.map(({ _existenteId, _cuentas, ...rest }) => rest);
        const creados = await proveedoresApi.bulkInsert(payload);
        creados.forEach((c, i) => {
          (nuevasFilas[i]._cuentas || []).forEach((cuenta) => {
            cuentasAInsertar.push({ id: uid(), proveedor_id: c.id, ...cuenta });
          });
        });
      }
      for (const r of actualizacionesFilas) {
        const { id, _existenteId, _cuentas, ...patch } = r;
        await proveedoresApi.update(_existenteId, patch);
        (_cuentas || []).forEach((cuenta) => {
          cuentasAInsertar.push({ id: uid(), proveedor_id: _existenteId, ...cuenta });
        });
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
    <Panel
      title="Carga masiva de proveedores"
      subtitle='Una fila por proveedor, con una columna "Compañía" (OSB/CTM/ISE) que dice a cuál pertenece — o usa hojas llamadas OSB/CTM/ISE como respaldo'
      right={<Button variant="ghost" onClick={descargarPlantilla}>Descargar plantilla</Button>}
    >
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
                <tr>{["","Unidad","Nombre","RFC","Cuentas"].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.slice(0, 15).map((r) => (
                  <tr key={r.id}>
                    <td style={tdStyle}>{r._existenteId ? <Pill tone="amber">Actualiza</Pill> : <Pill tone="teal">Nuevo</Pill>}</td>
                    <td style={tdStyle}>{r.unidad}</td>
                    <td style={tdStyle}>{r.nombre}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{r.rfc || "—"}</td>
                    <td style={tdStyle}>
                      {r._cuentas.length
                        ? r._cuentas.map((c) => c.banco || c.divisa).filter(Boolean).join(", ")
                        : "—"}
                    </td>
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

function ProveedoresPanel({ unidad, proveedoresApi, cuentasApi, perfilesApi }) {
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const blank = { nombre: "", rfc: "", id_sae: "", referencia: "", notas: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [saving, setSaving] = useState(false);
  const [buscar, setBuscar] = useState("");

  const cuentaBlank = { banco: "", sucursal: "", swift: "", clabe: "", numero_cuenta: "", divisa: "MXP" };
  const [nuevaCuenta, setNuevaCuenta] = useState(cuentaBlank);
  const cuentasDelProveedor = editId ? cuentasApi.rows.filter((c) => c.proveedor_id === editId) : [];

  const contarCuentas = (proveedorId) => cuentasApi.rows.filter((c) => c.proveedor_id === proveedorId).length;

  const filtrados = proveedoresUnidad.filter((p) => {
    if (!buscar.trim()) return true;
    const q = buscar.trim().toLowerCase();
    return [p.nombre, p.rfc, p.id_sae, p.referencia, p.notas].some((v) => (v || "").toLowerCase().includes(q));
  });

  const openNew = () => { setForm(blank); setEditId(null); setNuevaCuenta(cuentaBlank); setModalOpen(true); };
  const startEdit = (p) => { setForm(p); setEditId(p.id); setNuevaCuenta(cuentaBlank); setModalOpen(true); };
  const closeModal = () => { setModalOpen(false); setEditId(null); setForm(blank); setNuevaCuenta(cuentaBlank); };
  const remove = (id) => {
    const p = proveedoresApi.rows.find((x) => x.id === id);
    if (!confirm(`¿Eliminar al proveedor "${p?.nombre || id}"? Esto no se puede deshacer y también borrará sus cuentas bancarias.`)) return;
    proveedoresApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

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
  const eliminarCuenta = (id) => {
    if (!confirm("¿Eliminar esta cuenta bancaria? Esto no se puede deshacer.")) return;
    cuentasApi.remove(id).catch((err) => alert("No se pudo eliminar la cuenta: " + (err.message || err)));
  };

  // Exporta las tres compañías con UNA FILA POR CUENTA BANCARIA (un proveedor
  // con dos cuentas ocupa dos renglones). Así el preparador de transacciones
  // puede comparar la CLABE que viene en el correo contra la registrada.
  const exportarProveedores = async () => {
    const COLS = [
      { header: "Unidad",     width: 9,  get: (p, c) => p.unidad },
      { header: "Proveedor",  width: 38, get: (p, c) => p.nombre || "" },
      { header: "RFC",        width: 16, get: (p, c) => p.rfc || "" },
      { header: "Id SAE",     width: 11, get: (p, c) => p.id_sae || "" },
      { header: "Banco",      width: 20, get: (p, c) => c?.banco || "" },
      { header: "CLABE",      width: 22, get: (p, c) => c?.clabe || "" },
      { header: "No. Cuenta", width: 18, get: (p, c) => c?.numero_cuenta || "" },
      { header: "SWIFT",      width: 12, get: (p, c) => c?.swift || "" },
      { header: "Divisa",     width: 9,  get: (p, c) => c?.divisa || "" },
      { header: "Cuentas",    width: 9,  get: (p, c, n) => n },
    ];
    const wbx = new ExcelJS.Workbook();
    UNIDAD_KEYS.forEach((u) => {
      const delaUnidad = proveedoresApi.rows.filter((p) => p.unidad === u);
      if (!delaUnidad.length) return;
      const ws = wbx.addWorksheet(`Proveedores-${u}`);
      ws.columns = COLS.map((c) => ({ width: c.width }));
      const hr = ws.addRow(COLS.map((c) => c.header));
      formatearHojaDatos(ws, hr, COLS.length);
      delaUnidad.forEach((p) => {
        const ctas = cuentasApi.rows.filter((c) => c.proveedor_id === p.id);
        // Sin cuentas se escribe un renglón igual, para que el preparador
        // pueda marcar "este proveedor no tiene cuenta registrada".
        (ctas.length ? ctas : [null]).forEach((c) => {
          const row = ws.addRow(COLS.map((col) => col.get(p, c, ctas.length)));
          COLS.forEach((col, ci) => {
            const cell = row.getCell(ci + 1);
            cell.font = { name: "Calibri", size: 11 };
            // CLABE y cuenta SIEMPRE como texto: si Excel las toma por número,
            // pierde el cero inicial y los últimos dígitos.
            if (col.header === "CLABE" || col.header === "No. Cuenta") {
              cell.numFmt = "@";
              cell.alignment = { horizontal: "left" };
            }
          });
        });
      });
    });
    if (!wbx.worksheets.length) { alert("No hay proveedores que exportar."); return; }
    const buf = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Proveedores_OSB-CTM-ISE_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
    <Panel
      title={`Catálogo de proveedores — ${unidad}`}
      subtitle={`${proveedoresUnidad.length} registrados — cada compañía tiene el suyo. Un proveedor puede tener varias cuentas bancarias`}
      right={
        <div style={{ display: "flex", gap: 8 }}>
          <Button variant="ghost" onClick={exportarProveedores} title="Las tres compañías con sus cuentas bancarias, para el preparador de transacciones">Exportar las 3</Button>
          <Button onClick={openNew}>+ Nuevo proveedor</Button>
        </div>
      }
    >
      <div style={{ marginBottom: 14 }}>
        <Field label="Buscar">
          <TextInput value={buscar} onChange={(e) => setBuscar(e.target.value)} placeholder="Nombre, RFC, Id SAE…" style={{ width: 280 }} />
        </Field>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={tableStyle}>
          <thead>
            <tr>{["Nombre","RFC","Id SAE","Referencia","Notas","Cuentas",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
          </thead>
          <tbody>
            {filtrados.map((p) => (
              <tr key={p.id}>
                <td style={tdStyle}>{p.nombre}</td>
                <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{p.rfc || "—"}</td>
                <td style={{ ...tdStyle, fontFamily: T.fontMono, color: T.textDim }}>{p.id_sae || "—"}</td>
                <td style={tdStyle}>{p.referencia || "—"}</td>
                <td style={{ ...tdStyle, color: T.textDim }}>{p.notas || "—"}</td>
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
              <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin proveedores aún</td></tr>
            )}
            {proveedoresUnidad.length > 0 && !filtrados.length && (
              <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ningún proveedor coincide con la búsqueda</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {modalOpen && (
        <Modal title={editId ? "Editar proveedor" : "Nuevo proveedor"} subtitle={`Catálogo de ${unidad}`} onClose={closeModal} width={820}>
          <AutoriaCaption record={form} perfilesApi={perfilesApi} />
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
            <Field label="Referencia">
              <TextInput value={form.referencia} onChange={(e) => setForm({ ...form, referencia: e.target.value })} />
            </Field>
            <Field label="Notas" style={{ gridColumn: "span 3" }}>
              <TextInput value={form.notas} onChange={(e) => setForm({ ...form, notas: e.target.value })} />
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
                    <tr>{["Banco","Sucursal","SWIFT","CLABE","No. Cuenta","Divisa",""].map((h) => <th key={h} style={thStyle}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {cuentasDelProveedor.map((c) => (
                      <tr key={c.id}>
                        <td style={tdStyle}>{c.banco || "—"}</td>
                        <td style={tdStyle}>{c.sucursal || "—"}</td>
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.swift || "—"}</td>
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.clabe || "—"}</td>
                        <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{c.numero_cuenta || "—"}</td>
                        <td style={tdStyle}><Pill>{c.divisa || "MXP"}</Pill></td>
                        <td style={tdStyle}><Button variant="danger" onClick={() => eliminarCuenta(c.id)}>Eliminar</Button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 8 }}>
                <Field label="Banco">
                  <TextInput value={nuevaCuenta.banco} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, banco: e.target.value })} />
                </Field>
                <Field label="Sucursal">
                  <TextInput value={nuevaCuenta.sucursal} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, sucursal: e.target.value })} />
                </Field>
                <Field label="SWIFT">
                  <TextInput value={nuevaCuenta.swift} onChange={(e) => setNuevaCuenta({ ...nuevaCuenta, swift: e.target.value })} />
                </Field>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, alignItems: "end" }}>
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
// Sesión de Supabase Auth — undefined mientras carga, null si no hay sesión.
// `recovery` se activa cuando la sesión viene de un link de "olvidé mi
// contraseña" — en ese caso hay que pedir la contraseña nueva antes de
// dejar entrar a la app normal.
function useAuth() {
  const [session, setSession] = useState(undefined);
  const [recovery, setRecovery] = useState(false);
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((event, s) => {
      setSession(s);
      if (event === "PASSWORD_RECOVERY") setRecovery(true);
    });
    return () => listener.subscription.unsubscribe();
  }, []);
  return { session, recovery, clearRecovery: () => setRecovery(false) };
}

function SetNewPasswordScreen({ onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    if (password.length < 6) { setError("La contraseña debe tener al menos 6 caracteres."); return; }
    if (password !== confirm) { setError("Las contraseñas no coinciden."); return; }
    setLoading(true);
    try {
      const { error: err } = await supabase.auth.updateUser({ password });
      if (err) { setError(err.message); return; }
      onDone();
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: T.bg, fontFamily: T.fontUI }}>
      <form onSubmit={submit} style={{ ...panelStyle, width: 340 }}>
        <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>
          Control de presupuestos
        </div>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 18 }}>Elige tu nueva contraseña</div>
        <Field label="Nueva contraseña" style={{ marginBottom: 12 }}>
          <TextInput type="password" autoFocus value={password} onChange={(e) => setPassword(e.target.value)} required />
        </Field>
        <Field label="Confirmar contraseña" style={{ marginBottom: 16 }}>
          <TextInput type="password" value={confirm} onChange={(e) => setConfirm(e.target.value)} required />
        </Field>
        {error && <div style={{ fontSize: 12, color: T.red, marginBottom: 12 }}>{error}</div>}
        <Button type="submit" disabled={loading} style={{ width: "100%", justifyContent: "center" }}>
          {loading ? "Guardando…" : "Guardar contraseña"}
        </Button>
      </form>
    </div>
  );
}

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [aviso, setAviso] = useState("");
  const [loading, setLoading] = useState(false);
  const [modoOlvido, setModoOlvido] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setLoading(true);
    try {
      const { error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
      if (err) setError(err.message === "Invalid login credentials" ? "Correo o contraseña incorrectos." : err.message);
    } finally {
      setLoading(false);
    }
  };

  const enviarRecuperacion = async (e) => {
    e.preventDefault();
    setError(""); setAviso(""); setLoading(true);
    try {
      const { error: err } = await supabase.auth.resetPasswordForEmail(email.trim(), { redirectTo: window.location.origin });
      if (err) setError(err.message);
      else setAviso("Si ese correo tiene cuenta, te llegó un link para elegir una contraseña nueva.");
    } finally {
      setLoading(false);
    }
  };

  if (modoOlvido) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: T.bg, fontFamily: T.fontUI }}>
        <form onSubmit={enviarRecuperacion} style={{ ...panelStyle, width: 340 }}>
          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>
            Control de presupuestos
          </div>
          <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 8 }}>Recuperar contraseña</div>
          <div style={{ fontSize: 12, color: T.textFaint, marginBottom: 16 }}>Te mandamos un link a tu correo para elegir una nueva.</div>
          <Field label="Correo" style={{ marginBottom: 16 }}>
            <TextInput type="email" autoFocus value={email} onChange={(e) => setEmail(e.target.value)} required />
          </Field>
          {error && <div style={{ fontSize: 12, color: T.red, marginBottom: 12 }}>{error}</div>}
          {aviso && <div style={{ fontSize: 12, color: T.teal, marginBottom: 12 }}>{aviso}</div>}
          <Button type="submit" disabled={loading} style={{ width: "100%", justifyContent: "center" }}>
            {loading ? "Enviando…" : "Enviar link"}
          </Button>
          <button
            type="button"
            onClick={() => { setModoOlvido(false); setError(""); setAviso(""); }}
            style={{ background: "none", border: "none", color: T.accent, fontSize: 11.5, cursor: "pointer", padding: 0, marginTop: 14 }}
          >
            ← Volver a iniciar sesión
          </button>
        </form>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: T.bg, fontFamily: T.fontUI }}>
      <form onSubmit={submit} style={{ ...panelStyle, width: 340 }}>
        <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>
          Control de presupuestos
        </div>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 18 }}>Iniciar sesión</div>
        <Field label="Correo" style={{ marginBottom: 12 }}>
          <TextInput type="email" autoFocus value={email} onChange={(e) => setEmail(e.target.value)} required />
        </Field>
        <Field label="Contraseña" style={{ marginBottom: 16 }}>
          <TextInput type="password" value={password} onChange={(e) => setPassword(e.target.value)} required />
        </Field>
        {error && <div style={{ fontSize: 12, color: T.red, marginBottom: 12 }}>{error}</div>}
        <Button type="submit" disabled={loading} style={{ width: "100%", justifyContent: "center" }}>
          {loading ? "Entrando…" : "Entrar"}
        </Button>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14 }}>
          <span style={{ fontSize: 11, color: T.textFaint }}>
            ¿No tienes cuenta? Pídele a tu administrador que te dé de alta.
          </span>
        </div>
        <button
          type="button"
          onClick={() => { setModoOlvido(true); setError(""); }}
          style={{ background: "none", border: "none", color: T.accent, fontSize: 11.5, cursor: "pointer", padding: 0, marginTop: 10 }}
        >
          ¿Olvidaste tu contraseña?
        </button>
      </form>
    </div>
  );
}

/* ----------------------------------------------------------------------
   TAB: VEHÍCULOS  (flotilla + mantenimientos + dashboard)
---------------------------------------------------------------------- */

// Devuelve los vehículos ya filtrados por la barra de filtros compartida.
function filtrarVehiculos(vehiculos, f) {
  return vehiculos.filter((v) => {
    if (f.companias.length && !f.companias.includes(v.compania || SIN_DATO)) return false;
    if (f.ubicaciones.length && !f.ubicaciones.includes(v.ubicacion || SIN_DATO)) return false;
    if (f.funcional.length && !f.funcional.includes(v.estatus_funcional || SIN_DATO)) return false;
    if (f.admin.length && !f.admin.includes(v.estatus_administrativo || SIN_DATO)) return false;
    if (f.tipos.length && !f.tipos.includes(v.tipo || SIN_DATO)) return false;
    if (f.proyectos.length) {
      const ps = v.proyectos && v.proyectos.length ? v.proyectos : [SIN_DATO];
      if (!ps.some((p) => f.proyectos.includes(p))) return false;
    }
    return true;
  });
}

// Busca coincidencias en los campos de texto del vehículo (económico, placas,
// VIN, motor, modelo, responsables, notas).
function coincideBusqueda(v, q) {
  if (!q) return true;
  const t = q.toLowerCase();
  const campos = [
    v.no_economico, v.placas, v.vin, v.no_motor, v.fabricante,
    v.tipo, v.subtipo, v.modelo, v.notas, v.compania,
    ...(v.responsables || []), ...(v.proyectos || []),
  ];
  return campos.some((c) => c && String(c).toLowerCase().includes(t));
}

const VEH_COL_ESTATUS = {
  "Funcional": T.teal,
  "En Mtto": T.amber,
  "Mtto Requerido": T.amber,
  "Fuera de Servicio": T.red,
  "No Encontrado": T.red,
};

function tonoFuncional(e) {
  if (e === "Funcional") return "teal";
  if (e === "En Mtto" || e === "Mtto Requerido") return "amber";
  if (e === "Fuera de Servicio" || e === "No Encontrado") return "red";
  return "dim";
}
function tonoAdmin(e) {
  if (e === "Activa") return "teal";
  if (e === "Vendida") return "accent";
  if (e === "Baja") return "dim";
  return "red";
}
function tonoMtto(e) {
  if (e === "Concluido") return "teal";
  if (e === "En proceso") return "amber";
  if (e === "Cancelado") return "dim";
  return "red";
}

// Detecta el mismo VIN / placas / no. de motor en económicos distintos —
// son los registros que hay que revisar a mano antes de confiar en los conteos.
function posiblesDuplicados(vehiculos) {
  const salida = [];
  [["vin", "VIN"], ["placas", "Placas"], ["no_motor", "No. Motor"]].forEach(([campo, etiqueta]) => {
    const mapa = {};
    vehiculos.forEach((v) => {
      const val = (v[campo] || "").trim();
      if (!val) return;
      (mapa[val] = mapa[val] || []).push(v.no_economico);
    });
    Object.entries(mapa).forEach(([valor, ecos]) => {
      if (ecos.length > 1) salida.push({ campo: etiqueta, valor, economicos: ecos.sort() });
    });
  });
  return salida;
}

/* ------------------------------- DASHBOARD ------------------------------- */

function VehiculosDashboard({ vehiculos, mantenimientos, ubicaciones, onVerVehiculo }) {
  const nombreUbic = (cod) => (ubicaciones.find((u) => u.codigo === cod) || {}).nombre || cod || SIN_DATO;

  const activos = vehiculos.filter((v) => v.estatus_administrativo === "Activa");
  const cuenta = (pred) => activos.filter(pred).length;
  const funcionales = cuenta((v) => v.estatus_funcional === "Funcional");
  const enMtto = cuenta((v) => v.estatus_funcional === "En Mtto" || v.estatus_funcional === "Mtto Requerido");
  const fueraServicio = cuenta((v) => v.estatus_funcional === "Fuera de Servicio");
  const noEncontrados = cuenta((v) => v.estatus_funcional === "No Encontrado");
  const disponibilidad = activos.length ? (funcionales / activos.length) * 100 : 0;

  const idsVisibles = new Set(vehiculos.map((v) => v.id));
  const mttosVisibles = mantenimientos.filter((m) => idsVisibles.has(m.vehiculo_id));
  const mttosAbiertos = mttosVisibles.filter((m) => m.estatus === "Solicitado" || m.estatus === "En proceso");

  // Costo capturado de mantenimientos, separado por moneda.
  const costos = sumaPorMoneda(mttosVisibles, "costo", "moneda");

  // Barras apiladas: estatus funcional por compañía.
  const porCompania = useMemo(() => {
    const mapa = {};
    activos.forEach((v) => {
      const c = v.compania || SIN_DATO;
      mapa[c] = mapa[c] || { compania: c };
      const e = v.estatus_funcional || SIN_DATO;
      mapa[c][e] = (mapa[c][e] || 0) + 1;
    });
    return Object.values(mapa).sort((a, b) => a.compania.localeCompare(b.compania));
  }, [vehiculos]);

  const estatusPresentes = useMemo(() => {
    const s = new Set();
    activos.forEach((v) => s.add(v.estatus_funcional || SIN_DATO));
    return [...s].sort();
  }, [vehiculos]);

  const porTipo = useMemo(() => {
    const mapa = {};
    activos.forEach((v) => {
      const t = v.tipo || SIN_DATO;
      mapa[t] = (mapa[t] || 0) + 1;
    });
    return Object.entries(mapa).map(([tipo, n]) => ({ tipo, n })).sort((a, b) => b.n - a.n);
  }, [vehiculos]);

  const porUbicacion = useMemo(() => {
    const mapa = {};
    activos.forEach((v) => {
      const u = v.ubicacion || SIN_DATO;
      mapa[u] = mapa[u] || { ubicacion: u, total: 0, funcional: 0, mtto: 0, fuera: 0 };
      mapa[u].total += 1;
      if (v.estatus_funcional === "Funcional") mapa[u].funcional += 1;
      else if (v.estatus_funcional === "En Mtto" || v.estatus_funcional === "Mtto Requerido") mapa[u].mtto += 1;
      else mapa[u].fuera += 1;
    });
    return Object.values(mapa).sort((a, b) => b.total - a.total);
  }, [vehiculos]);

  // Vehículos con más SMI abiertas — los que están costando más atención.
  const rankingMtto = useMemo(() => {
    const mapa = {};
    mttosAbiertos.forEach((m) => { mapa[m.vehiculo_id] = (mapa[m.vehiculo_id] || 0) + 1; });
    return Object.entries(mapa)
      .map(([id, n]) => ({ vehiculo: vehiculos.find((v) => v.id === id), n }))
      .filter((r) => r.vehiculo)
      .sort((a, b) => b.n - a.n)
      .slice(0, 12);
  }, [mantenimientos, vehiculos]);

  const duplicados = useMemo(() => posiblesDuplicados(vehiculos), [vehiculos]);

  if (!vehiculos.length) {
    return <EmptyState title="Sin vehículos que mostrar" body="Ajusta los filtros de arriba, o carga la flotilla corriendo los scripts de migración en Supabase." />;
  }

  return (
    <div style={{ display: "grid", gap: 18 }}>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "stretch" }}>
        <KpiCard label="Flotilla activa" value={activos.length} />
        <KpiCard label="Funcionales" value={funcionales} accent={T.teal} />
        <KpiCard label="Requieren / en mtto" value={enMtto} accent={T.amber} />
        <KpiCard label="Fuera de servicio" value={fueraServicio} accent={T.red} />
        <KpiCard label="No encontrados" value={noEncontrados} accent={T.red} />
        <KpiCard label="SMI abiertas" value={mttosAbiertos.length} accent={T.blue} />
        <div style={{ ...panelStyle, padding: "10px 20px", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <div style={{ fontSize: 10.5, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em" }}>Disponibilidad</div>
          <Gauge pct={disponibilidad} size={86} />
        </div>
      </div>

      {duplicados.length > 0 && (
        <Panel title="Posibles duplicados" subtitle="Mismo VIN, placas o número de motor en económicos distintos — revisa si son la misma unidad capturada dos veces">
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {duplicados.map((d, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12, color: T.text }}>
                <Pill tone="amber">{d.campo}</Pill>
                <span style={{ fontFamily: T.fontMono, color: T.textDim }}>{d.valor}</span>
                <span style={{ color: T.textFaint }}>→</span>
                {d.economicos.map((e) => (
                  <button
                    key={e}
                    type="button"
                    onClick={() => onVerVehiculo && onVerVehiculo(e)}
                    style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: T.accent, fontFamily: T.fontMono, fontSize: 12, textDecoration: "underline" }}
                  >
                    {e}
                  </button>
                ))}
              </div>
            ))}
          </div>
        </Panel>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))", gap: 18 }}>
        <Panel title="Estatus funcional por compañía" subtitle="Solo unidades activas">
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={porCompania} margin={{ top: 8, right: 8, left: -18, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.borderSoft} vertical={false} />
              <XAxis dataKey="compania" tick={{ fontSize: 11, fill: T.textDim }} axisLine={{ stroke: T.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: T.textDim }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              {estatusPresentes.map((e) => (
                <Bar key={e} dataKey={e} stackId="a" fill={VEH_COL_ESTATUS[e] || T.textFaint} />
              ))}
            </BarChart>
          </ResponsiveContainer>
        </Panel>

        <Panel title="Composición de la flotilla" subtitle="Unidades activas por tipo">
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={porTipo} layout="vertical" margin={{ top: 8, right: 24, left: 30, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.borderSoft} horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11, fill: T.textDim }} axisLine={{ stroke: T.border }} tickLine={false} allowDecimals={false} />
              <YAxis type="category" dataKey="tipo" tick={{ fontSize: 11, fill: T.textDim }} axisLine={false} tickLine={false} width={90} />
              <Tooltip content={<ChartTooltip />} />
              <Bar dataKey="n" name="Unidades" fill={T.accent} radius={[0, 3, 3, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Panel>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))", gap: 18 }}>
        <Panel title="Distribución por ubicación" subtitle="Dónde está parada la flotilla activa">
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead>
              <tr>
                <th style={thStyle}>Ubicación</th>
                <th style={{ ...thStyle, textAlign: "right" }}>Total</th>
                <th style={{ ...thStyle, textAlign: "right" }}>Funcional</th>
                <th style={{ ...thStyle, textAlign: "right" }}>En mtto</th>
                <th style={{ ...thStyle, textAlign: "right" }}>Otros</th>
              </tr>
            </thead>
            <tbody>
              {porUbicacion.map((u) => (
                <tr key={u.ubicacion}>
                  <td style={tdStyle}>{nombreUbic(u.ubicacion)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono, fontWeight: 600 }}>{u.total}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono, color: T.teal }}>{u.funcional || "—"}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono, color: T.amber }}>{u.mtto || "—"}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono, color: T.textFaint }}>{u.fuera || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Panel>

        <Panel
          title="Unidades con más SMI abiertas"
          subtitle="Solicitadas o en proceso"
          right={costos.length > 0 && (
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 10, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.06em" }}>Costo capturado</div>
              {costos.map((c, i) => (
                <div key={i} style={{ fontSize: 12.5, fontFamily: T.fontMono, color: T.text }}>{c}</div>
              ))}
            </div>
          )}
        >
          {rankingMtto.length === 0 ? (
            <div style={{ fontSize: 12.5, color: T.textFaint, padding: "18px 0", textAlign: "center" }}>
              No hay SMI abiertas con los filtros actuales.
            </div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr>
                  <th style={thStyle}>Económico</th>
                  <th style={thStyle}>Compañía</th>
                  <th style={thStyle}>Estatus</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>SMI</th>
                </tr>
              </thead>
              <tbody>
                {rankingMtto.map(({ vehiculo, n }) => (
                  <tr key={vehiculo.id}>
                    <td style={tdStyle}>
                      <button
                        type="button"
                        onClick={() => onVerVehiculo && onVerVehiculo(vehiculo.no_economico)}
                        style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: T.accent, fontFamily: T.fontMono, fontSize: 12.5, textDecoration: "underline" }}
                      >
                        {vehiculo.no_economico}
                      </button>
                    </td>
                    <td style={tdStyle}>{vehiculo.compania || "—"}</td>
                    <td style={tdStyle}><Pill tone={tonoFuncional(vehiculo.estatus_funcional)}>{vehiculo.estatus_funcional || "—"}</Pill></td>
                    <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono, fontWeight: 600 }}>{n}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </Panel>
      </div>
    </div>
  );
}

/* ------------------------------- FORMULARIO ------------------------------ */

const VEH_FORM_VACIO = {
  no_economico: "", compania: "", proyectos: [], estatus_administrativo: "Activa",
  estatus_funcional: "", estatus_operacional: "", ubicacion: "", placas: "", vin: "",
  no_motor: "", fabricante: "", tipo: "", subtipo: "", modelo: "", anio: "",
  responsables: [], gps: "", notas: "",
};

function VehiculoModal({ inicial, ubicaciones, proyectosOpciones, companiasOpciones, onGuardar, onClose, perfilesApi, registro }) {
  const [f, setF] = useState({ ...VEH_FORM_VACIO, ...(inicial || {}), anio: inicial && inicial.anio != null ? String(inicial.anio) : "" });
  const [guardando, setGuardando] = useState(false);
  const set = (k, v) => setF((prev) => ({ ...prev, [k]: v }));

  const toggleProyecto = (p) =>
    set("proyectos", f.proyectos.includes(p) ? f.proyectos.filter((x) => x !== p) : [...f.proyectos, p]);

  const guardar = async () => {
    if (!f.no_economico.trim()) { alert("El número económico es obligatorio."); return; }
    setGuardando(true);
    try {
      await onGuardar({
        no_economico: f.no_economico.trim(),
        compania: f.compania || null,
        proyectos: f.proyectos,
        estatus_administrativo: f.estatus_administrativo || "Activa",
        estatus_funcional: f.estatus_funcional || null,
        estatus_operacional: f.estatus_operacional || null,
        ubicacion: f.ubicacion || null,
        placas: f.placas.trim() || null,
        vin: f.vin.trim().toUpperCase() || null,
        no_motor: f.no_motor.trim() || null,
        fabricante: f.fabricante.trim() || null,
        tipo: f.tipo || null,
        subtipo: f.subtipo.trim() || null,
        modelo: f.modelo.trim() || null,
        anio: f.anio && /^\d{4}$/.test(f.anio) ? Number(f.anio) : null,
        responsables: f.responsables,
        gps: f.gps || null,
        notas: f.notas.trim() || null,
      });
      onClose();
    } catch (err) {
      alert("No se pudo guardar: " + (err.message || err));
    } finally {
      setGuardando(false);
    }
  };

  return (
    <Modal
      title={inicial && inicial.id ? `Editar ${inicial.no_economico}` : "Nuevo vehículo"}
      subtitle={inicial && inicial.id ? "Los cambios quedan registrados con tu usuario" : "Alta manual de una unidad"}
      onClose={onClose}
      width={860}
    >
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
        <Field label="No. Económico *">
          <TextInput value={f.no_economico} onChange={(e) => set("no_economico", e.target.value)} placeholder="CTM-V001" />
        </Field>
        <Field label="Compañía">
          <Select value={f.compania} onChange={(e) => set("compania", e.target.value)}>
            <option value="">— Sin asignar —</option>
            {companiasOpciones.map((c) => <option key={c} value={c}>{c}</option>)}
          </Select>
        </Field>
        <Field label="Estatus administrativo">
          <Select value={f.estatus_administrativo} onChange={(e) => set("estatus_administrativo", e.target.value)}>
            {VEH_EST_ADMIN.map((e) => <option key={e} value={e}>{e}</option>)}
          </Select>
        </Field>
        <Field label="Estatus funcional">
          <Select value={f.estatus_funcional} onChange={(e) => set("estatus_funcional", e.target.value)}>
            <option value="">— Sin dato —</option>
            {VEH_EST_FUNCIONAL.map((e) => <option key={e} value={e}>{e}</option>)}
          </Select>
        </Field>
        <Field label="Estatus operacional">
          <Select value={f.estatus_operacional} onChange={(e) => set("estatus_operacional", e.target.value)}>
            <option value="">— Sin dato —</option>
            {VEH_EST_OPERACIONAL.map((e) => <option key={e} value={e}>{e}</option>)}
          </Select>
        </Field>
        <Field label="Ubicación">
          <Select value={f.ubicacion} onChange={(e) => set("ubicacion", e.target.value)}>
            <option value="">— Sin dato —</option>
            {ubicaciones.map((u) => <option key={u.codigo} value={u.codigo}>{u.codigo} · {u.nombre}</option>)}
          </Select>
        </Field>
        <Field label="Placas">
          <TextInput value={f.placas} onChange={(e) => set("placas", e.target.value)} />
        </Field>
        <Field label="VIN">
          <TextInput value={f.vin} onChange={(e) => set("vin", e.target.value)} />
        </Field>
        <Field label="No. Motor">
          <TextInput value={f.no_motor} onChange={(e) => set("no_motor", e.target.value)} />
        </Field>
        <Field label="Fabricante">
          <TextInput value={f.fabricante} onChange={(e) => set("fabricante", e.target.value)} placeholder="Toyota" />
        </Field>
        <Field label="Tipo">
          <Select value={f.tipo} onChange={(e) => set("tipo", e.target.value)}>
            <option value="">— Sin dato —</option>
            {VEH_TIPOS.map((t) => <option key={t} value={t}>{t}</option>)}
          </Select>
        </Field>
        <Field label="Subtipo">
          <TextInput value={f.subtipo} onChange={(e) => set("subtipo", e.target.value)} placeholder="Plataforma" />
        </Field>
        <Field label="Modelo">
          <TextInput value={f.modelo} onChange={(e) => set("modelo", e.target.value)} placeholder="Hilux Doble Cabina" />
        </Field>
        <Field label="Año">
          <TextInput value={f.anio} onChange={(e) => set("anio", e.target.value.replace(/\D/g, "").slice(0, 4))} placeholder="2019" />
        </Field>
        <Field label="GPS">
          <Select value={f.gps} onChange={(e) => set("gps", e.target.value)}>
            <option value="">— Sin GPS —</option>
            {VEH_GPS.map((g) => <option key={g} value={g}>{g}</option>)}
          </Select>
        </Field>
      </div>

      <div style={{ marginTop: 16 }}>
        <div style={{ textTransform: "uppercase", letterSpacing: "0.06em", fontSize: 10.5, color: T.textDim, marginBottom: 8 }}>Proyectos</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
          {proyectosOpciones.map((p) => (
            <label key={p} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: T.text, cursor: "pointer" }}>
              <input type="checkbox" checked={f.proyectos.includes(p)} onChange={() => toggleProyecto(p)} />
              {p}
            </label>
          ))}
        </div>
      </div>

      <div style={{ marginTop: 16, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Field label="Responsables (separados por coma)">
          <TextInput
            value={f.responsables.join(", ")}
            onChange={(e) => set("responsables", e.target.value.split(",").map((s) => s.trim()).filter(Boolean))}
            placeholder="Juan Pérez, María López"
          />
        </Field>
        <Field label="Notas generales">
          <TextInput value={f.notas} onChange={(e) => set("notas", e.target.value)} />
        </Field>
      </div>

      {registro && <div style={{ marginTop: 14 }}><AutoriaCaption record={registro} perfilesApi={perfilesApi} /></div>}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <Button variant="ghost" onClick={onClose}>Cancelar</Button>
        <Button onClick={guardar} disabled={guardando}>{guardando ? "Guardando…" : "Guardar"}</Button>
      </div>
    </Modal>
  );
}

function MantenimientoModal({ inicial, vehiculos, onGuardar, onClose, perfilesApi }) {
  const [f, setF] = useState({
    folio: "", vehiculo_id: "", zona: "", anio: "", fecha: "", tipo_mtto: "",
    descripcion: "", estatus: "Solicitado", proveedor: "", costo: "", moneda: "MXP", notas: "",
    ...(inicial || {}),
    anio: inicial && inicial.anio != null ? String(inicial.anio) : "",
    costo: inicial && inicial.costo != null ? String(inicial.costo) : "",
  });
  const [guardando, setGuardando] = useState(false);
  const set = (k, v) => setF((prev) => ({ ...prev, [k]: v }));

  const guardar = async () => {
    if (!f.folio.trim()) { alert("El folio de la SMI es obligatorio."); return; }
    if (!f.vehiculo_id) { alert("Elige a qué vehículo pertenece."); return; }
    setGuardando(true);
    try {
      await onGuardar({
        folio: f.folio.trim().toUpperCase(),
        vehiculo_id: f.vehiculo_id,
        zona: f.zona.trim() || null,
        anio: f.anio && /^\d{4}$/.test(f.anio) ? Number(f.anio) : null,
        fecha: f.fecha || null,
        tipo_mtto: f.tipo_mtto || null,
        descripcion: f.descripcion.trim() || null,
        estatus: f.estatus,
        proveedor: f.proveedor.trim() || null,
        costo: f.costo !== "" && !isNaN(Number(f.costo)) ? Number(f.costo) : null,
        moneda: f.moneda || "MXP",
        notas: f.notas.trim() || null,
      });
      onClose();
    } catch (err) {
      alert("No se pudo guardar: " + (err.message || err));
    } finally {
      setGuardando(false);
    }
  };

  const ordenados = [...vehiculos].sort((a, b) => a.no_economico.localeCompare(b.no_economico));

  return (
    <Modal
      title={inicial && inicial.id ? `Editar ${inicial.folio}` : "Nueva SMI"}
      subtitle="Solicitud de mantenimiento"
      onClose={onClose}
      width={780}
    >
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
        <Field label="Folio *">
          <TextInput value={f.folio} onChange={(e) => set("folio", e.target.value)} placeholder="MTTO-PR-2026-001" />
        </Field>
        <Field label="Vehículo *">
          <Select value={f.vehiculo_id} onChange={(e) => set("vehiculo_id", e.target.value)}>
            <option value="">— Elegir —</option>
            {ordenados.map((v) => (
              <option key={v.id} value={v.id}>{v.no_economico}{v.compania ? ` · ${v.compania}` : ""}</option>
            ))}
          </Select>
        </Field>
        <Field label="Zona">
          <TextInput value={f.zona} onChange={(e) => set("zona", e.target.value)} placeholder="PR-UP" />
        </Field>
        <Field label="Año">
          <TextInput value={f.anio} onChange={(e) => set("anio", e.target.value.replace(/\D/g, "").slice(0, 4))} placeholder="2026" />
        </Field>
        <Field label="Fecha">
          <TextInput type="date" value={f.fecha || ""} onChange={(e) => set("fecha", e.target.value)} />
        </Field>
        <Field label="Tipo">
          <Select value={f.tipo_mtto} onChange={(e) => set("tipo_mtto", e.target.value)}>
            <option value="">— Sin dato —</option>
            {MTTO_TIPOS.map((t) => <option key={t} value={t}>{t}</option>)}
          </Select>
        </Field>
        <Field label="Estatus">
          <Select value={f.estatus} onChange={(e) => set("estatus", e.target.value)}>
            {MTTO_ESTATUS.map((e) => <option key={e} value={e}>{e}</option>)}
          </Select>
        </Field>
        <Field label="Proveedor / taller">
          <TextInput value={f.proveedor} onChange={(e) => set("proveedor", e.target.value)} />
        </Field>
        <Field label="Costo">
          <TextInput value={f.costo} onChange={(e) => set("costo", e.target.value)} placeholder="0.00" />
        </Field>
        <Field label="Moneda">
          <Select value={f.moneda} onChange={(e) => set("moneda", e.target.value)}>
            {MONEDAS.map((m) => <option key={m} value={m}>{m}</option>)}
          </Select>
        </Field>
      </div>

      <div style={{ marginTop: 14, display: "grid", gap: 14 }}>
        <Field label="Descripción del trabajo">
          <TextInput value={f.descripcion} onChange={(e) => set("descripcion", e.target.value)} placeholder="Cambio de balatas y rectificado de discos" />
        </Field>
        <Field label="Notas">
          <TextInput value={f.notas} onChange={(e) => set("notas", e.target.value)} />
        </Field>
      </div>

      {inicial && inicial.id && <div style={{ marginTop: 14 }}><AutoriaCaption record={inicial} perfilesApi={perfilesApi} /></div>}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <Button variant="ghost" onClick={onClose}>Cancelar</Button>
        <Button onClick={guardar} disabled={guardando}>{guardando ? "Guardando…" : "Guardar"}</Button>
      </div>
    </Modal>
  );
}

/* --------------------------------- FLOTILLA ------------------------------- */

const COLUMNAS_VEHICULOS = [
  { key: "no_economico", label: "No. Económico" },
  { key: "compania", label: "Compañía" },
  { key: "proyectos", label: "Proyectos" },
  { key: "estatus_funcional", label: "Est. Funcional" },
  { key: "estatus_operacional", label: "Est. Operacional" },
  { key: "estatus_administrativo", label: "Est. Administrativo" },
  { key: "ubicacion", label: "Ubicación" },
  { key: "tipo", label: "Tipo" },
  { key: "subtipo", label: "Subtipo" },
  { key: "fabricante", label: "Fabricante" },
  { key: "modelo", label: "Modelo" },
  { key: "anio", label: "Año" },
  { key: "placas", label: "Placas" },
  { key: "vin", label: "VIN" },
  { key: "no_motor", label: "No. Motor" },
  { key: "gps", label: "GPS" },
  { key: "responsables", label: "Responsables" },
  { key: "smi", label: "SMI abiertas" },
  { key: "notas", label: "Notas" },
];

function FlotillaPanel({ vehiculos, vehiculosApi, mantenimientos, ubicaciones, proyectosOpciones, companiasOpciones, perfilesApi, expandido, setExpandido }) {
  const [busqueda, setBusqueda] = useSessionState("veh:busqueda", "");
  const [sort, setSort] = usePrefState("pref-veh-sort", { key: "no_economico", dir: "asc" }, sanearSort(["no_economico","compania","estatus_funcional","estatus_operacional","estatus_administrativo","ubicacion","placas","vin","marca","modelo","anio","tipo"]));
  const [editando, setEditando] = useState(null); // objeto vehículo, o {} para nuevo
  const { visible, hidden, toggle, showAll } = useColumnVisibility("veh:columnas", COLUMNAS_VEHICULOS);

  const nombreUbic = (cod) => (ubicaciones.find((u) => u.codigo === cod) || {}).nombre || cod;

  const smiPorVehiculo = useMemo(() => {
    const mapa = {};
    mantenimientos.forEach((m) => { (mapa[m.vehiculo_id] = mapa[m.vehiculo_id] || []).push(m); });
    return mapa;
  }, [mantenimientos]);

  const filas = useMemo(() => {
    const buscados = vehiculos.filter((v) => coincideBusqueda(v, busqueda));
    return sortRows(buscados, sort, {
      proyectos: (v) => (v.proyectos || []).join(", "),
      responsables: (v) => (v.responsables || []).join(", "),
      ubicacion: (v) => nombreUbic(v.ubicacion) || "",
      smi: (v) => (smiPorVehiculo[v.id] || []).filter((m) => m.estatus === "Solicitado" || m.estatus === "En proceso").length,
      anio: (v) => v.anio || 0,
    });
  }, [vehiculos, busqueda, sort, smiPorVehiculo, ubicaciones]);

  const guardar = async (datos) => {
    if (editando && editando.id) await vehiculosApi.update(editando.id, datos);
    else await vehiculosApi.insert({ id: uid(), ...datos });
  };

  const eliminar = (v) => {
    if (!confirm(`¿Eliminar ${v.no_economico}? También se borran sus SMI asociadas.`)) return;
    vehiculosApi.remove(v.id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  const celda = (v, key) => {
    switch (key) {
      case "no_economico":
        return <span style={{ fontFamily: T.fontMono, fontWeight: 600 }}>{v.no_economico}</span>;
      case "proyectos":
        return (v.proyectos || []).length
          ? <span style={{ display: "inline-flex", gap: 4, flexWrap: "wrap" }}>{v.proyectos.map((p) => <Pill key={p} tone="dim">{p}</Pill>)}</span>
          : <span style={{ color: T.textFaint }}>—</span>;
      case "estatus_funcional":
        return v.estatus_funcional ? <Pill tone={tonoFuncional(v.estatus_funcional)}>{v.estatus_funcional}</Pill> : <span style={{ color: T.textFaint }}>—</span>;
      case "estatus_administrativo":
        return <Pill tone={tonoAdmin(v.estatus_administrativo)}>{v.estatus_administrativo}</Pill>;
      case "ubicacion":
        return v.ubicacion ? nombreUbic(v.ubicacion) : <span style={{ color: T.textFaint }}>—</span>;
      case "responsables":
        return (v.responsables || []).length ? v.responsables.join(", ") : <span style={{ color: T.textFaint }}>—</span>;
      case "smi": {
        const abiertas = (smiPorVehiculo[v.id] || []).filter((m) => m.estatus === "Solicitado" || m.estatus === "En proceso").length;
        return abiertas ? <Pill tone="amber">{abiertas}</Pill> : <span style={{ color: T.textFaint }}>—</span>;
      }
      case "vin":
      case "no_motor":
      case "placas":
        return v[key] ? <span style={{ fontFamily: T.fontMono, fontSize: 11.5 }}>{v[key]}</span> : <span style={{ color: T.textFaint }}>—</span>;
      default:
        return v[key] != null && v[key] !== "" ? v[key] : <span style={{ color: T.textFaint }}>—</span>;
    }
  };

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <TextInput
          value={busqueda}
          onChange={(e) => setBusqueda(e.target.value)}
          placeholder="Buscar económico, placas, VIN, modelo, responsable…"
          style={{ width: 340 }}
        />
        <span style={{ fontSize: 12, color: T.textDim }}>
          {filas.length} {filas.length === 1 ? "unidad" : "unidades"}
        </span>
        <div style={{ flex: 1 }} />
        <ColumnVisibilityControl columns={COLUMNAS_VEHICULOS} hidden={hidden} onToggle={toggle} onShowAll={showAll} />
        <Button onClick={() => setEditando({})}>Nuevo vehículo</Button>
      </div>

      {filas.length === 0 ? (
        <EmptyState title="Sin resultados" body="Ningún vehículo coincide con los filtros y la búsqueda actuales." />
      ) : (
        <div style={{ ...panelStyle, padding: 0, overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead>
              <tr>
                <th style={{ ...thStyle, width: 34 }}></th>
                <th style={{ ...thStyle, width: 40 }}>#</th>
                {visible.map((c) => (
                  <SortableTh key={c.key} label={c.label} sortKey={c.key} sort={sort} setSort={setSort} />
                ))}
                <th style={{ ...thStyle, width: 76 }}>Acciones</th>
              </tr>
            </thead>
            <tbody>
              {filas.map((v, i) => {
                const smis = (smiPorVehiculo[v.id] || []).slice().sort((a, b) => (a.folio || "").localeCompare(b.folio || ""));
                const abierto = expandido === v.id;
                return (
                  <React.Fragment key={v.id}>
                    <tr>
                      <td style={{ ...tdStyle, textAlign: "center" }}>
                        {smis.length > 0 && (
                          <button
                            type="button"
                            onClick={() => setExpandido(abierto ? null : v.id)}
                            style={{ background: "none", border: "none", cursor: "pointer", color: T.textDim, fontSize: 11, padding: 0 }}
                            title={abierto ? "Ocultar SMI" : "Ver SMI"}
                          >
                            {abierto ? "▾" : "▸"}
                          </button>
                        )}
                      </td>
                      <td style={{ ...tdStyle, color: T.textFaint, fontFamily: T.fontMono, fontSize: 11 }}>{i + 1}</td>
                      {visible.map((c) => <td key={c.key} style={tdStyle}>{celda(v, c.key)}</td>)}
                      <td style={tdStyle}>
                        <div style={{ display: "flex", gap: 5 }}>
                          <IconButton icon="✎" label="Editar" onClick={() => setEditando(v)} />
                          <IconButton icon="✕" label="Eliminar" tone={T.red} onClick={() => eliminar(v)} />
                        </div>
                      </td>
                    </tr>
                    {abierto && (
                      <tr>
                        <td colSpan={visible.length + 3} style={{ padding: "10px 18px 16px 52px", background: T.panelAlt, borderBottom: `1px solid ${T.borderSoft}` }}>
                          <div style={{ fontSize: 10.5, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>
                            SMI de mantenimiento · {smis.length}
                          </div>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                            <thead>
                              <tr>
                                <th style={thStyle}>Folio</th>
                                <th style={thStyle}>Estatus</th>
                                <th style={thStyle}>Tipo</th>
                                <th style={thStyle}>Fecha</th>
                                <th style={thStyle}>Descripción</th>
                                <th style={{ ...thStyle, textAlign: "right" }}>Costo</th>
                              </tr>
                            </thead>
                            <tbody>
                              {smis.map((m) => (
                                <tr key={m.id}>
                                  <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{m.folio}</td>
                                  <td style={tdStyle}><Pill tone={tonoMtto(m.estatus)}>{m.estatus}</Pill></td>
                                  <td style={tdStyle}>{m.tipo_mtto || <span style={{ color: T.textFaint }}>—</span>}</td>
                                  <td style={tdStyle}>{m.fecha || <span style={{ color: T.textFaint }}>—</span>}</td>
                                  <td style={tdStyle}>{m.descripcion || <span style={{ color: T.textFaint }}>Pendiente de capturar</span>}</td>
                                  <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono }}>
                                    {m.costo != null ? money(m.costo, m.moneda || "MXP") : <span style={{ color: T.textFaint }}>—</span>}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {editando && (
        <VehiculoModal
          inicial={editando.id ? editando : null}
          registro={editando.id ? editando : null}
          ubicaciones={ubicaciones}
          proyectosOpciones={proyectosOpciones}
          companiasOpciones={companiasOpciones}
          perfilesApi={perfilesApi}
          onGuardar={guardar}
          onClose={() => setEditando(null)}
        />
      )}
    </div>
  );
}

/* ------------------------------ MANTENIMIENTOS ---------------------------- */

function MantenimientosPanel({ vehiculos, mantenimientos, mantenimientosApi, perfilesApi }) {
  const [estatusFiltro, setEstatusFiltro] = useSessionSetState("veh:mtto:estatus", []);
  const [aniosFiltro, setAniosFiltro] = useSessionSetState("veh:mtto:anios", []);
  const [sort, setSort] = usePrefState("pref-veh-mtto-sort", { key: "folio", dir: "asc" }, sanearSort(["folio","vehiculo","zona","anio","fecha","tipo_mtto","estatus","taller","costo"]));
  const [editando, setEditando] = useState(null);

  const porId = useMemo(() => {
    const m = {};
    vehiculos.forEach((v) => { m[v.id] = v; });
    return m;
  }, [vehiculos]);

  const aniosDisponibles = useMemo(
    () => [...new Set(mantenimientos.map((m) => m.anio).filter(Boolean))].sort((a, b) => b - a).map(String),
    [mantenimientos]
  );

  const filas = useMemo(() => {
    const f = mantenimientos.filter((m) => {
      if (estatusFiltro.size && !estatusFiltro.has(m.estatus)) return false;
      if (aniosFiltro.size && !aniosFiltro.has(String(m.anio))) return false;
      return true;
    });
    return sortRows(f, sort, {
      vehiculo: (m) => (porId[m.vehiculo_id] || {}).no_economico || "",
      compania: (m) => (porId[m.vehiculo_id] || {}).compania || "",
      costo: (m) => Number(m.costo) || 0,
    });
  }, [mantenimientos, estatusFiltro, aniosFiltro, sort, porId]);

  const totales = sumaPorMoneda(filas, "costo", "moneda");

  const guardar = async (datos) => {
    if (editando && editando.id) await mantenimientosApi.update(editando.id, datos);
    else await mantenimientosApi.insert({ id: uid(), ...datos });
  };

  const eliminar = (m) => {
    if (!confirm(`¿Eliminar la SMI ${m.folio}?`)) return;
    mantenimientosApi.remove(m.id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <MultiSelect
          opciones={MTTO_ESTATUS}
          seleccionados={[...estatusFiltro]}
          onChange={(vals) => setEstatusFiltro(new Set(vals))}
          todosLabel="Todos los estatus"
          unidadLabel=" estatus"
        />
        <MultiSelect
          opciones={aniosDisponibles}
          seleccionados={[...aniosFiltro]}
          onChange={(vals) => setAniosFiltro(new Set(vals))}
          todosLabel="Todos los años"
          unidadLabel=" años"
        />
        <span style={{ fontSize: 12, color: T.textDim }}>{filas.length} SMI</span>
        {totales.length > 0 && (
          <span style={{ fontSize: 12, color: T.teal, fontFamily: T.fontMono }}>Σ {totales.join(" · ")}</span>
        )}
        <div style={{ flex: 1 }} />
        <Button onClick={() => setEditando({})}>Nueva SMI</Button>
      </div>

      {filas.length === 0 ? (
        <EmptyState title="Sin SMI" body="No hay solicitudes de mantenimiento con los filtros actuales." />
      ) : (
        <div style={{ ...panelStyle, padding: 0, overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead>
              <tr>
                <th style={{ ...thStyle, width: 40 }}>#</th>
                <SortableTh label="Folio" sortKey="folio" sort={sort} setSort={setSort} />
                <SortableTh label="Vehículo" sortKey="vehiculo" sort={sort} setSort={setSort} />
                <SortableTh label="Compañía" sortKey="compania" sort={sort} setSort={setSort} />
                <SortableTh label="Zona" sortKey="zona" sort={sort} setSort={setSort} />
                <SortableTh label="Año" sortKey="anio" sort={sort} setSort={setSort} />
                <SortableTh label="Fecha" sortKey="fecha" sort={sort} setSort={setSort} />
                <SortableTh label="Tipo" sortKey="tipo_mtto" sort={sort} setSort={setSort} />
                <SortableTh label="Estatus" sortKey="estatus" sort={sort} setSort={setSort} />
                <SortableTh label="Descripción" sortKey="descripcion" sort={sort} setSort={setSort} />
                <SortableTh label="Proveedor" sortKey="proveedor" sort={sort} setSort={setSort} />
                <SortableTh label="Costo" sortKey="costo" sort={sort} setSort={setSort} sumLabel={totales} />
                <th style={{ ...thStyle, width: 76 }}>Acciones</th>
              </tr>
            </thead>
            <tbody>
              {filas.map((m, i) => {
                const v = porId[m.vehiculo_id];
                return (
                  <tr key={m.id}>
                    <td style={{ ...tdStyle, color: T.textFaint, fontFamily: T.fontMono, fontSize: 11 }}>{i + 1}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono, fontWeight: 600 }}>{m.folio}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{v ? v.no_economico : <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={tdStyle}>{(v && v.compania) || <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={tdStyle}>{m.zona || <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{m.anio || <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={tdStyle}>{m.fecha || <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={tdStyle}>{m.tipo_mtto || <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={tdStyle}><Pill tone={tonoMtto(m.estatus)}>{m.estatus}</Pill></td>
                    <td style={tdStyle}>{m.descripcion || <span style={{ color: T.textFaint }}>Pendiente de capturar</span>}</td>
                    <td style={tdStyle}>{m.proveedor || <span style={{ color: T.textFaint }}>—</span>}</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontFamily: T.fontMono }}>
                      {m.costo != null ? money(m.costo, m.moneda || "MXP") : <span style={{ color: T.textFaint }}>—</span>}
                    </td>
                    <td style={tdStyle}>
                      <div style={{ display: "flex", gap: 5 }}>
                        <IconButton icon="✎" label="Editar" onClick={() => setEditando(m)} />
                        <IconButton icon="✕" label="Eliminar" tone={T.red} onClick={() => eliminar(m)} />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {editando && (
        <MantenimientoModal
          inicial={editando.id ? editando : null}
          vehiculos={vehiculos}
          perfilesApi={perfilesApi}
          onGuardar={guardar}
          onClose={() => setEditando(null)}
        />
      )}
    </div>
  );
}

/* --------------------------------- TAB RAÍZ ------------------------------- */

function VehiculosTab({ vehiculos, vehiculosApi, mantenimientos, mantenimientosApi, ubicaciones, perfilesApi, unidadesPermitidas }) {
  const [sub, setSub] = useSessionState("veh:sub", "dashboard");
  const [companias, setCompanias] = useSessionSetState("veh:f:companias", []);
  const [ubics, setUbics] = useSessionSetState("veh:f:ubicaciones", []);
  const [funcional, setFuncional] = useSessionSetState("veh:f:funcional", []);
  const [admin, setAdmin] = useSessionSetState("veh:f:admin", ["Activa"]);
  const [tipos, setTipos] = useSessionSetState("veh:f:tipos", []);
  const [proyectos, setProyectos] = useSessionSetState("veh:f:proyectos", []);
  const [expandido, setExpandido] = useState(null);

  // Solo las compañías que este usuario tiene permitido ver. Si no tiene
  // restricción, ve las 5 (incluidas IZ2 y JEF, que no existen en el selector
  // global de unidad de negocio pero sí tienen vehículos).
  const companiasOpciones = useMemo(() => {
    const permitidas = unidadesPermitidas && unidadesPermitidas.length ? unidadesPermitidas : null;
    return VEH_COMPANIAS.filter((c) => !permitidas || permitidas.includes(c));
  }, [unidadesPermitidas]);

  const proyectosOpciones = useMemo(() => {
    const s = new Set();
    vehiculos.forEach((v) => (v.proyectos || []).forEach((p) => s.add(p)));
    return [...s].sort();
  }, [vehiculos]);

  const ubicacionesOpciones = useMemo(
    () => ubicaciones.slice().sort((a, b) => a.nombre.localeCompare(b.nombre)),
    [ubicaciones]
  );

  const filtrados = useMemo(
    () => filtrarVehiculos(vehiculos, {
      companias: [...companias], ubicaciones: [...ubics], funcional: [...funcional],
      admin: [...admin], tipos: [...tipos], proyectos: [...proyectos],
    }),
    [vehiculos, companias, ubics, funcional, admin, tipos, proyectos]
  );

  const idsFiltrados = useMemo(() => new Set(filtrados.map((v) => v.id)), [filtrados]);
  const mttosFiltrados = useMemo(
    () => mantenimientos.filter((m) => idsFiltrados.has(m.vehiculo_id)),
    [mantenimientos, idsFiltrados]
  );

  // Desde el Dashboard se puede saltar a una unidad concreta: cambia a Flotilla
  // y la deja expandida.
  const verVehiculo = (economico) => {
    const v = vehiculos.find((x) => x.no_economico === economico);
    setSub("flotilla");
    if (v) setExpandido(v.id);
  };

  const SUBS = [
    { id: "dashboard", label: "Dashboard" },
    { id: "flotilla", label: "Flotilla" },
    { id: "mantenimientos", label: "Mantenimientos" },
  ];

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ display: "flex", background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8, padding: 3 }}>
          {SUBS.map((s) => (
            <button
              key={s.id}
              onClick={() => setSub(s.id)}
              style={{
                padding: "7px 16px", borderRadius: 6, border: "none", cursor: "pointer",
                background: sub === s.id ? T.accent : "transparent",
                color: sub === s.id ? "#FFFFFF" : T.textDim,
                fontWeight: 600, fontSize: 12.5, fontFamily: T.fontUI,
              }}
            >
              {s.label}
            </button>
          ))}
        </div>

        <div style={{ flex: 1 }} />

        <MultiSelect opciones={companiasOpciones} seleccionados={[...companias]} onChange={(v) => setCompanias(new Set(v))} todosLabel="Todas las compañías" unidadLabel=" compañías" />
        <MultiSelect opciones={proyectosOpciones} seleccionados={[...proyectos]} onChange={(v) => setProyectos(new Set(v))} todosLabel="Todos los proyectos" unidadLabel=" proyectos" />
        <MultiSelect opciones={ubicacionesOpciones.map((u) => u.codigo)} seleccionados={[...ubics]} onChange={(v) => setUbics(new Set(v))} todosLabel="Todas las ubicaciones" unidadLabel=" ubic." />
        <MultiSelect opciones={VEH_TIPOS} seleccionados={[...tipos]} onChange={(v) => setTipos(new Set(v))} todosLabel="Todos los tipos" unidadLabel=" tipos" />
        <MultiSelect opciones={VEH_EST_FUNCIONAL} seleccionados={[...funcional]} onChange={(v) => setFuncional(new Set(v))} todosLabel="Todo est. funcional" unidadLabel=" estatus" />
        <MultiSelect opciones={VEH_EST_ADMIN} seleccionados={[...admin]} onChange={(v) => setAdmin(new Set(v))} todosLabel="Todo est. admin." unidadLabel=" estatus" />
      </div>

      {sub === "dashboard" && (
        <VehiculosDashboard
          vehiculos={filtrados}
          mantenimientos={mttosFiltrados}
          ubicaciones={ubicaciones}
          onVerVehiculo={verVehiculo}
        />
      )}
      {sub === "flotilla" && (
        <FlotillaPanel
          vehiculos={filtrados}
          vehiculosApi={vehiculosApi}
          mantenimientos={mttosFiltrados}
          ubicaciones={ubicaciones}
          proyectosOpciones={proyectosOpciones}
          companiasOpciones={companiasOpciones}
          perfilesApi={perfilesApi}
          expandido={expandido}
          setExpandido={setExpandido}
        />
      )}
      {sub === "mantenimientos" && (
        <MantenimientosPanel
          vehiculos={filtrados}
          mantenimientos={mttosFiltrados}
          mantenimientosApi={mantenimientosApi}
          perfilesApi={perfilesApi}
        />
      )}
    </div>
  );
}

export default function App() {
  const { session, recovery, clearRecovery } = useAuth();
  const proyectosApi = useCollection("proyectos");
  const zonasApi = useCollection("zonas");
  const rubrosApi = useCollection("rubros");
  const categoriasApi = useCollection("categorias");
  // Al cambiar el catálogo se reconstruye la variable de módulo y se fuerza
  // un render, para que los selectores muestren lo nuevo sin recargar.
  const [catalogoVersion, setCatalogoVersion] = useState(0);
  useEffect(() => {
    aplicarCatalogo(rubrosApi.rows, categoriasApi.rows);
    setCatalogoVersion((v) => v + 1);
  }, [rubrosApi.rows, categoriasApi.rows]);
  const partidasApi = useCollection("partidas", "created_at", { withAudit: true });
  const transaccionesApi = useCollection("transacciones", "created_at", { withAudit: true });
  const proveedoresApi = useCollection("proveedores", "created_at", { withAudit: true });
  const cuentasApi = useCollection("proveedor_cuentas", "created_at", { withAudit: true });
  const perfilesApi = useCollection("perfiles");
  const notasApi = useCollection("transaccion_notas");
  const vehiculosApi = useCollection("vehiculos", "no_economico", { withAudit: true });
  const vehMantenimientosApi = useCollection("vehiculo_mantenimientos", "folio", { withAudit: true });
  const vehUbicacionesApi = useCollection("vehiculo_ubicaciones", "nombre");
  const [unidad, setUnidad] = useState("CTM");
  const [tab, setTab] = useState("dashboard");

  const miPerfil = session ? perfilesApi.rows.find((p) => p.id === session.user.id) : null;
  const unidadesPermitidas = (miPerfil?.unidades_permitidas && miPerfil.unidades_permitidas.length)
    ? miPerfil.unidades_permitidas
    : UNIDAD_KEYS;

  // Si la unidad activa no está entre las permitidas (primera carga, o a alguien
  // le acaban de restringir el acceso), se cambia sola a la primera que sí puede ver.
  useEffect(() => {
    if (unidadesPermitidas.length && !unidadesPermitidas.includes(unidad)) {
      setUnidad(unidadesPermitidas[0]);
    }
  }, [unidadesPermitidas.join(","), unidad]);

  const unidades = useMemo(() => {
    const map = {};
    UNIDADES_BASE.filter((u) => unidadesPermitidas.includes(u)).forEach((u) => { map[u] = { proyectos: [] }; });
    proyectosApi.rows.forEach((p) => {
      if (!map[p.unidad]) return; // fuera del alcance permitido, se ignora
      map[p.unidad].proyectos.push(p);
    });
    return map;
  }, [proyectosApi.rows, unidadesPermitidas.join(",")]);

  const partidas = partidasApi.rows;
  // Solo las activas, en el orden del catálogo. Si la tabla no responde
  // —migración pendiente— se cae al respaldo en vez de quedarse en blanco.
  const zonas = zonasApi.rows.length
    ? zonasApi.rows.filter((z) => z.activa !== false)
        .sort((a, b) => (a.orden ?? 999) - (b.orden ?? 999) || String(a.nombre).localeCompare(String(b.nombre)))
        .map((z) => z.nombre)
    : ZONAS_RESPALDO;
  const transacciones = transaccionesApi.rows;
  const ready = proyectosApi.ready && partidasApi.ready && transaccionesApi.ready;

  const TABS = [
    { id: "dashboard", label: "Dashboard" },
    { id: "partidas", label: "Partidas" },
    { id: "transacciones", label: "Transacciones" },
    { id: "reporte", label: "Reporte de Pagos" },
    { id: "reporte-direccion", label: "Reporte Pagos Dirección" },
    { id: "vehiculos", label: "Vehículos" },
    { id: "catalogo", label: "Catálogo" },
  ];

  if (session === undefined) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: T.bg, fontFamily: T.fontUI, color: T.textFaint, fontSize: 13 }}>
        Cargando…
      </div>
    );
  }
  if (recovery) {
    return <SetNewPasswordScreen onDone={clearRecovery} />;
  }
  if (!session) {
    return <LoginScreen />;
  }

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
        table tbody tr:nth-child(even) > td { background-color: ${T.bg}; }
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
        <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
          <span style={{ fontSize: 11.5, color: T.textFaint }}>
            {perfilesApi.rows.find((p) => p.id === session.user.id)?.nombre || session.user.email}
          </span>
          <Button variant="ghost" onClick={() => supabase.auth.signOut()}>Cerrar sesión</Button>
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
          {tab === "partidas" && <PartidasTab zonas={zonas} unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} perfilesApi={perfilesApi} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
          {tab === "transacciones" && <TransaccionesTab zonas={zonas} unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} notasApi={notasApi} session={session} />}
          {tab === "reporte" && <ReportePagosTab unidad={unidad} partidas={partidas} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
          {tab === "reporte-direccion" && <ReportePagosDireccionTab unidad={unidad} partidas={partidas} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} />}
          {tab === "vehiculos" && (
            <VehiculosTab
              vehiculos={vehiculosApi.rows}
              vehiculosApi={vehiculosApi}
              mantenimientos={vehMantenimientosApi.rows}
              mantenimientosApi={vehMantenimientosApi}
              ubicaciones={vehUbicacionesApi.rows}
              perfilesApi={perfilesApi}
              unidadesPermitidas={miPerfil?.unidades_permitidas || []}
            />
          )}
          {tab === "catalogo" && <CatalogoTab key={catalogoVersion} unidad={unidad} unidades={unidades} proyectosApi={proyectosApi} zonasApi={zonasApi} rubrosApi={rubrosApi} categoriasApi={categoriasApi} partidas={partidas} transacciones={transacciones} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} />}
        </>
      )}
    </div>
  );
}