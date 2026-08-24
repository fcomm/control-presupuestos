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
const APP_VERSION = "1.56.0";
const CHANGELOG = [
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
function PartidaPickerButton({ partidas, transacciones = [], value, onChange, placeholder = "Elegir partida…", allowClear = false, partidasApi, unidad, proyectosOpciones = [] }) {
  const [open, setOpen] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const seleccionada = partidas.find((p) => p.id === value);

  const filtradas = partidas.filter((p) => {
    if (!busqueda.trim()) return true;
    const q = busqueda.trim().toLowerCase();
    return [p.concepto, p.folio, p.proyecto, p.rubro].some((v) => (v || "").toLowerCase().includes(q));
  });
  const meses = MESES.filter((m) => filtradas.some((p) => p.mes === m));
  const sinMes = filtradas.filter((p) => !p.mes);

  const elegir = (id) => {
    setOpen(false); setBusqueda(""); setCreando(false);
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
      const categoriaDefault = RUBROS.find((r) => r.rubro === nuevaPartida.rubro)?.categorias?.[0] || "Diversos";
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
                placeholder="Buscar…"
                style={{ width: "100%", marginBottom: 12 }}
              />
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

  const abrir = () => { setResaltadoId(value || ""); setEditando(null); setBusqueda(""); setOpen(true); };
  const cerrar = () => { setOpen(false); setBusqueda(""); setEditando(null); };

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
                <Button type="button" variant="ghost" onClick={() => setEditando(null)}>Volver a la lista</Button>
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
                  <div style={{ padding: 16, textAlign: "center", fontSize: 12, color: T.textFaint }}>Sin resultados</div>
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

  return { rows, sheetsFound, nuevas, actualizaciones, sinCompania };
}

function parseTransaccionesWorkbook(arrayBuffer, partidas, proveedores = [], cuentas = []) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const rows = [];
  const sheetsFound = [];
  let matched = 0, unmatched = 0, conCuenta = 0, sinCuenta = 0;

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

      rows.push({
        id: uid(),
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
        concepto_detallado: (col.concepto !== -1 && row[col.concepto]) ? String(row[col.concepto]).trim() : "",
        importe,
        moneda: (col.moneda !== -1 && row[col.moneda]) ? String(row[col.moneda]).trim().toUpperCase() : "MXP",
        status: (col.status !== -1 && row[col.status]) ? String(row[col.status]).trim() : "",
        folio_compra_sae: (col.folioCompraSae !== -1 && row[col.folioCompraSae]) ? String(row[col.folioCompraSae]).trim() : "",
        folio_factura: (col.folioFactura !== -1 && row[col.folioFactura]) ? String(row[col.folioFactura]).trim() : "",
        forma_pago: (col.formaPago !== -1 && row[col.formaPago]) ? String(row[col.formaPago]).trim() : "",
        metodo_pago: (col.metodoPago !== -1 && row[col.metodoPago]) ? String(row[col.metodoPago]).trim() : "",
        referencia_pago: (col.referenciaPago !== -1 && row[col.referenciaPago]) ? String(row[col.referenciaPago]).trim() : "",
      });
      count++;
    }
    if (count) sheetsFound.push({ sheetName, count });
  });

  return { rows, sheetsFound, matched, unmatched, conCuenta, sinCuenta };
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

  const mesesDisponibles = MESES.filter((m) => partidasUnidad.some((p) => p.mes === m));
  const [mesesSeleccionados, setMesesSeleccionados] = useSessionState("ss-dashboard-meses", []);
  const aniosDisponibles = [...new Set(partidasUnidad.map((p) => p.anio).filter(Boolean))].sort((a, b) => a - b);
  const [aniosSeleccionados, setAniosSeleccionados] = useSessionState("ss-dashboard-anios", []);
  const [monedaResumen, setMonedaResumen] = useSessionState("ss-dashboard-moneda-resumen", "MXP");
  const partidasFiltradasMes = partidasUnidad.filter((p) =>
    (!mesesSeleccionados.length || mesesSeleccionados.includes(p.mes)) &&
    (!aniosSeleccionados.length || aniosSeleccionados.includes(p.anio))
  );
  const mesLabel = (mesesSeleccionados.length || aniosSeleccionados.length)
    ? ` · ${[...mesesSeleccionados, ...aniosSeleccionados].join(", ")}`
    : "";
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

  const totalPresupuestadoMXN = partidasFiltradasMes.filter((p) => (p.moneda || "MXP") !== "USD").reduce((s, p) => s + (Number(p.monto_estimado) || 0), 0);
  const totalEjecutadoMXN = transFiltradasMes.filter((t) => (t.moneda || "MXP") !== "USD").reduce((s, t) => s + (Number(t.importe) || 0), 0);
  const totalPagadoMXN = transFiltradasMes.filter((t) => (t.moneda || "MXP") !== "USD" && t.status === "Pagado").reduce((s, t) => s + (Number(t.importe) || 0), 0);
  const totalPresupuestadoUSD = partidasFiltradasMes.filter((p) => p.moneda === "USD").reduce((s, p) => s + (Number(p.monto_estimado) || 0), 0);
  const totalEjecutadoUSD = transFiltradasMes.filter((t) => t.moneda === "USD").reduce((s, t) => s + (Number(t.importe) || 0), 0);
  const totalPagadoUSD = transFiltradasMes.filter((t) => t.moneda === "USD" && t.status === "Pagado").reduce((s, t) => s + (Number(t.importe) || 0), 0);

  // Filtro de Proyecto para los cuadros de montos — usa los importes YA prorrateados
  // (porProyectoPorMoneda), así que si una partida es "Desh Gral" y filtras por "Desh Marfo",
  // sí cuenta su parte correspondiente.
  const [proyectoKpi, setProyectoKpi] = useSessionState("ss-dashboard-proyecto", "Todos");
  // Si el proyecto guardado no existe en ESTA compañía (ej. veníamos de otra
  // unidad), se regresa solo a "Todos" en vez de quedarse "huérfano" en $0.00.
  useEffect(() => {
    if (proyectoKpi !== "Todos" && !proyectosUnidad.some((p) => p.nombre === proyectoKpi)) {
      setProyectoKpi("Todos");
    }
  }, [unidad, proyectosUnidad.map((p) => p.nombre).join(","), proyectoKpi]);

  const kpiDataDe = (moneda) => {
    if (proyectoKpi === "Todos") {
      return moneda === "USD"
        ? { presupuestado: totalPresupuestadoUSD, ejecutado: totalEjecutadoUSD, pagado: totalPagadoUSD }
        : { presupuestado: totalPresupuestadoMXN, ejecutado: totalEjecutadoMXN, pagado: totalPagadoMXN };
    }
    return porProyectoPorMoneda[moneda].find((p) => p.proyecto === proyectoKpi) || { presupuestado: 0, ejecutado: 0, pagado: 0 };
  };
  const proyectoKpiData = kpiDataDe("MXP");
  const kpiOcupado = proyectoKpiData.ejecutado;
  const kpiPagado = proyectoKpiData.pagado;
  const kpiPorPagar = kpiOcupado - kpiPagado;
  const kpiDisponible = proyectoKpiData.presupuestado - kpiOcupado;

  const proyectoKpiDataUSD = kpiDataDe("USD");
  const kpiOcupadoUSD = proyectoKpiDataUSD.ejecutado;
  const kpiPagadoUSD = proyectoKpiDataUSD.pagado;
  const kpiPorPagarUSD = kpiOcupadoUSD - kpiPagadoUSD;
  const kpiDisponibleUSD = proyectoKpiDataUSD.presupuestado - kpiOcupadoUSD;

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
      <Panel
        title="Resumen financiero"
        subtitle="Importes prorrateados por partida"
        right={
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end" }}>
            <Field label="Proyecto">
              <Select value={proyectoKpi} onChange={(e) => setProyectoKpi(e.target.value)} style={{ width: 190 }}>
                <option>Todos</option>
                {proyectosUnidad.map((p) => <option key={p.nombre}>{p.nombre}</option>)}
              </Select>
            </Field>
            <Field label="Mes">
              <MesMultiSelect mesesDisponibles={mesesDisponibles} seleccionados={mesesSeleccionados} onChange={setMesesSeleccionados} />
            </Field>
            <Field label="Año">
              <AnioMultiSelect aniosDisponibles={aniosDisponibles} seleccionados={aniosSeleccionados} onChange={setAniosSeleccionados} />
            </Field>
          </div>
        }
      >
        {(() => {
          const filas = [
            { moneda: "MXN", data: proyectoKpiData, ocupado: kpiOcupado, pagado: kpiPagado, porPagar: kpiPorPagar, disponible: kpiDisponible },
            { moneda: "USD", data: proyectoKpiDataUSD, ocupado: kpiOcupadoUSD, pagado: kpiPagadoUSD, porPagar: kpiPorPagarUSD, disponible: kpiDisponibleUSD },
          ];
          const colIcon = { textAlign: "center", color: T.textFaint, marginBottom: 6, display: "flex", justifyContent: "center" };
          const colVal = { textAlign: "center", fontFamily: T.fontMono, fontSize: 15, fontWeight: 700 };
          return (
            <table style={{ ...tableStyle, tableLayout: "fixed" }}>
              <colgroup>
                <col style={{ width: 110 }} />
                {[0, 1, 2, 3, 4].map((i) => <col key={i} />)}
              </colgroup>
              <thead>
                <tr>
                  <th style={thStyle}>Moneda</th>
                  <th style={{ ...thStyle, textAlign: "center" }}>Presupuestado</th>
                  <th style={{ ...thStyle, textAlign: "center" }}>Ocupado</th>
                  <th style={{ ...thStyle, textAlign: "center" }}>Pagado</th>
                  <th style={{ ...thStyle, textAlign: "center" }}>Por Pagar</th>
                  <th style={{ ...thStyle, textAlign: "center" }}>Disponible</th>
                </tr>
              </thead>
              <tbody>
                {filas.map((f) => (
                  <tr key={f.moneda}>
                    <td style={tdStyle}>
                      <span style={{
                        display: "inline-flex", alignItems: "center", justifyContent: "center",
                        width: 52, height: 52, borderRadius: "50%", background: T.accentBg,
                        color: T.accent, fontWeight: 700, fontSize: 12.5, fontFamily: T.fontMono,
                      }}>
                        {f.moneda}
                      </span>
                    </td>
                    <td style={tdStyle}>
                      <div style={colIcon}><Wallet size={20} color={T.accent} /></div>
                      <div style={colVal}>{money(f.data.presupuestado, f.moneda === "USD" ? "USD" : "MXP")}</div>
                    </td>
                    <td style={tdStyle}>
                      <div style={colIcon}><BarChart3 size={20} color={T.red} /></div>
                      <div style={{ ...colVal, color: f.data.presupuestado && f.ocupado / f.data.presupuestado > 1 ? T.red : T.text }}>
                        {money(f.ocupado, f.moneda === "USD" ? "USD" : "MXP")}
                      </div>
                    </td>
                    <td style={tdStyle}>
                      <div style={colIcon}><CheckCircle2 size={20} color={T.teal} /></div>
                      <div style={{ ...colVal, color: T.teal }}>{money(f.pagado, f.moneda === "USD" ? "USD" : "MXP")}</div>
                    </td>
                    <td style={tdStyle}>
                      <div style={colIcon}><FileEdit size={20} color={T.amber} /></div>
                      <div style={{ ...colVal, color: T.amber }}>{money(f.porPagar, f.moneda === "USD" ? "USD" : "MXP")}</div>
                    </td>
                    <td style={tdStyle}>
                      <div style={colIcon}>
                        {f.disponible < 0 ? <ArrowDownCircle size={20} color={T.red} /> : <ArrowUpCircle size={20} color={T.teal} />}
                      </div>
                      <div style={{ ...colVal, color: f.disponible < 0 ? T.red : T.teal }}>{money(f.disponible, f.moneda === "USD" ? "USD" : "MXP")}</div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          );
        })()}
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 14, fontSize: 11.5, color: T.textFaint }}>
          <Info size={14} /> Importes prorrateados por partida
        </div>
      </Panel>

      <ResumenComparativoPanel
        partidasUnidad={partidasUnidad}
        transacciones={transUnidad}
        proyectosUnidad={proyectosUnidad}
        proyectoKpi={proyectoKpi}
        setProyectoKpi={setProyectoKpi}
        monedaResumen={monedaResumen}
        setMonedaResumen={setMonedaResumen}
      />

      <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
        <Panel title="Presupuesto vs. ejecutado por proyecto" subtitle={`Gastos compartidos ya prorrateados según su marcador${mesLabel}`}>
          <div style={{ fontSize: 10.5, color: T.accent, letterSpacing: "0.08em", textTransform: "uppercase", fontFamily: T.fontMono, marginBottom: 4 }}>MXP</div>
          <ResponsiveContainer width="100%" height={Math.max(180, porProyecto.length * 34)}>
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
                <Bar dataKey="ejecutado" name="Ejecutado" fill={T.teal} radius={[0,3,3,0]} />
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
function ResumenComparativoPanel({
  partidasUnidad, transacciones, proyectosUnidad,
  proyectoKpi, setProyectoKpi,
  monedaResumen, setMonedaResumen,
}) {
  // Filtrado en cascada de tres niveles:
  //   1. Año     -> "YTD" o un año concreto con datos
  //   2. Periodo -> todo el año, o un rango de meses (no aplica en YTD)
  //   3. Desde/Hasta -> meses sueltos, porque el año ya quedó fijo arriba
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

  const partidaDe = (t) => partidasRango.find((p) => p.id === t.partida_id);
  const partidasMoneda = partidasRango.filter((p) => (p.moneda || "MXP") === monedaResumen);
  const pagadasMoneda = transacciones.filter((t) => t.status === "Pagado" && (t.moneda || "MXP") === monedaResumen && idsRango.has(t.partida_id));

  // Eje de meses COMPARTIDO por las dos tablas. Antes cada una lo armaba con
  // sus propios datos, así que un mes filtrado sin pagos simplemente perdía su
  // columna en "Pagado real" y las tablas dejaban de cuadrar. Ahora la columna
  // aparece en ambas (en ceros si no hay nada), que es el punto del panel.
  const aniosPresentes = [...new Set([
    ...partidasMoneda.map((p) => p.anio),
    ...pagadasMoneda.map((t) => partidaDe(t)?.anio),
  ].filter(Boolean))];
  const multiAnio = aniosPresentes.length > 1;

  const etiquetaColumna = (mes, anio) => (multiAnio && anio) ? `${mes} ${anio}` : mes;
  const columnaDePptado = (p) => etiquetaColumna(p.mes, p.anio);
  const columnaDePagado = (t) => {
    const p = partidaDe(t);
    return p ? etiquetaColumna(p.mes, p.anio) : null;
  };

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
      ...p, proyecto, monto_estimado: (Number(p.monto_estimado) || 0) * fraccion, _columna: columnaDePptado(p),
    }))
  );
  if (proyectoKpi !== "Todos") partidasResueltas = partidasResueltas.filter((p) => p.proyecto === proyectoKpi);

  let filasPagadoResueltas = pagadasMoneda.flatMap((t) => {
    const p = partidaDe(t);
    return resolverProrrateo(p.proyecto, proyectosUnidad).map(({ proyecto, fraccion }) => ({
      proyecto, rubro: p.rubro, concepto: t.concepto_detallado || t.proveedor || "—",
      importe: (Number(t.importe) || 0) * fraccion, _columna: columnaDePagado(t),
    }));
  });
  if (proyectoKpi !== "Todos") filasPagadoResueltas = filasPagadoResueltas.filter((f) => f.proyecto === proyectoKpi);

  const pivotPptado = pivotearPorMes(partidasResueltas, ["proyecto", "rubro", "concepto"], columnas, "monto_estimado", "_columna");
  const totalPptado = columnas.reduce((acc, m) => { acc[m] = 0; return acc; }, {});
  let granTotalPptado = 0;
  partidasResueltas.forEach((p) => {
    const v = Number(p.monto_estimado) || 0;
    if (totalPptado[p._columna] !== undefined) totalPptado[p._columna] += v;
    granTotalPptado += v;
  });

  const pivotPagado = pivotearPorMes(filasPagadoResueltas, ["proyecto", "rubro", "concepto"], columnas, "importe", "_columna");
  const totalPagado = columnas.reduce((acc, m) => { acc[m] = 0; return acc; }, {});
  let granTotalPagado = 0;
  filasPagadoResueltas.forEach((f) => {
    const v = Number(f.importe) || 0;
    if (totalPagado[f._columna] !== undefined) totalPagado[f._columna] += v;
    granTotalPagado += v;
  });

  // Un solo control para las dos tablas: si queda algo abierto en cualquiera de
  // las dos, contrae ambas; si ya están todas cerradas, las abre.
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
      subtitle="Filtros propios de este panel — aplican a las dos tablas de abajo"
      right={
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "flex-end" }}>
            <Button variant="ghost" onClick={toggleTodo} style={{ height: 34 }}>
              {todoContraido ? "Expandir todo" : "Contraer todo"}
            </Button>
          </div>
          <Field label="Proyecto">
            <Select value={proyectoKpi} onChange={(e) => setProyectoKpi(e.target.value)} style={{ width: 190 }}>
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
          <Field label="Moneda">
            <SlidingToggle opciones={["MXP", "USD"]} value={monedaResumen} onChange={setMonedaResumen} />
          </Field>
        </div>
      }
    >
      <div style={{ marginBottom: 8 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>Presupuestado por rubro</div>
      </div>
      {!partidasMoneda.length ? (
        <div style={{ textAlign: "center", color: T.textFaint, fontSize: 12.5, padding: 24 }}>Sin partidas en {monedaResumen} para estos filtros</div>
      ) : (
        <div style={{ overflowX: "auto", marginBottom: 28 }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                <th style={thStyle}>Proyecto</th>
                {columnas.map((m) => <th key={m} style={{ ...thStyle, textAlign: "right" }}>{m}</th>)}
                <th style={{ ...thStyle, textAlign: "right" }}>Total general</th>
              </tr>
            </thead>
            <tbody>
              {buildPivotTrs(pivotPptado, "", collapsedPptado, togglePptado, columnas, 0)}
              <tr style={{ background: T.panelAlt }}>
                <td style={{ ...tdStyle, fontWeight: 700 }}>Total general</td>
                {columnas.map((m) => <td key={m} style={{ ...tdStyle, fontFamily: T.fontMono, textAlign: "right", fontWeight: 700 }}>{money(totalPptado[m], monedaResumen)}</td>)}
                <td style={{ ...tdStyle, fontFamily: T.fontMono, textAlign: "right", fontWeight: 700 }}>{money(granTotalPptado, monedaResumen)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}

      <div style={{ marginBottom: 8 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>Pagado real (Status = Pagado)</div>
      </div>
      {!pagadasMoneda.length ? (
        <div style={{ textAlign: "center", color: T.textFaint, fontSize: 12.5, padding: 24 }}>Sin pagos en {monedaResumen} para estos filtros</div>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={tableStyle}>
            <thead>
              <tr>
                <th style={thStyle}>Proyecto</th>
                {columnas.map((m) => <th key={m} style={{ ...thStyle, textAlign: "right" }}>{m}</th>)}
                <th style={{ ...thStyle, textAlign: "right" }}>Total general</th>
              </tr>
            </thead>
            <tbody>
              {buildPivotTrs(pivotPagado, "", collapsedPagado, togglePagado, columnas, 0)}
              <tr style={{ background: T.panelAlt }}>
                <td style={{ ...tdStyle, fontWeight: 700 }}>Total general</td>
                {columnas.map((m) => <td key={m} style={{ ...tdStyle, fontFamily: T.fontMono, textAlign: "right", fontWeight: 700 }}>{money(totalPagado[m], monedaResumen)}</td>)}
                <td style={{ ...tdStyle, fontFamily: T.fontMono, textAlign: "right", fontWeight: 700 }}>{money(granTotalPagado, monedaResumen)}</td>
              </tr>
            </tbody>
          </table>
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
    entries: entries.map(([value, groupRows]) => ({ value, child: pivotearPorMes(groupRows, rest, meses, montoKey, colKey) })),
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
    const headers = ["Mes", "SMI", "Concepto", "Rubro", "Categoria", "Proyecto", "Sub Total MXN", "Sub Total USD", "Moneda", "ID"];
    const ejemplos = {
      OSB: ["Agosto", "", "Ejemplo: Servicio de energía eléctrica", "Servicios Básicos", "Energía eléctrica", "Todos", 12000, 0, "MXP", ""],
      CTM: ["Agosto", "", "Ejemplo: Renta de oficina base Poza Rica", "Servicios Operativos", "Arrendamientos", "Todos", 27000, 0, "MXP", ""],
      ISE: ["Agosto", "", "Ejemplo: Producto químico deshidratación", "Productos Químicos", "Productos químicos de operación", "Desh Gral", 0, 45000, "USD", ""],
    };
    UNIDAD_KEYS.forEach((u) => {
      const ws = wbx.addWorksheet(`RawData-${u}`);
      const headerRow = ws.addRow(headers);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
      });
      const ejemploRow = ws.addRow(ejemplos[u]);
      ejemploRow.eachCell((cell) => { cell.font = { italic: true, color: { argb: "FF8B99A6" } }; });
      ws.columns = [{ width: 10 }, { width: 8 }, { width: 45 }, { width: 24 }, { width: 30 }, { width: 16 }, { width: 14 }, { width: 14 }, { width: 9 }, { width: 14 }];
      ws.views = [{ state: "frozen", ySplit: 1 }];
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

// Modal ligero para editar una transacción sin salir de la vista en la que
// estás (ej. desde la fila expandida de una partida) — no cambia de pestaña.
// No incluye el selector de Partida (aquí ya se sabe a cuál pertenece).
function TransaccionQuickEditModal({ transaccion, onClose, transaccionesApi, proveedoresApi, cuentasApi, unidad }) {
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

function PartidasTab({ unidad, unidades, partidas, partidasApi, perfilesApi, transacciones, transaccionesApi, proveedoresApi, cuentasApi }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadores = marcadoresDisponibles(proyectosUnidad);
  const anioDefault = (() => {
    const anios = partidas.filter((p) => p.unidad === unidad).map((p) => p.anio).filter(Boolean);
    return anios.length ? Math.max(...anios) : new Date().getFullYear();
  })();
  const blank = { unidad, mes: "Agosto", anio: anioDefault, smi: "", concepto: "", rubro: RUBROS[0].rubro, categoria: RUBROS[0].categorias[0], proyecto: marcadores[0] || "", monto_estimado: "", moneda: "MXP", folio: "", es_recurrente: false };
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
      const enTexto = [p.concepto, p.folio, p.smi, p.categoria].some((v) => (v || "").toLowerCase().includes(q));
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

  const [sort, setSort] = useSessionState("ss-partidas-sort", { key: null, dir: "asc" });
  const partidasOrdenadas = sortRows(partidasFiltradas, sort, {
    mes: (r) => MESES.indexOf(r.mes),
    monto_estimado: (r) => Number(r.monto_estimado) || 0,
    anio: (r) => Number(r.anio) || 0,
  });

  const GROUP_OPCIONES = [
    { value: "anio", label: "Año" },
    { value: "mes", label: "Mes" },
    { value: "rubro", label: "Rubro" },
    { value: "categoria", label: "Categoría" },
    { value: "proyecto", label: "Proyecto" },
  ];
  const [groupBys, setGroupBys] = useSessionState("ss-partidas-groupbys", []);
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
    const COLS = [
      { header: "Unidad",    width: 9,  get: (p) => p.unidad },
      { header: "Folio",     width: 16, get: (p) => p.folio || "" },
      { header: "Mes",       width: 11, get: (p) => p.mes },
      { header: "Año",       width: 7,  get: (p) => p.anio },
      { header: "Concepto",  width: 50, get: (p) => p.concepto },
      { header: "Rubro",     width: 26, get: (p) => p.rubro },
      { header: "Categoria", width: 30, get: (p) => p.categoria },
      { header: "Proyecto",  width: 17, get: (p) => p.proyecto || "" },
      { header: "Moneda",    width: 9,  get: (p) => p.moneda || "MXP" },
      { header: "Monto",     width: 15, get: (p) => Number(p.monto_estimado) || 0, money: true },
      { header: "Usado",     width: 15, get: (p) => usadoDe(p), money: true },
      { header: "Disponible",width: 15, get: (p) => (Number(p.monto_estimado) || 0) - usadoDe(p), money: true },
      { header: "SMI",       width: 10, get: (p) => p.smi || "" },
      { header: "Recurrente",width: 12, get: (p) => (p.es_recurrente ? "Sí" : "No") },
    ];
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
      hr.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
      });
      filas.forEach((p) => {
        const row = ws.addRow(COLS.map((c) => c.get(p)));
        COLS.forEach((c, ci) => {
          const cell = row.getCell(ci + 1);
          cell.font = { name: "Calibri", size: 11 };
          if (c.money) cell.numFmt = '"$"#,##0.00';
          // El folio se fuerza a texto: si Excel lo interpreta como número,
          // se pierden ceros y el archivo deja de servir para reimportar.
          if (c.header === "Folio") cell.alignment = { horizontal: "left" };
        });
      });
      ws.views = [{ state: "frozen", ySplit: 1 }];
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

  const COLUMNAS_PARTIDA = [
    { key: "mes", label: "Mes", render: (p) => p.mes },
    { key: "anio", label: "Año", render: (p) => p.anio },
    {
      key: "concepto", label: "Concepto",
      render: (p) => (
        <span>
          {p.concepto}
          {p.es_recurrente && <span title="Recurrente — se repite cada mes" style={{ marginLeft: 6, fontSize: 11, color: T.accent }}>🔁</span>}
        </span>
      ),
    },
    { key: "rubro", label: "Rubro", render: (p) => <Pill>{p.rubro}</Pill> },
    { key: "categoria", label: "Categoría", render: (p) => <span style={{ color: T.textDim }}>{p.categoria}</span> },
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
    // Copia la partida como registro NUEVO. Se limpian dos cosas a propósito:
    //  - folio: submit() respeta el que venga en el formulario, así que heredarlo
    //    haría chocar el duplicado con la original. Vacío, se genera uno nuevo.
    //  - es_recurrente: si se heredara, "Generar recurrentes pendientes" crearía
    //    una segunda serie completa en todos los meses que faltan del año.
    const { id, folio, created_at, updated_at, created_by, updated_by, es_recurrente, ...resto } = p;
    const limpio = Object.fromEntries(Object.entries(resto).filter(([k]) => !k.startsWith("_")));
    setForm({ ...limpio, folio: "", es_recurrente: false });
    setEditId(null);
    setModalOpen(true);
  };
  const startEdit = (p) => { setForm(p); setEditId(p.id); setModalOpen(true); };

  const [generandoRecurrentes, setGenerandoRecurrentes] = useState(false);
  const partidasRecurrentes = partidasUnidad.filter((p) => p.es_recurrente);
  // Para cada partida recurrente, calcula qué meses le faltan (desde su propio
  // mes+1 hasta diciembre de su mismo año) que todavía no existan como partida
  // igual (mismo Proyecto+Rubro+Categoria+Concepto+Año) más adelante en el año.
  const faltantesRecurrentes = partidasRecurrentes.flatMap((base) => {
    const idxBase = MESES.indexOf(base.mes);
    if (idxBase === -1) return [];
    const meses = [];
    for (let i = idxBase + 1; i < 12; i++) {
      const yaExiste = partidasUnidad.some((p) =>
        p.anio === base.anio && p.mes === MESES[i] &&
        p.proyecto === base.proyecto && p.rubro === base.rubro &&
        p.categoria === base.categoria && p.concepto === base.concepto
      );
      if (!yaExiste) meses.push({ base, mes: MESES[i] });
    }
    return meses;
  });

  const generarRecurrentesPendientes = async () => {
    if (!faltantesRecurrentes.length) return;
    const confirmado = confirm(
      `Esto va a crear ${faltantesRecurrentes.length} partida(s) nueva(s) — los meses que faltan hasta diciembre para cada partida marcada como recurrente, con el mismo monto. ¿Continuar?`
    );
    if (!confirmado) return;
    setGenerandoRecurrentes(true);
    try {
      let folios = partidas.filter((p) => p.unidad === unidad).map((p) => p.folio);
      for (const { base, mes } of faltantesRecurrentes) {
        const folio = autoFolio(unidad, mes, base.anio, folios);
        folios = [...folios, folio];
        await partidasApi.insert({
          id: uid(), unidad, mes, anio: base.anio, smi: base.smi || "",
          concepto: base.concepto, rubro: base.rubro, categoria: base.categoria,
          proyecto: base.proyecto, monto_estimado: base.monto_estimado, moneda: base.moneda,
          folio, es_recurrente: true,
        });
      }
    } catch (err) {
      alert("No se pudo generar: " + (err.message || err));
    } finally {
      setGenerandoRecurrentes(false);
    }
  };
  const closeModal = () => { setModalOpen(false); setEditId(null); setForm({ ...blank, anio: anioDefault, proyecto: marcadores[0] || "" }); };
  const remove = (id) => {
    const p = partidasUnidad.find((x) => x.id === id);
    if (!confirm(`¿Eliminar la partida "${p?.concepto || id}" (folio ${p?.folio || "—"})? Esto no se puede deshacer.`)) return;
    partidasApi.remove(id).catch((err) => alert("No se pudo eliminar: " + (err.message || err)));
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel
        title={`Partidas de ${unidad}`}
        subtitle={filtrosActivos ? `${partidasFiltradas.length} de ${partidasUnidad.length} registradas` : `${partidasUnidad.length} registradas`}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            {faltantesRecurrentes.length > 0 && (
              <Button variant="ghost" onClick={generarRecurrentesPendientes} disabled={generandoRecurrentes}>
                {generandoRecurrentes ? "Generando…" : `Generar recurrentes pendientes (${faltantesRecurrentes.length})`}
              </Button>
            )}
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
            <Field label="Recurrente">
              <label style={{ display: "flex", alignItems: "center", gap: 8, height: 34, cursor: "pointer" }}>
                <input type="checkbox" checked={!!form.es_recurrente} onChange={(e) => setForm({ ...form, es_recurrente: e.target.checked })} />
                <span style={{ fontSize: 12, color: T.textDim }}>Se repite cada mes</span>
              </label>
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
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
      });
      const ejemploRow = ws.addRow(ejemplo);
      ejemploRow.eachCell((cell) => { cell.font = { italic: true, color: { argb: "FF8B99A6" } }; });
      ws.columns = headers.map(() => ({ width: 16 }));
      ws.getColumn(7).width = 32; ws.getColumn(8).width = 40;
      ws.getColumn(11).width = 22; ws.getColumn(12).width = 22;
      ws.views = [{ state: "frozen", ySplit: 1 }];
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
      const { rows, sheetsFound, matched, unmatched, conCuenta, sinCuenta } = parseTransaccionesWorkbook(buf, partidas, proveedores, cuentas);
      if (!rows.length) {
        setError('No encontré una hoja con columnas "Día", "Importe" y "A Partida" en este archivo.');
        return;
      }
      setPreview({ rows, sheetsFound, matched, unmatched, conCuenta, sinCuenta, fileName: file.name });
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
      const toInsert = preview.rows.map(({ _cuentasDisponibles, ...r }) => ({
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
          </div>
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

function TransaccionesTab({ unidad, unidades, partidas, partidasApi, transacciones, transaccionesApi, proveedoresApi, cuentasApi, perfilesApi, notasApi, session }) {
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadoresProyecto = marcadoresDisponibles(proyectosUnidad);
  const proveedoresUnidad = proveedoresApi.rows.filter((p) => p.unidad === unidad);
  const blank = {
    partida_id: partidasUnidad[0]?.id || "", unidad_detectada: unidad, dia: "", solicitante: "", smi: "", proyecto: "", zona: "", area: "",
    proveedor: "", proveedor_id: "", cuenta_id: "", concepto_detallado: "", importe: "", moneda: "MXP", status: "No Pagado", fecha_pago: "",
    folio_compra_sae: "", folio_factura: "", forma_pago: "", metodo_pago: "", referencia_pago: "",
  };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);
  const cuentasDelProveedorSeleccionado = form.proveedor_id ? cuentasApi.rows.filter((c) => c.proveedor_id === form.proveedor_id) : [];
  const transUnidad = transacciones.filter((t) => t.unidad_detectada === unidad);
  const sinVincular = transacciones.filter((t) => !t.partida_id && t.unidad_detectada === unidad);

  const [filtros, setFiltros] = useSessionState("ss-transacciones-filtros", { texto: "", fechaDesde: "", fechaHasta: "", reportado: "Todos", enviadoPagos: "Todos" });
  const [sort, setSort] = useSessionState("ss-transacciones-sort", { key: "dia", dir: "desc" });
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

  const GROUP_OPCIONES_TRANS = [
    { value: "dia", label: "Día de Pago Programado" },
    { value: "zona", label: "Zona" },
    { value: "area", label: "Área" },
    { value: "proveedor", label: "Proveedor" },
    { value: "proyecto", label: "Proyecto (transacción)" },
    { value: "status", label: "Status" },
    { value: "moneda", label: "Moneda" },
    { value: "_proyecto", label: "Proyecto (partida)" },
    { value: "_rubro", label: "Rubro (partida)" },
    { value: "_mes", label: "Mes (partida)" },
    { value: "_vinculo", label: "Vínculo" },
  ];
  const [groupBys, setGroupBys] = useSessionState("ss-transacciones-groupbys", []);
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
  const GROUP_OPCIONES_SINVINC = [
    { value: "dia", label: "Día de Pago Programado" },
    { value: "zona", label: "Zona" },
    { value: "area", label: "Área" },
    { value: "proveedor", label: "Proveedor" },
    { value: "proyecto", label: "Proyecto" },
    { value: "status", label: "Status" },
    { value: "moneda", label: "Moneda" },
  ];
  const [filtrosSV, setFiltrosSV] = useSessionState("ss-transacciones-sv-filtros", { fechaDesde: "", fechaHasta: "" });
  const [groupBysSV, setGroupBysSV] = useSessionState("ss-transacciones-sv-groupbys", []);
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
        <PartidaPickerButton
          partidas={partidasUnidad}
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
          partidas={partidasUnidad}
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

      {modalOpen && (
        <Modal
          title={editId ? "Editar transacción" : "Nueva transacción real"}
          subtitle={form.folio_transaccion ? `ID: ${form.folio_transaccion} — se vincula a una partida` : "Se vincula a una partida — una partida puede tener varias"}
          onClose={closeModal}
        >
          <AutoriaCaption record={form} perfilesApi={perfilesApi} />
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
            <Field label="Partida" style={{ gridColumn: "span 4" }}>
              <PartidaPickerButton
                partidas={partidasUnidad}
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
                {ZONAS.map((z) => <option key={z}>{z}</option>)}
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
  const [sort, setSort] = useSessionState("ss-reporte-sort", { key: "dia", dir: "desc" });
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
          cell.alignment = { horizontal: "center" };
          cell.font = { name: "Calibri", size: 11 };
        });
        fila += 1;

        filasGrupo.forEach((f) => {
          const row = ws.getRow(fila);
          columnasExcel.forEach((c, ci) => {
            const cell = row.getCell(ci + 1);
            cell.value = c.get(f);
            cell.alignment = { horizontal: "center" };
            cell.font = { name: "Calibri", size: 11 };
            if (c.money) cell.numFmt = '"$"#,##0.00';
          });
          fila += 1;
        });

        fila += 2; // espacio antes de la siguiente sección
      });
    });

    const buffer = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte-pagos-${unidad}-${new Date().toISOString().slice(0, 10)}.xlsx`;
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
const COLUMNAS_REPORTE_DIRECCION = [
  { key: "dia", label: "Día" },
  { key: "solicitante", label: "Solicitante" },
  { key: "proyecto", label: "Proyecto" },
  { key: "zona", label: "Zona" },
  { key: "proveedor", label: "Proveedor" },
  { key: "concepto", label: "Concepto" },
  { key: "importe", label: "Importe" },
  { key: "moneda", label: "Moneda" },
  { key: "a_partida", label: "A Partida" },
  { key: "status", label: "Status" },
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
  const [sort, setSort] = useSessionState("ss-reporte-direccion-sort", { key: "dia", dir: "desc" });
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
    const wbx = new ExcelJS.Workbook();
    const ws = wbx.addWorksheet("Reporte pagos direccion");

    ws.columns = [
      { width: 8.875 }, { width: 17.125 }, { width: 11.375 }, { width: 8.625 }, { width: 23.25 },
      { width: 41.125 }, { width: 11.125 }, { width: 11.375 }, { width: 11.875 }, { width: 9.625 },
    ];

    const diasOrdenados = filasOrdenadas.map((f) => f.dia).filter(Boolean).sort();
    const inicio = fechaDesde || diasOrdenados[0] || "";
    const fin = fechaHasta || diasOrdenados[diasOrdenados.length - 1] || "";

    const tituloCell = ws.getCell("B1");
    tituloCell.value = `Reporte de pagos a realizar del dia ${inicio} al dia ${fin} Compañía ${unidad}`;
    tituloCell.font = { bold: true, size: 16, name: "Calibri" };

    ws.getCell("F3").value = "Total a pagar MXP";
    ws.getCell("F3").font = { bold: true, name: "Calibri", size: 11 };
    ws.getCell("F3").alignment = { horizontal: "right" };
    ws.getCell("G3").value = totalMXN;
    ws.getCell("G3").numFmt = '"$"#,##0.00';
    ws.getCell("G3").font = { name: "Calibri", size: 11 };

    ws.getCell("F4").value = "Total a pagar USD";
    ws.getCell("F4").font = { bold: true, name: "Calibri", size: 11 };
    ws.getCell("F4").alignment = { horizontal: "right" };
    ws.getCell("G4").value = totalUSD;
    ws.getCell("G4").numFmt = '"$"#,##0.00';
    ws.getCell("G4").font = { name: "Calibri", size: 11 };

    const headers = ["Día", "Solicitante", "Proyecto", "Zona", "Proveedor", "Concepto", "Importe", "Moneda", "A Partida", "Status"];
    const headerRow = ws.getRow(6);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.alignment = { horizontal: "center" };
      cell.font = { name: "Calibri", size: 11 };
    });

    filasOrdenadas.forEach((f, i) => {
      const row = ws.getRow(7 + i);
      const valores = [f.dia, f.solicitante, f.proyecto, f.zona, f.proveedor, f.concepto, f.importe, f.moneda, f.a_partida, f.status];
      valores.forEach((v, ci) => {
        const cell = row.getCell(ci + 1);
        cell.value = v;
        cell.alignment = { horizontal: "center" };
        cell.font = { name: "Calibri", size: 11 };
        if (ci === 6) cell.numFmt = '"$"#,##0.00'; // columna Importe
      });
    });

    const buffer = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte-pagos-direccion-${unidad}-${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const [generandoReporte, setGenerandoReporte] = useState(false);
  const generarReportePDF = async () => {
    if (!filasOrdenadas.length) {
      alert("No hay transacciones en el filtro actual para generar el reporte.");
      return;
    }
    const confirmado = confirm(
      `Esto va a generar un PDF con las ${filasOrdenadas.length} transacción(es) que tienes filtradas ahora, y las va a marcar como "Reportadas". ¿Continuar?`
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

      autoTable(doc, {
        startY: 58,
        head: [["Día", "Solicitante", "Proyecto", "Zona", "Proveedor", "Concepto", "Importe", "Moneda", "A Partida", "Status"]],
        body: filasOrdenadas.map((f) => [
          f.dia || "", f.solicitante || "", f.proyecto || "", f.zona || "", f.proveedor || "",
          f.concepto || "", money(f.importe, f.moneda), f.moneda || "", f.a_partida || "", f.status || "",
        ]),
        styles: { fontSize: 7.5, cellPadding: 4 },
        headStyles: { fillColor: [62, 92, 118], textColor: 255, halign: "center" },
        bodyStyles: { halign: "center" },
        columnStyles: { 5: { halign: "left" } },
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

function CatalogoTab({ unidad, unidades, proyectosApi, proveedoresApi, cuentasApi, perfilesApi }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const [nuevo, setNuevo] = useState({ nombre: "", grupo: "", pct: "" });
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({ nombre: "", grupo: "", pct: "" });
  const [guardandoId, setGuardandoId] = useState(null);

  const empezarEditar = (p) => {
    setEditingId(p.id);
    setDraft({ nombre: p.nombre || "", grupo: p.grupo || "", pct: p.pct ?? "" });
  };
  const cancelarEditar = () => { setEditingId(null); setDraft({ nombre: "", grupo: "", pct: "" }); };
  const guardarEditar = async (id) => {
    setGuardandoId(id);
    try {
      await proyectosApi.update(id, { nombre: draft.nombre, grupo: draft.grupo, pct: Number(draft.pct) || 0 });
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
              {proyectosUnidad.map((p) => {
                const editando = editingId === p.id;
                return (
                  <tr key={p.id}>
                    {editando ? (
                      <>
                        <td style={tdStyle}><TextInput autoFocus value={draft.nombre} onChange={(e) => setDraft({ ...draft, nombre: e.target.value })} /></td>
                        <td style={tdStyle}><TextInput value={draft.grupo} onChange={(e) => setDraft({ ...draft, grupo: e.target.value })} placeholder="Desh / Prod / IMP" /></td>
                        <td style={tdStyle}><TextInput type="number" value={draft.pct} onChange={(e) => setDraft({ ...draft, pct: e.target.value })} style={{ width: 90 }} /></td>
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

      <ProveedoresPanel unidad={unidad} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} />
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
      hr.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E5C76" } };
      });
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
      ws.views = [{ state: "frozen", ySplit: 1 }];
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
  const [sort, setSort] = useSessionState("veh:sort", { key: "no_economico", dir: "asc" });
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
  const [sort, setSort] = useSessionState("veh:mtto:sort", { key: "folio", dir: "asc" });
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
          {tab === "partidas" && <PartidasTab unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} perfilesApi={perfilesApi} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
          {tab === "transacciones" && <TransaccionesTab unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} notasApi={notasApi} session={session} />}
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
          {tab === "catalogo" && <CatalogoTab unidad={unidad} unidades={unidades} proyectosApi={proyectosApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} />}
        </>
      )}
    </div>
  );
}
