import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
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
const APP_VERSION = "1.46.2";
const CHANGELOG = [
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
function nextFolioTransaccion(unidad, mes, existingTransacciones) {
  const abr = MES_ABR[mes] || "GEN";
  const prefix = `${unidad}-${abr}-`;
  let max = 0;
  existingTransacciones.forEach((t) => {
    if (t.folio_transaccion && t.folio_transaccion.startsWith(prefix)) {
      const n = parseInt(t.folio_transaccion.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return `${prefix}${String(max + 1).padStart(3, "0")}`;
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

  const nuevaPartidaBlank = { mes: MESES[0], concepto: "", rubro: RUBROS[0]?.rubro || "", proyecto: proyectosOpciones[0] || "", monto_estimado: "", moneda: "MXP" };
  const [creando, setCreando] = useState(false);
  const [nuevaPartida, setNuevaPartida] = useState(nuevaPartidaBlank);
  const [guardandoPartida, setGuardandoPartida] = useState(false);

  const crearPartida = async () => {
    if (!nuevaPartida.concepto.trim() || !nuevaPartida.monto_estimado) return;
    setGuardandoPartida(true);
    try {
      const anio = new Date().getFullYear();
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
  const [mesesSeleccionados, setMesesSeleccionados] = useSessionState("ss-dashboard-meses", []);
  const partidasFiltradasMes = mesesSeleccionados.length ? partidasUnidad.filter((p) => mesesSeleccionados.includes(p.mes)) : partidasUnidad;
  const mesLabel = mesesSeleccionados.length ? ` · ${mesesSeleccionados.join(", ")}` : "";
  const idsPartidasFiltradas = new Set(partidasFiltradasMes.map((p) => p.id));
  const transFiltradasMes = transUnidad.filter((t) => idsPartidasFiltradas.has(t.partida_id));

  const porProyecto = useMemo(() => {
    const map = {};
    proyectosUnidad.forEach((p) => { map[p.nombre] = { proyecto: p.nombre, presupuestado: 0, ejecutado: 0, pagado: 0 }; });

    partidasFiltradasMes.forEach((partida) => {
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!map[proyecto]) map[proyecto] = { proyecto, presupuestado: 0, ejecutado: 0, pagado: 0 };
        map[proyecto].presupuestado += (Number(partida.monto_estimado) || 0) * fraccion;
      });
    });

    transFiltradasMes.forEach((t) => {
      const partida = partidasFiltradasMes.find((p) => p.id === t.partida_id);
      if (!partida) return;
      const splits = resolverProrrateo(partida.proyecto, proyectosUnidad);
      splits.forEach(({ proyecto, fraccion }) => {
        if (!map[proyecto]) map[proyecto] = { proyecto, presupuestado: 0, ejecutado: 0, pagado: 0 };
        map[proyecto].ejecutado += (Number(t.importe) || 0) * fraccion;
        if (t.status === "Pagado") map[proyecto].pagado += (Number(t.importe) || 0) * fraccion;
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
  const totalPagado = transFiltradasMes.filter((t) => t.status === "Pagado").reduce((s, t) => s + (Number(t.importe) || 0), 0);

  // Filtro de Proyecto para los cuadros de montos — usa los importes YA prorrateados
  // (porProyecto), así que si una partida es "Desh Gral" y filtras por "Desh Marfo",
  // sí cuenta su parte correspondiente.
  const [proyectoKpi, setProyectoKpi] = useSessionState("ss-dashboard-proyecto", "Todos");
  // Si el proyecto guardado no existe en ESTA compañía (ej. veníamos de otra
  // unidad), se regresa solo a "Todos" en vez de quedarse "huérfano" en $0.00.
  useEffect(() => {
    if (proyectoKpi !== "Todos" && !proyectosUnidad.some((p) => p.nombre === proyectoKpi)) {
      setProyectoKpi("Todos");
    }
  }, [unidad, proyectosUnidad.map((p) => p.nombre).join(","), proyectoKpi]);
  const proyectoKpiData = proyectoKpi === "Todos"
    ? { presupuestado: totalPresupuestado, ejecutado: totalEjecutado, pagado: totalPagado }
    : (porProyecto.find((p) => p.proyecto === proyectoKpi) || { presupuestado: 0, ejecutado: 0, pagado: 0 });
  const kpiOcupado = proyectoKpiData.ejecutado;
  const kpiPagado = proyectoKpiData.pagado;
  const kpiPorPagar = kpiOcupado - kpiPagado;
  const kpiDisponible = proyectoKpiData.presupuestado - kpiOcupado;

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
        title="Resumen general"
        subtitle="Los importes ya vienen prorrateados según el marcador de cada partida"
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
          </div>
        }
      >
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          <KpiCard label="Presupuestado" value={money(proyectoKpiData.presupuestado)} />
          <KpiCard label="Ocupado" value={money(kpiOcupado)} accent={proyectoKpiData.presupuestado && kpiOcupado / proyectoKpiData.presupuestado > 1 ? T.red : T.amber} />
          <KpiCard label="Pagado" value={money(kpiPagado)} accent={T.teal} />
          <KpiCard label="Por Pagar" value={money(kpiPorPagar)} accent={T.amber} />
          <KpiCard label="Disponible" value={money(kpiDisponible)} accent={kpiDisponible < 0 ? T.red : T.text} />
        </div>
      </Panel>

      <ResumenPivotPanel partidasUnidad={partidasFiltradasMes} proyectosUnidad={proyectosUnidad} />

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

    </div>
  );
}

function ResumenPivotPanel({ partidasUnidad, proyectosUnidad }) {
  const partidasMXN = partidasUnidad.filter((p) => (p.moneda || "MXP") === "MXP");
  const meses = MESES.filter((m) => partidasMXN.some((p) => p.mes === m));
  const [collapsed, setCollapsed] = useState(new Set());
  const toggle = (path) => setCollapsed((prev) => {
    const next = new Set(prev);
    next.has(path) ? next.delete(path) : next.add(path);
    return next;
  });

  if (!partidasMXN.length) return null;

  // "Todos" / "<Grupo> Gral" son marcadores de prorrateo, no proyectos reales —
  // se reparten entre los proyectos reales antes de armar la tabla, para que
  // nunca aparezcan como su propia fila.
  const partidasResueltas = partidasMXN.flatMap((p) =>
    resolverProrrateo(p.proyecto, proyectosUnidad).map(({ proyecto, fraccion }) => ({
      ...p, proyecto, monto_estimado: (Number(p.monto_estimado) || 0) * fraccion,
    }))
  );

  const pivot = pivotearPorMes(partidasResueltas, ["proyecto", "rubro", "concepto"], meses);
  const totalGeneral = meses.reduce((acc, m) => { acc[m] = 0; return acc; }, {});
  let granTotal = 0;
  partidasMXN.forEach((p) => {
    const v = Number(p.monto_estimado) || 0;
    if (totalGeneral[p.mes] !== undefined) totalGeneral[p.mes] += v;
    granTotal += v;
  });

  return (
    <Panel title="Resumen presupuestado por proyecto y rubro" subtitle="Solo montos en MXP — clic en una fila para expandir/colapsar">
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
function ColumnVisibilityControl({ columns, hidden, onToggle, onShowAll }) {
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
        Columnas{ocultas ? ` (${columns.length - ocultas}/${columns.length})` : ""}
      </Button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 50,
          background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8,
          padding: 14, minWidth: 220, maxHeight: 360, overflowY: "auto",
          boxShadow: "0 8px 24px rgba(35,42,49,0.14)",
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: T.text }}>Columnas visibles</span>
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
              <input type="checkbox" checked={!hidden.has(c.key)} onChange={() => onToggle(c.key)} />
              {c.label}
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

function PartidasTab({ unidad, unidades, partidas, partidasApi, perfilesApi, transacciones }) {
  const proyectosUnidad = unidades[unidad]?.proyectos || [];
  const marcadores = marcadoresDisponibles(proyectosUnidad);
  const anioDefault = (() => {
    const anios = partidas.filter((p) => p.unidad === unidad).map((p) => p.anio).filter(Boolean);
    return anios.length ? Math.max(...anios) : new Date().getFullYear();
  })();
  const blank = { unidad, mes: "Agosto", anio: anioDefault, smi: "", concepto: "", rubro: RUBROS[0].rubro, categoria: RUBROS[0].categorias[0], proyecto: marcadores[0] || "", monto_estimado: "", moneda: "MXP", folio: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);

  const categoriasDisponibles = RUBROS.find((r) => r.rubro === form.rubro)?.categorias || [];
  const partidasUnidad = partidas.filter((p) => p.unidad === unidad);

  const [filtros, setFiltros] = useSessionState("ss-partidas-filtros", { texto: "", mes: [], rubro: "Todos", proyecto: "Todos" });
  const filtrosMes = Array.isArray(filtros.mes) ? filtros.mes : [];
  const rubrosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.rubro).filter(Boolean))].sort();
  const proyectosDisponiblesFiltro = [...new Set(partidasUnidad.map((p) => p.proyecto).filter(Boolean))].sort();
  const mesesDisponiblesFiltro = MESES.filter((m) => partidasUnidad.some((p) => p.mes === m));

  const partidasFiltradas = partidasUnidad.filter((p) => {
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const enTexto = [p.concepto, p.folio, p.smi, p.categoria].some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    if (filtrosMes.length && !filtrosMes.includes(p.mes)) return false;
    if (filtros.rubro !== "Todos" && p.rubro !== filtros.rubro) return false;
    if (filtros.proyecto !== "Todos" && p.proyecto !== filtros.proyecto) return false;
    return true;
  });
  const filtrosActivos = filtros.texto.trim() || filtrosMes.length > 0 || filtros.rubro !== "Todos" || filtros.proyecto !== "Todos";
  const limpiarFiltros = () => setFiltros({ texto: "", mes: [], rubro: "Todos", proyecto: "Todos" });

  const [sort, setSort] = useSessionState("ss-partidas-sort", { key: null, dir: "asc" });
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

  const COLUMNAS_PARTIDA = [
    { key: "mes", label: "Mes", render: (p) => p.mes },
    { key: "anio", label: "Año", render: (p) => p.anio },
    { key: "concepto", label: "Concepto", render: (p) => p.concepto },
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
  const toggleExpand = (id) => setExpandedIds((prev) => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });

  const renderRowTr = (p, depth = 0) => {
    const expandido = expandedIds.has(p.id);
    const transDeEsta = transacciones.filter((t) => t.partida_id === p.id);
    return (
      <React.Fragment key={p.id}>
        <tr>
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
            <div style={{ display: "flex", gap: 6 }}>
              <Button variant="ghost" onClick={() => startEdit(p)}>Editar</Button>
              <Button variant="danger" onClick={() => remove(p.id)}>Eliminar</Button>
            </div>
          </td>
        </tr>
        {expandido && (
          <tr>
            <td colSpan={columnasVisibles.length + 2} style={{ padding: "0 0 0 40px", background: T.panelAlt, borderBottom: `1px solid ${T.border}` }}>
              <table style={{ ...tableStyle, margin: "8px 0" }}>
                <thead>
                  <tr>{["Concepto","Monto","Status"].map((h) => <th key={h} style={{ ...thStyle, background: "transparent" }}>{h}</th>)}</tr>
                </thead>
                <tbody>
                  {transDeEsta.map((t) => (
                    <tr key={t.id}>
                      <td style={{ ...tdStyle, color: T.textDim }}>{t.concepto_detallado || "—"}</td>
                      <td style={{ ...tdStyle, fontFamily: T.fontMono }}>{money(t.importe, t.moneda)}</td>
                      <td style={tdStyle}>{t.status ? <Pill tone={t.status === "Pagado" ? "teal" : "amber"}>{t.status}</Pill> : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
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
  const startEdit = (p) => { setForm(p); setEditId(p.id); setModalOpen(true); };
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
            <MesMultiSelect mesesDisponibles={mesesDisponiblesFiltro} seleccionados={filtrosMes} onChange={(nuevo) => setFiltros({ ...filtros, mes: nuevo })} />
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
              <col style={{ width: 30 }} />
              {columnasVisibles.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
              <col style={{ width: 140 }} />
            </colgroup>
            <thead>
              <tr>
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
                ? buildGroupedTrs(grouped, "", collapsedGroups, toggleGroup, columnasVisibles.length + 2, 0, renderRowTr, Object.fromEntries(GROUP_OPCIONES.map((o) => [o.value, o.label])))
                : partidasOrdenadas.map((p) => renderRowTr(p))}
              {!partidasUnidad.length && (
                <tr><td colSpan={columnasVisibles.length + 2} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin partidas aún</td></tr>
              )}
              {partidasUnidad.length > 0 && !partidasFiltradas.length && (
                <tr><td colSpan={columnasVisibles.length + 2} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna partida coincide con estos filtros</td></tr>
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
    folio_compra_sae: "", folio_factura: "", forma_pago: "", metodo_pago: "",
  };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);
  const cuentasDelProveedorSeleccionado = form.proveedor_id ? cuentasApi.rows.filter((c) => c.proveedor_id === form.proveedor_id) : [];
  const transUnidad = transacciones.filter((t) => partidasUnidad.some((p) => p.id === t.partida_id));
  const sinVincular = transacciones.filter((t) => !t.partida_id && t.unidad_detectada === unidad);

  const [filtros, setFiltros] = useSessionState("ss-transacciones-filtros", { texto: "", fechaDesde: "", fechaHasta: "", reportado: "Todos" });
  const [sort, setSort] = useSessionState("ss-transacciones-sort", { key: "dia", dir: "desc" });
  const [seleccionadas, setSeleccionadas] = useState(new Set());
  const [marcandoReportado, setMarcandoReportado] = useState(false);

  const partidaDe = (t) => partidasUnidad.find((p) => p.id === t.partida_id);

  const transFiltradas = transUnidad.filter((t) => {
    if (filtros.fechaDesde && (!t.dia || t.dia < filtros.fechaDesde)) return false;
    if (filtros.fechaHasta && (!t.dia || t.dia > filtros.fechaHasta)) return false;
    if (filtros.reportado === "Reportado" && !t.reportado_at) return false;
    if (filtros.reportado === "No reportado" && t.reportado_at) return false;
    if (filtros.texto.trim()) {
      const q = filtros.texto.trim().toLowerCase();
      const partida = partidaDe(t);
      const enTexto = [t.proveedor, t.concepto_detallado, t.solicitante, t.zona, t.folio_transaccion, partida?.folio, partida?.concepto]
        .some((v) => (v || "").toLowerCase().includes(q));
      if (!enTexto) return false;
    }
    return true;
  });
  const filtrosActivos = filtros.texto.trim() || filtros.fechaDesde || filtros.fechaHasta || filtros.reportado !== "Todos";
  const limpiarFiltros = () => setFiltros({ texto: "", fechaDesde: "", fechaHasta: "", reportado: "Todos" });

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
    { value: "zona", label: "Zona" },
    { value: "area", label: "Área" },
    { value: "proveedor", label: "Proveedor" },
    { value: "proyecto", label: "Proyecto" },
    { value: "status", label: "Status" },
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
      key: "reportado_at", label: "Reportado",
      render: (t) => t.reportado_at
        ? <Pill tone="accent">{formatFechaHora(t.reportado_at)}</Pill>
        : <Pill tone="amber">No reportado</Pill>,
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
  const renderRowTr = (t, depth = 0) => (
    <tr key={t.id}>
      <td style={{ ...tdStyle, textAlign: "center" }}>
        <input type="checkbox" checked={seleccionadas.has(t.id)} onChange={() => toggleSeleccion(t.id)} />
      </td>
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
        const mesForm = MESES[(form.dia ? new Date(`${form.dia}T00:00:00`) : new Date()).getMonth()];
        const creada = await transaccionesApi.insert({ ...rest, id: uid(), folio_transaccion: nextFolioTransaccion(unidad, mesForm, transUnidad) });
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
          <Field label="Reportado">
            <Select value={filtros.reportado} onChange={(e) => setFiltros({ ...filtros, reportado: e.target.value })} style={{ width: 150 }}>
              <option>Todos</option>
              <option>Reportado</option>
              <option>No reportado</option>
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
          <ColumnVisibilityControl
            columns={COLUMNAS_TRANS}
            hidden={colVisibility.hidden}
            onToggle={colVisibility.toggle}
            onShowAll={colVisibility.showAll}
          />
        </div>

        {seleccionadas.size > 0 && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, padding: "10px 14px", background: T.accentBg, border: `1px solid ${T.accent}55`, borderRadius: 6 }}>
            <span style={{ fontSize: 12.5, color: T.text }}>{seleccionadas.size} seleccionada(s)</span>
            <Button onClick={() => marcarReportadas(true)} disabled={marcandoReportado}>
              {marcandoReportado ? "Marcando…" : "Marcar como reportadas"}
            </Button>
            <Button variant="ghost" onClick={() => marcarReportadas(false)} disabled={marcandoReportado}>Quitar marca de reportado</Button>
            <Button variant="ghost" onClick={() => setSeleccionadas(new Set())}>Cancelar selección</Button>
          </div>
        )}

        <div style={{ overflowX: "auto" }}>
          <table style={{ ...tableStyle, tableLayout: "fixed" }}>
            <colgroup>
              <col style={{ width: 28 }} />
              {columnasVisibles.map((c) => <col key={c.key} style={{ width: colWidths.getWidth(c.key) }} />)}
              <col style={{ width: 140 }} />
            </colgroup>
            <thead>
              <tr>
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
                ? buildGroupedTrs(grouped, "", collapsedGroups, toggleGroup, columnasVisibles.length + 2, 0, renderRowTr, Object.fromEntries(GROUP_OPCIONES_TRANS.map((o) => [o.value, o.label])))
                : transEnriquecidas.map((t) => renderRowTr(t))}
              {!transUnidad.length && (
                <tr><td colSpan={columnasVisibles.length + 2} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Sin transacciones aún</td></tr>
              )}
              {transUnidad.length > 0 && !transFiltradas.length && (
                <tr><td colSpan={columnasVisibles.length + 2} style={{ ...tdStyle, textAlign: "center", color: T.textFaint }}>Ninguna transacción coincide con estos filtros</td></tr>
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

      <ImportarTransaccionesPanel partidas={partidas} proveedores={proveedoresApi.rows} transaccionesApi={transaccionesApi} />

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
  { key: "referencia", label: "Referencia" },
  { key: "concepto", label: "Concepto de pago" },
  { key: "banco", label: "Banco" },
  { key: "clabe", label: "Cuenta CLABE" },
  { key: "numero_cuenta", label: "No. Cuenta" },
  { key: "swift", label: "SWIFT" },
  { key: "importe", label: "Importe" },
  { key: "moneda", label: "Moneda" },
  { key: "notas", label: "Notas" },
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
      zona: t.zona || "",
      numero_solicitud: t.smi || partida?.smi || "",
      no_sae: proveedor?.id_sae || "",
      folio_compra_sae: t.folio_compra_sae || "",
      folio_factura: t.folio_factura || "",
      forma_pago: FORMAS_PAGO.find((f) => f.value === t.forma_pago)?.label || t.forma_pago || "",
      metodo_pago: METODOS_PAGO.find((m) => m.value === t.metodo_pago)?.label || t.metodo_pago || "",
      proveedor: proveedor?.nombre || t.proveedor || "",
      referencia: proveedor?.referencia || "",
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
    return [f.solicitante, f.proveedor, f.concepto, f.folio_factura, f.folio_compra_sae, f.no_sae]
      .some((v) => (v || "").toString().toLowerCase().includes(q));
  });
  const filtrosFechaActivos = fechaDesde || fechaHasta;
  const limpiarFechas = () => { setFechaDesde(""); setFechaHasta(""); };
  const filasOrdenadas = sortRows(filasFiltradas, sort, { importe: (r) => r.importe });

  const colVisibility = useColumnVisibility("colv-reporte", COLUMNAS_REPORTE);
  const colWidths = useColumnWidths("colw-reporte");
  const { ordered: columnas, moveColumn } = useColumnOrder("colo-reporte", colVisibility.visible);
  const dragKeyRef = useRef(null);
  const onColDragStart = (e, key) => { dragKeyRef.current = key; e.dataTransfer.effectAllowed = "move"; };
  const onColDragOver = (e) => e.preventDefault();
  const onColDrop = (e, targetKey) => { e.preventDefault(); if (dragKeyRef.current) { moveColumn(dragKeyRef.current, targetKey); dragKeyRef.current = null; } };

  const exportarExcel = async () => {
    const wbx = new ExcelJS.Workbook();
    const ws = wbx.addWorksheet("Reporte de pagos");
    ws.columns = [
      { width: 8.875 }, { width: 10 }, { width: 8.125 }, { width: 10.125 }, { width: 10.125 },
      { width: 8.625 }, { width: 8.25 }, { width: 15 }, { width: 12 }, { width: 17.625 },
      { width: 11.25 }, { width: 13.625 }, { width: 15.5 }, { width: 18 }, { width: 12.5 },
      { width: 9 }, { width: 10.5 }, { width: 10.5 }, { width: 8.875 },
    ];

    const diasOrdenados = filasOrdenadas.map((f) => f.dia).filter(Boolean).sort();
    const inicio = fechaDesde || diasOrdenados[0] || "";
    const fin = fechaHasta || diasOrdenados[diasOrdenados.length - 1] || "";

    const zonas = [...new Set(filasOrdenadas.map((f) => f.zona).filter(Boolean))].sort();
    const ordenMoneda = (m) => (m === "MXP" ? 0 : m === "USD" ? 1 : 2);
    const headers = [
      "Día", "Solicitante", "Área", "No. Solicitud (SMI)", "No. SAE", "Folio Compra SAE", "Folio Factura",
      "Forma de Pago", "Método de Pago", "Proveedor", "Referencia", "Concepto de pago", "Banco",
      "Cuenta CLABE", "No. Cuenta", "SWIFT", "Importe", "Moneda", "Notas",
    ];

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
        headers.forEach((h, i) => {
          const cell = headerRow.getCell(i + 1);
          cell.value = h;
          cell.alignment = { horizontal: "center" };
          cell.font = { name: "Calibri", size: 11 };
        });
        fila += 1;

        filasGrupo.forEach((f) => {
          const row = ws.getRow(fila);
          const valores = [
            f.dia, f.solicitante, f.area, f.numero_solicitud, f.no_sae, f.folio_compra_sae, f.folio_factura,
            f.forma_pago, f.metodo_pago, f.proveedor, f.referencia, f.concepto, f.banco,
            f.clabe, f.numero_cuenta, f.swift, f.importe, f.moneda, f.notas,
          ];
          valores.forEach((v, ci) => {
            const cell = row.getCell(ci + 1);
            cell.value = v;
            cell.alignment = { horizontal: "center" };
            cell.font = { name: "Calibri", size: 11 };
            if (ci === 16) cell.numFmt = '"$"#,##0.00'; // columna Importe
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
            <Button onClick={exportarExcel}>Exportar a Excel</Button>
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

export default function App() {
  const { session, recovery, clearRecovery } = useAuth();
  const proyectosApi = useCollection("proyectos");
  const partidasApi = useCollection("partidas", "created_at", { withAudit: true });
  const transaccionesApi = useCollection("transacciones", "created_at", { withAudit: true });
  const proveedoresApi = useCollection("proveedores", "created_at", { withAudit: true });
  const cuentasApi = useCollection("proveedor_cuentas", "created_at", { withAudit: true });
  const perfilesApi = useCollection("perfiles");
  const notasApi = useCollection("transaccion_notas");
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
          {tab === "partidas" && <PartidasTab unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} perfilesApi={perfilesApi} transacciones={transacciones} />}
          {tab === "transacciones" && <TransaccionesTab unidad={unidad} unidades={unidades} partidas={partidas} partidasApi={partidasApi} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} notasApi={notasApi} session={session} />}
          {tab === "reporte" && <ReportePagosTab unidad={unidad} partidas={partidas} transacciones={transacciones} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} />}
          {tab === "reporte-direccion" && <ReportePagosDireccionTab unidad={unidad} partidas={partidas} transacciones={transacciones} transaccionesApi={transaccionesApi} proveedoresApi={proveedoresApi} />}
          {tab === "catalogo" && <CatalogoTab unidad={unidad} unidades={unidades} proyectosApi={proyectosApi} proveedoresApi={proveedoresApi} cuentasApi={cuentasApi} perfilesApi={perfilesApi} />}
        </>
      )}
    </div>
  );
}
