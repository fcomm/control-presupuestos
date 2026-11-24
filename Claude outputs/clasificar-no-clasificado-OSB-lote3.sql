-- =====================================================================
-- Clasificar No Clasificado de OSB — lote 3 (24 nov 2026)
-- 219 transacciones de 96 proveedores (MXP $1,878,553.30), según la tabla
-- que confirmó Paco (hospedaje, avión, combustible, casetas, mantenimiento
-- vehicular, TI, auditorías CLA, honorarios, mensajería, materiales, etc.).
-- Crea la categoría Materiales e Insumos › Material electrónico.
-- Cada transacción pasa a la partida GENERAL ("Todos") de su rubro + mes +
-- año + moneda (74 existentes, 25 nuevas) y conserva su proyecto.
-- Presupuesto: sale de la partida de origen solo lo que le queda; si ya estaba
-- sobregirada no queda en negativo (el sobregiro, $38,289.91, viaja con las
-- transacciones). El total de OSB no cambia. Todo o nada.
-- Quedan 65 en No Clasificado (los proveedores por definir).
-- Probado contra la base con reversa automática.
-- =====================================================================

-- Categoría nueva: Materiales e Insumos › Material electrónico
insert into public.categorias (rubro_id, nombre, orden)
select '586f08e6-a75d-4260-9b79-417f3c790228', 'Material electrónico',
 coalesce((select max(orden) from public.categorias where rubro_id = '586f08e6-a75d-4260-9b79-417f3c790228'), 0) + 1
where not exists (select 1 from public.categorias where rubro_id = '586f08e6-a75d-4260-9b79-417f3c790228' and nombre = 'Material electrónico');

create temp table _mov as
with mapa(proveedor, rubro, categoria) as (values
('PROMOTORA TURISTICA DE OAXACA SA DE CV','Viajes y Viáticos','Hospedaje'),
('ADMINISTRADORA DE HOTELES GRT','Viajes y Viáticos','Hospedaje'),
('INMOBILIARIA HOTELERA EL PRESIDENTE SAN JOSE DEL CABO','Viajes y Viáticos','Hospedaje'),
('ACCOR MEXICO HOTELS SA DE CV','Viajes y Viáticos','Hospedaje'),
('HOTELERA AZTLAN DE TAMPICO PROLONGACION AVE HIDALGO','Viajes y Viáticos','Hospedaje'),
('IMPULSORA DE EMPRESAS TURISTICAS','Viajes y Viáticos','Hospedaje'),
('CONSORCIO GRUPO HOTELERO T2','Viajes y Viáticos','Hospedaje'),
('HOTELES MILEX','Viajes y Viáticos','Hospedaje'),
('EXPLOITANT HOTELIER','Viajes y Viáticos','Hospedaje'),
('OPERACION Y CONTROL DE HOTELES','Viajes y Viáticos','Hospedaje'),
('INVERSIONES HOTELERAS DEL NORTE SA DE CV','Viajes y Viáticos','Hospedaje'),
('PROMOTORA HOTELERA DE VERACRUZ','Viajes y Viáticos','Hospedaje'),
('OPERADORA DE HOTELES HMR','Viajes y Viáticos','Hospedaje'),
('OPERADORA DECORI ORIZABA SA DE CV','Viajes y Viáticos','Hospedaje'),
('CALIDAD DE TABASCO S DE RL DE CV','Viajes y Viáticos','Hospedaje'),
('AEROENLACES NACIONALES SA DE CV','Viajes y Viáticos','Boletos de avión'),
('CONCESIONARIA VUELA COMPAÍA DE AVIACION SAPI DE CV','Viajes y Viáticos','Boletos de avión'),
('FLY BY WINGS SA DE CV','Viajes y Viáticos','Boletos de avión'),
('LINK CONEXIÓN AEREA SOCIEDAD ANONIMA DE CAPITAL VARIABLE','Viajes y Viáticos','Boletos de avión'),
('AUTOBUSES DE LA PIEDAD SA DE CV','Viajes y Viáticos','Transporte terrestre'),
('AUTOBUSES POBLANOS DEL CENTRO','Viajes y Viáticos','Transporte terrestre'),
('GREMIO DE TAXISTAS PARA EL TRANSPORTE TERRESTRE DE PASAJEROS SITIO 300','Viajes y Viáticos','Transporte terrestre'),
('GASTOS NO DEDUCIBLES','Viajes y Viáticos','Alimentación'),
('OPERADORA LA UNIQUITA','Viajes y Viáticos','Alimentación'),
('AEROPUERTO INTERNACIONAL DE LA CIUDAD DE MEXICO','Viajes y Viáticos','Viáticos diversos'),
('GPDC ESTACIONES DE SERVICIO SA DE CV','Vehículos','Combustible'),
('ESTACIONES DE SERVICIO SA DE CV','Vehículos','Combustible'),
('LAS MEJORES ESTACIONES SA DE CV','Vehículos','Combustible'),
('SUPERSERVICIO CARRETERO','Vehículos','Combustible'),
('SERVICIO GRUPO TEPEYAC','Vehículos','Combustible'),
('NOVOGAS JURIQUILLA BLVD. DE LAS CIENCIAS','Vehículos','Combustible'),
('PASE, SERVICIOS ELECTRONICOS LAGO ZURICH','Vehículos','Casetas'),
('PASE, SERVICIOS ELECTRONICOS','Vehículos','Casetas'),
('FONDO NACIONAL DE INFRAESTRUCTURA','Vehículos','Casetas'),
('FAME CORREGIDORA SA DE CV','Vehículos','Mantenimiento vehicular'),
('CAMINO DEL SOL AUTOMOTRIZ DE TIJUANA SRL DE CV','Vehículos','Mantenimiento vehicular'),
('NISSAUTO SONORA SA DE CV','Vehículos','Mantenimiento vehicular'),
('AUTOS ALAMEDA','Vehículos','Mantenimiento vehicular'),
('AUTOMOTORES DE LA LAGUNA','Vehículos','Mantenimiento vehicular'),
('BAQUI AUTOMOTRIZ S DE RL DE CV','Vehículos','Mantenimiento vehicular'),
('AUTOMOTRIZ HF','Vehículos','Mantenimiento vehicular'),
('GEISHA QUERETARO (ALDEN QUERETARO)','Vehículos','Mantenimiento vehicular'),
('MANTENIMIENTOS INDUSTRIALES LIZ-LEY','Vehículos','Mantenimiento vehicular'),
('DAYTONA MOTOS','Vehículos','Mantenimiento vehicular'),
('GRUPO LLANTERO DE POZA RICA','Vehículos','Refacciones'),
('LLANTAS VEGA','Vehículos','Refacciones'),
('ACUMULADORES Y SERVICIO QUERETARO SA DE CV','Vehículos','Refacciones'),
('ACCESORIOS Y REFACCIONES RAMOS','Vehículos','Refacciones'),
('VERIFIMETAB SA DE CV','Vehículos','Verificación'),
('GOOGLE INC','Tecnologías de Información','Suscripciones'),
('BT CLICKUP','Tecnologías de Información','Suscripciones'),
('MICROSOFT','Tecnologías de Información','Suscripciones'),
('ANTHROPIC,PBC','Tecnologías de Información','Suscripciones'),
('AMAZON WEB SERVICES, INC.','Tecnologías de Información','Servicios en la nube'),
('NEOTHEK SERVICOS WEB WEB HOSTING, SSL Y DOMINIOS','Tecnologías de Información','Hosting'),
('VITALWERKS INTERNET SOLUTIONS,LLC','Tecnologías de Información','Hosting'),
('ASPEL DE MEXICO SA DE CV','Tecnologías de Información','Licencias de Software'),
('CLA MX','Cumplimiento Legal','Auditorías'),
('CLA MX PS','Cumplimiento Legal','Auditorías'),
('JM REBOLLAR ABOGADOS SC','Servicios Administrativos','Honorarios profesionales'),
('JM REBOLLAR ABOGADOS SC FRAY JOSE DE LA CORUÑA','Servicios Administrativos','Honorarios profesionales'),
('V&CR ABOGADOS Y ASOCIADOS AV. 35 PONIENTE','Servicios Administrativos','Honorarios profesionales'),
('V&CR ABOGADOS Y ASOCIADOS','Servicios Administrativos','Honorarios profesionales'),
('ANA LILIA RODRIGUEZ GUERRERO','Servicios Administrativos','Honorarios profesionales'),
('ANA KAREN GARCIA POLO','Servicios Administrativos','Honorarios profesionales'),
('RRGL TRADUCCIONES SC PRIMERA DE CEDROS','Servicios Administrativos','Honorarios profesionales'),
('ZEPEDA STERLING SC AVENIDA LIC. MANUEL GÓMEZ MORIN 442 644 5054','Gastos Financieros e Impuestos','Gastos notariales'),
('ZEPEDA STERLING SC','Gastos Financieros e Impuestos','Gastos notariales'),
('BANCA MIFEL SA INSTITUCION DE BANCA MULTIPLE GRUPO FINANCIERO MIFEL','Gastos Financieros e Impuestos','Comisiones bancarias'),
('DHL EXPRESS MEXICO SA DE CV','Servicios Administrativos','Mensajería'),
('INTEGRALOGIS SA DE CV','Servicios Administrativos','Mensajería'),
('PACKSYS DISTRIBUCION','Servicios Administrativos','Mensajería'),
('INDEED','Servicios Administrativos','Reclutamiento'),
('PEM INSTITUTE','Servicios Administrativos','Capacitación'),
('AULACI','Servicios Administrativos','Capacitación'),
('AFISCA EDITORES','Servicios Administrativos','Capacitación'),
('RENTCO','Servicios Operativos','Arrendamientos'),
('CALIBRACIONES Y PRUEBAS EDALME','Servicios de Mantenimiento','Calibración'),
('INTERMEXICANA DE EXTINGUIDORES Y SEGURIDAD SA DE CV','Servicios de Mantenimiento','Mantenimiento de equipos'),
('SISTEMAS INTEGRALES CONTRA INCENDIOS Y SEGURIDAD INDUSTRIAL','Servicios de Mantenimiento','Mantenimiento de equipos'),
('ELEMENT14 S DE RL DE CV','Materiales e Insumos','Material electrónico'),
('WAVESHARE ELECTRONICS','Materiales e Insumos','Material electrónico'),
('CALVEK COMPONENTES BAJIO SA D ECV','Materiales e Insumos','Material electrónico'),
('ELECTRONICA STEREN','Materiales e Insumos','Material electrónico'),
('QILINSHOP','Materiales e Insumos','Material electrónico'),
('MANOMETROS DE JALISCO','Materiales e Insumos','Instrumentación'),
('360 TECNOLOGIA INDUSTRIAL','Materiales e Insumos','Instrumentación'),
('COFLI TUBE FITTINGS MEXICO SA DE CV','Materiales e Insumos','Instrumentación'),
('LA PALOMA COMPANIA DE METALES SA DE CV','Materiales e Insumos','Material mecánico'),
('ACEROS LEVINSON SA DE CV','Materiales e Insumos','Material mecánico'),
('ARIS SEAL DE MEXICO SA DE CV','Materiales e Insumos','Material mecánico'),
('MRO MARKET SA DE CV','Materiales e Insumos','Ferretería'),
('MATERIALES Y HERRAMIENTAS INDUSTRIALES DEL PUERTO','Materiales e Insumos','Ferretería'),
('SAFETY OFFSHORE OF MEXICO SA DE CV','Materiales e Insumos','EPP'),
('NATIONAL INSTRUMENTS CORP','Equipos y Herramientas','Equipos de medición'),
('SMART CUTTING TOOLS SA DE CV','Equipos y Herramientas','Herramienta especializada')
)
select t.id as tid, p.unidad as unidad, t.folio_transaccion, t.proveedor, t.importe, t.es_padre,
 m.rubro, m.categoria, p.id as origen_id, p.folio as origen_folio, p.rubro as rubro_origen,
 p.mes, p.anio, p.proyecto, p.moneda,
 case when p.rubro = m.rubro then 'categoria' else 'mover' end as accion
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.cancelada_en is null and t.transaccion_padre_id is null
 and p.unidad = 'OSB'
 and p.rubro ilike 'No Clasificado%'
 and not (p.rubro = m.rubro and coalesce(t.categoria, '') = m.categoria);

-- Lo que sale de cada partida de origen. Si la partida ya estaba sobregirada,
-- solo se mueve el presupuesto que le queda (nunca queda en negativo); el
-- sobregiro viaja con las transacciones a la partida de su rubro.
create temp table _g as
select m.origen_id, round(sum(m.importe), 2) as gasto,
 least(round(sum(m.importe), 2), greatest(p.monto_estimado, 0)) as monto
from _mov m join public.partidas p on p.id = m.origen_id
where m.accion = 'mover' group by m.origen_id, p.monto_estimado;

create temp table _od as -- presupuesto que pasa de cada origen a cada destino, en proporción al gasto
select m.origen_id, m.unidad, m.mes, m.anio, m.moneda, m.rubro, count(*) as n,
 sum(m.importe) * case when g.gasto = 0 then 0 else g.monto / g.gasto end as monto
from _mov m join _g g on g.origen_id = m.origen_id
where m.accion = 'mover' group by 1,2,3,4,5,6, g.gasto, g.monto;

create temp table _dest as -- un destino GENERAL por compañía + rubro + mes + año + moneda (partida sin proyecto)
select d.*, d.unidad || '-P-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], d.mes)] || right(d.anio::text, 2) || '-' as prefijo,
 (select p.id from public.partidas p where p.unidad = d.unidad and p.rubro = d.rubro and p.mes = d.mes
 and p.anio = d.anio and p.moneda = d.moneda and coalesce(p.proyecto, '') in ('Todos', '')
 order by p.folio limit 1) as existente_id
from (select unidad, mes, anio, moneda, rubro, sum(n) as n, round(sum(monto), 2) as monto
 from _od group by 1,2,3,4,5) d;

create temp table _nuevas as
select gen_random_uuid() as nid, d.*,
 d.prefijo || lpad((coalesce((select max(substring(p.folio from '-(\d+)$')::int) from public.partidas p
 where p.unidad = d.unidad and p.folio like d.prefijo || '%'), 0)
 + row_number() over (partition by d.prefijo order by d.rubro))::text, 3, '0') as folio_nuevo,
 d.rubro || ' ' || d.mes || ' ' || d.anio as concepto_nuevo
from _dest d where d.existente_id is null;

create temp table _final as
select d.unidad, d.mes, d.anio, d.moneda, d.rubro, d.n, d.monto,
 coalesce(d.existente_id, nv.nid) as destino_id,
 case when d.existente_id is null then 'nueva' else 'existente' end as tipo
from _dest d left join _nuevas nv
 on (nv.unidad, nv.mes, nv.anio, nv.moneda, nv.rubro) = (d.unidad, d.mes, d.anio, d.moneda, d.rubro);
do $$
declare n_mov int; n_upd int; t_osb numeric;
begin
 if exists (select 1 from public.partidas where folio !~ '^[A-Z]+-P-') then
  raise exception 'Corre primero renombrar-folios-partidas.sql. No se cambió nada.'; end if;
 select count(*) into n_mov from _mov;
 if n_mov <> 219 then raise exception 'Se esperaban 219 transacciones y hay %. No se cambió nada.', n_mov; end if;
 select round(sum(monto_estimado) filter (where unidad='OSB'),2) into t_osb from public.partidas;
 

 insert into public.partidas (id, unidad, mes, anio, smi, concepto, rubro, categoria, proyecto, monto_estimado, moneda, folio, created_by, extra)
 select nid, unidad, mes, anio, '', concepto_nuevo, rubro, null, 'Todos', monto, moneda, folio_nuevo, 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
 jsonb_build_object('origen', 'Clasificación No Clasificado OSB nov 2026 (lote 3)')
 from _nuevas;

 -- Las que se mueven de partida (y quedan con su categoría)
 update public.transacciones t
 set partida_id = f.destino_id, categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _mov m join _final f
 on (f.unidad, f.mes, f.anio, f.moneda, f.rubro) = (m.unidad, m.mes, m.anio, m.moneda, m.rubro)
 where m.accion = 'mover' and t.id = m.tid;
 get diagnostics n_upd = row_count;

 -- Las que ya están en una partida del rubro correcto: solo la categoría
 update public.transacciones t set categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _mov m where m.accion = 'categoria' and t.id = m.tid;
 n_upd := n_upd + (select count(*) from _mov where accion = 'categoria');
 if n_upd <> n_mov then raise exception 'desfase % vs %', n_upd, n_mov; end if;

 -- Parcialidades de compras PPD: la base les cambia la partida con su compra; falta la categoría
 update public.transacciones h set categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _mov m where m.es_padre and h.transaccion_padre_id = m.tid;

 -- Presupuesto: el monto movido sale de la partida de origen y entra a la existente que lo recibe
 update public.partidas p set monto_estimado = p.monto_estimado - g.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _g g where p.id = g.origen_id;
 update public.partidas p set monto_estimado = p.monto_estimado + f.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _final f where f.tipo = 'existente' and p.id = f.destino_id;
 if (select round(sum(monto_estimado),2) from public.partidas where unidad='OSB') <> t_osb
 then raise exception 'El presupuesto total de OSB cambió. No se cambió nada.'; end if;
 raise notice 'Listo: % transacciones clasificadas', n_mov;
end $$;

