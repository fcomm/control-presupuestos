-- =====================================================================
-- Clasificar los No Clasificado de OSB Sempra (24 nov 2026)
-- 44 transacciones (MXP $771,984.80 + USD 2,300.98), todas pagadas:
--   Adobe, Microsoft ............... Tecnologías de Información › Suscripciones
--   Gerardo Celaya, Christian Alemán, Andrés Gutiérrez, Sanigap, Motum
--                                  ... Servicios Operativos › Arrendamientos
--   Liberty Fianzas ................ Gastos Financieros e Impuestos › Fianzas
--   Comercializadora Empresarial ... Servicios de Mantenimiento › Calibración
--   Salazar Collado ................ Servicios Operativos › Laboratorios
--   Lenovo ......................... Equipos y Herramientas › Equipos de cómputo
--   Hercon (reparación Genie) ...... Servicios de Mantenimiento › Reparaciones
--   Grupo Mexicano de Seguros ...... Gastos Financieros e Impuestos › Seguros No Vehiculares
--   Marco Antonio Sánchez (llantas)  Vehículos › Llantas
--   Ing. en Seguridad Vehicular SEI  Vehículos › Accesorios (roll bar)
-- Cada transacción pasa a la partida GENERAL ("Todos") de su rubro + mes + año
-- + moneda (18 ya existen, 6 se crean) con su monto de presupuesto; la
-- transacción conserva su proyecto Sempra. Todo o nada; el total de OSB no cambia.
-- Requiere haber corrido antes renombrar-folios-partidas.sql (folios OSB-P-...).
-- Probado contra la base con reversa automática.
-- =====================================================================

create temp table _mov as
with mapa(proveedor, rubro, categoria) as (values
('ADOBE INC','Tecnologías de Información','Suscripciones'),
('MICROSOFT','Tecnologías de Información','Suscripciones'),
('GERARDO CELAYA CRUZ','Servicios Operativos','Arrendamientos'),
('CHRISTIAN ALEMAN FLORES JUNCO','Servicios Operativos','Arrendamientos'),
('ANDRES GUTIERREZ LOPEZ','Servicios Operativos','Arrendamientos'),
('SANIGAP','Servicios Operativos','Arrendamientos'),
('MOTUM SERVICES','Servicios Operativos','Arrendamientos'),
('LIBERTY FIANZAS','Gastos Financieros e Impuestos','Fianzas'),
('COMERCIALIZADORA EMPRESARIAL Y DE SERVICIOS','Servicios de Mantenimiento','Calibración'),
('SALAZAR COLLADO PEDRO JESUS','Servicios Operativos','Laboratorios'),
('LENOVO MEXICO, S DE RL DE CV','Equipos y Herramientas','Equipos de cómputo'),
('HERCON MAQUINARIA QUERETAROSA DE CV','Servicios de Mantenimiento','Reparaciones'),
('GRUPO MEXICANO DE SEGUROS','Gastos Financieros e Impuestos','Seguros No Vehiculares'),
('MARCO ANTONIO SANCHEZ GOMEZ','Vehículos','Llantas'),
('INGENIERIA EN SEGURIDAD VEHICULAR SEI','Vehículos','Accesorios')
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
 and (t.proyecto ilike '%sempra%' or p.proyecto ilike '%sempra%')
 and (p.rubro ilike 'No Clasificado%' or coalesce(t.categoria, '') = '')
 and not (p.rubro = m.rubro and coalesce(t.categoria, '') = m.categoria);

create temp table _g as -- lo que sale de cada partida de origen (solo las que se mueven)
select origen_id, round(sum(importe), 2) as monto from _mov where accion = 'mover' group by 1;

create temp table _dest as -- un destino GENERAL por compañía + rubro + mes + año + moneda (partida sin proyecto)
select d.*, d.unidad || '-P-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], d.mes)] || right(d.anio::text, 2) || '-' as prefijo,
 (select p.id from public.partidas p where p.unidad = d.unidad and p.rubro = d.rubro and p.mes = d.mes
 and p.anio = d.anio and p.moneda = d.moneda and coalesce(p.proyecto, '') in ('Todos', '')
 order by p.folio limit 1) as existente_id
from (select unidad, mes, anio, moneda, rubro, count(*) as n, round(sum(importe), 2) as monto
 from _mov where accion = 'mover' group by 1,2,3,4,5) d;

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
 if n_mov <> 44 then raise exception 'Se esperaban 44 transacciones y hay %. No se cambió nada.', n_mov; end if;
 select round(sum(monto_estimado) filter (where unidad='OSB'),2) into t_osb from public.partidas;
 

 insert into public.partidas (id, unidad, mes, anio, smi, concepto, rubro, categoria, proyecto, monto_estimado, moneda, folio, created_by, extra)
 select nid, unidad, mes, anio, '', concepto_nuevo, rubro, null, 'Todos', monto, moneda, folio_nuevo, 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
 jsonb_build_object('origen', 'Clasificación Sempra OSB nov 2026')
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

