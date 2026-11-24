-- Reclasificar gastos "No Clasificado" de OSB por proveedor — lote 4 (24/11/26)
-- 89 transacciones · $300,619.45
-- 5 proveedores (6 variantes de nombre), según lo que definió Paco:
--   AT&T Comercialización Móvil = como Starlink · Multimóvil = alquiler de oficinas móviles
--   Pérez Muñoz = como Telmex · PITS = cargas de combustible · Ballesteros = agua en pipa para campo
--
-- El PASO 2:
--   1. Por cada rubro + proyecto + mes + año + moneda: si ya existe una partida de ese
--      rubro (p. ej. de la reclasificación anterior), la usa y le suma el presupuesto;
--      si no, crea "<Rubro> <Proyecto> <Mes> <Año>" con el siguiente folio del mes.
--   2. Mueve las transacciones y les pone su categoría (las parcialidades PPD siguen a su compra).
--   3. Resta lo movido al presupuesto de la partida "No Clasificado" de origen:
--      el total presupuestado de OSB no cambia.
--   No toca status, importes, folios de transacción ni adjuntos. Todo o nada.

------------------------------------------------------------------------
-- PASO 1 · PREVIO (solo lectura): una fila por partida destino.
------------------------------------------------------------------------
with mapa(proveedor, rubro, categoria) as (values
  ('AT&T COMERCIALIZACION MOVIL S DE RL DE CV',         'Servicios Operativos',  'Servicios de Telecomunicaciones'),
  ('MULTIMOVIL SA DE CV',                               'Servicios Operativos',  'Arrendamientos'),
  ('PEREZ MUÑOZ LUIS ALBERTO',                          'Servicios Básicos',     'Telefonía fija'),
  ('PEREZ MUÑOZ LUIS ALBERTO VATE CARRILLO 4423546470', 'Servicios Básicos',     'Telefonía fija'),
  ('PITS ABASTECIMIENTOS',                              'Vehículos',             'Combustible'),
  ('JESUS BALLESTEROS VILLA',                           'Servicios Básicos',     'Agua')
),
mov as (
select t.importe, m.rubro, p.mes, p.anio, p.proyecto, p.moneda
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.unidad_detectada = 'OSB' and t.cancelada_en is null and t.transaccion_padre_id is null
  and p.rubro = 'No Clasificado'
),
d as (
  select mes, anio, proyecto, moneda, rubro, count(*) as n, round(sum(importe), 2) as monto, 'OSB-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], mes)] || right(anio::text, 2) || '-' as prefijo
  from mov group by 1,2,3,4,5
),
e as (
  select d.*, (select p.folio || ' · ' || p.concepto from public.partidas p where p.unidad = 'OSB' and p.rubro = d.rubro and p.mes = d.mes
                 and p.anio = d.anio and p.proyecto = d.proyecto and p.moneda = d.moneda order by p.folio limit 1) as existente
  from d
)
select coalesce(e.existente,
         e.prefijo || lpad((coalesce((select max(substring(p.folio from '-(\d+)$')::int) from public.partidas p
                                      where p.unidad = 'OSB' and p.folio like e.prefijo || '%'), 0)
                            + row_number() over (partition by e.prefijo, e.existente is null order by e.proyecto, e.rubro))::text, 3, '0')
         || ' · ' || e.rubro || ' ' || e.proyecto || ' ' || e.mes || ' ' || e.anio) as partida_destino,
       case when e.existente is null then 'nueva' else 'existente' end as tipo,
       e.n as transacciones, e.monto
from e
order by e.anio, array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], e.mes), e.proyecto, e.rubro;

------------------------------------------------------------------------
-- PASO 2 · APLICAR. Quita /* y */ y ejecuta el bloque completo.
------------------------------------------------------------------------
/*
create temp table _mov as
with mapa(proveedor, rubro, categoria) as (values
  ('AT&T COMERCIALIZACION MOVIL S DE RL DE CV',         'Servicios Operativos',  'Servicios de Telecomunicaciones'),
  ('MULTIMOVIL SA DE CV',                               'Servicios Operativos',  'Arrendamientos'),
  ('PEREZ MUÑOZ LUIS ALBERTO',                          'Servicios Básicos',     'Telefonía fija'),
  ('PEREZ MUÑOZ LUIS ALBERTO VATE CARRILLO 4423546470', 'Servicios Básicos',     'Telefonía fija'),
  ('PITS ABASTECIMIENTOS',                              'Vehículos',             'Combustible'),
  ('JESUS BALLESTEROS VILLA',                           'Servicios Básicos',     'Agua')
)
select t.id as tid, t.folio_transaccion, t.proveedor, t.importe, t.es_padre, m.rubro, m.categoria,
       p.id as origen_id, p.folio as origen_folio, p.mes, p.anio, p.proyecto, p.moneda
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.unidad_detectada = 'OSB' and t.cancelada_en is null and t.transaccion_padre_id is null
  and p.rubro = 'No Clasificado';

create temp table _g as   -- lo que sale de cada partida de origen hacia cada rubro
select origen_id, mes, anio, proyecto, moneda, rubro, round(sum(importe), 2) as monto
from _mov group by 1,2,3,4,5,6;

create temp table _dest as  -- un destino por rubro + proyecto + mes + año + moneda
select d.*, 'OSB-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], mes)] || right(anio::text, 2) || '-' as prefijo,
       (select p.id from public.partidas p where p.unidad = 'OSB' and p.rubro = d.rubro and p.mes = d.mes
          and p.anio = d.anio and p.proyecto = d.proyecto and p.moneda = d.moneda order by p.folio limit 1) as existente_id,
       (select count(*) from public.partidas p where p.unidad = 'OSB' and p.rubro = d.rubro and p.mes = d.mes
          and p.anio = d.anio and p.proyecto = d.proyecto and p.moneda = d.moneda) as existentes
from (select mes, anio, proyecto, moneda, rubro, count(*) as n, round(sum(importe), 2) as monto
      from _mov group by 1,2,3,4,5) d;

create temp table _nuevas as
select gen_random_uuid() as nid, d.*,
       d.prefijo || lpad((coalesce((select max(substring(p.folio from '-(\d+)$')::int) from public.partidas p
                                    where p.unidad = 'OSB' and p.folio like d.prefijo || '%'), 0)
                          + row_number() over (partition by d.prefijo order by d.proyecto, d.rubro))::text, 3, '0') as folio_nuevo,
       d.rubro || ' ' || d.proyecto || ' ' || d.mes || ' ' || d.anio as concepto_nuevo
from _dest d where d.existente_id is null;

create temp table _final as
select d.mes, d.anio, d.proyecto, d.moneda, d.rubro, d.n, d.monto,
       coalesce(d.existente_id, nv.nid) as destino_id,
       case when d.existente_id is null then 'nueva' else 'existente' end as tipo,
       coalesce(nv.folio_nuevo, (select folio from public.partidas where id = d.existente_id)) as folio_destino,
       coalesce(nv.concepto_nuevo, (select concepto from public.partidas where id = d.existente_id)) as concepto_destino
from _dest d left join _nuevas nv on (nv.mes, nv.anio, nv.proyecto, nv.moneda, nv.rubro) = (d.mes, d.anio, d.proyecto, d.moneda, d.rubro);

do $$
declare n_mov int; n_upd int;
begin
  select count(*) into n_mov from _mov;
  if n_mov <> 89 then raise exception 'Se esperaban 89 transacciones y hay %. No se cambió nada.', n_mov; end if;

  insert into public.partidas (id, unidad, mes, anio, smi, concepto, rubro, categoria, proyecto, monto_estimado, moneda, folio, created_by, extra)
  select nid, 'OSB', mes, anio, '', concepto_nuevo, rubro, null, proyecto, monto, moneda, folio_nuevo, 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
         jsonb_build_object('origen', 'Reclasificación No Clasificado nov 2026 (lote 4)')
  from _nuevas;

  update public.transacciones t
  set partida_id = f.destino_id, categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m join _final f on (f.mes, f.anio, f.proyecto, f.moneda, f.rubro) = (m.mes, m.anio, m.proyecto, m.moneda, m.rubro)
  where t.id = m.tid;
  get diagnostics n_upd = row_count;
  if n_upd <> 89 then raise exception 'Se movieron % en vez de 89. No se cambió nada.', n_upd; end if;

  -- Parcialidades de compras PPD: la base ya les cambió la partida; falta la categoría
  update public.transacciones h set categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m where m.es_padre and h.transaccion_padre_id = m.tid;

  -- Presupuesto: sale de la partida No Clasificado de origen y entra a la existente que lo recibe
  update public.partidas p set monto_estimado = p.monto_estimado - x.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from (select origen_id, sum(monto) monto from _g group by 1) x where p.id = x.origen_id;
  update public.partidas p set monto_estimado = p.monto_estimado + f.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _final f where f.tipo = 'existente' and p.id = f.destino_id;
end $$;

select count(*) as transacciones, round(sum(t.importe), 2) as importe, count(distinct t.partida_id) as partidas_destino
from public.transacciones t join _mov m on m.tid = t.id;

drop table _mov; drop table _g; drop table _dest; drop table _nuevas; drop table _final;
*/
