-- =====================================================================
-- Consolidar en partidas GENERALES (sin proyecto) las partidas que crearon
-- los scripts de limpieza de hoy (24 nov 2026):
--   · Homologación alimentos AMEX
--   · Reclasificación No Clasificado (lotes 1 y 2)
--   · Homologación por proveedor (OSB, CTM, ISE)
-- Esas partidas se crearon una por proyecto. Con el nuevo criterio (presupuesto
-- por rubro; el proyecto lo lleva la transacción) se agrupan en UNA partida por
-- compañía + rubro + mes + año + moneda, con proyecto "Todos".
--
-- Qué hace:
--   1. Si ya hay una partida general para ese grupo, todo se va a ella.
--      Si no, la primera partida del grupo (por folio) se vuelve general:
--      proyecto "Todos" y concepto "<Rubro> <Mes> <Año>"; conserva su folio.
--   2. Mueve las transacciones (las parcialidades PPD siguen a su compra).
--   3. Pasa el presupuesto completo a la partida general.
--   4. Elimina las partidas que quedaron vacías (solo las creadas por esos scripts).
-- Resultado esperado: 249 partidas → 151 generales (150 convertidas, 99 eliminadas).
-- El presupuesto total de cada compañía no cambia; si cambiara, no se aplica nada.
-- No toca partidas capturadas por ustedes. Las transacciones conservan su proyecto.
-- Probado contra la base con reversa automática.
-- =====================================================================

-- PASO 1 (opcional): previo de los grupos
-- select unidad, rubro, mes, anio, moneda, count(*) partidas, sum(monto_estimado) monto
-- from public.partidas
-- where extra->>'origen' in ('Homologación alimentos AMEX nov 2026', 'Homologación por proveedor nov 2026',
--  'Reclasificación No Clasificado nov 2026', 'Reclasificación No Clasificado nov 2026 (lote 2)')
--  and coalesce(proyecto, '') not in ('Todos', '')
-- group by 1,2,3,4,5 order by 1,4,3,2;

-- PASO 2: aplicar (todo o nada)
-- Partidas creadas hoy por los scripts de limpieza que quedaron ligadas a un proyecto
create temp table _src as
select p.id, p.folio, p.unidad, p.rubro, p.mes, p.anio, p.moneda, p.monto_estimado
from public.partidas p
where p.extra->>'origen' in ('Homologación alimentos AMEX nov 2026', 'Homologación por proveedor nov 2026',
 'Reclasificación No Clasificado nov 2026', 'Reclasificación No Clasificado nov 2026 (lote 2)')
 and coalesce(p.proyecto, '') not in ('Todos', '');

-- Un destino general por compañía + rubro + mes + año + moneda:
-- 1) una partida general que ya exista; 2) si no hay, la primera partida del grupo se convierte en general
create temp table _grp as
select g.*, coalesce(
 (select p.id from public.partidas p where p.unidad = g.unidad and p.rubro = g.rubro and p.mes = g.mes
  and p.anio = g.anio and p.moneda = g.moneda and coalesce(p.proyecto, '') in ('Todos', '') order by p.folio limit 1),
 g.primera) as destino_id,
 exists (select 1 from public.partidas p where p.unidad = g.unidad and p.rubro = g.rubro and p.mes = g.mes
  and p.anio = g.anio and p.moneda = g.moneda and coalesce(p.proyecto, '') in ('Todos', '')) as hay_general
from (select unidad, rubro, mes, anio, moneda, count(*) as n,
  (array_agg(id order by folio))[1] as primera
 from _src group by 1,2,3,4,5) g;

create temp table _map as -- cada partida de origen y su destino (excluye la que se convierte)
select s.id as src_id, s.monto_estimado, g.destino_id
from _src s join _grp g using (unidad, rubro, mes, anio, moneda)
where s.id <> g.destino_id;
do $$
declare n_src int; n_conv int; n_del int; n_rest int;
 t_osb numeric; t_ctm numeric; t_ise numeric;
begin
 select round(sum(monto_estimado) filter (where unidad='OSB'),2), round(sum(monto_estimado) filter (where unidad='CTM'),2),
  round(sum(monto_estimado) filter (where unidad='ISE'),2) into t_osb, t_ctm, t_ise from public.partidas;
 select count(*) into n_src from _src;
 if n_src <> 249 then raise exception 'Se esperaban 249 partidas con proyecto y hay %. No se cambió nada.', n_src; end if;

 -- La primera partida de cada grupo sin general existente se vuelve general
 update public.partidas p set proyecto = 'Todos', concepto = p.rubro || ' ' || p.mes || ' ' || p.anio,
  updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _grp g where not g.hay_general and p.id = g.destino_id;
 get diagnostics n_conv = row_count;

 -- Transacciones: primero las principales (las parcialidades PPD siguen a su compra), luego cualquier resto
 update public.transacciones t set partida_id = m.destino_id, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _map m where t.partida_id = m.src_id and t.transaccion_padre_id is null;
 update public.transacciones t set partida_id = m.destino_id, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _map m where t.partida_id = m.src_id;
 select count(*) into n_rest from public.transacciones t join _map m on m.src_id = t.partida_id;
 if n_rest > 0 then raise exception 'Quedaron % transacciones en partidas de origen', n_rest; end if;

 -- Presupuesto: pasa completo a la partida general
 update public.partidas p set monto_estimado = p.monto_estimado + x.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from (select destino_id, sum(monto_estimado) as monto from _map group by 1) x where p.id = x.destino_id;

 -- Las partidas de origen quedan vacías (sin transacciones): se eliminan
 delete from public.partidas p using _map m where p.id = m.src_id;
 get diagnostics n_del = row_count;

 if (select round(sum(monto_estimado),2) from public.partidas where unidad='OSB') <> t_osb
  or (select round(sum(monto_estimado),2) from public.partidas where unidad='CTM') <> t_ctm
  or (select round(sum(monto_estimado),2) from public.partidas where unidad='ISE') <> t_ise
 then raise exception 'El presupuesto total cambió. No se cambió nada.'; end if;
 raise notice 'Listo: % partidas convertidas a generales, % partidas consolidadas y eliminadas', n_conv, n_del;
end $$;
