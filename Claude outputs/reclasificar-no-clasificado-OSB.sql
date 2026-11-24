-- Reclasificar gastos "No Clasificado" de OSB por proveedor (tabla de Paco, 24/11/26).
-- 15 proveedores · 215 transacciones · $52,000,751.53 · 85 partidas nuevas
-- Semiconductores Querétaro queda FUERA hasta definir su categoría.
--
-- El PASO 2:
--   1. Crea una partida por rubro + proyecto + mes + año + moneda, con concepto
--      "<Rubro> <Proyecto> <Mes> <Año>", el siguiente folio del mes y presupuesto =
--      lo que se le mueve. (Hoy no existe ninguna partida de esos rubros en esos
--      meses y proyectos, por eso todas son nuevas.)
--   2. Mueve las transacciones y les pone su categoría. Las compras PPD se mueven con
--      sus parcialidades (la base las arrastra) y éstas reciben la misma categoría.
--   3. Resta lo movido al presupuesto de cada partida "No Clasificado" de origen, así
--      el total presupuestado de OSB no cambia ($137,463,586.45 antes y después).
--   No toca status, importes, folios de transacción ni adjuntos.
--   Todo o nada: si no son exactamente 215, se detiene y no cambia nada.
--   Probado contra la base el 24/11/26 en una transacción revertida: 215 / $52,000,751.53 / 85.

------------------------------------------------------------------------
-- PASO 1 · PREVIO (solo lectura): una fila por partida nueva
------------------------------------------------------------------------
with mapa(proveedor, rubro, categoria) as (values
  ('SITELSOFT',                                   'Servicios Operativos',       'Servicios especializados'),
  ('INNOVARQ URBITECH',                           'Servicios Operativos',       'Servicios especializados'),
  ('SONDER AKTION DE MEXICO',                     'Servicios Operativos',       'Servicios especializados'),
  ('INMOBILIARIA IZ2 S DE RL DE CV',              'Servicios Operativos',       'Arrendamientos'),
  ('ESPECIALIDADES EN TURBO MAQUINARIA',          'Servicios Operativos',       'Servicios especializados'),
  ('RICARDO POBLANO PAZ',                         'Servicios Operativos',       'Servicios especializados'),
  ('SOLUCIONES INDUSTRIALES SERCOIN',             'Servicios Operativos',       'Servicios especializados'),
  ('RADIOMOVIL DIPSA SA DE CV',                   'Servicios Básicos',          'Telefonía móvil'),
  ('INSTRU CONTROL',                              'Materiales e Insumos',       'Instrumentación'),
  ('AT&T COMUNICACIONES DIGITALES S DE RL DE CV', 'Servicios Básicos',          'Telefonía móvil'),
  ('ORACLE DE MEXICO SA DE CV',                   'Tecnologías de Información', 'Servicios en la nube'),
  ('MRS DEL BAJIO',                               'Servicios Operativos',       'Servicios especializados'),
  ('CONSULTORES ADMINISTRATIVOS MARS',            'Servicios Administrativos',  'Consultoría'),
  ('HINOJOSA GOMEZ HUMBERTO',                     'Servicios Administrativos',  'Consultoría'),
  ('TRACTOCAMIONES BASURTO SA DE CV',             'Servicios Operativos',       'Arrendamientos')
),
mov as (
select t.id as tid, t.folio_transaccion, t.proveedor, t.importe, t.es_padre, m.rubro, m.categoria,
       p.id as origen_id, p.folio as origen_folio, p.mes, p.anio, p.proyecto, p.moneda
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.unidad_detectada = 'OSB' and t.cancelada_en is null and t.transaccion_padre_id is null
  and p.rubro = 'No Clasificado'
),
n as (
  select mes, anio, proyecto, moneda, rubro, 'OSB-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], mes)] || right(anio::text, 2) || '-' as prefijo,
         count(*) as transacciones, round(sum(importe), 2) as monto,
         string_agg(distinct origen_folio, ', ') as partidas_origen,
         string_agg(distinct upper(trim(proveedor)), ', ') as proveedores
  from mov group by 1,2,3,4,5,6
)
select n.prefijo || lpad((coalesce((select max(substring(p.folio from '-(\d+)$')::int) from public.partidas p
                             where p.unidad = 'OSB' and p.folio like n.prefijo || '%'), 0)
                   + row_number() over (partition by n.prefijo order by n.proyecto, n.rubro))::text, 3, '0') as folio_nuevo,
       n.rubro || ' ' || n.proyecto || ' ' || n.mes || ' ' || n.anio as concepto_nuevo,
       n.transacciones, n.monto, n.partidas_origen, n.proveedores
from n
order by n.anio, array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], n.mes), n.proyecto, n.rubro;

------------------------------------------------------------------------
-- PASO 2 · APLICAR. Quita /* y */ y ejecuta el bloque completo.
------------------------------------------------------------------------
/*
create temp table _mov as
with mapa(proveedor, rubro, categoria) as (values
  ('SITELSOFT',                                   'Servicios Operativos',       'Servicios especializados'),
  ('INNOVARQ URBITECH',                           'Servicios Operativos',       'Servicios especializados'),
  ('SONDER AKTION DE MEXICO',                     'Servicios Operativos',       'Servicios especializados'),
  ('INMOBILIARIA IZ2 S DE RL DE CV',              'Servicios Operativos',       'Arrendamientos'),
  ('ESPECIALIDADES EN TURBO MAQUINARIA',          'Servicios Operativos',       'Servicios especializados'),
  ('RICARDO POBLANO PAZ',                         'Servicios Operativos',       'Servicios especializados'),
  ('SOLUCIONES INDUSTRIALES SERCOIN',             'Servicios Operativos',       'Servicios especializados'),
  ('RADIOMOVIL DIPSA SA DE CV',                   'Servicios Básicos',          'Telefonía móvil'),
  ('INSTRU CONTROL',                              'Materiales e Insumos',       'Instrumentación'),
  ('AT&T COMUNICACIONES DIGITALES S DE RL DE CV', 'Servicios Básicos',          'Telefonía móvil'),
  ('ORACLE DE MEXICO SA DE CV',                   'Tecnologías de Información', 'Servicios en la nube'),
  ('MRS DEL BAJIO',                               'Servicios Operativos',       'Servicios especializados'),
  ('CONSULTORES ADMINISTRATIVOS MARS',            'Servicios Administrativos',  'Consultoría'),
  ('HINOJOSA GOMEZ HUMBERTO',                     'Servicios Administrativos',  'Consultoría'),
  ('TRACTOCAMIONES BASURTO SA DE CV',             'Servicios Operativos',       'Arrendamientos')
)
select t.id as tid, t.folio_transaccion, t.proveedor, t.importe, t.es_padre, m.rubro, m.categoria,
       p.id as origen_id, p.folio as origen_folio, p.mes, p.anio, p.proyecto, p.moneda
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.unidad_detectada = 'OSB' and t.cancelada_en is null and t.transaccion_padre_id is null
  and p.rubro = 'No Clasificado';

create temp table _g as
select origen_id, mes, anio, proyecto, moneda, rubro, round(sum(importe), 2) as monto, 'OSB-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], mes)] || right(anio::text, 2) || '-' as prefijo
from _mov group by 1,2,3,4,5,6;

create temp table _nuevas as
select gen_random_uuid() as nid, n.*, n.prefijo || lpad((coalesce((select max(substring(p.folio from '-(\d+)$')::int) from public.partidas p
                             where p.unidad = 'OSB' and p.folio like n.prefijo || '%'), 0)
                   + row_number() over (partition by n.prefijo order by n.proyecto, n.rubro))::text, 3, '0') as folio_nuevo,
       n.rubro || ' ' || n.proyecto || ' ' || n.mes || ' ' || n.anio as concepto_nuevo
from (select mes, anio, proyecto, moneda, rubro, prefijo, sum(monto) as monto from _g group by 1,2,3,4,5,6) n;

do $$
declare n_mov int; n_upd int;
begin
  select count(*) into n_mov from _mov;
  if n_mov <> 215 then raise exception 'Se esperaban 215 transacciones y hay %. No se cambió nada.', n_mov; end if;

  insert into public.partidas (id, unidad, mes, anio, smi, concepto, rubro, categoria, proyecto, monto_estimado, moneda, folio, created_by, extra)
  select nid, 'OSB', mes, anio, '', concepto_nuevo, rubro, null, proyecto, monto, moneda, folio_nuevo, 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
         jsonb_build_object('origen', 'Reclasificación No Clasificado nov 2026')
  from _nuevas;

  update public.transacciones t
  set partida_id = n.nid, categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m join _nuevas n on (n.mes, n.anio, n.proyecto, n.moneda, n.rubro) = (m.mes, m.anio, m.proyecto, m.moneda, m.rubro)
  where t.id = m.tid;
  get diagnostics n_upd = row_count;
  if n_upd <> 215 then raise exception 'Se movieron % en vez de 215. No se cambió nada.', n_upd; end if;

  -- Parcialidades de las compras PPD: la base ya les cambió la partida; falta la categoría
  update public.transacciones h
  set categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m where m.es_padre and h.transaccion_padre_id = m.tid;

  -- El presupuesto movido sale de la partida "No Clasificado" de origen
  update public.partidas p
  set monto_estimado = p.monto_estimado - x.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from (select origen_id, sum(monto) monto from _g group by 1) x where p.id = x.origen_id;
end $$;

-- Resultado: 215 · $52,000,751.53 · 85
select count(*) as transacciones, round(sum(t.importe), 2) as importe, count(distinct t.partida_id) as partidas_nuevas
from public.transacciones t join _mov m on m.tid = t.id;

drop table _mov; drop table _g; drop table _nuevas;
*/
