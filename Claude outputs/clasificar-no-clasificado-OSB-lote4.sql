-- =====================================================================
-- Clasificar No Clasificado de OSB — lote 4 (24 nov 2026)
-- 31 transacciones de la tabla B que confirmó Paco:
--   Implementación y Control de Procesos Ind. .. Servicios Operativos › Servicios especializados
--   Poder Ejecutivo de Querétaro, Gob. Tabasco y Veracruz .. Gastos Financieros e Impuestos › Derechos
--   Mario Daniel Hernández Leyton ............. Servicios Administrativos › Honorarios profesionales
--   Aquiles México ............................ Tecnologías de Información › Suscripciones
--   Gas Express Nieto ......................... Servicios Básicos › Gas
--   Amazon México, Amazon Export, Liverpool, Acomee y "Compras AMEX" (sin proveedor)
--                                           ... Equipos y Herramientas (sin categoría, por definir)
-- Mismo criterio que el lote 3: partida general por rubro + mes + año + moneda,
-- sin dejar partidas en negativo; el total de OSB no cambia. Todo o nada.
-- Quedan 34 en No Clasificado (los pendientes).
-- Probado contra la base con reversa automática.
-- =====================================================================

create temp table _mov as
with mapa(proveedor, rubro, categoria) as (values
('IMPLEMENTACION Y CONTROL DE PROCESOS INDUSTRIALES','Servicios Operativos','Servicios especializados'),
('PODER EJECUTIVO DEL ESTADO DE QUERETARO','Gastos Financieros e Impuestos','Derechos'),
('PODER EJECUTIVO DEL ESTADO DE QUERETARO 5 D EMAYO','Gastos Financieros e Impuestos','Derechos'),
('GOBIERNO DEL ESTADO DE TABASCO','Gastos Financieros e Impuestos','Derechos'),
('GOBIERNO DEL ESTADO DE VERACRUZ','Gastos Financieros e Impuestos','Derechos'),
('MARIO DANIEL HERNANDEZ LEYTON','Servicios Administrativos','Honorarios profesionales'),
('AQUILES MEXICO SA DE CV','Tecnologías de Información','Suscripciones'),
('GAS EXPRESS NIETO','Servicios Básicos','Gas'),
('','Equipos y Herramientas',''),
('DISTRIBUIDORA LIVERPOOL SA DE CV','Equipos y Herramientas',''),
('AMAZON EXPORT SALES LLC','Equipos y Herramientas',''),
('SERVICIOS COMERCIALES AMAZON MEXICO S DE RL DE CV','Equipos y Herramientas',''),
('GRUPO COMERCIAL ACOMEE','Equipos y Herramientas','')
)
select t.id as tid, p.unidad as unidad, t.folio_transaccion, t.proveedor, t.importe, t.es_padre,
 m.rubro, m.categoria, p.id as origen_id, p.folio as origen_folio, p.rubro as rubro_origen,
 p.mes, p.anio, p.proyecto, p.moneda,
 case when p.rubro = m.rubro then 'categoria' else 'mover' end as accion
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(coalesce(t.proveedor, '')))
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
 if n_mov <> 31 then raise exception 'Se esperaban 31 transacciones y hay %. No se cambió nada.', n_mov; end if;
 select round(sum(monto_estimado) filter (where unidad='OSB'),2) into t_osb from public.partidas;
 

 insert into public.partidas (id, unidad, mes, anio, smi, concepto, rubro, categoria, proyecto, monto_estimado, moneda, folio, created_by, extra)
 select nid, unidad, mes, anio, '', concepto_nuevo, rubro, null, 'Todos', monto, moneda, folio_nuevo, 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
 jsonb_build_object('origen', 'Clasificación No Clasificado OSB nov 2026 (lote 4)')
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

