-- =====================================================================
-- Corregir partidas que quedaron en NEGATIVO con la limpieza del 24 nov 2026
--
-- Por qué pasó: esas partidas ya tenían gastado más de lo presupuestado.
-- Al reclasificar, sus transacciones se fueron a la partida de su rubro
-- correcto llevándose su monto completo, y el sobregiro se quedó en la
-- partida de origen como presupuesto negativo.
--
-- Qué hace (por compañía + mes + año + moneda):
--   1. Cada partida negativa queda en 0.
--   2. Ese monto se descuenta de las partidas de OTRO rubro del mismo
--      mes/año/moneda que recibieron transacciones ese día, en proporción a
--      lo que recibió cada una. Así el sobregiro queda donde está el gasto.
--   3. El presupuesto total de cada compañía y moneda no cambia.
-- Si alguna partida receptora quedara negativa, o el total cambiara, no se
-- aplica nada. Correr DESPUÉS de clasificar-no-clasificado-sempra-OSB.sql.
-- Probado contra la base con reversa automática.
-- =====================================================================

create temp table _u as
select partida_id, sum(importe) as usado,
 sum(importe) filter (where updated_at >= '2026-11-24' and updated_at < '2026-11-25') as recibido
from public.transacciones
where cancelada_en is null and transaccion_padre_id is null
group by 1;

create temp table _neg as
select p.id, p.folio, p.unidad, p.rubro, p.mes, p.anio, p.moneda, -p.monto_estimado as ajuste
from public.partidas p where p.monto_estimado < 0;

create temp table _rec as
select n.id as neg_id, r.id as rec_id, r.folio, u.recibido,
 round(n.ajuste * u.recibido / sum(u.recibido) over (partition by n.id), 2) as descuento
from _neg n
join public.partidas r on r.unidad = n.unidad and r.mes = n.mes and r.anio = n.anio and r.moneda = n.moneda
 and r.id <> n.id and r.rubro <> n.rubro and r.monto_estimado >= 0
join _u u on u.partida_id = r.id and coalesce(u.recibido, 0) > 0
where r.updated_at >= '2026-11-24';

-- Que el redondeo no deje centavos sueltos: la diferencia va a la receptora más grande
update _rec x set descuento = x.descuento + d.dif
from (select n.id as neg_id, n.ajuste - sum(r.descuento) as dif,
  (select r2.rec_id from _rec r2 where r2.neg_id = n.id order by r2.recibido desc limit 1) as mayor
  from _neg n join _rec r on r.neg_id = n.id group by n.id, n.ajuste) d
where x.neg_id = d.neg_id and x.rec_id = d.mayor;

do $$
declare n_neg int; t_antes numeric; t_despues numeric;
begin
 select count(*) into n_neg from _neg;
 if n_neg = 0 then raise exception 'No hay partidas negativas. No se cambió nada.'; end if;
 if exists (select 1 from _neg n where not exists (select 1 from _rec r where r.neg_id = n.id)) then
  raise exception 'Hay partidas negativas sin receptora identificable: %. No se cambió nada.',
   (select string_agg(folio, ', ') from _neg n where not exists (select 1 from _rec r where r.neg_id = n.id));
 end if;
 select sum(monto_estimado) into t_antes from public.partidas;

 update public.partidas p set monto_estimado = 0, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from _neg n where p.id = n.id;
 update public.partidas p set monto_estimado = p.monto_estimado - x.descuento, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from (select rec_id, sum(descuento) as descuento from _rec group by 1) x where p.id = x.rec_id;

 select sum(monto_estimado) into t_despues from public.partidas;
 if round(t_antes, 2) <> round(t_despues, 2) then raise exception 'El presupuesto total cambió. No se cambió nada.'; end if;
 if exists (select 1 from public.partidas where monto_estimado < 0) then
  raise exception 'Quedarían partidas negativas: %. No se cambió nada.',
   (select string_agg(folio, ', ') from public.partidas where monto_estimado < 0);
 end if;
 raise notice 'Listo: % partidas negativas corregidas', n_neg;
end $$;
