-- =====================================================================
-- Folios de PARTIDA con "P" para no confundirlos con los de transacción
--   Antes:  OSB-OCT26-071        (igual que una transacción)
--   Ahora:  OSB-P-OCT26-071
-- Aplica a las 900 partidas de OSB, CTM e ISE; conserva el número.
-- Las transacciones NO cambian (su folio está en Drive, Pagos, etc.).
-- Las transacciones apuntan a la partida por su id, no por el folio: no se desligan.
-- Los reportes ya enviados a Dirección conservan el folio con el que se enviaron;
-- la comparación entre versiones empata por id, así que no marca cambios falsos.
-- Correr ANTES de clasificar-no-clasificado-sempra-OSB.sql y junto con la app v3.17.0.
-- Todo o nada. Probado contra la base con reversa automática.
-- =====================================================================

do $$
declare n_total int; n_upd int;
begin
 select count(*) into n_total from public.partidas where folio ~ '^[A-Z]+-[A-Z]{3}[0-9]{2}-[0-9]+$';
 if n_total <> 900 then
  raise exception 'Se esperaban 900 partidas con folio del formato anterior y hay %. No se cambió nada.', n_total;
 end if;

 update public.partidas
 set folio = regexp_replace(folio, '^([A-Z]+)-([A-Z]{3}[0-9]{2}-[0-9]+)$', '\1-P-\2')
 where folio ~ '^[A-Z]+-[A-Z]{3}[0-9]{2}-[0-9]+$';
 get diagnostics n_upd = row_count;

 if n_upd <> n_total
  or exists (select 1 from public.partidas where folio !~ '^[A-Z]+-P-[A-Z]{3}[0-9]{2}-[0-9]+$') then
  raise exception 'Algún folio quedó fuera del formato nuevo. No se cambió nada.';
 end if;
 raise notice 'Listo: % folios de partida renombrados', n_upd;
end $$;
