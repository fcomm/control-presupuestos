-- =====================================================================
-- Compra PPD Luis Borioli — ON00143 Item 2 (cotización LAA-SEMPRAPROY/2025 R4)
-- Crea la transacción PADRE (siguiente folio OSB-NOV25-###, hoy OSB-NOV25-045)
-- por USD 537,205.11 y convierte en sus parcialidades 1 a 5 (20 % cada una,
-- en orden de fecha) los 5 pagos que ya existen:
--   1 OSB-NOV25-043 (10/11/25)   2 OSB-NOV25-042 (24/11/25)
--   3 OSB-FEB26-043 (04/02/26)   4 OSB-JUN26-058 (18/06/26)
--   5 OSB-JUN26-057 (23/06/26)
-- Todo queda en la partida OSB-P-OCT25-008 (Servicios Operativos USD, nov 2025).
-- Las parcialidades conservan folio, estatus Pagado, fecha de pago y adjuntos;
-- su tipo de pago pasa de "Anticipo" a "Parcialidad". El padre queda Pagado
-- (100 %) con última fecha de pago 23/06/26, calculado por la base.
-- El padre se crea con los datos de OSB-NOV25-043 (proveedor, cuenta, área,
-- zona, solicitante, proyecto, categoría y concepto).
-- Todo o nada. Probado contra la base con reversa automática.
-- =====================================================================

do $$
declare v_padre uuid := gen_random_uuid(); v_base public.transacciones%rowtype; n int; v_folio text;
begin
 if (select count(*) from public.transacciones
     where folio_transaccion in ('OSB-NOV25-043','OSB-NOV25-042','OSB-FEB26-043','OSB-JUN26-058','OSB-JUN26-057')
       and cancelada_en is null and not es_padre and transaccion_padre_id is null) <> 5 then
  raise exception 'No están las 5 transacciones libres (vigentes, sin compra PPD). No se cambió nada.';
 end if;

 select * into v_base from public.transacciones where folio_transaccion = 'OSB-NOV25-043' and cancelada_en is null;
 select 'OSB-NOV25-' || lpad((max(substring(folio_transaccion from '^OSB-NOV25-(\d+)$')::int) + 1)::text, 3, '0') into v_folio
 from public.transacciones where folio_transaccion ~ '^OSB-NOV25-\d+$';

 -- Padre (compra PPD)
 insert into public.transacciones (id, folio_transaccion, unidad_detectada, dia, smi, area, zona, moneda, importe, es_padre, proyecto, categoria,
  cuenta_id, proveedor, proveedor_id, tipo_pago, forma_pago, metodo_pago, partida_id, solicitante, folio_factura, folio_compra_sae, referencia_pago,
  concepto_detallado, registrada_en, created_by, updated_by, es_mensual, historial_cancelacion)
 select v_padre, v_folio, b.unidad_detectada, b.dia, '', b.area, b.zona, b.moneda, 537205.105, true, b.proyecto, b.categoria,
  b.cuenta_id, b.proveedor, b.proveedor_id, '', b.forma_pago, 'PPD', b.partida_id, b.solicitante, '', '', '',
  b.concepto_detallado, now(), 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb', 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb', false, '[]'::jsonb
 from (select v_base.*) b;

 -- Parcialidades 1 a 5, en orden de fecha
 with h as (select id, row_number() over (order by dia, folio_transaccion) as num from public.transacciones
   where folio_transaccion in ('OSB-NOV25-043','OSB-NOV25-042','OSB-FEB26-043','OSB-JUN26-058','OSB-JUN26-057')
     and cancelada_en is null and not es_padre and transaccion_padre_id is null)
 update public.transacciones t set transaccion_padre_id = v_padre, num_parcialidad = h.num, pct_pago = 20,
   tipo_pago = 'Parcialidad', partida_id = v_base.partida_id, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
 from h where t.id = h.id;
 get diagnostics n = row_count;

 if n <> 5 or (select sum(pct_pago) from public.transacciones where transaccion_padre_id = v_padre) <> 100
  or (select status from public.transacciones where id = v_padre) <> 'Pagado' then
  raise exception 'Algo no cuadró al ligar las parcialidades. No se cambió nada.';
 end if;
 raise notice 'Listo: compra PPD % con 5 parcialidades (100 %%), Pagada', v_folio;
end $$;
