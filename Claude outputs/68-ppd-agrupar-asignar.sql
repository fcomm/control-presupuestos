-- =====================================================================
-- 68 — Compras PPD a partir de transacciones existentes (app v3.19.0)
--   ppd_agrupar(ids, total, partida, concepto)  crea la compra PPD (padre) y
--                                                liga las transacciones como
--                                                sus parcialidades 1..n
--   ppd_asignar(ids, padre)                      liga transacciones a una
--                                                compra PPD que ya existe
--   ppd_quitar(id)                               regresa una parcialidad a
--                                                transacción normal
-- Solo administradores. Cada función hace todo o nada.
-- Las parcialidades conservan su folio, estatus, fecha de pago y adjuntos;
-- pasan a la partida del padre (la base lo exige) con tipo de pago "Parcialidad".
-- Correr una vez en el SQL editor de Supabase.
-- =====================================================================

-- Valida un grupo de transacciones candidatas a parcialidad. Devuelve el número
-- de filas o lanza el error con un mensaje para la persona.
create or replace function public.ppd_validar_candidatas(p_ids uuid[])
returns int language plpgsql stable set search_path = public as $$
declare v_n int := coalesce(array_length(p_ids, 1), 0); v_txt text;
begin
  if v_n = 0 then raise exception 'No hay transacciones seleccionadas.'; end if;
  if (select count(*) from transacciones where id = any(p_ids)) <> v_n then
    raise exception 'Alguna de las transacciones ya no existe. Recarga e intenta de nuevo.';
  end if;
  select string_agg(coalesce(folio_transaccion, '(sin folio)'), ', ') into v_txt
    from transacciones where id = any(p_ids) and cancelada_en is not null;
  if v_txt is not null then raise exception 'Están canceladas: %.', v_txt; end if;
  select string_agg(coalesce(folio_transaccion, '(sin folio)'), ', ') into v_txt
    from transacciones where id = any(p_ids) and es_padre;
  if v_txt is not null then raise exception 'Ya son compras PPD (padre): %.', v_txt; end if;
  select string_agg(coalesce(folio_transaccion, '(sin folio)'), ', ') into v_txt
    from transacciones where id = any(p_ids) and transaccion_padre_id is not null;
  if v_txt is not null then raise exception 'Ya son parcialidades de otra compra PPD: %. Quítalas primero.', v_txt; end if;
  if (select count(distinct unidad_detectada) from transacciones where id = any(p_ids)) > 1 then
    raise exception 'Todas deben ser de la misma compañía.';
  end if;
  if (select count(distinct coalesce(moneda, 'MXP')) from transacciones where id = any(p_ids)) > 1 then
    raise exception 'Todas deben estar en la misma moneda.';
  end if;
  if (select count(distinct coalesce(proveedor_id::text, upper(trim(coalesce(proveedor, ''))))) from transacciones where id = any(p_ids)) > 1 then
    raise exception 'Todas deben ser del mismo proveedor.';
  end if;
  return v_n;
end $$;

-- Liga las transacciones (ya validadas) al padre, en orden de fecha, a partir
-- de p_num_ini y con lo que quede libre del 100 %.
create or replace function public.ppd_ligar(p_ids uuid[], p_padre uuid, p_num_ini int, p_libre numeric, p_cierra boolean)
returns int language plpgsql set search_path = public as $$
declare r record; v_total numeric; v_partida uuid; v_pct numeric; v_acum numeric := 0; v_i int := 0; v_n int;
begin
  if not public.es_admin() then raise exception 'Solo un administrador puede ligar parcialidades.'; end if;
  select importe, partida_id into v_total, v_partida from transacciones where id = p_padre;
  if coalesce(v_total, 0) <= 0 then raise exception 'La compra PPD no tiene importe.'; end if;
  v_n := array_length(p_ids, 1);
  for r in select id, importe from transacciones where id = any(p_ids) order by dia nulls last, folio_transaccion loop
    v_i := v_i + 1;
    v_pct := round(coalesce(r.importe, 0) / v_total * 100, 4);
    -- La última cierra exacto: así la suma no queda en 100.0001 % por redondeo
    if v_i = v_n and (p_cierra or v_acum + v_pct > p_libre) and abs(v_acum + v_pct - p_libre) < 0.01 then
      v_pct := p_libre - v_acum;
    end if;
    v_acum := v_acum + v_pct;
    if v_acum > p_libre + 0.00001 then
      raise exception 'Con estas transacciones la compra pasaría de 100 %%: solo queda % %% libre (% de lo seleccionado).',
        round(p_libre, 2), round(v_acum, 2);
    end if;
    update transacciones set transaccion_padre_id = p_padre, num_parcialidad = p_num_ini + v_i - 1, pct_pago = v_pct,
           tipo_pago = 'Parcialidad', partida_id = v_partida, updated_by = auth.uid()
     where id = r.id;
  end loop;
  return v_i;
end $$;

create or replace function public.ppd_agrupar(p_ids uuid[], p_total numeric, p_partida uuid default null, p_concepto text default null)
returns jsonb language plpgsql set search_path = public as $$
declare v_n int; b transacciones%rowtype; v_suma numeric; v_partida uuid; v_pmon text; v_punidad text;
        v_prefijo text; v_folio text; v_padre uuid := gen_random_uuid(); v_fecha date;
begin
  if not public.es_admin() then raise exception 'Solo un administrador puede agrupar en compra PPD.'; end if;
  v_n := public.ppd_validar_candidatas(p_ids);
  select sum(importe) into v_suma from transacciones where id = any(p_ids);
  if p_total is null or p_total <= 0 then raise exception 'Falta el total de la compra.'; end if;
  if p_total < v_suma - 0.005 then
    raise exception 'El total de la compra (%) no puede ser menor que la suma de las transacciones (%).', round(p_total, 2), round(v_suma, 2);
  end if;

  select * into b from transacciones where id = any(p_ids) order by dia nulls last, folio_transaccion limit 1;
  v_partida := coalesce(p_partida, b.partida_id);
  if v_partida is null then raise exception 'Elige la partida de la compra.'; end if;
  select coalesce(moneda, 'MXP'), unidad into v_pmon, v_punidad from partidas where id = v_partida;
  if v_punidad is null then raise exception 'La partida elegida no existe.'; end if;
  if v_punidad is distinct from b.unidad_detectada then raise exception 'La partida es de otra compañía.'; end if;
  if v_pmon <> coalesce(b.moneda, 'MXP') then raise exception 'La partida es en % y las transacciones en %.', v_pmon, coalesce(b.moneda, 'MXP'); end if;

  -- Folio del padre: el siguiente del mes de la transacción más antigua
  v_fecha := coalesce(b.dia, current_date);
  v_prefijo := b.unidad_detectada || '-' ||
    (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[extract(month from v_fecha)::int] ||
    to_char(v_fecha, 'YY') || '-';
  select v_prefijo || lpad((coalesce(max(substring(folio_transaccion from '-(\d+)$')::int), 0) + 1)::text, 3, '0') into v_folio
    from transacciones where folio_transaccion ~ ('^' || v_prefijo || '\d+$');

  insert into transacciones (id, folio_transaccion, unidad_detectada, dia, smi, area, zona, moneda, importe, es_padre,
    proyecto, categoria, cuenta_id, proveedor, proveedor_id, tipo_pago, forma_pago, metodo_pago, partida_id, solicitante,
    folio_factura, folio_compra_sae, referencia_pago, concepto_detallado, registrada_en, created_by, updated_by,
    es_mensual, historial_cancelacion)
  values (v_padre, v_folio, b.unidad_detectada, b.dia, '', b.area, b.zona, b.moneda, round(p_total, 4), true,
    b.proyecto, b.categoria, b.cuenta_id, b.proveedor, b.proveedor_id, '', coalesce(b.forma_pago, '03'), 'PPD', v_partida, b.solicitante,
    '', '', '', coalesce(nullif(trim(p_concepto), ''), b.concepto_detallado), now(), auth.uid(), auth.uid(),
    false, '[]'::jsonb);

  perform public.ppd_ligar(p_ids, v_padre, 1, 100, abs(p_total - v_suma) < 0.01);
  return jsonb_build_object('id', v_padre, 'folio', v_folio, 'parcialidades', v_n);
end $$;

create or replace function public.ppd_asignar(p_ids uuid[], p_padre uuid)
returns jsonb language plpgsql set search_path = public as $$
declare v_n int; p transacciones%rowtype; v_usado numeric; v_num int; v_suma numeric; v_ligadas int;
begin
  if not public.es_admin() then raise exception 'Solo un administrador puede asignar a una compra PPD.'; end if;
  v_n := public.ppd_validar_candidatas(p_ids);
  select * into p from transacciones where id = p_padre;
  if p.id is null or not p.es_padre then raise exception 'La transacción elegida no es una compra PPD.'; end if;
  if p.cancelada_en is not null then raise exception 'La compra PPD % está cancelada.', p.folio_transaccion; end if;
  if exists (select 1 from transacciones t where t.id = any(p_ids) and (
       t.unidad_detectada is distinct from p.unidad_detectada
    or coalesce(t.moneda, 'MXP') <> coalesce(p.moneda, 'MXP')
    or coalesce(t.proveedor_id::text, upper(trim(coalesce(t.proveedor, '')))) <> coalesce(p.proveedor_id::text, upper(trim(coalesce(p.proveedor, '')))))) then
    raise exception 'Las transacciones deben ser de la misma compañía, moneda y proveedor que la compra PPD %.', p.folio_transaccion;
  end if;
  select coalesce(sum(pct_pago), 0) into v_usado
    from transacciones where transaccion_padre_id = p_padre and cancelada_en is null;
  select coalesce(max(num_parcialidad), 0) into v_num from transacciones where transaccion_padre_id = p_padre;
  select sum(importe) into v_suma from transacciones where id = any(p_ids);
  v_ligadas := public.ppd_ligar(p_ids, p_padre, v_num + 1, 100 - v_usado,
    abs(v_suma - (p.importe * (100 - v_usado) / 100)) < 0.01);
  return jsonb_build_object('folio', p.folio_transaccion, 'parcialidades', v_ligadas);
end $$;

create or replace function public.ppd_quitar(p_id uuid)
returns jsonb language plpgsql set search_path = public as $$
declare h transacciones%rowtype; v_folio text;
begin
  if not public.es_admin() then raise exception 'Solo un administrador puede quitar parcialidades.'; end if;
  select * into h from transacciones where id = p_id;
  if h.id is null or h.transaccion_padre_id is null then raise exception 'La transacción no es parcialidad de una compra PPD.'; end if;
  select folio_transaccion into v_folio from transacciones where id = h.transaccion_padre_id;
  update transacciones set transaccion_padre_id = null, num_parcialidad = null, pct_pago = null,
         tipo_pago = '', updated_by = auth.uid()
   where id = p_id;
  return jsonb_build_object('folio_padre', v_folio);
end $$;

grant execute on function public.ppd_agrupar(uuid[], numeric, uuid, text) to authenticated;
grant execute on function public.ppd_asignar(uuid[], uuid) to authenticated;
grant execute on function public.ppd_quitar(uuid) to authenticated;
