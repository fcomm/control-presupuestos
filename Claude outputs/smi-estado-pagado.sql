-- ============================================================================
-- SMI: nuevo estado "Pagado" — v3.6.0
--
-- Ejecutar en: Supabase Dashboard > SQL Editor > New query > Run
--
-- QUÉ HACE
-- 1. Acepta el estado 'pagado' en solicitudes (si la tabla tiene una regla
--    que limita los estados, se reemplaza por una que lo incluye).
-- 2. Lo pone solo: cuando TODAS las transacciones vigentes (no canceladas) de
--    una SMI quedan en status Pagado, la SMI pasa de "Agendado a pago" a
--    "Pagado". Si alguna vuelve a No Pagado, la SMI regresa a "Agendado a
--    pago". En una compra PPD esto pasa cuando se paga el 100 % (la compra
--    se marca Pagada sola al completarse sus parcialidades).
--    Solo mueve SMI que estén en Agendado a pago o Pagado: una cancelada o
--    regresada a mano a otro paso no se toca.
-- 3. Pone en Pagado las SMI que hoy ya tienen todo pagado.
-- ============================================================================

begin;

-- 1. Regla de estados -------------------------------------------------------
do $$
declare r record;
begin
  for r in
    select c.conname from pg_constraint c
    where c.conrelid = 'public.solicitudes'::regclass and c.contype = 'c'
      and pg_get_constraintdef(c.oid) ilike '%estado%'
  loop
    execute format('alter table public.solicitudes drop constraint %I', r.conname);
  end loop;
end $$;

alter table public.solicitudes add constraint solicitudes_estado_check check (estado in (
  'entrada', 'en_revision', 'con_solicitante', 'cotizando', 'por_autorizar',
  'por_programar_pago', 'agendado_pago', 'pagado', 'rechazada', 'cancelada'));


-- 2. Sincronización automática ----------------------------------------------
create or replace function public.smi_sincronizar_pagado(p_solicitud text)
returns void language plpgsql security definer set search_path = public as $$
declare
  n_vigentes int;
  n_pagadas  int;
begin
  if p_solicitud is null then return; end if;
  select count(*), count(*) filter (where status = 'Pagado')
    into n_vigentes, n_pagadas
    from public.transacciones
   where solicitud_id::text = p_solicitud and cancelada_en is null;

  if n_vigentes > 0 and n_pagadas = n_vigentes then
    update public.solicitudes set estado = 'pagado'
     where id::text = p_solicitud and estado = 'agendado_pago';
  else
    update public.solicitudes set estado = 'agendado_pago'
     where id::text = p_solicitud and estado = 'pagado';
  end if;
end $$;

create or replace function public.trg_smi_pagado() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if tg_op in ('UPDATE', 'DELETE') then
    perform public.smi_sincronizar_pagado(old.solicitud_id::text);
  end if;
  if tg_op in ('INSERT', 'UPDATE')
     and (tg_op = 'INSERT' or new.solicitud_id is distinct from old.solicitud_id
          or new.status is distinct from old.status or new.cancelada_en is distinct from old.cancelada_en) then
    perform public.smi_sincronizar_pagado(new.solicitud_id::text);
  end if;
  return null;
end $$;

drop trigger if exists trg_smi_pagado on public.transacciones;
create trigger trg_smi_pagado
  after insert or update or delete on public.transacciones
  for each row execute function public.trg_smi_pagado();


-- 3. Las que ya están pagadas hoy -------------------------------------------
select public.smi_sincronizar_pagado(s.id::text)
  from public.solicitudes s where s.estado = 'agendado_pago';

commit;


-- VERIFICACIÓN: SMI por estado
select unidad, estado, count(*) from public.solicitudes
 where estado in ('agendado_pago', 'pagado') group by 1, 2 order by 1, 2;
