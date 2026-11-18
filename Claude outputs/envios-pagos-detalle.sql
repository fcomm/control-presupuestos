-- ============================================================================
-- Detalle de cada envío a Pagos — v3.7.0
--
-- Ejecutar en: Supabase Dashboard > SQL Editor > New query > Run
--
-- Guarda QUÉ pagos iban en cada envío (R0, R1…) de una fecha, con su importe,
-- proveedor, concepto y cuenta tal como se mandaron. Con eso, al reenviar la
-- misma fecha la app dice qué es nuevo, qué cambió y qué se quitó.
-- Los envíos hechos antes de esta migración no tienen detalle: el primer
-- reenvío de esas fechas sale sin comparación (la app lo avisa).
-- ============================================================================

begin;

create table if not exists public.envios_pagos_detalle (
  id              text primary key,
  envio_id        text not null,          -- envios_pagos.id
  unidad          text not null,
  fecha_pago      date not null,
  revision        int  not null,
  transaccion_id  text not null,
  folio           text,
  proveedor       text,
  concepto        text,
  zona            text,
  importe         numeric(14,2) not null default 0,
  moneda          text not null default 'MXP',
  forma_pago      text,
  referencia_pago text,
  banco           text,
  clabe           text,
  numero_cuenta   text,
  folio_factura   text,
  created_at      timestamptz not null default now()
);
create index if not exists envios_pagos_detalle_envio_idx on public.envios_pagos_detalle (envio_id);
create index if not exists envios_pagos_detalle_fecha_idx on public.envios_pagos_detalle (unidad, fecha_pago);

comment on table public.envios_pagos_detalle is
  'Lo que iba en cada envío a Pagos, congelado al enviarse. Sirve para comparar un reenvío contra el envío anterior.';

alter table public.envios_pagos_detalle enable row level security;

-- Mismo acceso por compañía que el resto de la app, y mismo permiso de
-- edición (Solo lectura no puede enviar) si ya se corrió permisos-usuario.sql.
do $$
declare acceso text; edicion text;
begin
  acceso  := case when to_regprocedure('public.tiene_acceso_unidad(text)') is not null
                  then 'tiene_acceso_unidad(unidad)' else 'true' end;
  edicion := case when to_regprocedure('public.puede_editar_unidad(text)') is not null
                  then 'puede_editar_unidad(unidad)' else 'true' end;
  execute 'drop policy if exists envios_pagos_detalle_rls on public.envios_pagos_detalle';
  execute format('create policy envios_pagos_detalle_rls on public.envios_pagos_detalle for all to authenticated using (%s) with check (%s)', acceso, acceso);
  execute 'drop policy if exists permiso_edicion_ins on public.envios_pagos_detalle';
  execute format('create policy permiso_edicion_ins on public.envios_pagos_detalle as restrictive for insert to authenticated with check (%s)', edicion);
end $$;

commit;

-- VERIFICACIÓN: debe existir y estar vacía.
select count(*) as filas from public.envios_pagos_detalle;
