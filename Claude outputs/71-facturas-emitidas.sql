-- 71 · Facturas emitidas a clientes — ya aplicada en Supabase el 25/11/26
create table if not exists public.facturas_emitidas (
  id              uuid primary key default gen_random_uuid(),
  unidad          text not null,
  folio           text not null,
  cliente         text,
  proyecto        text,
  importe         numeric,
  moneda          text not null default 'MXP' check (moneda in ('MXP','USD')),
  status          text not null default 'Emitida' check (status in ('Emitida','Pagada','Cancelada')),
  fecha_emision   date not null,
  fecha_pago      date,
  notas           text,
  drive_folder_id text,
  created_at      timestamptz not null default now(),
  created_by      uuid default auth.uid(),
  updated_at      timestamptz not null default now(),
  updated_by      uuid,
  constraint facturas_pagada_con_fecha check (status <> 'Pagada' or fecha_pago is not null)
);
create unique index if not exists facturas_emitidas_folio_unico on public.facturas_emitidas (unidad, upper(trim(folio)));
create index if not exists facturas_emitidas_unidad_idx on public.facturas_emitidas (unidad, fecha_emision);
create trigger trg_facturas_emitidas_updated_at before update on public.facturas_emitidas
  for each row execute function set_updated_at();
alter table public.facturas_emitidas enable row level security;
create policy "facturas_emitidas: leer" on public.facturas_emitidas
  for select to authenticated using (tiene_acceso_unidad(unidad));
create policy "facturas_emitidas: alta" on public.facturas_emitidas
  for insert to authenticated with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "facturas_emitidas: cambio" on public.facturas_emitidas
  for update to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad))
  with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "facturas_emitidas: borrar" on public.facturas_emitidas
  for delete to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
alter publication supabase_realtime add table public.facturas_emitidas;
