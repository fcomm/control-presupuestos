-- 69 · Comprobantes de pago leídos por la app (Transacciones › Importar / Exportar › Comprobantes de pago)
-- Bitácora de cada comprobante bancario que se reconoció contra una transacción.
-- Sirve para dos cosas: no reconocer dos veces el mismo comprobante (si se vuelve
-- a subir, sale como "Ya cargado") y saber de qué comprobante salió cada Pagado.
-- El número de tarjeta de los pagos de tarjeta de crédito NO se lee ni se guarda.
create table if not exists public.comprobantes_pago (
  id              uuid primary key default gen_random_uuid(),
  unidad          text not null,
  transaccion_id  uuid references public.transacciones(id) on delete set null,
  adjunto_id      text,
  autorizacion    text,
  clave_rastreo   text,
  formato         text,           -- detalle | transferencia | servicio | lista | sat
  estatus_banco   text,           -- Aceptada | Aplicada
  importe         numeric not null,
  moneda          text not null,
  fecha           date,
  cuenta_destino  text,
  beneficiario    text,
  concepto        text,
  archivo         text,
  pagina          integer,
  marco_pagado    boolean not null default false,  -- true: la app cambió la transacción a Pagado
  created_at      timestamptz not null default now(),
  created_by      uuid default auth.uid()
);
create unique index if not exists comprobantes_pago_unico
  on public.comprobantes_pago (unidad, autorizacion, importe, fecha) where autorizacion is not null;
create index if not exists comprobantes_pago_transaccion_idx on public.comprobantes_pago (transaccion_id);

alter table public.comprobantes_pago enable row level security;
create policy "comprobantes_pago: leer" on public.comprobantes_pago
  for select to authenticated using (tiene_acceso_unidad(unidad));
create policy "comprobantes_pago: alta" on public.comprobantes_pago
  for insert to authenticated with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "comprobantes_pago: borrar" on public.comprobantes_pago
  for delete to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
