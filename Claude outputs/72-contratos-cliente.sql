-- 72 · Contratos con clientes, adendas/extensiones y asignación de gastos y facturas
-- Ya aplicada en Supabase el 25/11/26. Las tablas del generador anterior
-- (contratos_parametros, contratos_proveedor_legal, contratos_proveedor_unidad,
-- contratos_instrumentos, contratos_clausulas) NO se tocan: quedan sin uso.
create table if not exists public.contratos_cliente (
  id uuid primary key default gen_random_uuid(),
  unidad text not null, numero text not null, cliente text, proyecto text, objeto text,
  moneda text not null default 'MXP' check (moneda in ('MXP','USD')),
  monto_original numeric, fecha_inicio date, fecha_fin_original date,
  status text not null default 'Vigente' check (status in ('Vigente','Concluido','Cancelado')),
  notas text, drive_folder_id text,
  created_at timestamptz not null default now(), created_by uuid default auth.uid(),
  updated_at timestamptz not null default now(), updated_by uuid
);
create unique index if not exists contratos_cliente_numero_unico on public.contratos_cliente (unidad, upper(trim(numero)));
create trigger trg_contratos_cliente_updated_at before update on public.contratos_cliente for each row execute function set_updated_at();
create table if not exists public.contratos_cliente_modificaciones (
  id uuid primary key default gen_random_uuid(),
  contrato_id uuid not null references public.contratos_cliente(id) on delete cascade,
  unidad text not null,
  tipo text not null default 'Adenda' check (tipo in ('Adenda','Extensión')),
  numero text, fecha date not null, monto_cambio numeric, nueva_fecha_fin date,
  created_at timestamptz not null default now(), created_by uuid default auth.uid(),
  updated_at timestamptz not null default now(), updated_by uuid,
  constraint modificacion_con_cambio check (monto_cambio is not null or nueva_fecha_fin is not null)
);
create index if not exists contratos_mod_contrato_idx on public.contratos_cliente_modificaciones (contrato_id);
create trigger trg_contratos_mod_updated_at before update on public.contratos_cliente_modificaciones for each row execute function set_updated_at();
alter table public.transacciones add column if not exists contrato_id uuid references public.contratos_cliente(id) on delete set null;
alter table public.facturas_emitidas add column if not exists contrato_id uuid references public.contratos_cliente(id) on delete set null;
create index if not exists transacciones_contrato_idx on public.transacciones (contrato_id) where contrato_id is not null;
create index if not exists facturas_contrato_idx on public.facturas_emitidas (contrato_id) where contrato_id is not null;
alter table public.contratos_cliente enable row level security;
alter table public.contratos_cliente_modificaciones enable row level security;
create policy "contratos_cliente: leer" on public.contratos_cliente for select to authenticated using (tiene_acceso_unidad(unidad));
create policy "contratos_cliente: alta" on public.contratos_cliente for insert to authenticated with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "contratos_cliente: cambio" on public.contratos_cliente for update to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad)) with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "contratos_cliente: borrar" on public.contratos_cliente for delete to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "contratos_mod: leer" on public.contratos_cliente_modificaciones for select to authenticated using (tiene_acceso_unidad(unidad));
create policy "contratos_mod: alta" on public.contratos_cliente_modificaciones for insert to authenticated with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "contratos_mod: cambio" on public.contratos_cliente_modificaciones for update to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad)) with check (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
create policy "contratos_mod: borrar" on public.contratos_cliente_modificaciones for delete to authenticated using (tiene_acceso_unidad(unidad) and puede_editar_unidad(unidad));
alter publication supabase_realtime add table public.contratos_cliente;
alter publication supabase_realtime add table public.contratos_cliente_modificaciones;
