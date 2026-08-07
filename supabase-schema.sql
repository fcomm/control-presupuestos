-- Control de Presupuestos — esquema Supabase
-- Ejecutar completo en: Supabase Dashboard > SQL Editor > New query > Run

create extension if not exists "pgcrypto";

-- ---------------------------------------------------------------
-- PROYECTOS (catálogo de proyectos por unidad de negocio + % de prorrateo)
-- ---------------------------------------------------------------
create table if not exists proyectos (
  id uuid primary key default gen_random_uuid(),
  unidad text not null,               -- 'OSB' | 'CTM' | 'ISE'
  nombre text not null,
  grupo text,                         -- 'Desh' | 'Prod' | 'IMP' | etc.
  pct numeric not null default 0,
  created_at timestamptz not null default now()
);

-- ---------------------------------------------------------------
-- PARTIDAS (presupuesto)
-- ---------------------------------------------------------------
create table if not exists partidas (
  id uuid primary key default gen_random_uuid(),
  unidad text not null,
  mes text,
  anio int,
  smi text,
  concepto text not null,
  rubro text,
  categoria text,
  proyecto text,                      -- proyecto directo o marcador (Desh Gral / Prod Gral / Todos)
  monto_estimado numeric not null default 0,
  moneda text not null default 'MXN',
  folio text,                         -- UNIDAD-MESAÑO-### — liga con transacciones.folio_original
  created_at timestamptz not null default now()
);

create index if not exists idx_partidas_unidad on partidas (unidad);
create index if not exists idx_partidas_folio on partidas (folio);

-- ---------------------------------------------------------------
-- TRANSACCIONES (pagos reales ejecutados)
-- ---------------------------------------------------------------
create table if not exists transacciones (
  id uuid primary key default gen_random_uuid(),
  partida_id uuid references partidas(id) on delete set null,
  folio_original text,                -- folio tal como venía en el archivo importado
  unidad_detectada text,              -- unidad inferida del prefijo del folio
  dia date,
  solicitante text,
  area text,
  proveedor text,
  concepto_detallado text,
  importe numeric not null default 0,
  moneda text not null default 'MXN',
  created_at timestamptz not null default now()
);

create index if not exists idx_transacciones_partida on transacciones (partida_id);

-- ---------------------------------------------------------------
-- ROW LEVEL SECURITY
-- MVP: todo mundo con la anon key (el equipo, vía la app) puede leer/escribir.
-- No hay login todavía — cualquiera con la URL de la app tiene acceso completo.
-- Antes de compartir la app fuera del equipo, cambiar estas políticas por unas
-- que exijan autenticación (auth.uid() is not null).
-- ---------------------------------------------------------------
alter table proyectos enable row level security;
alter table partidas enable row level security;
alter table transacciones enable row level security;

create policy "proyectos_all_anon" on proyectos for all using (true) with check (true);
create policy "partidas_all_anon" on partidas for all using (true) with check (true);
create policy "transacciones_all_anon" on transacciones for all using (true) with check (true);

-- ---------------------------------------------------------------
-- REALTIME — para que los cambios de una persona se vean en vivo en las demás
-- ---------------------------------------------------------------
alter publication supabase_realtime add table proyectos;
alter publication supabase_realtime add table partidas;
alter publication supabase_realtime add table transacciones;

-- ---------------------------------------------------------------
-- SEED — proyectos de CTM ya acordados
-- ---------------------------------------------------------------
insert into proyectos (unidad, nombre, grupo, pct) values
  ('CTM', 'Desh Marfo',      'Desh', 10),
  ('CTM', 'Desh Pozoleo',    'Desh', 10),
  ('CTM', 'Desh Tamaulipas', 'Desh', 20),
  ('CTM', 'Desh Naranjos',   'Desh', 10),
  ('CTM', 'Desh Cacalilao',  'Desh', 10),
  ('CTM', 'Desh Poza Rica',  'Desh', 10),
  ('CTM', 'Prod Poza Rica',  'Prod', 10),
  ('CTM', 'Prod Veracruz',   'Prod', 10),
  ('CTM', 'IMP',             'IMP',  10)
on conflict do nothing;
