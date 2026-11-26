-- 73 · Contratos con varios proyectos (centros de costo) — ya aplicada en Supabase el 26/11/26
alter table public.contratos_cliente add column if not exists proyectos text[] not null default '{}';
update public.contratos_cliente set proyectos = array[proyecto] where proyecto is not null and trim(proyecto) <> '' and cardinality(proyectos) = 0;
