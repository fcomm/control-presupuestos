-- 70 · Nueva unidad de negocio CT — ya aplicada en Supabase el 25/11/26
-- 1) puede_editar_unidad: la lista de compañías por omisión incluye CT.
create or replace function public.puede_editar_unidad(u text)
 returns boolean language sql stable security definer set search_path to 'public'
as $function$
  select coalesce((
    select case
      when coalesce(u, '') = '' then exists (
        select 1
        from unnest(case when p.unidades_permitidas is null or cardinality(p.unidades_permitidas) = 0
                         then array['OSB','CTM','ISE','CT'] else p.unidades_permitidas end) x
        where coalesce(p.permisos->>x, 'edicion') <> 'lectura')
      else coalesce(p.permisos->>u, 'edicion') <> 'lectura'
    end
    from perfiles p where p.id = auth.uid()
  ), false);
$function$;
-- 2) Vehículos: CT entre las compañías válidas.
do $$
declare c text;
begin
  select conname into c from pg_constraint
   where conrelid = 'public.vehiculos'::regclass and contype = 'c' and pg_get_constraintdef(oid) ilike '%IZ2%';
  if c is not null then execute format('alter table public.vehiculos drop constraint %I', c); end if;
end $$;
alter table public.vehiculos add constraint vehiculos_compania_check
  check (compania is null or compania = any (array['OSB','CTM','ISE','CT','IZ2','JEF']));
