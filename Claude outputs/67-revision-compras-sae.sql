-- 67 · Revisión de compras de Aspel SAE (TEMPORAL)  — ya aplicada en Supabase el 19/11/26
-- Tabla aparte, sin liga con ninguna tabla de la app. Solo administradores.
-- Para retirarla:  drop function public.revision_compras_sae_vaciar(text);
--                  drop table public.revision_compras_sae;
create table if not exists public.revision_compras_sae (
  id            bigint generated always as identity primary key,
  compania      text not null default 'OSB',
  cve_doc text, serie text, folio integer, status text, tipo text,
  cve_clpv integer, proveedor text, rfc text, su_refer text,
  fecha_doc date, fecha_rec date, fecha_pag date, fecha_cancela date, fecha_elab timestamp,
  num_alma integer,
  subtotal numeric, descuento numeric, ret_isr numeric, ret_iva numeric, iva numeric, importe numeric,
  moneda text, tipcamb numeric, importe_usd numeric,
  escfd text, bloq text, cve_obs integer, obs_cond text, doc_ant text, forma_envio text, metodo_pago text,
  cargado_en timestamptz not null default now()
);
create index if not exists revision_compras_sae_compania_idx on public.revision_compras_sae (compania);

alter table public.revision_compras_sae enable row level security;
create policy "revision_compras_sae: admin lee" on public.revision_compras_sae
  for select to authenticated using (public.es_admin());
create policy "revision_compras_sae: admin carga" on public.revision_compras_sae
  for insert to authenticated with check (public.es_admin());
create policy "revision_compras_sae: admin borra" on public.revision_compras_sae
  for delete to authenticated using (public.es_admin());

-- Vacía los datos de una compañía antes de recargar su archivo.
create or replace function public.revision_compras_sae_vaciar(p_compania text)
returns integer language plpgsql security definer set search_path = public as $$
declare n integer;
begin
  if not es_admin() then
    raise exception 'Solo un administrador puede recargar la revisión de compras.';
  end if;
  delete from public.revision_compras_sae where compania = p_compania;
  get diagnostics n = row_count;
  return n;
end $$;
revoke all on function public.revision_compras_sae_vaciar(text) from public, anon;
grant execute on function public.revision_compras_sae_vaciar(text) to authenticated;
