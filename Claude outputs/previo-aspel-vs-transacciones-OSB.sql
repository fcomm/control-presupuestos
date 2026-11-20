-- PREVIO (solo lectura): compras de Aspel (revision_compras_sae) contra las
-- transacciones de OSB en la app. NO modifica nada.
--
-- Cómo se emparejan, uno a uno:
--   1) Mismo folio de compra SAE (en una compra PPD se toma el de sus parcialidades).
--   2) Mismo proveedor (RFC del catálogo o nombre), importe a ±$1 y fecha a ±60 días.
-- Las parcialidades PPD no se comparan solas: van con su compra padre.
with
compras as (
  select * from public.revision_compras_sae
  where compania = 'OSB' and fecha_doc >= '2025-01-01'          -- periodo a comparar
),
corte as (select max(fecha_doc) as ultima from compras),
trans as (
  select t.id, t.folio_transaccion, t.dia, t.proveedor, t.importe, t.moneda, t.status,
         t.es_padre,
         upper(trim(coalesce(nullif(t.folio_compra_sae, ''),
               (select max(nullif(h.folio_compra_sae, '')) from public.transacciones h
                 where h.transaccion_padre_id = t.id and h.folio_compra_sae ~* '^O[NRE]\d+$')))) as folio_sae,
         upper(trim(coalesce(pr.rfc, ''))) as rfc,
         upper(trim(t.proveedor)) as nombre
  from public.transacciones t
  left join public.proveedores pr on pr.id = t.proveedor_id
  where t.unidad_detectada = 'OSB' and t.cancelada_en is null
    and t.transaccion_padre_id is null
),
candidatos as (
  select t.id as t_id, c.id as c_id,
         case when t.folio_sae = c.cve_doc then 1 else 2 end as regla,
         abs(c.fecha_doc - t.dia) as dias
  from trans t
  join compras c
    on t.folio_sae = c.cve_doc
    or ((t.rfc <> '' and t.rfc = c.rfc) or t.nombre = upper(c.proveedor))
       and abs(c.importe - t.importe) < 1 and abs(c.fecha_doc - t.dia) <= 60
),
pares as (
  select * from (
    select k.*,
           row_number() over (partition by t_id order by regla, dias) as rt,
           row_number() over (partition by c_id order by regla, dias) as rc
    from candidatos k
  ) x where rt = 1 and rc = 1
)
select
  case
    when p.t_id is not null and t.es_padre and abs(c.importe - t.importe) >= 1
                                                    then '1b. PPD · coincide por folio, importe distinto'
    when p.t_id is not null and t.es_padre          then '1. PPD · coincide'
    when p.t_id is not null and abs(c.importe - t.importe) >= 1
                                                    then '2b. Coincide por folio, importe distinto'
    when p.t_id is not null                         then '2. Coincide'
    when t.id is not null and t.es_padre            then '3. PPD · solo en la app'
    when t.id is not null and t.dia > (select ultima from corte)
                                                    then '4b. Solo en la app · posterior al corte de Aspel'
    when t.id is not null                           then '4. Solo en la app'
    else                                                 '5. Solo en Aspel'
  end                                               as resultado,
  case p.regla when 1 then 'folio SAE' when 2 then 'proveedor+importe+fecha' end as como_coincide,
  t.folio_transaccion,
  t.dia                                             as dia_app,
  t.importe                                         as importe_app,
  t.status                                          as status_app,
  (select count(*) from public.transacciones h
    where h.transaccion_padre_id = t.id and h.cancelada_en is null) as parcialidades,
  c.cve_doc,
  c.fecha_doc                                       as fecha_sae,
  c.importe                                         as importe_sae,
  c.status                                          as status_sae,
  coalesce(t.proveedor, c.proveedor)                as proveedor,
  c.rfc
from trans t
full join pares p   on p.t_id = t.id
full join compras c on c.id = p.c_id
order by resultado, coalesce(t.dia, c.fecha_doc) desc;
