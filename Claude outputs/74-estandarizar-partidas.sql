-- 74: Estandarizacion de nombres de partidas (YA APLICADO el 27/11/2026). Respaldos: respaldo_partidas_20261127, respaldo_tx_partida_20261127, respaldo_mapa_partidas_20261127.

create table respaldo_partidas_20261127 as select * from partidas;
create table respaldo_tx_partida_20261127 as select id, partida_id from transacciones;
create table respaldo_mapa_partidas_20261127 as
with base as (
  select p.id, p.unidad u, p.anio, p.mes, p.rubro, coalesce(p.folio,'') folio,
         case when coalesce(p.moneda,'MXP')='USD' then 'USD' else 'MXP' end mon,
         upper(translate(trim(coalesce(p.zona,'')),'áéíóúÁÉÍÓÚ','aeiouAEIOU')) zn,
         upper(translate(coalesce(p.concepto,''),'áéíóúÁÉÍÓÚñÑ','aeiouAEIOUnN')) cn,
         (select count(*) from transacciones t where t.partida_id=p.id) ntx
  from partidas p),
zmap(z,pref) as (values ('QUERETARO','QRO'),('QRO','QRO'),('CDMX','CDMX'),('POZA RICA','PR'),('COTAXTLA','PR'),('ALTAMIRA','PR'),
  ('CERRO AZUL','PR'),('TAMAULIPAS','PR'),('PARAISO','PAR'),('VILLA HERMOSA','PAR')),
txz as (
  select t.partida_id, zm.pref, count(*) n, min(t.created_at) primero
  from transacciones t join zmap zm on zm.z = upper(translate(trim(t.zona),'áéíóúÁÉÍÓÚ','aeiouAEIOU'))
  where t.cancelada_en is null and t.categoria is not null and coalesce(trim(t.zona),'')<>''
  group by 1,2),
txz1 as (select distinct on (partida_id) partida_id, pref from txz order by partida_id, n desc, primero),
cat as (
  select t.partida_id, count(*) n,
    count(*) filter (where upper(translate(t.categoria,'áéíóúÁÉÍÓÚ','aeiouAEIOU')) like '%COMBUST%') comb,
    count(*) filter (where upper(translate(t.categoria,'áéíóúÁÉÍÓÚ','aeiouAEIOU')) ~ 'MANT|REFACC|LLANTA|ACUMUL') mant
  from transacciones t where t.cancelada_en is null and t.categoria is not null group by 1),
z as (
  select b.*, coalesce(
    case when b.u='OSB' then 'QRO' end,
    (select pref from zmap where z=b.zn),
    case substring(b.cn from '^\s*(REVPR|REVQRO|PAR|PR|QRO)\s*-') when 'REVPR' then 'PR' when 'REVQRO' then 'QRO' else substring(b.cn from '^\s*(REVPR|REVQRO|PAR|PR|QRO)\s*-') end,
    (select pref from txz1 where txz1.partida_id=b.id),
    case when b.cn ~ '\mQRO\M|QUERET' then 'QRO'
         when b.cn ~ 'POZA RICA|COTAXTLA|ALTAMIRA|CERRO AZUL|VERACRUZ|VER\.' then 'PR'
         when b.cn ~ 'PARAISO|HALLIBURTON|HOKCHI|SCHLUMBER|MARINA|\mSUR\M' then 'PAR' end,
    'QRO') zp0,
    case when b.rubro <> 'Vehículos' then ''
         when c.n > 0 then case when c.comb::numeric/c.n >= 0.8 then 'Combustibles' when c.mant::numeric/c.n >= 0.8 then 'Mantenimientos' else '' end
         when b.cn like '%COMBUST%' then 'Combustibles'
         when b.cn ~ 'MANT|MTTO|REFACC' then 'Mantenimientos' else '' end det
  from base b left join cat c on c.partida_id=b.id),
z2 as (select z.*, case when z.id='e2c5bb3c-7663-49dd-8de1-717aaa980cbc' then 'QRO' else z.zp0 end zp from z),
n as (select z.*, z.zp||'-'||z.rubro||case when z.det<>'' then ' · '||z.det else '' end||case when z.mon='USD' then ' USD' else '' end nuevo from z2 z),
mapa as (
  select n.id, n.u, first_value(n.id) over (partition by n.u, n.anio, n.mes, n.nuevo order by n.ntx desc, n.folio) sob, n.nuevo
  from n)
select id, sob, nuevo, u from mapa;
alter table respaldo_partidas_20261127 enable row level security;
alter table respaldo_tx_partida_20261127 enable row level security;
alter table respaldo_mapa_partidas_20261127 enable row level security;
do $$
declare a int; b int;
begin
  select count(*), count(distinct sob) into a, b from respaldo_mapa_partidas_20261127;
  if a <> 947 or b <> 486 then raise exception 'mapa inesperado % %', a, b; end if;
end $$;
update transacciones t set partida_id = m.sob
  from respaldo_mapa_partidas_20261127 m
 where t.partida_id = m.id and m.id <> m.sob and t.transaccion_padre_id is null;
update partidas p set monto_estimado = s.total, concepto = s.nuevo
  from (select m.sob, min(m.nuevo) nuevo, sum(r.monto_estimado) total
          from respaldo_mapa_partidas_20261127 m join respaldo_partidas_20261127 r on r.id = m.id
         group by m.sob) s
 where p.id = s.sob;
do $$
declare n int;
begin
  select count(*) into n from transacciones t join respaldo_mapa_partidas_20261127 m on m.id = t.partida_id where m.id <> m.sob;
  if n > 0 then raise exception 'quedan % transacciones en partidas fusionadas', n; end if;
end $$;
delete from partidas p using respaldo_mapa_partidas_20261127 m where p.id = m.id and m.id <> m.sob;
