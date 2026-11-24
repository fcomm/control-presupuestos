-- HOMOLOGACIÓN por proveedor — OSB, CTM e ISE (24/11/26)
-- Aplica a TODAS las transacciones vigentes de los 70 nombres de proveedor de los lotes 1–4
-- la clasificación acordada. Sustituye a los lotes 3 y 4 (no hace falta correrlos).
-- Cambio respecto al lote 1: Telcel y AT&T Comunicaciones Digitales pasan a
-- Servicios Operativos › Servicios de Telecomunicaciones (decisión de Paco 24/11).
--
-- 587 transacciones · $18,843,641.19
--   OSB: 508 se mueven de partida · 29 solo cambian categoría
--   CTM:  28 se mueven            ·  7 solo cambian categoría
--   ISE:   8 se mueven            ·  7 solo cambian categoría
--
-- Qué hace el PASO 2:
--   · Si la transacción ya está en una partida del rubro correcto: solo le pone la categoría.
--   · Si no (No Clasificado u otro rubro): la mueve a la partida de ese rubro del mismo mes,
--     año, proyecto y moneda (la que ya exista; si no, crea "<Rubro> <Proyecto> <Mes> <Año>"
--     con el siguiente folio del mes de esa compañía) y le pone la categoría.
--   · Presupuesto: lo movido se resta de la partida de origen y se suma a la existente que lo
--     recibe (las nuevas nacen con lo que reciben). El total de cada compañía no cambia.
--   · Parcialidades PPD: siguen a su compra y reciben la misma categoría.
--   Todo o nada. Probado en una transacción revertida: 587 / 104 partidas nuevas /
--   76 existentes / 0 inconsistentes al final / presupuestos de OSB, CTM e ISE sin cambio.

------------------------------------------------------------------------
-- PASO 1 · PREVIO (solo lectura): una fila por transacción a cambiar
------------------------------------------------------------------------
with mapa(proveedor, rubro, categoria) as (values
('SITELSOFT','Servicios Operativos','Servicios especializados'),
('INNOVARQ URBITECH','Servicios Operativos','Servicios especializados'),
('SONDER AKTION DE MEXICO','Servicios Operativos','Servicios especializados'),
('INMOBILIARIA IZ2 S DE RL DE CV','Servicios Operativos','Arrendamientos'),
('ESPECIALIDADES EN TURBO MAQUINARIA','Servicios Operativos','Servicios especializados'),
('RICARDO POBLANO PAZ','Servicios Operativos','Servicios especializados'),
('SOLUCIONES INDUSTRIALES SERCOIN','Servicios Operativos','Servicios especializados'),
('RADIOMOVIL DIPSA SA DE CV','Servicios Operativos','Servicios de Telecomunicaciones'),
('INSTRU CONTROL','Materiales e Insumos','Instrumentación'),
('AT&T COMUNICACIONES DIGITALES S DE RL DE CV','Servicios Operativos','Servicios de Telecomunicaciones'),
('ORACLE DE MEXICO SA DE CV','Tecnologías de Información','Servicios en la nube'),
('MRS DEL BAJIO','Servicios Operativos','Servicios especializados'),
('CONSULTORES ADMINISTRATIVOS MARS','Servicios Administrativos','Consultoría'),
('HINOJOSA GOMEZ HUMBERTO','Servicios Administrativos','Consultoría'),
('TRACTOCAMIONES BASURTO SA DE CV','Servicios Operativos','Arrendamientos'),
('SEMICONDUCTORES QUERETARO','Servicios Operativos','Servicios especializados'),
('CIDRONIO CIPRIANO JIMENEZ GOVEA','Servicios Administrativos','Consultoría'),
('CIDRONIO CIPRIANO JIMENEZ GOVEA JOSEFA ORTIZ DE DOMINGUEZ 5530318783','Servicios Administrativos','Consultoría'),
('HERNANDEZ GONZALEZ RENE MOISES','Servicios Administrativos','Consultoría'),
('HERNANDEZ GONZALEZ RENE MOISES CALLE 5 55 5106 3257','Servicios Administrativos','Consultoría'),
('SUPER FLETES 23','Servicios Operativos','Transporte y logística'),
('RUBEN RODRIGUEZ OVIL','Servicios Operativos','Arrendamientos'),
('TELECOMUNICACION Y MERCADOTECNIA DE MONTERREY','Equipos y Herramientas','Equipos de medición'),
('MOUSER ELECTRONICS','Materiales e Insumos','Material eléctrico'),
('OCULUS HOLDING','Servicios Operativos','Servicios especializados'),
('HDI SEGUROS SA DE CV','Gastos Financieros e Impuestos','Seguros No Vehiculares'),
('SERVICIOS INDUSTRIALES OBRAS NAVALES','Servicios Operativos','Servicios especializados'),
('CHAMPION TECHNOLOGIES DE MEXICO SA DE CV','Servicios Operativos','Servicios especializados'),
('CHAMPION TECHNOLOGIES DE MEXICO','Servicios Operativos','Servicios especializados'),
('ENRIQUE SALAS','Servicios Administrativos','Consultoría'),
('DELOITE IMPUESTOS Y SERVICIOS LEGALES','Servicios Administrativos','Consultoría'),
('JOSE ALBERTO HERRERA ALARCON','Servicios Básicos','Recolección de residuos'),
('SERVICIOS CENTRALES DE COBRANZA HOTELERA SA DE CV','Viajes y Viáticos','Hospedaje'),
('MARISELA MARTINEZ RIOS','Servicios Operativos','Arrendamientos'),
('MARISELA MARTINEZ RIOS PRIVADA SAN CARLOS','Servicios Operativos','Arrendamientos'),
('PROVEEDORA INDUSTRIAL FIRE','Equipos y Herramientas','Herramienta especializada'),
('FEGUECHA SERVICIOS','Servicios Administrativos','Consultoría'),
('JOAQUIN SALCIDO HERNANDEZ','Servicios Operativos','Servicios especializados'),
('JOAQUIN SALCIDO HERNANDEZ PASEO SAN ISIDRO','Servicios Operativos','Servicios especializados'),
('SISTEMAS AMBIENTALES INDUSTRIALES DE MEXICO','Servicios Operativos','Servicios especializados'),
('SISTEMAS AMBIENTALES INDUSTRIALES DE MEXICO MARTIRES DE CANANEA 9931-16-23-02','Servicios Operativos','Servicios especializados'),
('INSTRUMENTOS WIKA MEXICO SA DE CV','Materiales e Insumos','Instrumentación'),
('INMOBILIARIA BALMAC','Servicios Operativos','Arrendamientos'),
('GUSTAVO CARRION RUIZ','Materiales e Insumos','Ensambles y Subensambles Metalicos'),
('EDWIN GOMEZ MARTINEZ','Servicios Operativos','Servicios especializados'),
('SISTEMAS Y SERVICIOS DE COMUNICACION SA DE CV','Equipos y Herramientas','Equipos de comunicación'),
('STARLINK SATELLITE SYSTEMS MEXICO','Servicios Operativos','Servicios de Telecomunicaciones'),
('ADT PRIVATE SECURITY SERVICE DE MEXICO SA DE CV','Servicios Administrativos','Vigilancia'),
('INTEGRON SERVICES SA DE CV','Servicios Operativos','Servicios especializados'),
('BANCO NACIONAL DE MEXICO SA','Gastos Financieros e Impuestos','Comisiones bancarias'),
('BANCO NACIONAL DE MEXICO SA ISABEL LA CATOLICA','Gastos Financieros e Impuestos','Comisiones bancarias'),
('TELEFONOS DE MEXICO SAB DE CV','Servicios Básicos','Telefonía fija'),
('WEBFLOW','Tecnologías de Información','Suscripciones'),
('OPENAI, LLC','Tecnologías de Información','Suscripciones'),
('AEROVIAS DE MEXICO SA DE CV','Viajes y Viáticos','Boletos de avión'),
('ZOHOCORP MEXICO','Tecnologías de Información','Suscripciones'),
('GOOGLE CLOUD MEXICO','Tecnologías de Información','Servicios en la nube'),
('ELIGIO DURAN MARISCAL','Servicios Operativos','Arrendamientos'),
('ELIGIO DURAN MARISCAL PASEO DE HABSBURGO','Servicios Operativos','Arrendamientos'),
('CYBERPUERTA SA DE CV','Equipos y Herramientas','Equipos de cómputo'),
('OPERADORA DE SERVICIOS HOTELEROS OSH SA DE CV','Viajes y Viáticos','Hospedaje'),
('OFFICE DEPOT DE MEXICO SA DE CV','Materiales e Insumos','Papelería'),
('TOTAL CLOUD','Tecnologías de Información','Servicios en la nube'),
('TOTAL CLOUD MARGARITA 8130542935','Tecnologías de Información','Servicios en la nube'),
('AT&T COMERCIALIZACION MOVIL S DE RL DE CV','Servicios Operativos','Servicios de Telecomunicaciones'),
('MULTIMOVIL SA DE CV','Servicios Operativos','Arrendamientos'),
('PEREZ MUÑOZ LUIS ALBERTO','Servicios Básicos','Telefonía fija'),
('PEREZ MUÑOZ LUIS ALBERTO VATE CARRILLO 4423546470','Servicios Básicos','Telefonía fija'),
('PITS ABASTECIMIENTOS','Vehículos','Combustible'),
('JESUS BALLESTEROS VILLA','Servicios Básicos','Agua')
)
select t.unidad_detectada as compania, t.folio_transaccion, t.dia, t.proveedor, t.importe, t.moneda,
       p.folio as partida_actual, p.rubro as rubro_actual, coalesce(t.categoria, '') as categoria_actual,
       m.rubro as rubro_nuevo, m.categoria as categoria_nueva,
       case when p.rubro = m.rubro then 'solo categoría' else 'mover de partida' end as accion
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.cancelada_en is null and t.transaccion_padre_id is null
  and not (p.rubro = m.rubro and coalesce(t.categoria, '') = m.categoria)
order by 1, accion, t.proveedor, t.dia;

------------------------------------------------------------------------
-- PASO 2 · APLICAR. Quita /* y */ y ejecuta el bloque completo.
------------------------------------------------------------------------
/*
create temp table _mov as
with mapa(proveedor, rubro, categoria) as (values
('SITELSOFT','Servicios Operativos','Servicios especializados'),
('INNOVARQ URBITECH','Servicios Operativos','Servicios especializados'),
('SONDER AKTION DE MEXICO','Servicios Operativos','Servicios especializados'),
('INMOBILIARIA IZ2 S DE RL DE CV','Servicios Operativos','Arrendamientos'),
('ESPECIALIDADES EN TURBO MAQUINARIA','Servicios Operativos','Servicios especializados'),
('RICARDO POBLANO PAZ','Servicios Operativos','Servicios especializados'),
('SOLUCIONES INDUSTRIALES SERCOIN','Servicios Operativos','Servicios especializados'),
('RADIOMOVIL DIPSA SA DE CV','Servicios Operativos','Servicios de Telecomunicaciones'),
('INSTRU CONTROL','Materiales e Insumos','Instrumentación'),
('AT&T COMUNICACIONES DIGITALES S DE RL DE CV','Servicios Operativos','Servicios de Telecomunicaciones'),
('ORACLE DE MEXICO SA DE CV','Tecnologías de Información','Servicios en la nube'),
('MRS DEL BAJIO','Servicios Operativos','Servicios especializados'),
('CONSULTORES ADMINISTRATIVOS MARS','Servicios Administrativos','Consultoría'),
('HINOJOSA GOMEZ HUMBERTO','Servicios Administrativos','Consultoría'),
('TRACTOCAMIONES BASURTO SA DE CV','Servicios Operativos','Arrendamientos'),
('SEMICONDUCTORES QUERETARO','Servicios Operativos','Servicios especializados'),
('CIDRONIO CIPRIANO JIMENEZ GOVEA','Servicios Administrativos','Consultoría'),
('CIDRONIO CIPRIANO JIMENEZ GOVEA JOSEFA ORTIZ DE DOMINGUEZ 5530318783','Servicios Administrativos','Consultoría'),
('HERNANDEZ GONZALEZ RENE MOISES','Servicios Administrativos','Consultoría'),
('HERNANDEZ GONZALEZ RENE MOISES CALLE 5 55 5106 3257','Servicios Administrativos','Consultoría'),
('SUPER FLETES 23','Servicios Operativos','Transporte y logística'),
('RUBEN RODRIGUEZ OVIL','Servicios Operativos','Arrendamientos'),
('TELECOMUNICACION Y MERCADOTECNIA DE MONTERREY','Equipos y Herramientas','Equipos de medición'),
('MOUSER ELECTRONICS','Materiales e Insumos','Material eléctrico'),
('OCULUS HOLDING','Servicios Operativos','Servicios especializados'),
('HDI SEGUROS SA DE CV','Gastos Financieros e Impuestos','Seguros No Vehiculares'),
('SERVICIOS INDUSTRIALES OBRAS NAVALES','Servicios Operativos','Servicios especializados'),
('CHAMPION TECHNOLOGIES DE MEXICO SA DE CV','Servicios Operativos','Servicios especializados'),
('CHAMPION TECHNOLOGIES DE MEXICO','Servicios Operativos','Servicios especializados'),
('ENRIQUE SALAS','Servicios Administrativos','Consultoría'),
('DELOITE IMPUESTOS Y SERVICIOS LEGALES','Servicios Administrativos','Consultoría'),
('JOSE ALBERTO HERRERA ALARCON','Servicios Básicos','Recolección de residuos'),
('SERVICIOS CENTRALES DE COBRANZA HOTELERA SA DE CV','Viajes y Viáticos','Hospedaje'),
('MARISELA MARTINEZ RIOS','Servicios Operativos','Arrendamientos'),
('MARISELA MARTINEZ RIOS PRIVADA SAN CARLOS','Servicios Operativos','Arrendamientos'),
('PROVEEDORA INDUSTRIAL FIRE','Equipos y Herramientas','Herramienta especializada'),
('FEGUECHA SERVICIOS','Servicios Administrativos','Consultoría'),
('JOAQUIN SALCIDO HERNANDEZ','Servicios Operativos','Servicios especializados'),
('JOAQUIN SALCIDO HERNANDEZ PASEO SAN ISIDRO','Servicios Operativos','Servicios especializados'),
('SISTEMAS AMBIENTALES INDUSTRIALES DE MEXICO','Servicios Operativos','Servicios especializados'),
('SISTEMAS AMBIENTALES INDUSTRIALES DE MEXICO MARTIRES DE CANANEA 9931-16-23-02','Servicios Operativos','Servicios especializados'),
('INSTRUMENTOS WIKA MEXICO SA DE CV','Materiales e Insumos','Instrumentación'),
('INMOBILIARIA BALMAC','Servicios Operativos','Arrendamientos'),
('GUSTAVO CARRION RUIZ','Materiales e Insumos','Ensambles y Subensambles Metalicos'),
('EDWIN GOMEZ MARTINEZ','Servicios Operativos','Servicios especializados'),
('SISTEMAS Y SERVICIOS DE COMUNICACION SA DE CV','Equipos y Herramientas','Equipos de comunicación'),
('STARLINK SATELLITE SYSTEMS MEXICO','Servicios Operativos','Servicios de Telecomunicaciones'),
('ADT PRIVATE SECURITY SERVICE DE MEXICO SA DE CV','Servicios Administrativos','Vigilancia'),
('INTEGRON SERVICES SA DE CV','Servicios Operativos','Servicios especializados'),
('BANCO NACIONAL DE MEXICO SA','Gastos Financieros e Impuestos','Comisiones bancarias'),
('BANCO NACIONAL DE MEXICO SA ISABEL LA CATOLICA','Gastos Financieros e Impuestos','Comisiones bancarias'),
('TELEFONOS DE MEXICO SAB DE CV','Servicios Básicos','Telefonía fija'),
('WEBFLOW','Tecnologías de Información','Suscripciones'),
('OPENAI, LLC','Tecnologías de Información','Suscripciones'),
('AEROVIAS DE MEXICO SA DE CV','Viajes y Viáticos','Boletos de avión'),
('ZOHOCORP MEXICO','Tecnologías de Información','Suscripciones'),
('GOOGLE CLOUD MEXICO','Tecnologías de Información','Servicios en la nube'),
('ELIGIO DURAN MARISCAL','Servicios Operativos','Arrendamientos'),
('ELIGIO DURAN MARISCAL PASEO DE HABSBURGO','Servicios Operativos','Arrendamientos'),
('CYBERPUERTA SA DE CV','Equipos y Herramientas','Equipos de cómputo'),
('OPERADORA DE SERVICIOS HOTELEROS OSH SA DE CV','Viajes y Viáticos','Hospedaje'),
('OFFICE DEPOT DE MEXICO SA DE CV','Materiales e Insumos','Papelería'),
('TOTAL CLOUD','Tecnologías de Información','Servicios en la nube'),
('TOTAL CLOUD MARGARITA 8130542935','Tecnologías de Información','Servicios en la nube'),
('AT&T COMERCIALIZACION MOVIL S DE RL DE CV','Servicios Operativos','Servicios de Telecomunicaciones'),
('MULTIMOVIL SA DE CV','Servicios Operativos','Arrendamientos'),
('PEREZ MUÑOZ LUIS ALBERTO','Servicios Básicos','Telefonía fija'),
('PEREZ MUÑOZ LUIS ALBERTO VATE CARRILLO 4423546470','Servicios Básicos','Telefonía fija'),
('PITS ABASTECIMIENTOS','Vehículos','Combustible'),
('JESUS BALLESTEROS VILLA','Servicios Básicos','Agua')
)
select t.id as tid, t.unidad_detectada as unidad, t.folio_transaccion, t.proveedor, t.importe, t.es_padre,
       m.rubro, m.categoria, p.id as origen_id, p.folio as origen_folio, p.rubro as rubro_origen,
       p.mes, p.anio, p.proyecto, p.moneda,
       case when p.rubro = m.rubro then 'categoria' else 'mover' end as accion
from public.transacciones t
join public.partidas p on p.id = t.partida_id
join mapa m on m.proveedor = upper(trim(t.proveedor))
where t.cancelada_en is null and t.transaccion_padre_id is null
  and not (p.rubro = m.rubro and coalesce(t.categoria, '') = m.categoria);

create temp table _g as   -- lo que sale de cada partida de origen (solo las que se mueven)
select origen_id, round(sum(importe), 2) as monto from _mov where accion = 'mover' group by 1;

create temp table _dest as  -- un destino por compañía + rubro + proyecto + mes + año + moneda
select d.*, d.unidad || '-' || (array['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])[array_position(array['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'], d.mes)] || right(d.anio::text, 2) || '-' as prefijo,
       (select p.id from public.partidas p where p.unidad = d.unidad and p.rubro = d.rubro and p.mes = d.mes
          and p.anio = d.anio and p.proyecto = d.proyecto and p.moneda = d.moneda order by p.folio limit 1) as existente_id
from (select unidad, mes, anio, proyecto, moneda, rubro, count(*) as n, round(sum(importe), 2) as monto
      from _mov where accion = 'mover' group by 1,2,3,4,5,6) d;

create temp table _nuevas as
select gen_random_uuid() as nid, d.*,
       d.prefijo || lpad((coalesce((select max(substring(p.folio from '-(\d+)$')::int) from public.partidas p
                                    where p.unidad = d.unidad and p.folio like d.prefijo || '%'), 0)
                          + row_number() over (partition by d.prefijo order by d.proyecto, d.rubro))::text, 3, '0') as folio_nuevo,
       d.rubro || ' ' || d.proyecto || ' ' || d.mes || ' ' || d.anio as concepto_nuevo
from _dest d where d.existente_id is null;

create temp table _final as
select d.unidad, d.mes, d.anio, d.proyecto, d.moneda, d.rubro, d.n, d.monto,
       coalesce(d.existente_id, nv.nid) as destino_id,
       case when d.existente_id is null then 'nueva' else 'existente' end as tipo
from _dest d left join _nuevas nv
  on (nv.unidad, nv.mes, nv.anio, nv.proyecto, nv.moneda, nv.rubro) = (d.unidad, d.mes, d.anio, d.proyecto, d.moneda, d.rubro);

do $$
declare n_mov int; n_upd int;
begin
  select count(*) into n_mov from _mov;
  if n_mov <> 587 then raise exception 'Se esperaban 587 transacciones por homologar y hay %. No se cambió nada.', n_mov; end if;

  insert into public.partidas (id, unidad, mes, anio, smi, concepto, rubro, categoria, proyecto, monto_estimado, moneda, folio, created_by, extra)
  select nid, unidad, mes, anio, '', concepto_nuevo, rubro, null, proyecto, monto, moneda, folio_nuevo, 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
         jsonb_build_object('origen', 'Homologación por proveedor nov 2026')
  from _nuevas;

  -- Las que se mueven de partida (y quedan con su categoría)
  update public.transacciones t
  set partida_id = f.destino_id, categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m join _final f
    on (f.unidad, f.mes, f.anio, f.proyecto, f.moneda, f.rubro) = (m.unidad, m.mes, m.anio, m.proyecto, m.moneda, m.rubro)
  where m.accion = 'mover' and t.id = m.tid;
  get diagnostics n_upd = row_count;

  -- Las que ya están en una partida del rubro correcto: solo la categoría
  update public.transacciones t set categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m where m.accion = 'categoria' and t.id = m.tid;
  n_upd := n_upd + (select count(*) from _mov where accion = 'categoria');
  if n_upd <> 587 then raise exception 'Se actualizaron % en vez de 587. No se cambió nada.', n_upd; end if;

  -- Parcialidades de compras PPD: la base les cambia la partida con su compra; falta la categoría
  update public.transacciones h set categoria = m.categoria, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _mov m where m.es_padre and h.transaccion_padre_id = m.tid;

  -- Presupuesto: el monto movido sale de la partida de origen y entra a la existente que lo recibe
  update public.partidas p set monto_estimado = p.monto_estimado - g.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _g g where p.id = g.origen_id;
  update public.partidas p set monto_estimado = p.monto_estimado + f.monto, updated_by = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb'
  from _final f where f.tipo = 'existente' and p.id = f.destino_id;
end $$;

-- Resultado: 587 transacciones, todas con el rubro y la categoría de su proveedor
select t.unidad_detectada as compania, count(*) as transacciones, round(sum(t.importe), 2) as importe
from public.transacciones t join _mov m on m.tid = t.id group by 1 order by 1;

drop table _mov; drop table _g; drop table _dest; drop table _nuevas; drop table _final;
*/
