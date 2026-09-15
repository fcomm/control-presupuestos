-- ============================================================================
-- Migración 24 — Módulo de Contratos
-- Requiere: migración 23 (zonas-por-companias)
-- Alcance: OSB, CTM, ISE. IZ2 y JEF quedan fuera, como en Vehículos.
--
-- BORRADOR. Antes de aplicar, confirmar:
--   1) nombre real de la tabla de transacciones y su columna de proveedor y monto
--      (marcado como TODO en la vista contratos_acumulado_proveedor)
--   2) nombre y tipo de la columna de unidad de negocio usada en el resto del esquema
-- ============================================================================

begin;

-- ---------------------------------------------------------------------------
-- 1. Datos legales por unidad de negocio
-- ---------------------------------------------------------------------------
create table if not exists contratos_datos_unidad (
  unidad                text primary key,
  razon_social          text not null,
  rfc                   text not null,
  domicilio             text not null,
  representante         text not null,
  cargo_representante   text,
  escritura             text,
  notario               text,
  correo_notificaciones text,
  correo_facturacion    text,
  jurisdiccion          text not null default 'Zapopan, Jalisco',
  actualizado_en        timestamptz not null default now(),
  actualizado_por       uuid references auth.users(id)
);

comment on table contratos_datos_unidad is
  'Datos legales de la empresa contratante. Una fila por unidad de negocio.';

-- ---------------------------------------------------------------------------
-- 2. Parámetros versionados (umbrales, penas, plazos)
--    Versionados por vigente_desde: los instrumentos ya emitidos conservan
--    los parámetros con los que se generaron.
-- ---------------------------------------------------------------------------
create table if not exists contratos_parametros (
  id                      bigserial primary key,
  unidad                  text not null references contratos_datos_unidad(unidad),
  vigente_desde           date not null,

  umbral_operacion        numeric(14,2) not null default 50000,
  umbral_acumulado        numeric(14,2) not null default 200000,
  ventana_meses           int           not null default 12,
  alerta_pct              numeric(5,2)  not null default 80,
  ventana_antifrac_dias   int           not null default 30,

  pena_diaria_pct         numeric(5,2)  not null default 1,
  pena_tope_pct           numeric(5,2)  not null default 10,
  plazo_pago_dias         int           not null default 30,
  plazo_aceptacion_dias   int           not null default 3,
  dias_rechazo            int           not null default 5,
  garantia_pct            numeric(5,2)  not null default 10,
  garantia_plazo_dias     int           not null default 10,
  garantia_vicios_meses   int           not null default 12,
  confidencialidad_anos   int           not null default 5,
  auditoria_anos          int           not null default 5,
  aviso_terminacion_dias  int           not null default 15,

  creado_en               timestamptz not null default now(),
  creado_por              uuid references auth.users(id)
);

create unique index if not exists contratos_parametros_unidad_fecha_idx
  on contratos_parametros (unidad, vigente_desde);

comment on column contratos_parametros.vigente_desde is
  'Nunca editar una fila vigente: insertar una nueva con fecha posterior.';

-- ---------------------------------------------------------------------------
-- 3. Datos legales del proveedor
--    Complementa proveedores_maestro (ASPEL) con lo que ese catálogo no trae.
--    Índice compuesto por unidad, siguiendo el criterio de la migración 23.
-- ---------------------------------------------------------------------------
create table if not exists contratos_proveedor_legal (
  id                   bigserial primary key,
  unidad               text not null references contratos_datos_unidad(unidad),
  rfc                  text not null,
  razon_social         text not null,
  personalidad         text,
  domicilio            text,
  representante        text,
  escritura            text,
  contacto             text,
  correo               text,

  repse_registro       text,
  repse_vigencia       date,
  opinion_32d_fecha    date,
  opinion_32d_sentido  text check (opinion_32d_sentido in ('positiva','negativa','sin_obligaciones','suspendido')),
  beneficiario_ctrl    boolean not null default false,
  nivel_dd             text not null default 'estandar'
                         check (nivel_dd in ('simplificada','estandar','reforzada')),
  grupo_economico      text,

  notas                text,
  actualizado_en       timestamptz not null default now(),
  actualizado_por      uuid references auth.users(id)
);

create unique index if not exists contratos_proveedor_legal_unidad_rfc_idx
  on contratos_proveedor_legal (unidad, upper(trim(rfc)));

create index if not exists contratos_proveedor_legal_grupo_idx
  on contratos_proveedor_legal (unidad, grupo_economico);

comment on column contratos_proveedor_legal.grupo_economico is
  'Agrupa RFCs del mismo beneficiario controlador para el umbral acumulado.';

-- ---------------------------------------------------------------------------
-- 4. Instrumentos generados
--    respuestas: jsonb con las 11 casillas del árbol, para auditoría.
--    parametros_id: congela los umbrales usados al momento de generar.
-- ---------------------------------------------------------------------------
create table if not exists contratos_instrumentos (
  id                bigserial primary key,
  folio             text not null,
  unidad            text not null references contratos_datos_unidad(unidad),
  rfc_proveedor     text not null,

  tipo              text not null check (tipo in (
                      'orden_compra',
                      'contrato_especifico',
                      'contrato_especifico_flowdown',
                      'contrato_marco'
                    )),
  objeto            text not null,
  monto             numeric(14,2),
  moneda            text not null default 'MXN',
  vigencia_inicio   date,
  vigencia_fin      date,
  monto_maximo      numeric(14,2),

  respuestas        jsonb not null default '{}'::jsonb,
  ruta_decision     text,
  parametros_id     bigint references contratos_parametros(id),

  estado            text not null default 'borrador' check (estado in (
                      'borrador','generado','en_firma','vigente','concluido','rescindido'
                    )),
  archivo_path      text,
  administrador     text,
  centro_costo      text,

  creado_en         timestamptz not null default now(),
  creado_por        uuid references auth.users(id)
);

create unique index if not exists contratos_instrumentos_folio_idx
  on contratos_instrumentos (unidad, folio);

create index if not exists contratos_instrumentos_prov_idx
  on contratos_instrumentos (unidad, upper(trim(rfc_proveedor)), vigencia_fin);

comment on column contratos_instrumentos.ruta_decision is
  'Texto legible de la compuerta que determinó el instrumento. Para auditoría.';

-- ---------------------------------------------------------------------------
-- 5. Órdenes emitidas al amparo de un contrato marco
-- ---------------------------------------------------------------------------
create table if not exists contratos_ordenes (
  id              bigserial primary key,
  folio           text not null,
  instrumento_id  bigint not null references contratos_instrumentos(id) on delete restrict,
  descripcion     text not null,
  importe         numeric(14,2) not null,
  moneda          text not null default 'MXN',
  lugar_entrega   text,
  fecha_entrega   date,
  estado          text not null default 'emitida' check (estado in (
                    'emitida','aceptada','recibida','pagada','cancelada'
                  )),
  archivo_path    text,
  creado_en       timestamptz not null default now(),
  creado_por      uuid references auth.users(id)
);

create unique index if not exists contratos_ordenes_folio_idx
  on contratos_ordenes (instrumento_id, folio);

-- ---------------------------------------------------------------------------
-- 6. Plantillas en Storage
-- ---------------------------------------------------------------------------
create table if not exists contratos_plantillas (
  id            bigserial primary key,
  tipo          text not null,
  version       text not null,
  storage_path  text not null,
  activa        boolean not null default false,
  notas         text,
  creado_en     timestamptz not null default now()
);

create unique index if not exists contratos_plantillas_tipo_version_idx
  on contratos_plantillas (tipo, version);

create unique index if not exists contratos_plantillas_activa_idx
  on contratos_plantillas (tipo) where activa;

-- ---------------------------------------------------------------------------
-- 7. Acumulado por proveedor en ventana móvil
--    TODO: sustituir <tabla_transacciones>, <col_rfc>, <col_monto>, <col_fecha>
--    y <col_unidad> por los nombres reales del esquema.
-- ---------------------------------------------------------------------------
-- create or replace view contratos_acumulado_proveedor as
-- select
--   t.<col_unidad>                                as unidad,
--   upper(trim(t.<col_rfc>))                      as rfc,
--   coalesce(pl.grupo_economico, upper(trim(t.<col_rfc>))) as grupo,
--   sum(t.<col_monto>)                            as acumulado_12m,
--   count(*)                                      as operaciones_12m,
--   max(t.<col_fecha>)                            as ultima_operacion
-- from <tabla_transacciones> t
-- left join contratos_proveedor_legal pl
--   on pl.unidad = t.<col_unidad>
--  and upper(trim(pl.rfc)) = upper(trim(t.<col_rfc>))
-- where t.<col_fecha> >= (current_date - interval '12 months')
-- group by 1, 2, 3;

-- ---------------------------------------------------------------------------
-- 8. RLS
-- ---------------------------------------------------------------------------
alter table contratos_datos_unidad      enable row level security;
alter table contratos_parametros        enable row level security;
alter table contratos_proveedor_legal   enable row level security;
alter table contratos_instrumentos      enable row level security;
alter table contratos_ordenes           enable row level security;
alter table contratos_plantillas        enable row level security;

do $$
declare t text;
begin
  foreach t in array array[
    'contratos_datos_unidad','contratos_parametros','contratos_proveedor_legal',
    'contratos_instrumentos','contratos_ordenes','contratos_plantillas'
  ] loop
    execute format(
      'create policy %I on %I for all to authenticated using (true) with check (true)',
      t || '_auth_all', t
    );
  end loop;
end $$;

-- ---------------------------------------------------------------------------
-- 9. Semilla de parámetros
-- ---------------------------------------------------------------------------
insert into contratos_parametros (unidad, vigente_desde)
select u.unidad, current_date
from (values ('OSB'), ('CTM'), ('ISE')) as u(unidad)
where exists (select 1 from contratos_datos_unidad d where d.unidad = u.unidad)
on conflict do nothing;

commit;
