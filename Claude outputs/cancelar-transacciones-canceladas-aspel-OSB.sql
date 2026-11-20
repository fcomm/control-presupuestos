-- Transacciones de OSB que solo coinciden con documentos CANCELADOS en Aspel
-- (archivo ComprasOsbog5-ExceptoCompras). 28 transacciones, todas Pagado.
-- Cruce: mismo proveedor (RFC o nombre), importe ±$1 y fecha ±60 días; OSB-SEP26-003 por su folio SAE
-- (ahí el importe NO coincide: $423.98 en la app contra $7,192.70 en Aspel).
-- Se excluyeron las que además coinciden con una compra válida de Aspel.

------------------------------------------------------------------------
-- PASO 1 · PREVIO (solo lectura). Revisa y borra de la lista los folios
-- que NO quieras cancelar, aquí y en el paso 2 (es la misma lista).
------------------------------------------------------------------------
with lista(folio, doc_aspel, importe_aspel, fecha_aspel) as (values
  ('OSB-MAR25-056', 'ON03842', 8622.44, date '2025-03-31'),
  ('OSB-ABR25-050', 'ON03904', 4610072.27, date '2025-04-29'),
  ('OSB-ABR25-056', 'ON03901', 5520.93, date '2025-04-30'),
  ('OSB-ABR25-076', 'ON03881', 1045.6, date '2025-04-30'),
  ('OSB-MAY25-004', 'ON03915', 867316.17, date '2025-05-02'),
  ('OSB-MAY25-006', 'ON03914', 962676.07, date '2025-05-02'),
  ('OSB-MAY25-014', 'ON03985', 2380.0, date '2025-05-08'),
  ('OSB-MAY25-048', 'ON04041', 11581.92, date '2025-05-30'),
  ('OSB-AGO25-014', 'ON04154', 19744.79, date '2025-08-07'),
  ('OSB-OCT25-027', 'ON04319', 5153.93, date '2025-10-13'),
  ('OSB-OCT25-062', 'ON04348', 1411.0, date '2025-10-31'),
  ('OSB-NOV25-010', 'ON04418', 3367.1, date '2025-11-24'),
  ('OSB-DIC25-001', 'ON04495', 1618.34, date '2025-12-29'),
  ('OSB-DIC25-088', 'ON04611', 35357.86, date '2025-12-31'),
  ('OSB-ENE26-052', 'ON04642', 34253.36, date '2026-01-31'),
  ('OSB-FEB26-033', 'ON04670', 1317.1, date '2026-02-22'),
  ('OSB-FEB26-037', 'ON04717', 79301.58, date '2026-02-25'),
  ('OSB-MAR26-024', 'ON04728', 2554.39, date '2026-03-11'),
  ('OSB-ABR26-037', 'ON04801', 186.76, date '2026-04-16'),
  ('OSB-ABR26-096', 'ON04869', 10895.38, date '2026-04-30'),
  ('OSB-MAY26-008', 'ON04933', 4319.84, date '2026-05-08'),
  ('OSB-MAY26-027', 'ON04892', 92220.0, date '2026-05-22'),
  ('OSB-MAY26-034', 'ON04928', 9138.0, date '2026-05-27'),
  ('OSB-MAY26-039', 'ON04905', 1597370.29, date '2026-05-26'),
  ('OSB-JUN26-012', 'ON04911', 1639671.6, date '2026-06-05'),
  ('OSB-JUL26-057', 'ON05035', 24824.0, date '2026-07-22'),
  ('OSB-JUL26-075', 'ON05046', 16792.89, date '2026-07-31'),
  ('OSB-SEP26-003', 'ON05110', 7192.7, date '2026-08-31')
)
select t.folio_transaccion, t.dia, t.proveedor, t.concepto_detallado, t.importe, t.status,
       t.folio_factura, t.folio_compra_sae, p.folio as partida,
       l.doc_aspel, l.fecha_aspel, l.importe_aspel,
       (select count(*) from public.adjuntos a
         where a.entidad = 'transaccion' and a.entidad_id = t.id::text) as adjuntos
from lista l
join public.transacciones t on t.folio_transaccion = l.folio and t.unidad_detectada = 'OSB'
left join public.partidas p on p.id = t.partida_id
order by t.dia;

------------------------------------------------------------------------
-- PASO 2 · CANCELAR (no borra). Igual que el botón Cancelar de la app:
-- conserva folio y expediente, sale de reportes y pagos, y se puede
-- reactivar desde Transacciones. Corre TODO el bloque junto.
------------------------------------------------------------------------
/*
begin;

with lista(folio) as (values
  ('OSB-MAR25-056'),
  ('OSB-ABR25-050'),
  ('OSB-ABR25-056'),
  ('OSB-ABR25-076'),
  ('OSB-MAY25-004'),
  ('OSB-MAY25-006'),
  ('OSB-MAY25-014'),
  ('OSB-MAY25-048'),
  ('OSB-AGO25-014'),
  ('OSB-OCT25-027'),
  ('OSB-OCT25-062'),
  ('OSB-NOV25-010'),
  ('OSB-DIC25-001'),
  ('OSB-DIC25-088'),
  ('OSB-ENE26-052'),
  ('OSB-FEB26-033'),
  ('OSB-FEB26-037'),
  ('OSB-MAR26-024'),
  ('OSB-ABR26-037'),
  ('OSB-ABR26-096'),
  ('OSB-MAY26-008'),
  ('OSB-MAY26-027'),
  ('OSB-MAY26-034'),
  ('OSB-MAY26-039'),
  ('OSB-JUN26-012'),
  ('OSB-JUL26-057'),
  ('OSB-JUL26-075'),
  ('OSB-SEP26-003')
)
update public.transacciones t
set cancelada_en = now(),
    cancelada_por = 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',  -- francisco.munoz@osbog.com
    motivo_cancelacion = 'Documento cancelado en Aspel (depuración nov 2026)',
    historial_cancelacion = coalesce(t.historial_cancelacion, '[]'::jsonb) || jsonb_build_array(jsonb_build_object(
      'accion', 'cancelada', 'fecha', now(), 'usuario', 'efee41b3-e8f6-4ad6-a5b5-6294b788f0fb',
      'motivo', 'Documento cancelado en Aspel (depuración nov 2026)'))
from lista l
where t.folio_transaccion = l.folio
  and t.unidad_detectada = 'OSB'
  and t.cancelada_en is null
returning t.folio_transaccion, t.proveedor, t.importe;

-- Si el número de renglones que regresa es el que esperabas:  commit;
-- Si no:                                                      rollback;
*/
