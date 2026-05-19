-- =============================================================================
-- Reset de datos de PRUEBA — María Emilia Estética
-- =============================================================================
--
-- Ejecutar en Neon (SQL Editor) sobre la base que usa Vercel (DATABASE_URL).
-- Recomendado: crear un branch de respaldo en Neon antes de correr esto.
--
-- BORRA:
--   • Caja: ventas, líneas, auditoría, sesiones de caja
--   • Agenda: turnos (appointments)
--
-- CONSERVA (no se toca):
--   • services (catálogo de servicios, precios, anticipos, categorías)
--   • sedes, business_hour_slots (horarios)
--   • service_packages, package_services (catálogo de paquetes)
--   • service_availability_dates, agenda_events, agenda_event_services
--
-- OPCIONAL (descomentar al final si también querés vaciar clientes/paquetes vendidos):
--   • client_profiles, client_packages
--
-- =============================================================================

-- --- 1) Vista previa (opcional; podés ejecutar solo esto primero) ------------
SELECT 'sales' AS tabla, COUNT(*)::bigint AS filas FROM sales
UNION ALL SELECT 'sale_lines', COUNT(*) FROM sale_lines
UNION ALL SELECT 'sale_audit_log', COUNT(*) FROM sale_audit_log
UNION ALL SELECT 'cash_sessions', COUNT(*) FROM cash_sessions
UNION ALL SELECT 'appointments', COUNT(*) FROM appointments;

-- --- 2) Borrado + reinicio de numeración de tickets ---------------------------
BEGIN;

-- Ventas (en cascada: sale_lines, sale_audit_log)
DELETE FROM sales;

-- Turnos de caja (aperturas / cierres)
DELETE FROM cash_sessions;

-- Agenda: turnos confirmados, pendientes de anticipo, cancelados, etc.
DELETE FROM appointments;

-- Próximo comprobante en caja será ticket #1 (formato #000001 en la app)
SELECT setval('sales_numero_ticket_seq', 0, false);

COMMIT;

-- --- 3) Comprobación ------------------------------------------------------------
SELECT 'sales' AS tabla, COUNT(*)::bigint AS filas_restantes FROM sales
UNION ALL SELECT 'sale_lines', COUNT(*) FROM sale_lines
UNION ALL SELECT 'sale_audit_log', COUNT(*) FROM sale_audit_log
UNION ALL SELECT 'cash_sessions', COUNT(*) FROM cash_sessions
UNION ALL SELECT 'appointments', COUNT(*) FROM appointments;

SELECT
  last_value AS ultimo_valor_seq,
  is_called AS seq_usada
FROM sales_numero_ticket_seq;

-- Servicios intactos (debería ser > 0 si ya cargaste el catálogo):
SELECT COUNT(*)::bigint AS servicios_en_catalogo FROM services;

-- =============================================================================
-- OPCIONAL: descomentar en un segundo paso si querés fichas y paquetes asignados
-- =============================================================================
--
-- BEGIN;
-- DELETE FROM client_packages;
-- DELETE FROM client_profiles;
-- COMMIT;
--
-- =============================================================================
