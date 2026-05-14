-- Ejemplo: ejecutar en Neon (SQL Editor) después de aplicar la migración `0000_init.sql`.
-- Ajustá nombre, horarios y responsable a tu negocio.

INSERT INTO services (nombre, duracion_min, responsable, capacidad, horario_inicio, horario_fin)
VALUES
  ('Manicura', 45, 'María Emilia', 2, '09:00', '18:00'),
  ('Depilación', 30, 'María Emilia', 1, '10:00', '17:00');
