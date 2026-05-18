-- Franjas de atención del local por día (Luxon: 1=lunes … 7=domingo)
CREATE TABLE IF NOT EXISTS business_hour_slots (
  id SERIAL PRIMARY KEY,
  dia_semana SMALLINT NOT NULL CHECK (dia_semana >= 1 AND dia_semana <= 7),
  horario_inicio TEXT NOT NULL,
  horario_fin TEXT NOT NULL,
  activo BOOLEAN NOT NULL DEFAULT true
);

CREATE INDEX IF NOT EXISTS business_hour_slots_dia_idx
  ON business_hour_slots (dia_semana);

-- Lun–sáb: mañana y tarde (ajustable en admin)
INSERT INTO business_hour_slots (dia_semana, horario_inicio, horario_fin)
SELECT d.dia, f.inicio, f.fin
FROM (VALUES (1), (2), (3), (4), (5), (6)) AS d(dia)
CROSS JOIN (
  VALUES ('09:00', '13:00'), ('17:00', '21:00')
) AS f(inicio, fin)
WHERE NOT EXISTS (SELECT 1 FROM business_hour_slots LIMIT 1);
