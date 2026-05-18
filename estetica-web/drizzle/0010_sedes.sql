-- Sedes de atención (Ensenada, Bavio, Magdalena)
CREATE TABLE IF NOT EXISTS sedes (
  id SERIAL PRIMARY KEY,
  nombre TEXT NOT NULL,
  activo BOOLEAN NOT NULL DEFAULT true,
  orden SMALLINT NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

INSERT INTO sedes (nombre, orden)
SELECT v.nombre, v.orden
FROM (
  VALUES ('Ensenada', 1), ('Bartolomé Bavio', 2), ('Magdalena', 3)
) AS v(nombre, orden)
WHERE NOT EXISTS (SELECT 1 FROM sedes LIMIT 1);

-- Horarios de atención por sede
ALTER TABLE business_hour_slots
  ADD COLUMN IF NOT EXISTS sede_id INTEGER REFERENCES sedes (id) ON DELETE CASCADE;

UPDATE business_hour_slots
SET sede_id = (SELECT id FROM sedes ORDER BY orden, id LIMIT 1)
WHERE sede_id IS NULL;

INSERT INTO business_hour_slots (dia_semana, horario_inicio, horario_fin, activo, sede_id)
SELECT b.dia_semana, b.horario_inicio, b.horario_fin, b.activo, s.id
FROM business_hour_slots b
CROSS JOIN sedes s
WHERE b.sede_id = (SELECT id FROM sedes ORDER BY orden, id LIMIT 1)
  AND s.id <> b.sede_id
  AND NOT EXISTS (
    SELECT 1 FROM business_hour_slots x WHERE x.sede_id = s.id LIMIT 1
  );

ALTER TABLE business_hour_slots
  ALTER COLUMN sede_id SET NOT NULL;

CREATE INDEX IF NOT EXISTS business_hour_slots_sede_dia_idx
  ON business_hour_slots (sede_id, dia_semana);

-- Turnos y eventos por sede
ALTER TABLE appointments
  ADD COLUMN IF NOT EXISTS sede_id INTEGER REFERENCES sedes (id);

UPDATE appointments
SET sede_id = (SELECT id FROM sedes ORDER BY orden, id LIMIT 1)
WHERE sede_id IS NULL;

ALTER TABLE appointments
  ALTER COLUMN sede_id SET NOT NULL;

CREATE INDEX IF NOT EXISTS appointments_sede_fecha_idx
  ON appointments (sede_id, fecha);

ALTER TABLE agenda_events
  ADD COLUMN IF NOT EXISTS sede_id INTEGER REFERENCES sedes (id);

UPDATE agenda_events
SET sede_id = (SELECT id FROM sedes ORDER BY orden, id LIMIT 1)
WHERE sede_id IS NULL;

ALTER TABLE agenda_events
  ALTER COLUMN sede_id SET NOT NULL;

CREATE INDEX IF NOT EXISTS agenda_events_sede_fecha_idx
  ON agenda_events (sede_id, fecha);
