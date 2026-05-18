-- Fechas habilitadas por servicio (si hay filas, solo esas fechas admiten turnos)
CREATE TABLE IF NOT EXISTS service_availability_dates (
  service_id INTEGER NOT NULL REFERENCES services (id) ON DELETE CASCADE,
  fecha DATE NOT NULL,
  PRIMARY KEY (service_id, fecha)
);

CREATE INDEX IF NOT EXISTS service_availability_dates_fecha_idx
  ON service_availability_dates (fecha);

-- Eventos especiales: en esa fecha solo se reservan servicios del evento
CREATE TABLE IF NOT EXISTS agenda_events (
  id SERIAL PRIMARY KEY,
  nombre TEXT NOT NULL,
  descripcion TEXT,
  fecha DATE NOT NULL,
  horario_inicio TEXT NOT NULL DEFAULT '09:00',
  horario_fin TEXT NOT NULL DEFAULT '18:00',
  activo BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS agenda_events_fecha_idx ON agenda_events (fecha);

CREATE TABLE IF NOT EXISTS agenda_event_services (
  event_id INTEGER NOT NULL REFERENCES agenda_events (id) ON DELETE CASCADE,
  service_id INTEGER NOT NULL REFERENCES services (id) ON DELETE CASCADE,
  PRIMARY KEY (event_id, service_id)
);
