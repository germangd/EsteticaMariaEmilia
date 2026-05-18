-- Varios eventos por día (franjas distintas) + precio y cliente del evento
ALTER TABLE agenda_events
  ADD COLUMN IF NOT EXISTS precio_pesos INTEGER NOT NULL DEFAULT 0,
  ADD COLUMN IF NOT EXISTS cliente_telefono TEXT,
  ADD COLUMN IF NOT EXISTS cliente_nombre TEXT;
