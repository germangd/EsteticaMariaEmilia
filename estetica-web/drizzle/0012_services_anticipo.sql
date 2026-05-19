-- Anticipo opcional por servicio (porcentaje individual)
-- Ejecutar el bloque completo en Neon (no solo las líneas de tipo de dato).
ALTER TABLE services
  ADD COLUMN IF NOT EXISTS anticipo_requerido boolean NOT NULL DEFAULT false,
  ADD COLUMN IF NOT EXISTS anticipo_porcentaje integer NOT NULL DEFAULT 0;
