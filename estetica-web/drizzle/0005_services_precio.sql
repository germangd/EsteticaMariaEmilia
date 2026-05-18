-- Precio sugerido de referencia por servicio (caja / cobros)
ALTER TABLE services
ADD COLUMN IF NOT EXISTS precio_pesos integer NOT NULL DEFAULT 0;
