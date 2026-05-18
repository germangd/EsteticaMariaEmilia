-- Categorías de servicio (ej. Depilación) con sub-servicios (cavado, axilas, etc.)
ALTER TABLE services
ADD COLUMN IF NOT EXISTS parent_id integer REFERENCES services (id) ON DELETE SET NULL;

ALTER TABLE services
ADD COLUMN IF NOT EXISTS es_grupo boolean NOT NULL DEFAULT false;

CREATE INDEX IF NOT EXISTS services_parent_id_idx ON services (parent_id);
