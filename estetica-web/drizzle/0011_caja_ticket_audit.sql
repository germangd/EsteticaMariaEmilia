-- Número de ticket global secuencial + registro de auditoría

ALTER TABLE sales ADD COLUMN IF NOT EXISTS numero_ticket INTEGER;

UPDATE sales s
SET numero_ticket = sub.rn
FROM (
  SELECT id, ROW_NUMBER() OVER (ORDER BY created_at ASC, id ASC) AS rn
  FROM sales
) sub
WHERE s.id = sub.id AND s.numero_ticket IS NULL;

ALTER TABLE sales ALTER COLUMN numero_ticket SET NOT NULL;

CREATE UNIQUE INDEX IF NOT EXISTS sales_numero_ticket_idx ON sales (numero_ticket);

CREATE SEQUENCE IF NOT EXISTS sales_numero_ticket_seq;
SELECT setval(
  'sales_numero_ticket_seq',
  COALESCE((SELECT MAX(numero_ticket) FROM sales), 0)
);

CREATE TABLE IF NOT EXISTS sale_audit_log (
  id SERIAL PRIMARY KEY,
  sale_id INTEGER NOT NULL REFERENCES sales (id) ON DELETE CASCADE,
  accion TEXT NOT NULL,
  detalle TEXT,
  datos_antes TEXT,
  datos_despues TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS sale_audit_log_sale_id_idx ON sale_audit_log (sale_id, created_at DESC);
