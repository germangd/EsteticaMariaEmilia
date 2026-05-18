-- Caja: sesiones, ventas y líneas (ticket interno)
CREATE TABLE IF NOT EXISTS cash_sessions (
  id SERIAL PRIMARY KEY,
  opened_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  closed_at TIMESTAMPTZ,
  opening_amount_pesos INTEGER NOT NULL DEFAULT 0,
  closing_amount_pesos INTEGER,
  notes TEXT,
  status TEXT NOT NULL DEFAULT 'abierta'
);

CREATE INDEX IF NOT EXISTS cash_sessions_status_idx ON cash_sessions (status);

CREATE TABLE IF NOT EXISTS sales (
  id SERIAL PRIMARY KEY,
  session_id INTEGER NOT NULL REFERENCES cash_sessions (id),
  numero INTEGER NOT NULL,
  cliente_telefono TEXT,
  cliente_nombre TEXT,
  subtotal_pesos INTEGER NOT NULL DEFAULT 0,
  descuento_pesos INTEGER NOT NULL DEFAULT 0,
  total_pesos INTEGER NOT NULL DEFAULT 0,
  metodo_pago TEXT NOT NULL DEFAULT 'efectivo',
  notas TEXT,
  estado TEXT NOT NULL DEFAULT 'completada',
  appointment_id INTEGER REFERENCES appointments (id) ON DELETE SET NULL,
  client_package_id INTEGER REFERENCES client_packages (id) ON DELETE SET NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS sales_session_numero_idx ON sales (session_id, numero);
CREATE INDEX IF NOT EXISTS sales_session_created_idx ON sales (session_id, created_at DESC);

CREATE TABLE IF NOT EXISTS sale_lines (
  id SERIAL PRIMARY KEY,
  sale_id INTEGER NOT NULL REFERENCES sales (id) ON DELETE CASCADE,
  tipo TEXT NOT NULL DEFAULT 'otro',
  descripcion TEXT NOT NULL,
  cantidad INTEGER NOT NULL DEFAULT 1,
  precio_unitario_pesos INTEGER NOT NULL DEFAULT 0,
  total_linea_pesos INTEGER NOT NULL DEFAULT 0,
  service_id INTEGER REFERENCES services (id) ON DELETE SET NULL,
  service_package_id INTEGER REFERENCES service_packages (id) ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS sale_lines_sale_id_idx ON sale_lines (sale_id);
