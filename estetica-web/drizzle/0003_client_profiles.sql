CREATE TABLE IF NOT EXISTS client_profiles (
  telefono text PRIMARY KEY NOT NULL,
  nombre text,
  email text,
  notas text,
  updated_at timestamptz DEFAULT now() NOT NULL
);
