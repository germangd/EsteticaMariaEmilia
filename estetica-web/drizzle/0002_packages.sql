CREATE TABLE IF NOT EXISTS "service_packages" (
  "id" serial PRIMARY KEY NOT NULL,
  "nombre" text NOT NULL,
  "descripcion" text,
  "precio_pesos" integer DEFAULT 0 NOT NULL,
  "sesiones_total" integer DEFAULT 1 NOT NULL,
  "activo" boolean DEFAULT true NOT NULL,
  "created_at" timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE IF NOT EXISTS "package_services" (
  "package_id" integer NOT NULL,
  "service_id" integer NOT NULL,
  CONSTRAINT "package_services_package_id_service_id_pk" PRIMARY KEY("package_id","service_id")
);

CREATE TABLE IF NOT EXISTS "client_packages" (
  "id" serial PRIMARY KEY NOT NULL,
  "package_id" integer NOT NULL,
  "nombre_cliente" text NOT NULL,
  "telefono" text NOT NULL,
  "sesiones_iniciales" integer NOT NULL,
  "sesiones_restantes" integer NOT NULL,
  "precio_cobrado_pesos" integer,
  "notas" text,
  "estado" text DEFAULT 'activo' NOT NULL,
  "fecha_compra" date NOT NULL,
  "created_at" timestamp with time zone DEFAULT now() NOT NULL
);

DO $$ BEGIN
  ALTER TABLE "package_services" ADD CONSTRAINT "package_services_package_id_service_packages_id_fk"
    FOREIGN KEY ("package_id") REFERENCES "service_packages"("id") ON DELETE cascade ON UPDATE no action;
EXCEPTION WHEN duplicate_object THEN null; END $$;

DO $$ BEGIN
  ALTER TABLE "package_services" ADD CONSTRAINT "package_services_service_id_services_id_fk"
    FOREIGN KEY ("service_id") REFERENCES "services"("id") ON DELETE cascade ON UPDATE no action;
EXCEPTION WHEN duplicate_object THEN null; END $$;

DO $$ BEGIN
  ALTER TABLE "client_packages" ADD CONSTRAINT "client_packages_package_id_service_packages_id_fk"
    FOREIGN KEY ("package_id") REFERENCES "service_packages"("id") ON DELETE no action ON UPDATE no action;
EXCEPTION WHEN duplicate_object THEN null; END $$;
