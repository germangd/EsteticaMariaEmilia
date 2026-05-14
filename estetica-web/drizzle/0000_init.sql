CREATE TABLE "services" (
	"id" serial PRIMARY KEY NOT NULL,
	"nombre" text NOT NULL,
	"duracion_min" integer DEFAULT 30 NOT NULL,
	"responsable" text DEFAULT 'No asignado' NOT NULL,
	"capacidad" integer DEFAULT 1 NOT NULL,
	"horario_inicio" text DEFAULT '09:00' NOT NULL,
	"horario_fin" text DEFAULT '18:00' NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
