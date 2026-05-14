CREATE TABLE "appointments" (
	"id" serial PRIMARY KEY NOT NULL,
	"fecha" date NOT NULL,
	"hora" text NOT NULL,
	"nombre_cliente" text NOT NULL,
	"telefono" text NOT NULL,
	"email" text,
	"servicio_nombre" text NOT NULL,
	"responsable" text NOT NULL,
	"codigo_cancelacion" text NOT NULL,
	"estado" text DEFAULT 'activo' NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL,
	CONSTRAINT "appointments_codigo_cancelacion_unique" UNIQUE("codigo_cancelacion")
);
