import {
  date,
  integer,
  pgTable,
  serial,
  text,
  timestamp,
} from "drizzle-orm/pg-core";

/** Paridad con hoja `Config` (columnas A–F en `Código.gs`). */
export const services = pgTable("services", {
  id: serial("id").primaryKey(),
  nombre: text("nombre").notNull(),
  duracionMin: integer("duracion_min").notNull().default(30),
  responsable: text("responsable").notNull().default("No asignado"),
  capacidad: integer("capacidad").notNull().default(1),
  horarioInicio: text("horario_inicio").notNull().default("09:00"),
  horarioFin: text("horario_fin").notNull().default("18:00"),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

/**
 * Paridad con hoja `Turnos` (fila `appendRow` en `Código.gs`):
 * id, fecha, hora, nombre, teléfono, servicio, responsable, código, estado.
 * Añadimos `email` opcional (no estaba en la fila del sheet; el mail se enviaba igual).
 */
export const appointments = pgTable("appointments", {
  id: serial("id").primaryKey(),
  fecha: date("fecha", { mode: "string" }).notNull(),
  hora: text("hora").notNull(),
  nombreCliente: text("nombre_cliente").notNull(),
  telefono: text("telefono").notNull(),
  email: text("email"),
  servicioNombre: text("servicio_nombre").notNull(),
  responsable: text("responsable").notNull(),
  codigoCancelacion: text("codigo_cancelacion").notNull().unique(),
  estado: text("estado").notNull().default("activo"),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

export type ServiceRow = typeof services.$inferSelect;
export type AppointmentRow = typeof appointments.$inferSelect;
