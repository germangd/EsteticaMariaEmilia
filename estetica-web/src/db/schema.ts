import {
  boolean,
  date,
  integer,
  pgTable,
  primaryKey,
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

/** Catálogo de paquetes (ej. 3 sesiones depilación). */
export const servicePackages = pgTable("service_packages", {
  id: serial("id").primaryKey(),
  nombre: text("nombre").notNull(),
  descripcion: text("descripcion"),
  precioPesos: integer("precio_pesos").notNull().default(0),
  sesionesTotal: integer("sesiones_total").notNull().default(1),
  activo: boolean("activo").notNull().default(true),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

/** Servicios incluidos en un paquete. */
export const packageServices = pgTable(
  "package_services",
  {
    packageId: integer("package_id")
      .notNull()
      .references(() => servicePackages.id, { onDelete: "cascade" }),
    serviceId: integer("service_id")
      .notNull()
      .references(() => services.id, { onDelete: "cascade" }),
  },
  (t) => [primaryKey({ columns: [t.packageId, t.serviceId] })]
);

/** Paquete vendido / asignado a un cliente (control de sesiones). */
export const clientPackages = pgTable("client_packages", {
  id: serial("id").primaryKey(),
  packageId: integer("package_id")
    .notNull()
    .references(() => servicePackages.id),
  nombreCliente: text("nombre_cliente").notNull(),
  telefono: text("telefono").notNull(),
  sesionesIniciales: integer("sesiones_iniciales").notNull(),
  sesionesRestantes: integer("sesiones_restantes").notNull(),
  precioCobradoPesos: integer("precio_cobrado_pesos"),
  notas: text("notas"),
  estado: text("estado").notNull().default("activo"),
  fechaCompra: date("fecha_compra", { mode: "string" }).notNull(),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

export type ServiceRow = typeof services.$inferSelect;
export type AppointmentRow = typeof appointments.$inferSelect;
export type ServicePackageRow = typeof servicePackages.$inferSelect;
export type ClientPackageRow = typeof clientPackages.$inferSelect;
