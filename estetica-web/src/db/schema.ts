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
  /** Precio de referencia en ARS (caja; 0 = sin precio fijo). */
  precioPesos: integer("precio_pesos").notNull().default(0),
  /** Categoría padre (ej. Depilación). Null si es suelto o es categoría raíz. */
  /** FK a `services.id` (ver migración 0006; sin `.references` por autorreferencia). */
  parentId: integer("parent_id"),
  /** true = agrupa otros servicios; no se reserva ni entra en paquetes como ítem. */
  esGrupo: boolean("es_grupo").notNull().default(false),
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
/** Datos extra del cliente (clave: teléfono normalizado). */
export const clientProfiles = pgTable("client_profiles", {
  telefono: text("telefono").primaryKey(),
  nombre: text("nombre"),
  email: text("email"),
  notas: text("notas"),
  updatedAt: timestamp("updated_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

export type ServicePackageRow = typeof servicePackages.$inferSelect;
export type ClientPackageRow = typeof clientPackages.$inferSelect;
export type ClientProfileRow = typeof clientProfiles.$inferSelect;

/** Sesión de caja (apertura / cierre del día). */
export const cashSessions = pgTable("cash_sessions", {
  id: serial("id").primaryKey(),
  openedAt: timestamp("opened_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
  closedAt: timestamp("closed_at", { withTimezone: true }),
  openingAmountPesos: integer("opening_amount_pesos").notNull().default(0),
  closingAmountPesos: integer("closing_amount_pesos"),
  notes: text("notes"),
  status: text("status").notNull().default("abierta"),
});

/** Comprobante / venta registrada en caja. */
export const sales = pgTable("sales", {
  id: serial("id").primaryKey(),
  sessionId: integer("session_id")
    .notNull()
    .references(() => cashSessions.id),
  numero: integer("numero").notNull(),
  clienteTelefono: text("cliente_telefono"),
  clienteNombre: text("cliente_nombre"),
  subtotalPesos: integer("subtotal_pesos").notNull().default(0),
  descuentoPesos: integer("descuento_pesos").notNull().default(0),
  totalPesos: integer("total_pesos").notNull().default(0),
  metodoPago: text("metodo_pago").notNull().default("efectivo"),
  notas: text("notas"),
  estado: text("estado").notNull().default("completada"),
  appointmentId: integer("appointment_id").references(() => appointments.id, {
    onDelete: "set null",
  }),
  clientPackageId: integer("client_package_id").references(
    () => clientPackages.id,
    { onDelete: "set null" }
  ),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

export const saleLines = pgTable("sale_lines", {
  id: serial("id").primaryKey(),
  saleId: integer("sale_id")
    .notNull()
    .references(() => sales.id, { onDelete: "cascade" }),
  tipo: text("tipo").notNull().default("otro"),
  descripcion: text("descripcion").notNull(),
  cantidad: integer("cantidad").notNull().default(1),
  precioUnitarioPesos: integer("precio_unitario_pesos").notNull().default(0),
  totalLineaPesos: integer("total_linea_pesos").notNull().default(0),
  serviceId: integer("service_id").references(() => services.id, {
    onDelete: "set null",
  }),
  servicePackageId: integer("service_package_id").references(
    () => servicePackages.id,
    { onDelete: "set null" }
  ),
});

export type CashSessionRow = typeof cashSessions.$inferSelect;
export type SaleRow = typeof sales.$inferSelect;
export type SaleLineRow = typeof saleLines.$inferSelect;
