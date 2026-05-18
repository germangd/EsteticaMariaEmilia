import { and, desc, eq, max, sql } from "drizzle-orm";
import { getDb } from "@/db/client";
import {
  appointments,
  cashSessions,
  clientPackages,
  saleLines,
  sales,
  servicePackages,
  services,
} from "@/db/schema";
import { normalizarTelefono } from "@/lib/clientes-repo";

export const METODOS_PAGO = [
  "efectivo",
  "transferencia",
  "debito",
  "credito",
  "otro",
] as const;

export type MetodoPago = (typeof METODOS_PAGO)[number];

export type ArqueoMetodo = {
  metodoPago: string;
  totalPesos: number;
  cantidad: number;
};

export type SesionCaja = {
  id: number;
  openedAt: string;
  closedAt: string | null;
  openingAmountPesos: number;
  closingAmountPesos: number | null;
  notes: string | null;
  status: string;
  totalVentasPesos: number;
  cantidadVentas: number;
  arqueoPorMetodo: ArqueoMetodo[];
  /** Fondo inicial + ventas en efectivo (para contar el cajón). */
  efectivoEsperadoEnCajon: number;
};

export type PrefillCobroTurno = {
  appointmentId: number;
  clienteNombre: string;
  clienteTelefono: string;
  servicioNombre: string;
  fecha: string;
  hora: string;
  serviceId: number | null;
  yaCobrado: boolean;
  ventaId: number | null;
};

export type LineaVentaInput = {
  tipo: "servicio" | "paquete" | "otro";
  descripcion: string;
  cantidad: number;
  precioUnitarioPesos: number;
  serviceId?: number;
  servicePackageId?: number;
};

export type CrearVentaInput = {
  sessionId?: number;
  clienteTelefono?: string;
  clienteNombre?: string;
  descuentoPesos?: number;
  metodoPago: string;
  notas?: string | null;
  appointmentId?: number;
  clientPackageId?: number;
  lineas: LineaVentaInput[];
};

export type VentaLinea = {
  id: number;
  tipo: string;
  descripcion: string;
  cantidad: number;
  precioUnitarioPesos: number;
  totalLineaPesos: number;
};

export type VentaResumen = {
  id: number;
  sessionId: number;
  numero: number;
  clienteTelefono: string | null;
  clienteNombre: string | null;
  subtotalPesos: number;
  descuentoPesos: number;
  totalPesos: number;
  metodoPago: string;
  notas: string | null;
  estado: string;
  createdAt: string;
};

export type VentaDetalle = VentaResumen & {
  lineas: VentaLinea[];
  sessionOpenedAt: string;
};

function pesos(n: unknown): number {
  const x = Number(n);
  if (!Number.isFinite(x)) return 0;
  return Math.max(0, Math.round(x));
}

function normalizeLinea(l: LineaVentaInput): LineaVentaInput | null {
  const descripcion = l.descripcion.trim();
  if (!descripcion) return null;
  const cantidad = Math.max(1, Math.round(l.cantidad));
  const precioUnitarioPesos = pesos(l.precioUnitarioPesos);
  return {
    tipo: l.tipo === "servicio" || l.tipo === "paquete" ? l.tipo : "otro",
    descripcion,
    cantidad,
    precioUnitarioPesos,
    serviceId: l.serviceId && l.serviceId > 0 ? l.serviceId : undefined,
    servicePackageId:
      l.servicePackageId && l.servicePackageId > 0
        ? l.servicePackageId
        : undefined,
  };
}

async function totalesSesion(
  sessionId: number
): Promise<{ total: number; count: number }> {
  const db = getDb();
  if (!db) return { total: 0, count: 0 };

  const [row] = await db
    .select({
      total: sql<number>`coalesce(sum(${sales.totalPesos}), 0)::int`,
      count: sql<number>`count(*)::int`,
    })
    .from(sales)
    .where(
      and(eq(sales.sessionId, sessionId), eq(sales.estado, "completada"))
    );

  return {
    total: Number(row?.total ?? 0),
    count: Number(row?.count ?? 0),
  };
}

export async function resumenArqueoSesion(
  sessionId: number
): Promise<ArqueoMetodo[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select({
      metodoPago: sales.metodoPago,
      totalPesos: sql<number>`coalesce(sum(${sales.totalPesos}), 0)::int`,
      cantidad: sql<number>`count(*)::int`,
    })
    .from(sales)
    .where(
      and(eq(sales.sessionId, sessionId), eq(sales.estado, "completada"))
    )
    .groupBy(sales.metodoPago)
    .orderBy(sales.metodoPago);

  return rows.map((r) => ({
    metodoPago: r.metodoPago,
    totalPesos: Number(r.totalPesos ?? 0),
    cantidad: Number(r.cantidad ?? 0),
  }));
}

async function buildSesionStats(
  row: typeof cashSessions.$inferSelect
): Promise<SesionCaja> {
  const stats = await totalesSesion(row.id);
  const arqueoRaw = await resumenArqueoSesion(row.id);
  const arqueoPorMetodo = Array.isArray(arqueoRaw) ? arqueoRaw : [];
  const efectivoVentas =
    arqueoPorMetodo.find((a) => a.metodoPago === "efectivo")?.totalPesos ?? 0;

  return {
    id: row.id,
    openedAt: row.openedAt.toISOString(),
    closedAt: row.closedAt?.toISOString() ?? null,
    openingAmountPesos: row.openingAmountPesos,
    closingAmountPesos: row.closingAmountPesos,
    notes: row.notes,
    status: row.status,
    totalVentasPesos: stats.total,
    cantidadVentas: stats.count,
    arqueoPorMetodo,
    efectivoEsperadoEnCajon: row.openingAmountPesos + efectivoVentas,
  };
}

export async function obtenerSesionAbierta(): Promise<
  SesionCaja | null | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [row] = await db
    .select()
    .from(cashSessions)
    .where(eq(cashSessions.status, "abierta"))
    .orderBy(desc(cashSessions.openedAt))
    .limit(1);

  if (!row) return null;
  return buildSesionStats(row);
}

export async function abrirSesionCaja(input: {
  openingAmountPesos?: number;
  notes?: string | null;
}): Promise<
  | { ok: true; sesion: SesionCaja }
  | { ok: false; reason: "no_db" | "ya_abierta" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const abierta = await obtenerSesionAbierta();
  if (abierta && typeof abierta === "object" && "ok" in abierta) {
    return { ok: false, reason: "no_db" };
  }
  if (abierta) {
    return { ok: false, reason: "ya_abierta" };
  }

  const openingAmountPesos = pesos(input.openingAmountPesos ?? 0);

  const [inserted] = await db
    .insert(cashSessions)
    .values({
      openingAmountPesos,
      notes: input.notes?.trim() || null,
      status: "abierta",
    })
    .returning();

  return {
    ok: true,
    sesion: await buildSesionStats(inserted),
  };
}

export async function cerrarSesionCaja(input: {
  sessionId: number;
  closingAmountPesos: number;
  notes?: string | null;
}): Promise<
  | { ok: true; sesion: SesionCaja }
  | { ok: false; reason: "no_db" | "not_found" | "cerrada" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const closingAmountPesos = pesos(input.closingAmountPesos);
  if (input.sessionId < 1) return { ok: false, reason: "invalido" };

  const [row] = await db
    .select()
    .from(cashSessions)
    .where(eq(cashSessions.id, input.sessionId))
    .limit(1);

  if (!row) return { ok: false, reason: "not_found" };
  if (row.status !== "abierta") return { ok: false, reason: "cerrada" };

  const [updated] = await db
    .update(cashSessions)
    .set({
      status: "cerrada",
      closedAt: new Date(),
      closingAmountPesos,
      notes: input.notes?.trim() || row.notes,
    })
    .where(eq(cashSessions.id, input.sessionId))
    .returning();

  return { ok: true, sesion: await buildSesionStats(updated) };
}

export async function listarVentasSesion(
  sessionId: number
): Promise<VentaResumen[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(sales)
    .where(eq(sales.sessionId, sessionId))
    .orderBy(desc(sales.numero));

  return rows.map((r) => ({
    id: r.id,
    sessionId: r.sessionId,
    numero: r.numero,
    clienteTelefono: r.clienteTelefono,
    clienteNombre: r.clienteNombre,
    subtotalPesos: r.subtotalPesos,
    descuentoPesos: r.descuentoPesos,
    totalPesos: r.totalPesos,
    metodoPago: r.metodoPago,
    notas: r.notas,
    estado: r.estado,
    createdAt: r.createdAt.toISOString(),
  }));
}

export async function crearVenta(
  input: CrearVentaInput
): Promise<
  | { ok: true; id: number; numero: number }
  | {
      ok: false;
      reason:
        | "no_db"
        | "sin_sesion"
        | "sesion_cerrada"
        | "invalido"
        | "sin_lineas";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const lineasNorm = input.lineas
    .map(normalizeLinea)
    .filter((l): l is LineaVentaInput => l !== null);

  if (lineasNorm.length === 0) return { ok: false, reason: "sin_lineas" };

  let sessionId = input.sessionId;
  if (!sessionId || sessionId < 1) {
    const abierta = await obtenerSesionAbierta();
    if (abierta && typeof abierta === "object" && "ok" in abierta) {
      return { ok: false, reason: "no_db" };
    }
    if (!abierta) {
      return { ok: false, reason: "sin_sesion" };
    }
    sessionId = abierta.id;
  }

  const [ses] = await db
    .select()
    .from(cashSessions)
    .where(eq(cashSessions.id, sessionId))
    .limit(1);

  if (!ses) return { ok: false, reason: "sin_sesion" };
  if (ses.status !== "abierta") return { ok: false, reason: "sesion_cerrada" };

  const metodoPago = METODOS_PAGO.includes(input.metodoPago as MetodoPago)
    ? input.metodoPago
    : "efectivo";

  const lineasCalc = lineasNorm.map((l) => ({
    ...l,
    totalLineaPesos: l.cantidad * l.precioUnitarioPesos,
  }));

  const subtotalPesos = lineasCalc.reduce((s, l) => s + l.totalLineaPesos, 0);
  const descuentoPesos = Math.min(
    subtotalPesos,
    pesos(input.descuentoPesos ?? 0)
  );
  const totalPesos = subtotalPesos - descuentoPesos;

  const tel = input.clienteTelefono?.trim()
    ? normalizarTelefono(input.clienteTelefono)
    : null;
  const nombre = input.clienteNombre?.trim() || null;

  const [maxRow] = await db
    .select({ n: max(sales.numero) })
    .from(sales)
    .where(eq(sales.sessionId, sessionId));

  const numero = (maxRow?.n ?? 0) + 1;

  const [venta] = await db
    .insert(sales)
    .values({
      sessionId,
      numero,
      clienteTelefono: tel,
      clienteNombre: nombre,
      subtotalPesos,
      descuentoPesos,
      totalPesos,
      metodoPago,
      notas: input.notas?.trim() || null,
      appointmentId:
        input.appointmentId && input.appointmentId > 0
          ? input.appointmentId
          : null,
      clientPackageId:
        input.clientPackageId && input.clientPackageId > 0
          ? input.clientPackageId
          : null,
      estado: "completada",
    })
    .returning({ id: sales.id });

  await db.insert(saleLines).values(
    lineasCalc.map((l) => ({
      saleId: venta.id,
      tipo: l.tipo,
      descripcion: l.descripcion,
      cantidad: l.cantidad,
      precioUnitarioPesos: l.precioUnitarioPesos,
      totalLineaPesos: l.totalLineaPesos,
      serviceId: l.serviceId ?? null,
      servicePackageId: l.servicePackageId ?? null,
    }))
  );

  return { ok: true, id: venta.id, numero };
}

export async function obtenerVentaDetalle(
  id: number
): Promise<VentaDetalle | null | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [venta] = await db
    .select({
      sale: sales,
      sessionOpenedAt: cashSessions.openedAt,
    })
    .from(sales)
    .innerJoin(cashSessions, eq(sales.sessionId, cashSessions.id))
    .where(eq(sales.id, id))
    .limit(1);

  if (!venta) return null;

  const lineas = await db
    .select()
    .from(saleLines)
    .where(eq(saleLines.saleId, id))
    .orderBy(saleLines.id);

  const s = venta.sale;
  return {
    id: s.id,
    sessionId: s.sessionId,
    numero: s.numero,
    clienteTelefono: s.clienteTelefono,
    clienteNombre: s.clienteNombre,
    subtotalPesos: s.subtotalPesos,
    descuentoPesos: s.descuentoPesos,
    totalPesos: s.totalPesos,
    metodoPago: s.metodoPago,
    notas: s.notas,
    estado: s.estado,
    createdAt: s.createdAt.toISOString(),
    sessionOpenedAt: venta.sessionOpenedAt.toISOString(),
    lineas: lineas.map((l) => ({
      id: l.id,
      tipo: l.tipo,
      descripcion: l.descripcion,
      cantidad: l.cantidad,
      precioUnitarioPesos: l.precioUnitarioPesos,
      totalLineaPesos: l.totalLineaPesos,
    })),
  };
}

export async function anularVenta(
  id: number
): Promise<
  | { ok: true }
  | { ok: false; reason: "no_db" | "not_found" | "ya_anulada" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [row] = await db
    .select()
    .from(sales)
    .where(eq(sales.id, id))
    .limit(1);

  if (!row) return { ok: false, reason: "not_found" };
  if (row.estado === "anulada") return { ok: false, reason: "ya_anulada" };

  await db
    .update(sales)
    .set({ estado: "anulada" })
    .where(eq(sales.id, id));

  return { ok: true };
}

export type CatalogoCaja = {
  servicios: { id: number; nombre: string }[];
  paquetes: { id: number; nombre: string; precioPesos: number }[];
};

export async function obtenerCatalogoCaja(): Promise<
  CatalogoCaja | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const servs = await db
    .select({ id: services.id, nombre: services.nombre })
    .from(services)
    .orderBy(services.nombre);

  const packs = await db
    .select({
      id: servicePackages.id,
      nombre: servicePackages.nombre,
      precioPesos: servicePackages.precioPesos,
    })
    .from(servicePackages)
    .where(eq(servicePackages.activo, true))
    .orderBy(servicePackages.nombre);

  return {
    servicios: servs,
    paquetes: packs,
  };
}

async function ventaActivaPorReferencia(opts: {
  appointmentId?: number;
  clientPackageId?: number;
}): Promise<{ id: number } | null> {
  const db = getDb();
  if (!db) return null;

  if (opts.appointmentId && opts.appointmentId > 0) {
    const [row] = await db
      .select({ id: sales.id })
      .from(sales)
      .where(
        and(
          eq(sales.appointmentId, opts.appointmentId),
          eq(sales.estado, "completada")
        )
      )
      .limit(1);
    if (row) return row;
  }

  if (opts.clientPackageId && opts.clientPackageId > 0) {
    const [row] = await db
      .select({ id: sales.id })
      .from(sales)
      .where(
        and(
          eq(sales.clientPackageId, opts.clientPackageId),
          eq(sales.estado, "completada")
        )
      )
      .limit(1);
    if (row) return row;
  }

  return null;
}

export async function obtenerPrefillCobroTurno(
  appointmentId: number
): Promise<PrefillCobroTurno | null | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(appointmentId) || appointmentId < 1) return null;

  const [turno] = await db
    .select()
    .from(appointments)
    .where(eq(appointments.id, appointmentId))
    .limit(1);

  if (!turno || turno.estado !== "activo") return null;

  const [svc] = await db
    .select({ id: services.id })
    .from(services)
    .where(eq(services.nombre, turno.servicioNombre))
    .limit(1);

  const venta = await ventaActivaPorReferencia({ appointmentId });

  return {
    appointmentId: turno.id,
    clienteNombre: turno.nombreCliente,
    clienteTelefono: turno.telefono,
    servicioNombre: turno.servicioNombre,
    fecha: turno.fecha,
    hora: turno.hora,
    serviceId: svc?.id ?? null,
    yaCobrado: Boolean(venta),
    ventaId: venta?.id ?? null,
  };
}

export async function crearVentaDesdePaquete(params: {
  clientPackageId: number;
  metodoPago: string;
  notas?: string | null;
}): Promise<
  | { ok: true; id: number; numero: number }
  | {
      ok: false;
      reason:
        | "no_db"
        | "not_found"
        | "ya_cobrado"
        | "sin_sesion"
        | "sesion_cerrada"
        | "sin_monto";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [asig] = await db
    .select({
      id: clientPackages.id,
      packageId: clientPackages.packageId,
      nombreCliente: clientPackages.nombreCliente,
      telefono: clientPackages.telefono,
      precioCobradoPesos: clientPackages.precioCobradoPesos,
      paqueteNombre: servicePackages.nombre,
    })
    .from(clientPackages)
    .innerJoin(
      servicePackages,
      eq(clientPackages.packageId, servicePackages.id)
    )
    .where(eq(clientPackages.id, params.clientPackageId))
    .limit(1);

  if (!asig) return { ok: false, reason: "not_found" };

  const existente = await ventaActivaPorReferencia({
    clientPackageId: asig.id,
  });
  if (existente) return { ok: false, reason: "ya_cobrado" };

  const monto = pesos(asig.precioCobradoPesos ?? 0);
  if (monto <= 0) return { ok: false, reason: "sin_monto" };

  const venta = await crearVenta({
    clienteTelefono: asig.telefono,
    clienteNombre: asig.nombreCliente,
    metodoPago: params.metodoPago,
    notas: params.notas,
    clientPackageId: asig.id,
    lineas: [
      {
        tipo: "paquete",
        descripcion: `Paquete: ${asig.paqueteNombre}`,
        cantidad: 1,
        precioUnitarioPesos: monto,
        servicePackageId: asig.packageId,
      },
    ],
  });
  if (venta.ok) return venta;
  if (venta.reason === "sin_lineas" || venta.reason === "invalido") {
    return { ok: false, reason: "sin_monto" };
  }
  if (
    venta.reason === "sin_sesion" ||
    venta.reason === "sesion_cerrada" ||
    venta.reason === "no_db"
  ) {
    return { ok: false, reason: venta.reason };
  }
  return { ok: false, reason: "sin_monto" };
}

export async function crearVentaDesdeTurno(params: {
  appointmentId: number;
  precioPesos: number;
  metodoPago: string;
  descuentoPesos?: number;
  notas?: string | null;
}): Promise<
  | { ok: true; id: number; numero: number }
  | {
      ok: false;
      reason:
        | "no_db"
        | "not_found"
        | "ya_cobrado"
        | "sin_sesion"
        | "sesion_cerrada"
        | "sin_monto"
        | "invalido";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [turno] = await db
    .select()
    .from(appointments)
    .where(eq(appointments.id, params.appointmentId))
    .limit(1);

  if (!turno || turno.estado !== "activo") {
    return { ok: false, reason: "not_found" };
  }

  const existente = await ventaActivaPorReferencia({
    appointmentId: turno.id,
  });
  if (existente) return { ok: false, reason: "ya_cobrado" };

  const monto = pesos(params.precioPesos);
  if (monto <= 0) return { ok: false, reason: "sin_monto" };

  const [svc] = await db
    .select({ id: services.id })
    .from(services)
    .where(eq(services.nombre, turno.servicioNombre))
    .limit(1);

  const fechaHora = `${turno.fecha} ${turno.hora}`;

  const venta = await crearVenta({
    clienteTelefono: turno.telefono,
    clienteNombre: turno.nombreCliente,
    metodoPago: params.metodoPago,
    descuentoPesos: params.descuentoPesos,
    notas: params.notas,
    appointmentId: turno.id,
    lineas: [
      {
        tipo: "servicio",
        descripcion: `${turno.servicioNombre} (${fechaHora})`,
        cantidad: 1,
        precioUnitarioPesos: monto,
        serviceId: svc?.id,
      },
    ],
  });
  if (venta.ok) return venta;
  if (venta.reason === "sin_lineas" || venta.reason === "invalido") {
    return { ok: false, reason: "sin_monto" };
  }
  if (
    venta.reason === "sin_sesion" ||
    venta.reason === "sesion_cerrada" ||
    venta.reason === "no_db"
  ) {
    return { ok: false, reason: venta.reason };
  }
  return { ok: false, reason: "sin_monto" };
}
