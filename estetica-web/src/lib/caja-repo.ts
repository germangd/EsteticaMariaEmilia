import { and, desc, eq, ilike, max, or, sql, sum } from "drizzle-orm";
import { alias } from "drizzle-orm/pg-core";
import { getDb } from "@/db/client";
import {
  appointments,
  cashSessions,
  clientPackages,
  saleAuditLog,
  saleLines,
  sales,
  servicePackages,
  services,
} from "@/db/schema";
import { normalizarTelefono } from "@/lib/clientes-repo";
import { ESTADO_TURNO } from "@/lib/appointment-estado";
import { calcularAnticipoPesos } from "@/lib/servicio-anticipo";

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
  precioSugeridoPesos: number;
  anticipoRequerido: boolean;
  anticipoPorcentaje: number;
  anticipoSugeridoPesos: number;
  /** Turno web aún sin confirmar: sugerir cobrar solo el anticipo. */
  pendienteAnticipo: boolean;
  /** Suma de ventas completadas vinculadas al turno. */
  totalAbonadoPesos: number;
  /** Precio del tratamiento menos lo ya cobrado. */
  saldoPendientePesos: number;
  /** Importe sugerido para la línea en caja (anticipo, saldo o total). */
  importeCobroSugeridoPesos: number;
  yaCobrado: boolean;
  ventaId: number | null;
};

export type PrefillCobroPaquete = {
  clientPackageId: number;
  packageId: number;
  clienteNombre: string;
  clienteTelefono: string;
  paqueteNombre: string;
  fechaCompra: string;
  precioSugeridoPesos: number;
  yaCobrado: boolean;
  ventaId: number | null;
};

export type VentaHistorial = VentaResumen & {
  sessionOpenedAt: string;
  sessionStatus: string;
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
  numeroTicket: number;
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

export type VentaAuditoria = {
  id: number;
  saleId: number;
  accion: string;
  detalle: string | null;
  createdAt: string;
  datosAntes: unknown | null;
  datosDespues: unknown | null;
};

/** Desglose de anticipo / saldo cuando la venta está vinculada a un turno con anticipo. */
export type ResumenAnticipoTicket = {
  servicioNombre: string;
  precioTratamientoPesos: number;
  anticipoPorcentaje: number;
  anticipoReferenciaPesos: number;
  totalAbonadoPesos: number;
  importeEsteComprobantePesos: number;
  saldoPendientePesos: number;
};

export type VentaDetalle = VentaResumen & {
  lineas: VentaLinea[];
  sessionOpenedAt: string;
  auditoria: VentaAuditoria[];
  resumenAnticipo: ResumenAnticipoTicket | null;
};

export type ModificarVentaInput = {
  clienteTelefono?: string;
  clienteNombre?: string;
  descuentoPesos?: number;
  metodoPago: string;
  notas?: string | null;
  lineas: LineaVentaInput[];
  motivo?: string | null;
};

export type BuscarVentasFiltros = {
  fechaDesde?: string;
  fechaHasta?: string;
  cliente?: string;
  limite?: number;
  offset?: number;
  /** Sin tope de paginación (máx. 5000) para exportar CSV. */
  sinLimite?: boolean;
};

function parseAuditoriaJson(raw: string | null): unknown | null {
  if (!raw) return null;
  try {
    return JSON.parse(raw) as unknown;
  } catch {
    return raw;
  }
}

function condicionesHistorialVentas(filtros: BuscarVentasFiltros) {
  const conds = [];
  if (filtros.fechaDesde?.trim()) {
    conds.push(
      sql`(${sales.createdAt}::date >= ${filtros.fechaDesde.trim()}::date)`
    );
  }
  if (filtros.fechaHasta?.trim()) {
    conds.push(
      sql`(${sales.createdAt}::date <= ${filtros.fechaHasta.trim()}::date)`
    );
  }
  const clienteQ = filtros.cliente?.trim();
  if (clienteQ) {
    const like = `%${clienteQ}%`;
    conds.push(
      or(
        ilike(sales.clienteNombre, like),
        ilike(sales.clienteTelefono, like)
      )!
    );
  }
  return conds.length > 0 ? and(...conds) : undefined;
}

function pesos(n: unknown): number {
  const x = Number(n);
  if (!Number.isFinite(x)) return 0;
  return Math.max(0, Math.round(x));
}

type Db = NonNullable<ReturnType<typeof getDb>>;

function mapSaleRow(r: typeof sales.$inferSelect): VentaResumen {
  return {
    id: r.id,
    sessionId: r.sessionId,
    numero: r.numero,
    numeroTicket: r.numeroTicket,
    clienteTelefono: r.clienteTelefono,
    clienteNombre: r.clienteNombre,
    subtotalPesos: r.subtotalPesos,
    descuentoPesos: r.descuentoPesos,
    totalPesos: r.totalPesos,
    metodoPago: r.metodoPago,
    notas: r.notas,
    estado: r.estado,
    createdAt: r.createdAt.toISOString(),
  };
}

async function siguienteNumeroTicket(db: Db): Promise<number> {
  try {
    const res = await db.execute(
      sql`SELECT nextval('sales_numero_ticket_seq')::int AS n`
    );
    const row =
      res && typeof res === "object" && "rows" in res && Array.isArray(res.rows)
        ? (res.rows[0] as { n?: number })
        : Array.isArray(res)
          ? (res[0] as { n?: number })
          : undefined;
    const n = Number(row?.n);
    if (Number.isFinite(n) && n > 0) return n;
  } catch {
    /* secuencia aún no migrada */
  }
  const [maxRow] = await db.select({ n: max(sales.numeroTicket) }).from(sales);
  return (maxRow?.n ?? 0) + 1;
}

async function registrarAuditoriaVenta(
  db: Db,
  params: {
    saleId: number;
    accion: "creada" | "modificada" | "anulada";
    detalle?: string | null;
    antes?: unknown;
    despues?: unknown;
  }
): Promise<void> {
  await db.insert(saleAuditLog).values({
    saleId: params.saleId,
    accion: params.accion,
    detalle: params.detalle?.trim() || null,
    datosAntes: params.antes != null ? JSON.stringify(params.antes) : null,
    datosDespues: params.despues != null ? JSON.stringify(params.despues) : null,
  });
}

export async function listarAuditoriaVenta(
  saleId: number
): Promise<VentaAuditoria[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(saleAuditLog)
    .where(eq(saleAuditLog.saleId, saleId))
    .orderBy(desc(saleAuditLog.createdAt));

  return rows.map((r) => ({
    id: r.id,
    saleId: r.saleId,
    accion: r.accion,
    detalle: r.detalle,
    createdAt: r.createdAt.toISOString(),
    datosAntes: parseAuditoriaJson(r.datosAntes),
    datosDespues: parseAuditoriaJson(r.datosDespues),
  }));
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

  return rows.map(mapSaleRow);
}

/** Busca ventas con filtros (historial de movimientos). */
export async function buscarVentasHistorial(
  filtros: BuscarVentasFiltros = {}
): Promise<
  | { ventas: VentaHistorial[]; total: number }
  | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const cap = filtros.sinLimite
    ? 5000
    : Math.min(200, Math.max(1, Math.round(filtros.limite ?? 50)));
  const offset = filtros.sinLimite
    ? 0
    : Math.max(0, Math.round(filtros.offset ?? 0));

  const whereClause = condicionesHistorialVentas(filtros);

  const [countRow] = await db
    .select({ c: sql<number>`count(*)::int` })
    .from(sales)
    .where(whereClause);

  const rows = await db
    .select({
      sale: sales,
      sessionOpenedAt: cashSessions.openedAt,
      sessionStatus: cashSessions.status,
    })
    .from(sales)
    .innerJoin(cashSessions, eq(sales.sessionId, cashSessions.id))
    .where(whereClause)
    .orderBy(desc(sales.createdAt))
    .limit(cap)
    .offset(offset);

  const ventas: VentaHistorial[] = rows.map((r) => ({
    ...mapSaleRow(r.sale),
    sessionOpenedAt: r.sessionOpenedAt.toISOString(),
    sessionStatus: r.sessionStatus,
  }));

  return { ventas, total: Number(countRow?.c ?? 0) };
}

/** CSV del historial con los mismos filtros que la búsqueda en pantalla. */
export function generarCsvHistorialVentas(
  ventas: VentaHistorial[],
  filtros: BuscarVentasFiltros
): string {
  const lines: string[] = [];
  lines.push("HISTORIAL DE VENTAS - CAJA");
  if (filtros.fechaDesde) lines.push(csvRow(["Desde", filtros.fechaDesde]));
  if (filtros.fechaHasta) lines.push(csvRow(["Hasta", filtros.fechaHasta]));
  if (filtros.cliente?.trim()) {
    lines.push(csvRow(["Cliente", filtros.cliente.trim()]));
  }
  lines.push(csvRow(["Registros exportados", ventas.length]));
  lines.push("");
  lines.push(
    csvRow([
      "Ticket N",
      "Fecha hora",
      "Estado",
      "Cliente",
      "Telefono",
      "Metodo pago",
      "Subtotal ARS",
      "Descuento ARS",
      "Total ARS",
      "Notas",
      "Sesion ID",
      "Num sesion",
    ])
  );
  for (const v of ventas) {
    lines.push(
      csvRow([
        v.numeroTicket,
        v.createdAt,
        v.estado,
        v.clienteNombre ?? "",
        v.clienteTelefono ?? "",
        v.metodoPago,
        v.subtotalPesos,
        v.descuentoPesos,
        v.totalPesos,
        v.notas ?? "",
        v.sessionId,
        v.numero,
      ])
    );
  }
  return `\uFEFF${lines.join("\r\n")}\r\n`;
}

export async function exportarHistorialVentasCsv(
  filtros: BuscarVentasFiltros
): Promise<
  | { ok: true; csv: string; total: number; exportados: number }
  | { ok: false; reason: "no_db" }
> {
  const res = await buscarVentasHistorial({ ...filtros, sinLimite: true });
  if (!("ventas" in res)) return { ok: false, reason: "no_db" };
  const csv = generarCsvHistorialVentas(res.ventas, filtros);
  return {
    ok: true,
    csv,
    total: res.total,
    exportados: res.ventas.length,
  };
}

/** Últimas ventas (sin filtros). */
export async function listarVentasRecientes(
  limite = 40
): Promise<VentaHistorial[] | { ok: false; reason: "no_db" }> {
  const res = await buscarVentasHistorial({ limite });
  if (!("ventas" in res)) return { ok: false, reason: "no_db" };
  return res.ventas;
}

export async function crearVenta(
  input: CrearVentaInput
): Promise<
  | { ok: true; id: number; numero: number; numeroTicket: number }
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
  const numeroTicket = await siguienteNumeroTicket(db);

  const [venta] = await db
    .insert(sales)
    .values({
      sessionId,
      numero,
      numeroTicket,
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

  const detalleNuevo = {
    numeroTicket,
    clienteNombre: nombre,
    clienteTelefono: tel,
    subtotalPesos,
    descuentoPesos,
    totalPesos,
    metodoPago,
    notas: input.notas?.trim() || null,
    lineas: lineasCalc,
  };
  await registrarAuditoriaVenta(db, {
    saleId: venta.id,
    accion: "creada",
    detalle: "Venta registrada en caja",
    despues: detalleNuevo,
  });

  return { ok: true, id: venta.id, numero, numeroTicket };
}

async function resolverResumenAnticipoTicket(
  db: Db,
  appointmentId: number,
  ventaId: number,
  ventaEstado: string,
  ventaTotalPesos: number
): Promise<ResumenAnticipoTicket | null> {
  const [turno] = await db
    .select({
      servicioNombre: appointments.servicioNombre,
    })
    .from(appointments)
    .where(eq(appointments.id, appointmentId))
    .limit(1);

  if (!turno) return null;

  const [svc] = await db
    .select({
      precioPesos: services.precioPesos,
      anticipoRequerido: services.anticipoRequerido,
      anticipoPorcentaje: services.anticipoPorcentaje,
    })
    .from(services)
    .where(eq(services.nombre, turno.servicioNombre))
    .limit(1);

  const precioTratamientoPesos = svc?.precioPesos ?? 0;
  if (!svc?.anticipoRequerido || precioTratamientoPesos <= 0) {
    return null;
  }

  const [abonadoRow] = await db
    .select({ total: sum(sales.totalPesos) })
    .from(sales)
    .where(
      and(
        eq(sales.appointmentId, appointmentId),
        eq(sales.estado, "completada")
      )
    );

  const totalAbonadoPesos = pesos(abonadoRow?.total ?? 0);
  const anticipoPorcentaje = svc.anticipoPorcentaje ?? 0;
  const importeEsteComprobantePesos =
    ventaEstado === "completada" ? ventaTotalPesos : 0;

  return {
    servicioNombre: turno.servicioNombre,
    precioTratamientoPesos,
    anticipoPorcentaje,
    anticipoReferenciaPesos: calcularAnticipoPesos(
      precioTratamientoPesos,
      true,
      anticipoPorcentaje
    ),
    totalAbonadoPesos,
    importeEsteComprobantePesos,
    saldoPendientePesos: Math.max(0, precioTratamientoPesos - totalAbonadoPesos),
  };
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

  const auditoriaRaw = await listarAuditoriaVenta(id);
  const auditoria = Array.isArray(auditoriaRaw) ? auditoriaRaw : [];

  const s = venta.sale;
  const resumenAnticipo =
    s.appointmentId != null && s.appointmentId > 0
      ? await resolverResumenAnticipoTicket(
          db,
          s.appointmentId,
          s.id,
          s.estado,
          s.totalPesos
        )
      : null;

  return {
    ...mapSaleRow(s),
    sessionOpenedAt: venta.sessionOpenedAt.toISOString(),
    lineas: lineas.map((l) => ({
      id: l.id,
      tipo: l.tipo,
      descripcion: l.descripcion,
      cantidad: l.cantidad,
      precioUnitarioPesos: l.precioUnitarioPesos,
      totalLineaPesos: l.totalLineaPesos,
    })),
    auditoria,
    resumenAnticipo,
  };
}

export async function anularVenta(
  id: number,
  motivo?: string | null
): Promise<
  | { ok: true }
  | { ok: false; reason: "no_db" | "not_found" | "ya_anulada" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const antes = await obtenerVentaDetalle(id);
  if (antes && typeof antes === "object" && "ok" in antes) {
    return { ok: false, reason: "no_db" };
  }
  if (!antes) return { ok: false, reason: "not_found" };
  if (antes.estado === "anulada") return { ok: false, reason: "ya_anulada" };

  await db
    .update(sales)
    .set({ estado: "anulada" })
    .where(eq(sales.id, id));

  await registrarAuditoriaVenta(db, {
    saleId: id,
    accion: "anulada",
    detalle: motivo?.trim() || "Comprobante anulado",
    antes,
    despues: { ...antes, estado: "anulada" },
  });

  return { ok: true };
}

export async function modificarVenta(
  id: number,
  input: ModificarVentaInput
): Promise<
  | { ok: true }
  | {
      ok: false;
      reason:
        | "no_db"
        | "not_found"
        | "anulada"
        | "sin_lineas"
        | "invalido";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const antes = await obtenerVentaDetalle(id);
  if (antes && typeof antes === "object" && "ok" in antes) {
    return { ok: false, reason: "no_db" };
  }
  if (!antes) return { ok: false, reason: "not_found" };
  if (antes.estado === "anulada") return { ok: false, reason: "anulada" };

  const lineasNorm = input.lineas
    .map(normalizeLinea)
    .filter((l): l is LineaVentaInput => l !== null);
  if (lineasNorm.length === 0) return { ok: false, reason: "sin_lineas" };

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

  await db
    .update(sales)
    .set({
      clienteTelefono: tel,
      clienteNombre: nombre,
      subtotalPesos,
      descuentoPesos,
      totalPesos,
      metodoPago,
      notas: input.notas?.trim() || null,
    })
    .where(eq(sales.id, id));

  await db.delete(saleLines).where(eq(saleLines.saleId, id));
  await db.insert(saleLines).values(
    lineasCalc.map((l) => ({
      saleId: id,
      tipo: l.tipo,
      descripcion: l.descripcion,
      cantidad: l.cantidad,
      precioUnitarioPesos: l.precioUnitarioPesos,
      totalLineaPesos: l.totalLineaPesos,
      serviceId: l.serviceId ?? null,
      servicePackageId: l.servicePackageId ?? null,
    }))
  );

  const despues = await obtenerVentaDetalle(id);
  if (!despues || (despues && typeof despues === "object" && "ok" in despues)) {
    return { ok: false, reason: "no_db" };
  }

  await registrarAuditoriaVenta(db, {
    saleId: id,
    accion: "modificada",
    detalle: input.motivo?.trim() || "Comprobante modificado",
    antes,
    despues,
  });

  return { ok: true };
}

export type CatalogoCaja = {
  servicios: {
    id: number;
    nombre: string;
    precioPesos: number;
    categoriaNombre: string | null;
    anticipoRequerido: boolean;
    anticipoPorcentaje: number;
  }[];
  paquetes: { id: number; nombre: string; precioPesos: number }[];
};

export async function obtenerCatalogoCaja(): Promise<
  CatalogoCaja | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const categoria = alias(services, "categoria_caja");

  const servs = await db
    .select({
      id: services.id,
      nombre: services.nombre,
      precioPesos: services.precioPesos,
      categoriaNombre: categoria.nombre,
      anticipoRequerido: services.anticipoRequerido,
      anticipoPorcentaje: services.anticipoPorcentaje,
    })
    .from(services)
    .leftJoin(categoria, eq(services.parentId, categoria.id))
    .where(eq(services.esGrupo, false))
    .orderBy(categoria.nombre, services.nombre);

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

  if (
    !turno ||
    (turno.estado !== "activo" && turno.estado !== "pendiente_anticipo")
  ) {
    return null;
  }

  const [svc] = await db
    .select({
      id: services.id,
      precioPesos: services.precioPesos,
      anticipoRequerido: services.anticipoRequerido,
      anticipoPorcentaje: services.anticipoPorcentaje,
    })
    .from(services)
    .where(
      sql`lower(trim(${services.nombre})) = lower(trim(${turno.servicioNombre}))`
    )
    .limit(1);

  const [abonadoRow] = await db
    .select({ total: sum(sales.totalPesos) })
    .from(sales)
    .where(
      and(
        eq(sales.appointmentId, appointmentId),
        eq(sales.estado, "completada")
      )
    );

  const [ultimaVenta] = await db
    .select({ id: sales.id })
    .from(sales)
    .where(
      and(
        eq(sales.appointmentId, appointmentId),
        eq(sales.estado, "completada")
      )
    )
    .orderBy(desc(sales.id))
    .limit(1);

  const precioSugeridoPesos = svc?.precioPesos ?? 0;
  const anticipoRequerido = Boolean(svc?.anticipoRequerido);
  const anticipoPorcentaje = svc?.anticipoPorcentaje ?? 0;
  const anticipoSugeridoPesos = calcularAnticipoPesos(
    precioSugeridoPesos,
    anticipoRequerido,
    anticipoPorcentaje
  );
  const pendienteAnticipo =
    turno.estado === ESTADO_TURNO.PENDIENTE_ANTICIPO;
  const totalAbonadoPesos = pesos(abonadoRow?.total ?? 0);
  const saldoPendientePesos = Math.max(
    0,
    precioSugeridoPesos - totalAbonadoPesos
  );

  let importeCobroSugeridoPesos = precioSugeridoPesos;
  if (pendienteAnticipo) {
    importeCobroSugeridoPesos =
      anticipoSugeridoPesos > 0 ? anticipoSugeridoPesos : precioSugeridoPesos;
  } else if (anticipoRequerido && totalAbonadoPesos > 0 && saldoPendientePesos > 0) {
    importeCobroSugeridoPesos = saldoPendientePesos;
  } else if (totalAbonadoPesos > 0 && saldoPendientePesos <= 0) {
    importeCobroSugeridoPesos = 0;
  }

  const yaCobrado = pendienteAnticipo
    ? totalAbonadoPesos > 0 &&
      (anticipoSugeridoPesos <= 0 ||
        totalAbonadoPesos >= anticipoSugeridoPesos)
    : totalAbonadoPesos > 0 && saldoPendientePesos <= 0;

  return {
    appointmentId: turno.id,
    clienteNombre: turno.nombreCliente,
    clienteTelefono: turno.telefono,
    servicioNombre: turno.servicioNombre,
    fecha: turno.fecha,
    hora: turno.hora,
    serviceId: svc?.id ?? null,
    precioSugeridoPesos,
    anticipoRequerido,
    anticipoPorcentaje,
    anticipoSugeridoPesos,
    pendienteAnticipo,
    totalAbonadoPesos,
    saldoPendientePesos,
    importeCobroSugeridoPesos,
    yaCobrado,
    ventaId: ultimaVenta?.id ?? null,
  };
}

export async function obtenerPrefillCobroPaquete(
  clientPackageId: number
): Promise<PrefillCobroPaquete | null | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(clientPackageId) || clientPackageId < 1) return null;

  const [asig] = await db
    .select({
      id: clientPackages.id,
      packageId: clientPackages.packageId,
      nombreCliente: clientPackages.nombreCliente,
      telefono: clientPackages.telefono,
      precioCobradoPesos: clientPackages.precioCobradoPesos,
      fechaCompra: clientPackages.fechaCompra,
      estado: clientPackages.estado,
      paqueteNombre: servicePackages.nombre,
      precioCatalogo: servicePackages.precioPesos,
    })
    .from(clientPackages)
    .innerJoin(
      servicePackages,
      eq(clientPackages.packageId, servicePackages.id)
    )
    .where(eq(clientPackages.id, clientPackageId))
    .limit(1);

  if (!asig || asig.estado !== "activo") return null;

  const venta = await ventaActivaPorReferencia({
    clientPackageId: asig.id,
  });

  const precioSugeridoPesos = Math.max(
    0,
    asig.precioCobradoPesos ?? asig.precioCatalogo ?? 0
  );

  return {
    clientPackageId: asig.id,
    packageId: asig.packageId,
    clienteNombre: asig.nombreCliente,
    clienteTelefono: asig.telefono,
    paqueteNombre: asig.paqueteNombre,
    fechaCompra: asig.fechaCompra,
    precioSugeridoPesos,
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

function csvEscape(s: string): string {
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function csvRow(cells: (string | number)[]): string {
  return cells.map((c) => csvEscape(String(c))).join(",");
}

export type DatosCierreSesion = {
  sesion: SesionCaja;
  ventas: VentaResumen[];
};

export async function obtenerDatosCierreSesion(
  sessionId: number
): Promise<DatosCierreSesion | null | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(sessionId) || sessionId < 1) return null;

  const [row] = await db
    .select()
    .from(cashSessions)
    .where(eq(cashSessions.id, sessionId))
    .limit(1);

  if (!row) return null;

  const ventasRaw = await listarVentasSesion(sessionId);
  if (!Array.isArray(ventasRaw)) return { ok: false, reason: "no_db" };

  const sesion = await buildSesionStats(row);
  return { sesion, ventas: ventasRaw };
}

export function generarCsvCierreSesion(data: DatosCierreSesion): string {
  const { sesion, ventas } = data;
  const lines: string[] = [];

  lines.push("REPORTE DE CIERRE DE CAJA");
  lines.push(csvRow(["Sesion ID", sesion.id]));
  lines.push(csvRow(["Apertura", sesion.openedAt]));
  lines.push(csvRow(["Cierre", sesion.closedAt ?? "-"]));
  lines.push(csvRow(["Estado", sesion.status]));
  lines.push(csvRow(["Fondo inicial ARS", sesion.openingAmountPesos]));
  lines.push(
    csvRow([
      "Monto contado al cerrar ARS",
      sesion.closingAmountPesos ?? "",
    ])
  );
  lines.push(csvRow(["Total ventas ARS", sesion.totalVentasPesos]));
  lines.push(csvRow(["Cantidad ventas", sesion.cantidadVentas]));
  lines.push(
    csvRow(["Efectivo esperado en cajon ARS", sesion.efectivoEsperadoEnCajon])
  );

  if (sesion.closingAmountPesos != null) {
    const diff = sesion.closingAmountPesos - sesion.efectivoEsperadoEnCajon;
    lines.push(csvRow(["Diferencia efectivo ARS", diff]));
  }

  lines.push("");
  lines.push("ARQUEO POR FORMA DE PAGO");
  lines.push(csvRow(["Metodo", "Operaciones", "Total ARS"]));
  for (const a of sesion.arqueoPorMetodo) {
    lines.push(csvRow([a.metodoPago, a.cantidad, a.totalPesos]));
  }
  lines.push(
    csvRow([
      "TOTAL",
      sesion.cantidadVentas,
      sesion.totalVentasPesos,
    ])
  );

  lines.push("");
  lines.push("DETALLE DE VENTAS");
  lines.push(
    csvRow([
      "Ticket N",
      "Num sesion",
      "Fecha hora",
      "Estado",
      "Cliente",
      "Telefono",
      "Metodo pago",
      "Subtotal",
      "Descuento",
      "Total",
      "Notas",
    ])
  );

  for (const v of [...ventas].reverse()) {
    lines.push(
      csvRow([
        v.numeroTicket,
        v.numero,
        v.createdAt,
        v.estado,
        v.clienteNombre ?? "",
        v.clienteTelefono ?? "",
        v.metodoPago,
        v.subtotalPesos,
        v.descuentoPesos,
        v.totalPesos,
        v.notas ?? "",
      ])
    );
  }

  return `\uFEFF${lines.join("\r\n")}\r\n`;
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

export type ResumenVentasBucket = {
  periodo: string;
  etiqueta: string;
  totalPesos: number;
  cantidad: number;
};

export type ResumenVentasPeriodo = {
  agrupacion: "dia" | "semana" | "mes";
  fechaDesde: string;
  fechaHasta: string;
  buckets: ResumenVentasBucket[];
  totalPesos: number;
  cantidad: number;
  anuladas: number;
  porMetodo: ArqueoMetodo[];
};

function condicionesFechaVentasCompletadas(
  fechaDesde?: string,
  fechaHasta?: string
) {
  const parts = [eq(sales.estado, "completada")];
  if (fechaDesde?.trim()) {
    parts.push(sql`(${sales.createdAt}::date >= ${fechaDesde.trim()}::date)`);
  }
  if (fechaHasta?.trim()) {
    parts.push(sql`(${sales.createdAt}::date <= ${fechaHasta.trim()}::date)`);
  }
  return and(...parts);
}

function condicionesFechaSesiones(fechaDesde?: string, fechaHasta?: string) {
  const parts = [];
  if (fechaDesde?.trim()) {
    parts.push(
      sql`(${cashSessions.openedAt}::date >= ${fechaDesde.trim()}::date)`
    );
  }
  if (fechaHasta?.trim()) {
    parts.push(
      sql`(${cashSessions.openedAt}::date <= ${fechaHasta.trim()}::date)`
    );
  }
  return parts.length > 0 ? and(...parts) : undefined;
}

export async function obtenerResumenVentasPeriodo(params: {
  fechaDesde: string;
  fechaHasta: string;
  agrupacion: "dia" | "semana" | "mes";
  etiquetaBucket: (
    periodo: string,
    agrupacion: "dia" | "semana" | "mes"
  ) => string;
}): Promise<ResumenVentasPeriodo | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const trunc =
    params.agrupacion === "dia"
      ? "day"
      : params.agrupacion === "semana"
        ? "week"
        : "month";
  const truncSql = sql.raw(`'${trunc}'`);

  const whereVentas = condicionesFechaVentasCompletadas(
    params.fechaDesde,
    params.fechaHasta
  );

  const bucketRows = await db
    .select({
      periodo: sql<string>`date_trunc(${truncSql}, ${sales.createdAt})::date::text`,
      totalPesos: sql<number>`coalesce(sum(${sales.totalPesos}), 0)::int`,
      cantidad: sql<number>`count(*)::int`,
    })
    .from(sales)
    .where(whereVentas)
    .groupBy(sql`date_trunc(${truncSql}, ${sales.createdAt})`)
    .orderBy(sql`date_trunc(${truncSql}, ${sales.createdAt})`);

  const [totRow] = await db
    .select({
      totalPesos: sql<number>`coalesce(sum(${sales.totalPesos}), 0)::int`,
      cantidad: sql<number>`count(*)::int`,
    })
    .from(sales)
    .where(whereVentas);

  const metodoRows = await db
    .select({
      metodoPago: sales.metodoPago,
      totalPesos: sql<number>`coalesce(sum(${sales.totalPesos}), 0)::int`,
      cantidad: sql<number>`count(*)::int`,
    })
    .from(sales)
    .where(whereVentas)
    .groupBy(sales.metodoPago)
    .orderBy(sales.metodoPago);

  const anuladasParts = [eq(sales.estado, "anulada")];
  if (params.fechaDesde?.trim()) {
    anuladasParts.push(
      sql`(${sales.createdAt}::date >= ${params.fechaDesde.trim()}::date)`
    );
  }
  if (params.fechaHasta?.trim()) {
    anuladasParts.push(
      sql`(${sales.createdAt}::date <= ${params.fechaHasta.trim()}::date)`
    );
  }
  const [anulRow] = await db
    .select({ c: sql<number>`count(*)::int` })
    .from(sales)
    .where(and(...anuladasParts));

  return {
    agrupacion: params.agrupacion,
    fechaDesde: params.fechaDesde,
    fechaHasta: params.fechaHasta,
    buckets: bucketRows.map((r) => ({
      periodo: r.periodo,
      etiqueta: params.etiquetaBucket(r.periodo, params.agrupacion),
      totalPesos: Number(r.totalPesos ?? 0),
      cantidad: Number(r.cantidad ?? 0),
    })),
    totalPesos: Number(totRow?.totalPesos ?? 0),
    cantidad: Number(totRow?.cantidad ?? 0),
    anuladas: Number(anulRow?.c ?? 0),
    porMetodo: metodoRows.map((r) => ({
      metodoPago: r.metodoPago,
      totalPesos: Number(r.totalPesos ?? 0),
      cantidad: Number(r.cantidad ?? 0),
    })),
  };
}

export async function listarSesionesCaja(params: {
  fechaDesde?: string;
  fechaHasta?: string;
  limite?: number;
}): Promise<SesionCaja[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const limite = Math.min(80, Math.max(1, Math.round(params.limite ?? 30)));
  const whereSes = condicionesFechaSesiones(
    params.fechaDesde,
    params.fechaHasta
  );

  const rows = await db
    .select()
    .from(cashSessions)
    .where(whereSes)
    .orderBy(desc(cashSessions.openedAt))
    .limit(limite);

  const out: SesionCaja[] = [];
  for (const row of rows) {
    out.push(await buildSesionStats(row));
  }
  return out;
}
