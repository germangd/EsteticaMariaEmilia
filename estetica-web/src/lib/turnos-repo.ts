import { and, asc, eq, gte, inArray, lte, ne } from "drizzle-orm";
import { ESTADO_TURNO, ESTADOS_OCUPAN_CUPO } from "@/lib/appointment-estado";
import { getDb } from "@/db/client";
import {
  appointments,
  sedes,
  services,
  type AppointmentRow,
} from "@/db/schema";
import {
  contarSolapamiento,
  generarCodigo,
  horaAMinutos,
  normalizarHora,
  type TurnoOcupado,
} from "@/lib/agenda";
import { cupoCompartidoEntreSedes } from "@/lib/cupo-sedes";
import { dedupeNombresServicio } from "@/lib/servicio-format";

function isUniqueViolation(e: unknown): boolean {
  const msg = String(e);
  return (
    msg.includes("23505") ||
    msg.toLowerCase().includes("unique") ||
    msg.toLowerCase().includes("duplicate")
  );
}

/**
 * Turnos activos del día con duración del catálogo (cupo por solapamiento).
 * Por defecto incluye todas las sedes (ver `cupoCompartidoEntreSedes`).
 */
export async function listarActivosConDuracion(
  fecha: string,
  responsable: string,
  sedeId?: number
): Promise<TurnoOcupado[]> {
  const db = getDb();
  if (!db) return [];

  const parts = [
    eq(appointments.fecha, fecha),
    eq(appointments.responsable, responsable),
    inArray(appointments.estado, [...ESTADOS_OCUPAN_CUPO]),
  ];
  if (!cupoCompartidoEntreSedes() && sedeId != null && sedeId > 0) {
    parts.push(eq(appointments.sedeId, sedeId));
  }

  const rows = await db
    .select({
      hora: appointments.hora,
      duracionMin: services.duracionMin,
    })
    .from(appointments)
    .innerJoin(services, eq(appointments.servicioNombre, services.nombre))
    .where(and(...parts));

  return rows.map((r) => ({
    hora: normalizarHora(r.hora),
    duracionMin: Math.max(5, r.duracionMin),
  }));
}

/** Inserta turno si hay cupo considerando la duración del servicio. */
export type TurnoListado = AppointmentRow & { sedeNombre: string };

export async function insertarTurnoSiHayCupo(params: {
  fecha: string;
  hora: string;
  nombre: string;
  telefono: string;
  email: string | null;
  servicioNombre: string;
  responsable: string;
  sedeId: number;
  capacidad: number;
  duracionMin: number;
  estado?: typeof ESTADO_TURNO.ACTIVO | typeof ESTADO_TURNO.PENDIENTE_ANTICIPO;
}): Promise<
  | { ok: true; id: number; codigo: string }
  | { ok: false; reason: "no_db" | "cupo" | "codigo_duplicado" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const hora = normalizarHora(params.hora);
  const inicio = horaAMinutos(hora);
  if (inicio < 0) return { ok: false, reason: "cupo" };

  const duracion = Math.max(5, Math.round(params.duracionMin));
  const ocupados = await listarActivosConDuracion(
    params.fecha,
    params.responsable,
    params.sedeId
  );
  if (contarSolapamiento(inicio, duracion, ocupados) >= params.capacidad) {
    return { ok: false, reason: "cupo" };
  }

  const email = params.email?.trim() || null;

  for (let intento = 0; intento < 8; intento++) {
    const codigo = generarCodigo();
    try {
      const [row] = await db
        .insert(appointments)
        .values({
          fecha: params.fecha,
          hora,
          nombreCliente: params.nombre,
          telefono: params.telefono,
          email,
          servicioNombre: params.servicioNombre,
          responsable: params.responsable,
          sedeId: params.sedeId,
          codigoCancelacion: codigo,
          estado: params.estado ?? ESTADO_TURNO.ACTIVO,
        })
        .returning({ id: appointments.id, codigoCancelacion: appointments.codigoCancelacion });
      if (row) return { ok: true, id: row.id, codigo: row.codigoCancelacion };
    } catch (e) {
      if (isUniqueViolation(e)) continue;
      throw e;
    }
  }

  return { ok: false, reason: "codigo_duplicado" };
}

export async function cancelarTurnoPorCodigo(
  codigo: string
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  const c = codigo.trim().toUpperCase();
  if (!c) return { ok: false, reason: "not_found" };

  const updated = await db
    .update(appointments)
    .set({ estado: "cancelado" })
    .where(
      and(
        eq(appointments.codigoCancelacion, c),
        ne(appointments.estado, "cancelado")
      )
    )
    .returning({ id: appointments.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}


/** Nombres de servicio del catálogo (para filtros en admin). */
export async function listarNombresServiciosCatalogo(): Promise<string[]> {
  const db = getDb();
  if (!db) return [];
  const rows = await db
    .select({ n: services.nombre })
    .from(services)
    .where(eq(services.esGrupo, false))
    .orderBy(asc(services.nombre));
  return dedupeNombresServicio(rows.map((r) => r.n.trim()).filter(Boolean));
}

/**
 * Turnos activos con filtros opcionales; orden por fecha y hora.
 * `fechaDesde` / `fechaHasta`: ISO `yyyy-MM-dd`.
 */
export async function listarTurnosActivosFiltrados(params: {
  fechaDesde: string;
  fechaHasta?: string | null;
  servicioNombre?: string | null;
  sedeId?: number | null;
}): Promise<TurnoListado[]> {
  const db = getDb();
  if (!db) return [];
  const parts = [
    eq(appointments.estado, ESTADO_TURNO.ACTIVO),
    gte(appointments.fecha, params.fechaDesde),
  ];
  const hasta = params.fechaHasta?.trim();
  if (hasta) parts.push(lte(appointments.fecha, hasta));
  const svc = params.servicioNombre?.trim();
  if (svc) parts.push(eq(appointments.servicioNombre, svc));
  if (params.sedeId != null && params.sedeId > 0) {
    parts.push(eq(appointments.sedeId, params.sedeId));
  }
  const rows = await db
    .select({
      turno: appointments,
      sedeNombre: sedes.nombre,
    })
    .from(appointments)
    .innerJoin(sedes, eq(appointments.sedeId, sedes.id))
    .where(and(...parts))
    .orderBy(asc(appointments.fecha), asc(appointments.hora));
  return rows.map((r) => ({ ...r.turno, sedeNombre: r.sedeNombre }));
}

/** Turnos con anticipo pendiente de confirmación (ocupan cupo hasta confirmar o cancelar). */
export async function listarTurnosPendientesAnticipo(params: {
  fechaDesde: string;
  fechaHasta?: string | null;
  sedeId?: number | null;
}): Promise<TurnoListado[]> {
  const db = getDb();
  if (!db) return [];
  const parts = [
    eq(appointments.estado, ESTADO_TURNO.PENDIENTE_ANTICIPO),
    gte(appointments.fecha, params.fechaDesde),
  ];
  const hasta = params.fechaHasta?.trim();
  if (hasta) parts.push(lte(appointments.fecha, hasta));
  if (params.sedeId != null && params.sedeId > 0) {
    parts.push(eq(appointments.sedeId, params.sedeId));
  }
  const rows = await db
    .select({
      turno: appointments,
      sedeNombre: sedes.nombre,
    })
    .from(appointments)
    .innerJoin(sedes, eq(appointments.sedeId, sedes.id))
    .where(and(...parts))
    .orderBy(asc(appointments.fecha), asc(appointments.hora));
  return rows.map((r) => ({ ...r.turno, sedeNombre: r.sedeNombre }));
}

export async function confirmarTurnoAnticipo(
  appointmentId: number
): Promise<
  | { ok: true; turno: TurnoListado }
  | { ok: false; reason: "no_db" | "not_found" | "no_pendiente" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(appointmentId) || appointmentId < 1) {
    return { ok: false, reason: "not_found" };
  }

  const [row] = await db
    .update(appointments)
    .set({ estado: ESTADO_TURNO.ACTIVO })
    .where(
      and(
        eq(appointments.id, appointmentId),
        eq(appointments.estado, ESTADO_TURNO.PENDIENTE_ANTICIPO)
      )
    )
    .returning();

  if (!row) return { ok: false, reason: "no_pendiente" };

  const [sede] = await db
    .select({ nombre: sedes.nombre })
    .from(sedes)
    .where(eq(sedes.id, row.sedeId))
    .limit(1);

  return {
    ok: true,
    turno: { ...row, sedeNombre: sede?.nombre ?? "—" },
  };
}

export async function obtenerTurnoPorId(
  id: number
): Promise<TurnoListado | null> {
  const db = getDb();
  if (!db || !Number.isFinite(id) || id < 1) return null;
  const [r] = await db
    .select({
      turno: appointments,
      sedeNombre: sedes.nombre,
    })
    .from(appointments)
    .innerJoin(sedes, eq(appointments.sedeId, sedes.id))
    .where(eq(appointments.id, id))
    .limit(1);
  if (!r) return null;
  return { ...r.turno, sedeNombre: r.sedeNombre };
}
