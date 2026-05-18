import { and, asc, eq, inArray } from "drizzle-orm";
import { getDb } from "@/db/client";
import {
  agendaEventServices,
  agendaEvents,
  services,
  type AgendaEventRow,
} from "@/db/schema";
import { horaAMinutos } from "@/lib/agenda";
import { normalizarTelefono } from "@/lib/clientes-repo";
import {
  esFechaIsoValida,
  franjasHorariasSeSolapan,
} from "@/lib/disponibilidad-repo";
import { padHoraHHmm } from "@/lib/servicio-format";

export type EventoInput = {
  nombre: string;
  descripcion?: string | null;
  fecha: string;
  horarioInicio: string;
  horarioFin: string;
  precioPesos?: number;
  clienteTelefono?: string | null;
  clienteNombre?: string | null;
  activo?: boolean;
  serviceIds: number[];
};

export type EventoConServicios = {
  id: number;
  nombre: string;
  descripcion: string | null;
  fecha: string;
  horarioInicio: string;
  horarioFin: string;
  precioPesos: number;
  clienteTelefono: string | null;
  clienteNombre: string | null;
  activo: boolean;
  servicios: { id: number; nombre: string }[];
};

function normalizeInput(input: EventoInput): EventoInput {
  const tel = input.clienteTelefono?.trim()
    ? normalizarTelefono(input.clienteTelefono)
    : null;
  const nombreCliente = input.clienteNombre?.trim() || null;
  const precio = Math.max(0, Math.round(Number(input.precioPesos) || 0));

  return {
    nombre: input.nombre.trim(),
    descripcion: input.descripcion?.trim() || null,
    fecha: input.fecha.trim(),
    horarioInicio: padHoraHHmm(input.horarioInicio || "09:00"),
    horarioFin: padHoraHHmm(input.horarioFin || "18:00"),
    precioPesos: precio,
    clienteTelefono: tel,
    clienteNombre: nombreCliente,
    activo: input.activo !== false,
    serviceIds: [...new Set(input.serviceIds.filter((id) => id > 0))],
  };
}

function franjaHorariaValida(inicio: string, fin: string): boolean {
  const i = horaAMinutos(inicio);
  const f = horaAMinutos(fin);
  return i >= 0 && f > i;
}

async function franjaSolapaConOtroEvento(
  fecha: string,
  horarioInicio: string,
  horarioFin: string,
  excludeId?: number
): Promise<boolean> {
  const db = getDb();
  if (!db) return false;

  const rows = await db
    .select({
      id: agendaEvents.id,
      horarioInicio: agendaEvents.horarioInicio,
      horarioFin: agendaEvents.horarioFin,
    })
    .from(agendaEvents)
    .where(and(eq(agendaEvents.fecha, fecha), eq(agendaEvents.activo, true)));

  for (const row of rows) {
    if (excludeId != null && row.id === excludeId) continue;
    if (
      franjasHorariasSeSolapan(
        horarioInicio,
        horarioFin,
        padHoraHHmm(row.horarioInicio),
        padHoraHHmm(row.horarioFin)
      )
    ) {
      return true;
    }
  }
  return false;
}

async function attachServicios(
  eventIds: number[]
): Promise<Map<number, { id: number; nombre: string }[]>> {
  const db = getDb();
  const map = new Map<number, { id: number; nombre: string }[]>();
  if (!db || eventIds.length === 0) return map;

  const links = await db
    .select({
      eventId: agendaEventServices.eventId,
      id: services.id,
      nombre: services.nombre,
    })
    .from(agendaEventServices)
    .innerJoin(services, eq(agendaEventServices.serviceId, services.id))
    .where(inArray(agendaEventServices.eventId, eventIds));

  for (const row of links) {
    const list = map.get(row.eventId) ?? [];
    list.push({ id: row.id, nombre: row.nombre });
    map.set(row.eventId, list);
  }
  return map;
}

function rowToEvento(
  r: AgendaEventRow,
  servicios: { id: number; nombre: string }[]
): EventoConServicios {
  return {
    id: r.id,
    nombre: r.nombre,
    descripcion: r.descripcion,
    fecha: r.fecha,
    horarioInicio: padHoraHHmm(r.horarioInicio),
    horarioFin: padHoraHHmm(r.horarioFin),
    precioPesos: r.precioPesos ?? 0,
    clienteTelefono: r.clienteTelefono,
    clienteNombre: r.clienteNombre,
    activo: r.activo,
    servicios,
  };
}

export async function listarEventosAdmin(): Promise<
  EventoConServicios[] | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(agendaEvents)
    .orderBy(asc(agendaEvents.fecha), asc(agendaEvents.horarioInicio), asc(agendaEvents.id));

  const svcMap = await attachServicios(rows.map((r) => r.id));
  return rows.map((r) => rowToEvento(r, svcMap.get(r.id) ?? []));
}

export async function crearEvento(
  input: EventoInput
): Promise<
  | { ok: true; id: number }
  | { ok: false; reason: "no_db" | "invalido" | "franja_solapada" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const data = normalizeInput(input);
  if (!data.nombre || !esFechaIsoValida(data.fecha)) {
    return { ok: false, reason: "invalido" };
  }
  if (data.serviceIds.length === 0) return { ok: false, reason: "invalido" };
  if (!franjaHorariaValida(data.horarioInicio, data.horarioFin)) {
    return { ok: false, reason: "invalido" };
  }

  const svcs = await db
    .select({ id: services.id, esGrupo: services.esGrupo })
    .from(services)
    .where(inArray(services.id, data.serviceIds));
  if (svcs.length !== data.serviceIds.length || svcs.some((s) => s.esGrupo)) {
    return { ok: false, reason: "invalido" };
  }

  if (
    await franjaSolapaConOtroEvento(
      data.fecha,
      data.horarioInicio,
      data.horarioFin
    )
  ) {
    return { ok: false, reason: "franja_solapada" };
  }

  const [row] = await db
    .insert(agendaEvents)
    .values({
      nombre: data.nombre,
      descripcion: data.descripcion,
      fecha: data.fecha,
      horarioInicio: data.horarioInicio,
      horarioFin: data.horarioFin,
      precioPesos: data.precioPesos,
      clienteTelefono: data.clienteTelefono,
      clienteNombre: data.clienteNombre,
      activo: data.activo,
    })
    .returning({ id: agendaEvents.id });

  await db.insert(agendaEventServices).values(
    data.serviceIds.map((serviceId) => ({
      eventId: row.id,
      serviceId,
    }))
  );

  return { ok: true, id: row.id };
}

export async function actualizarEvento(
  id: number,
  input: EventoInput
): Promise<
  | { ok: true }
  | { ok: false; reason: "no_db" | "not_found" | "invalido" | "franja_solapada" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const data = normalizeInput(input);
  if (!data.nombre || !esFechaIsoValida(data.fecha)) {
    return { ok: false, reason: "invalido" };
  }
  if (data.serviceIds.length === 0) return { ok: false, reason: "invalido" };
  if (!franjaHorariaValida(data.horarioInicio, data.horarioFin)) {
    return { ok: false, reason: "invalido" };
  }

  if (
    await franjaSolapaConOtroEvento(
      data.fecha,
      data.horarioInicio,
      data.horarioFin,
      id
    )
  ) {
    return { ok: false, reason: "franja_solapada" };
  }

  const updated = await db
    .update(agendaEvents)
    .set({
      nombre: data.nombre,
      descripcion: data.descripcion,
      fecha: data.fecha,
      horarioInicio: data.horarioInicio,
      horarioFin: data.horarioFin,
      precioPesos: data.precioPesos,
      clienteTelefono: data.clienteTelefono,
      clienteNombre: data.clienteNombre,
      activo: data.activo,
    })
    .where(eq(agendaEvents.id, id))
    .returning({ id: agendaEvents.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };

  await db.delete(agendaEventServices).where(eq(agendaEventServices.eventId, id));
  await db.insert(agendaEventServices).values(
    data.serviceIds.map((serviceId) => ({ eventId: id, serviceId }))
  );

  return { ok: true };
}

export async function eliminarEvento(
  id: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const deleted = await db
    .delete(agendaEvents)
    .where(eq(agendaEvents.id, id))
    .returning({ id: agendaEvents.id });

  if (deleted.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}
