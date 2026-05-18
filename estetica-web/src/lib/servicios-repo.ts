import { asc, eq } from "drizzle-orm";
import { getDb } from "@/db/client";
import { services, type ServiceRow } from "@/db/schema";
import { padHoraHHmm } from "@/lib/servicio-format";

export type ServicioInput = {
  nombre: string;
  duracionMin: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
  precioPesos?: number;
};

function normalizeInput(input: ServicioInput): ServicioInput {
  return {
    nombre: input.nombre.trim(),
    duracionMin: Math.max(5, Math.round(input.duracionMin)),
    responsable: input.responsable.trim() || "No asignado",
    capacidad: Math.max(1, Math.round(input.capacidad)),
    horarioInicio: padHoraHHmm(input.horarioInicio || "09:00"),
    horarioFin: padHoraHHmm(input.horarioFin || "18:00"),
    precioPesos: Math.max(0, Math.round(input.precioPesos ?? 0)),
  };
}

export async function listarServiciosAdmin(): Promise<
  ServiceRow[] | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  const rows = await db
    .select()
    .from(services)
    .orderBy(asc(services.nombre), asc(services.id));
  return rows;
}

export async function crearServicio(
  input: ServicioInput
): Promise<
  | { ok: true; id: number }
  | { ok: false; reason: "no_db" | "duplicado" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  const data = normalizeInput(input);
  if (!data.nombre) return { ok: false, reason: "invalido" };

  const todos = await db.select({ id: services.id, nombre: services.nombre }).from(services);
  const key = data.nombre.toLowerCase();
  if (todos.some((r) => r.nombre.trim().toLowerCase() === key)) {
    return { ok: false, reason: "duplicado" };
  }

  const [row] = await db
    .insert(services)
    .values({
      nombre: data.nombre,
      duracionMin: data.duracionMin,
      responsable: data.responsable,
      capacidad: data.capacidad,
      horarioInicio: data.horarioInicio,
      horarioFin: data.horarioFin,
      precioPesos: data.precioPesos ?? 0,
    })
    .returning({ id: services.id });

  return { ok: true, id: row.id };
}

export async function actualizarServicio(
  id: number,
  input: ServicioInput
): Promise<
  | { ok: true }
  | { ok: false; reason: "no_db" | "not_found" | "duplicado" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const data = normalizeInput(input);
  if (!data.nombre) return { ok: false, reason: "invalido" };

  const todos = await db.select({ id: services.id, nombre: services.nombre }).from(services);
  const key = data.nombre.toLowerCase();
  if (todos.some((r) => r.id !== id && r.nombre.trim().toLowerCase() === key)) {
    return { ok: false, reason: "duplicado" };
  }

  const updated = await db
    .update(services)
    .set({
      nombre: data.nombre,
      duracionMin: data.duracionMin,
      responsable: data.responsable,
      capacidad: data.capacidad,
      horarioInicio: data.horarioInicio,
      horarioFin: data.horarioFin,
      precioPesos: data.precioPesos ?? 0,
    })
    .where(eq(services.id, id))
    .returning({ id: services.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

export async function eliminarServicio(
  id: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const deleted = await db
    .delete(services)
    .where(eq(services.id, id))
    .returning({ id: services.id });

  if (deleted.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}
