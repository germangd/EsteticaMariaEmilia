import { asc, eq } from "drizzle-orm";
import { getDb } from "@/db/client";
import { services, type ServiceRow } from "@/db/schema";
import { padHoraHHmm } from "@/lib/servicio-format";
import { normalizarAnticipoPorcentaje } from "@/lib/servicio-anticipo";

export type ServicioInput = {
  nombre: string;
  duracionMin: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
  precioPesos?: number;
  parentId?: number | null;
  esGrupo?: boolean;
  anticipoRequerido?: boolean;
  anticipoPorcentaje?: number;
};

function normalizeInput(input: ServicioInput): ServicioInput {
  const esGrupo = input.esGrupo === true;
  return {
    nombre: input.nombre.trim(),
    duracionMin: esGrupo ? 0 : Math.max(5, Math.round(input.duracionMin)),
    responsable: input.responsable.trim() || "No asignado",
    capacidad: esGrupo ? 0 : Math.max(1, Math.round(input.capacidad)),
    horarioInicio: padHoraHHmm(input.horarioInicio || "09:00"),
    horarioFin: padHoraHHmm(input.horarioFin || "18:00"),
    precioPesos: Math.max(0, Math.round(input.precioPesos ?? 0)),
    parentId: esGrupo ? null : input.parentId ?? null,
    esGrupo,
    anticipoRequerido: esGrupo ? false : input.anticipoRequerido === true,
    anticipoPorcentaje: esGrupo
      ? 0
      : normalizarAnticipoPorcentaje(
          input.anticipoRequerido === true,
          input.anticipoPorcentaje ?? 0
        ),
  };
}

async function validarJerarquia(
  data: ServicioInput,
  id?: number
): Promise<
  | { ok: true }
  | { ok: false; reason: "invalido" | "parent_invalido" | "tiene_hijos" }
> {
  if (!data.nombre) return { ok: false, reason: "invalido" };

  if (data.esGrupo) {
    if (data.parentId != null) return { ok: false, reason: "invalido" };
    return { ok: true };
  }

  if (data.parentId == null) return { ok: true };

  const db = getDb();
  if (!db) return { ok: false, reason: "invalido" };

  if (id != null && data.parentId === id) {
    return { ok: false, reason: "parent_invalido" };
  }

  const [parent] = await db
    .select({ id: services.id, esGrupo: services.esGrupo })
    .from(services)
    .where(eq(services.id, data.parentId))
    .limit(1);

  if (!parent?.esGrupo) return { ok: false, reason: "parent_invalido" };

  if (id != null) {
    const hijos = await db
      .select({ id: services.id })
      .from(services)
      .where(eq(services.parentId, id))
      .limit(1);
    if (hijos.length > 0) return { ok: false, reason: "tiene_hijos" };
  }

  return { ok: true };
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
  | {
      ok: false;
      reason:
        | "no_db"
        | "duplicado"
        | "invalido"
        | "parent_invalido"
        | "tiene_hijos";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  const data = normalizeInput(input);

  const val = await validarJerarquia(data);
  if (!val.ok) return { ok: false, reason: val.reason };

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
      parentId: data.parentId,
      esGrupo: data.esGrupo ?? false,
      anticipoRequerido: data.anticipoRequerido,
      anticipoPorcentaje: data.anticipoPorcentaje,
    })
    .returning({ id: services.id });

  return { ok: true, id: row.id };
}

export async function actualizarServicio(
  id: number,
  input: ServicioInput
): Promise<
  | { ok: true }
  | {
      ok: false;
      reason:
        | "no_db"
        | "not_found"
        | "duplicado"
        | "invalido"
        | "parent_invalido"
        | "tiene_hijos";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const data = normalizeInput(input);

  const val = await validarJerarquia(data, id);
  if (!val.ok) return { ok: false, reason: val.reason };

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
      parentId: data.parentId,
      esGrupo: data.esGrupo ?? false,
      anticipoRequerido: data.anticipoRequerido,
      anticipoPorcentaje: data.anticipoPorcentaje,
    })
    .where(eq(services.id, id))
    .returning({ id: services.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

export async function actualizarPrecioServicio(
  id: number,
  precioPesos: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" | "es_grupo" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const [row] = await db
    .select({ esGrupo: services.esGrupo })
    .from(services)
    .where(eq(services.id, id))
    .limit(1);
  if (!row) return { ok: false, reason: "not_found" };
  if (row.esGrupo) return { ok: false, reason: "es_grupo" };

  const updated = await db
    .update(services)
    .set({ precioPesos: Math.max(0, Math.round(precioPesos)) })
    .where(eq(services.id, id))
    .returning({ id: services.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

export async function eliminarServicio(
  id: number
): Promise<
  { ok: true } | { ok: false; reason: "no_db" | "not_found" | "tiene_hijos" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const hijos = await db
    .select({ id: services.id })
    .from(services)
    .where(eq(services.parentId, id))
    .limit(1);
  if (hijos.length > 0) return { ok: false, reason: "tiene_hijos" };

  const deleted = await db
    .delete(services)
    .where(eq(services.id, id))
    .returning({ id: services.id });

  if (deleted.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

/** IDs de servicios que pueden reservarse o incluirse en paquetes. */
export async function listarIdsServiciosReservables(): Promise<number[]> {
  const db = getDb();
  if (!db) return [];
  const rows = await db
    .select({ id: services.id })
    .from(services)
    .where(eq(services.esGrupo, false));
  return rows.map((r) => r.id);
}
