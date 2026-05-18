import { asc, eq } from "drizzle-orm";
import { getDb } from "@/db/client";
import { sedes, type SedeRow } from "@/db/schema";

export type SedePublica = { id: number; nombre: string };

function rowToPublica(r: SedeRow): SedePublica {
  return { id: r.id, nombre: r.nombre };
}

export async function listarSedesActivas(): Promise<
  SedePublica[] | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(sedes)
    .where(eq(sedes.activo, true))
    .orderBy(asc(sedes.orden), asc(sedes.id));

  return rows.map(rowToPublica);
}

export async function listarSedesAdmin(): Promise<
  SedeRow[] | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  return db
    .select()
    .from(sedes)
    .orderBy(asc(sedes.orden), asc(sedes.id));
}

export async function obtenerSedePorId(
  id: number
): Promise<SedeRow | null> {
  const db = getDb();
  if (!db || !Number.isFinite(id) || id < 1) return null;

  const [row] = await db
    .select()
    .from(sedes)
    .where(eq(sedes.id, id))
    .limit(1);

  return row ?? null;
}

export async function crearSede(input: {
  nombre: string;
  orden?: number;
  activo?: boolean;
}): Promise<
  | { ok: true; id: number }
  | { ok: false; reason: "no_db" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const nombre = input.nombre.trim();
  if (!nombre) return { ok: false, reason: "invalido" };

  const [row] = await db
    .insert(sedes)
    .values({
      nombre,
      orden: Math.max(0, Math.round(Number(input.orden) || 0)),
      activo: input.activo !== false,
    })
    .returning({ id: sedes.id });

  return { ok: true, id: row.id };
}

export async function actualizarSede(
  id: number,
  input: { nombre: string; orden?: number; activo?: boolean }
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" | "invalido" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const nombre = input.nombre.trim();
  if (!nombre) return { ok: false, reason: "invalido" };

  const updated = await db
    .update(sedes)
    .set({
      nombre,
      orden: Math.max(0, Math.round(Number(input.orden) || 0)),
      activo: input.activo !== false,
    })
    .where(eq(sedes.id, id))
    .returning({ id: sedes.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

export async function eliminarSede(
  id: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const updated = await db
    .update(sedes)
    .set({ activo: false })
    .where(eq(sedes.id, id))
    .returning({ id: sedes.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}
