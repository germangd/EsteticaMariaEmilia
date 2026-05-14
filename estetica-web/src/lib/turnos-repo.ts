import { and, eq, ne, sql } from "drizzle-orm";
import { getDb } from "@/db/client";
import { appointments } from "@/db/schema";
import { generarCodigo, normalizarHora } from "@/lib/agenda";
import { getNeonSql } from "@/lib/db";

function isUniqueViolation(e: unknown): boolean {
  const msg = String(e);
  return (
    msg.includes("23505") ||
    msg.toLowerCase().includes("unique") ||
    msg.toLowerCase().includes("duplicate")
  );
}

/** Inserta turno solo si cupo < capacidad (una sentencia en DB). */
export async function insertarTurnoSiHayCupo(params: {
  fecha: string;
  hora: string;
  nombre: string;
  telefono: string;
  email: string | null;
  servicioNombre: string;
  responsable: string;
  capacidad: number;
}): Promise<
  | { ok: true; id: number; codigo: string }
  | { ok: false; reason: "no_db" | "cupo" | "codigo_duplicado" }
> {
  const neonSql = getNeonSql();
  if (!neonSql) return { ok: false, reason: "no_db" };

  const email = params.email?.trim() || null;

  for (let intento = 0; intento < 8; intento++) {
    const codigo = generarCodigo();
    try {
      const rows = await neonSql`
        WITH ocupacion AS (
          SELECT COUNT(*)::int AS c
          FROM appointments
          WHERE fecha = ${params.fecha}::date
            AND hora = ${params.hora}
            AND servicio_nombre = ${params.servicioNombre}
            AND responsable = ${params.responsable}
            AND estado = 'activo'
        )
        INSERT INTO appointments (
          fecha, hora, nombre_cliente, telefono, email,
          servicio_nombre, responsable, codigo_cancelacion, estado
        )
        SELECT
          ${params.fecha}::date,
          ${params.hora},
          ${params.nombre},
          ${params.telefono},
          ${email},
          ${params.servicioNombre},
          ${params.responsable},
          ${codigo},
          'activo'
        FROM ocupacion
        WHERE ocupacion.c < ${params.capacidad}
        RETURNING id, codigo_cancelacion
      `;
      const list = rows as unknown as {
        id: number;
        codigo_cancelacion: string;
      }[];
      const row = list[0];
      if (row) return { ok: true, id: row.id, codigo: row.codigo_cancelacion };
      return { ok: false, reason: "cupo" };
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

/** Cupos activos por hora para un servicio/responsable/fecha (paridad `obtenerHorariosDisponibles`). */
export async function contarActivosPorHora(
  fecha: string,
  servicioNombre: string,
  responsable: string
): Promise<Map<string, number>> {
  const db = getDb();
  if (!db) return new Map();

  const rows = await db
    .select({
      hora: appointments.hora,
      n: sql<number>`count(*)::int`.as("n"),
    })
    .from(appointments)
    .where(
      and(
        eq(appointments.fecha, fecha),
        eq(appointments.servicioNombre, servicioNombre),
        eq(appointments.responsable, responsable),
        eq(appointments.estado, "activo")
      )
    )
    .groupBy(appointments.hora);

  const map = new Map<string, number>();
  for (const r of rows) {
    map.set(normalizarHora(r.hora), Number(r.n));
  }
  return map;
}
