import { desc, eq, sql, type SQL } from "drizzle-orm";
import type { AnyColumn } from "drizzle-orm";
import { getDb } from "@/db/client";
import {
  appointments,
  clientPackages,
  clientProfiles,
  servicePackages,
  type AppointmentRow,
} from "@/db/schema";
import { getNeonSql } from "@/lib/db";
import type { AsignacionPaquete } from "@/lib/paquetes-repo";

export function normalizarTelefono(raw: string): string {
  return raw.trim().replace(/\s+/g, "");
}

function telefonoCoincide(column: AnyColumn, telefono: string): SQL {
  return sql`regexp_replace(trim(${column}), '\\s', '', 'g') = ${telefono}`;
}

export type ClienteResumen = {
  telefono: string;
  nombre: string;
  email: string | null;
  turnosTotal: number;
  turnosActivos: number;
  paquetesActivos: number;
  ultimaFecha: string | null;
  tieneNotas: boolean;
};

export type ClienteDetalle = {
  telefono: string;
  perfil: {
    nombre: string | null;
    email: string | null;
    notas: string | null;
  };
  turnos: AppointmentRow[];
  paquetes: (AsignacionPaquete & { sesionesUsadas: number })[];
};

export async function listarClientesResumen(
  busqueda?: string
): Promise<ClienteResumen[] | { ok: false; reason: "no_db" }> {
  const neonSql = getNeonSql();
  if (!neonSql) return { ok: false, reason: "no_db" };

  const q = busqueda?.trim() ?? "";
  const like = q ? `%${q.replace(/%/g, "")}%` : null;

  const rows = await neonSql`
    WITH telefonos AS (
      SELECT DISTINCT trim(telefono) AS telefono
      FROM appointments
      WHERE trim(telefono) <> ''
      UNION
      SELECT DISTINCT trim(telefono) AS telefono
      FROM client_packages
      WHERE trim(telefono) <> ''
    ),
    turnos AS (
      SELECT
        trim(telefono) AS telefono,
        count(*)::int AS turnos_total,
        count(*) FILTER (WHERE estado = 'activo')::int AS turnos_activos,
        max(fecha)::text AS ultima_fecha,
        (array_agg(nombre_cliente ORDER BY fecha DESC, id DESC))[1] AS nombre_reciente,
        (array_agg(email ORDER BY fecha DESC, id DESC) FILTER (WHERE email IS NOT NULL AND trim(email) <> ''))[1] AS email_reciente
      FROM appointments
      WHERE trim(telefono) <> ''
      GROUP BY trim(telefono)
    ),
    paquetes AS (
      SELECT
        trim(telefono) AS telefono,
        count(*) FILTER (WHERE estado = 'activo')::int AS paquetes_activos
      FROM client_packages
      WHERE trim(telefono) <> ''
      GROUP BY trim(telefono)
    )
    SELECT
      t.telefono,
      coalesce(p.nombre, tr.nombre_reciente, 'Sin nombre') AS nombre,
      coalesce(p.email, tr.email_reciente) AS email,
      coalesce(tr.turnos_total, 0) AS turnos_total,
      coalesce(tr.turnos_activos, 0) AS turnos_activos,
      coalesce(pq.paquetes_activos, 0) AS paquetes_activos,
      tr.ultima_fecha,
      (p.notas IS NOT NULL AND trim(p.notas) <> '') AS tiene_notas
    FROM telefonos t
    LEFT JOIN turnos tr ON tr.telefono = t.telefono
    LEFT JOIN paquetes pq ON pq.telefono = t.telefono
    LEFT JOIN client_profiles p ON p.telefono = t.telefono
    WHERE (
      ${like}::text IS NULL
      OR t.telefono ILIKE ${like}
      OR coalesce(p.nombre, tr.nombre_reciente, '') ILIKE ${like}
      OR coalesce(p.email, tr.email_reciente, '') ILIKE ${like}
    )
    ORDER BY tr.ultima_fecha DESC NULLS LAST, t.telefono ASC
  `;

  return (rows as unknown as {
    telefono: string;
    nombre: string;
    email: string | null;
    turnos_total: number;
    turnos_activos: number;
    paquetes_activos: number;
    ultima_fecha: string | null;
    tiene_notas: boolean;
  }[]).map((r) => ({
    telefono: r.telefono,
    nombre: r.nombre,
    email: r.email,
    turnosTotal: r.turnos_total,
    turnosActivos: r.turnos_activos,
    paquetesActivos: r.paquetes_activos,
    ultimaFecha: r.ultima_fecha,
    tieneNotas: r.tiene_notas,
  }));
}

export async function obtenerClienteDetalle(
  telefonoRaw: string
): Promise<ClienteDetalle | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const telefono = normalizarTelefono(telefonoRaw);
  if (!telefono) return { ok: false, reason: "not_found" };

  const turnos = await db
    .select()
    .from(appointments)
    .where(telefonoCoincide(appointments.telefono, telefono))
    .orderBy(desc(appointments.fecha), desc(appointments.hora));

  if (turnos.length === 0) {
    const [soloPaquete] = await db
      .select({ telefono: clientPackages.telefono })
      .from(clientPackages)
      .where(telefonoCoincide(clientPackages.telefono, telefono))
      .limit(1);
    if (!soloPaquete) return { ok: false, reason: "not_found" };
  }

  const [perfilRow] = await db
    .select()
    .from(clientProfiles)
    .where(eq(clientProfiles.telefono, telefono))
    .limit(1);

  const paquetesRows = await db
    .select({
      id: clientPackages.id,
      packageId: clientPackages.packageId,
      paqueteNombre: servicePackages.nombre,
      nombreCliente: clientPackages.nombreCliente,
      telefono: clientPackages.telefono,
      sesionesIniciales: clientPackages.sesionesIniciales,
      sesionesRestantes: clientPackages.sesionesRestantes,
      precioCobradoPesos: clientPackages.precioCobradoPesos,
      notas: clientPackages.notas,
      estado: clientPackages.estado,
      fechaCompra: clientPackages.fechaCompra,
    })
    .from(clientPackages)
    .innerJoin(servicePackages, eq(clientPackages.packageId, servicePackages.id))
    .where(telefonoCoincide(clientPackages.telefono, telefono));

  const paquetes = paquetesRows
    .map((p) => ({
      ...p,
      sesionesUsadas: p.sesionesIniciales - p.sesionesRestantes,
    }))
    .sort((a, b) => b.fechaCompra.localeCompare(a.fechaCompra));

  const nombreReciente =
    turnos.length > 0
      ? turnos[turnos.length - 1]?.nombreCliente
      : paquetesRows[0]?.nombreCliente;

  return {
    telefono,
    perfil: {
      nombre: perfilRow?.nombre ?? nombreReciente ?? null,
      email: perfilRow?.email ?? turnos.find((t) => t.email)?.email ?? null,
      notas: perfilRow?.notas ?? null,
    },
    turnos,
    paquetes,
  };
}

export async function guardarPerfilCliente(params: {
  telefono: string;
  nombre?: string | null;
  email?: string | null;
  notas?: string | null;
}): Promise<{ ok: true; telefono: string } | { ok: false; reason: "no_db" | "invalido" }> {
  const res = await actualizarCliente({
    telefono: params.telefono,
    nombre: params.nombre,
    email: params.email,
    notas: params.notas,
  });
  if (!res.ok) {
    if (res.reason === "no_db") return { ok: false, reason: "no_db" };
    return { ok: false, reason: "invalido" };
  }
  return { ok: true, telefono: res.telefono };
}

export async function actualizarCliente(params: {
  telefono: string;
  telefonoNuevo?: string | null;
  nombre?: string | null;
  email?: string | null;
  notas?: string | null;
}): Promise<
  | { ok: true; telefono: string }
  | {
      ok: false;
      reason: "no_db" | "invalido" | "not_found" | "telefono_ocupado";
    }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const telefonoViejo = normalizarTelefono(params.telefono);
  const telefonoNuevo = normalizarTelefono(
    params.telefonoNuevo?.trim() ? params.telefonoNuevo : telefonoViejo
  );
  if (!telefonoViejo || !telefonoNuevo) return { ok: false, reason: "invalido" };

  const detalle = await obtenerClienteDetalle(telefonoViejo);
  if ("reason" in detalle) {
    return {
      ok: false,
      reason: detalle.reason === "not_found" ? "not_found" : "no_db",
    };
  }

  if (telefonoNuevo !== telefonoViejo) {
    const otro = await obtenerClienteDetalle(telefonoNuevo);
    if (!("reason" in otro)) {
      return { ok: false, reason: "telefono_ocupado" };
    }
  }

  const nombreTrim = params.nombre?.trim() ?? "";
  const nombreTurno =
    nombreTrim || detalle.perfil.nombre?.trim() || "Sin nombre";
  const emailVal = params.email?.trim() || null;
  const notasVal = params.notas?.trim() || null;

  await db
    .update(appointments)
    .set({
      telefono: telefonoNuevo,
      nombreCliente: nombreTurno,
      email: emailVal,
    })
    .where(telefonoCoincide(appointments.telefono, telefonoViejo));

  await db
    .update(clientPackages)
    .set({
      telefono: telefonoNuevo,
      nombreCliente: nombreTurno,
    })
    .where(telefonoCoincide(clientPackages.telefono, telefonoViejo));

  if (telefonoNuevo !== telefonoViejo) {
    await db
      .delete(clientProfiles)
      .where(eq(clientProfiles.telefono, telefonoViejo));
  }

  await db
    .insert(clientProfiles)
    .values({
      telefono: telefonoNuevo,
      nombre: nombreTrim || null,
      email: emailVal,
      notas: notasVal,
    })
    .onConflictDoUpdate({
      target: clientProfiles.telefono,
      set: {
        nombre: nombreTrim || null,
        email: emailVal,
        notas: notasVal,
        updatedAt: new Date(),
      },
    });

  return { ok: true, telefono: telefonoNuevo };
}

export async function eliminarCliente(
  telefonoRaw: string
): Promise<
  | { ok: true; eliminados: { turnos: number; paquetes: number } }
  | { ok: false; reason: "no_db" | "not_found" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const telefono = normalizarTelefono(telefonoRaw);
  if (!telefono) return { ok: false, reason: "invalido" };

  const detalle = await obtenerClienteDetalle(telefono);
  if ("reason" in detalle) {
    return {
      ok: false,
      reason: detalle.reason === "not_found" ? "not_found" : "no_db",
    };
  }

  const turnosEliminados = await db
    .delete(appointments)
    .where(telefonoCoincide(appointments.telefono, telefono))
    .returning({ id: appointments.id });

  const paquetesEliminados = await db
    .delete(clientPackages)
    .where(telefonoCoincide(clientPackages.telefono, telefono))
    .returning({ id: clientPackages.id });

  await db
    .delete(clientProfiles)
    .where(eq(clientProfiles.telefono, telefono));

  return {
    ok: true,
    eliminados: {
      turnos: turnosEliminados.length,
      paquetes: paquetesEliminados.length,
    },
  };
}
