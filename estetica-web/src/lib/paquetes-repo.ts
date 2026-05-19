import { and, asc, desc, eq, inArray } from "drizzle-orm";
import { getDb } from "@/db/client";
import {
  clientPackages,
  packageServices,
  servicePackages,
  services,
} from "@/db/schema";

export type PaqueteInput = {
  nombre: string;
  descripcion?: string | null;
  precioPesos: number;
  sesionesTotal: number;
  activo?: boolean;
  serviceIds: number[];
};

export type PaqueteConServicios = {
  id: number;
  nombre: string;
  descripcion: string | null;
  precioPesos: number;
  sesionesTotal: number;
  activo: boolean;
  servicios: { id: number; nombre: string }[];
};

export type AsignacionPaquete = {
  id: number;
  packageId: number;
  paqueteNombre: string;
  nombreCliente: string;
  telefono: string;
  sesionesIniciales: number;
  sesionesRestantes: number;
  precioCobradoPesos: number | null;
  notas: string | null;
  estado: string;
  fechaCompra: string;
};

function normalizePaqueteInput(input: PaqueteInput): PaqueteInput {
  return {
    nombre: input.nombre.trim(),
    descripcion: input.descripcion?.trim() || null,
    precioPesos: Math.max(0, Math.round(input.precioPesos)),
    sesionesTotal: Math.max(1, Math.round(input.sesionesTotal)),
    activo: input.activo !== false,
    serviceIds: [...new Set(input.serviceIds.filter((id) => id > 0))],
  };
}

async function attachServicios(
  packageIds: number[]
): Promise<Map<number, { id: number; nombre: string }[]>> {
  const db = getDb();
  const map = new Map<number, { id: number; nombre: string }[]>();
  if (!db || packageIds.length === 0) return map;

  const links = await db
    .select({
      packageId: packageServices.packageId,
      serviceId: services.id,
      nombre: services.nombre,
    })
    .from(packageServices)
    .innerJoin(services, eq(packageServices.serviceId, services.id))
    .where(inArray(packageServices.packageId, packageIds));

  for (const row of links) {
    const list = map.get(row.packageId) ?? [];
    list.push({ id: row.serviceId, nombre: row.nombre });
    map.set(row.packageId, list);
  }
  return map;
}

export type PaquetePublico = {
  id: number;
  nombre: string;
  descripcion: string | null;
  precioPesos: number;
  sesionesTotal: number;
  serviciosIncluidos: string[];
};

/** Combos activos para reserva web (sin datos de clientes). */
export async function listarPaquetesPublicos(): Promise<
  PaquetePublico[] | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(servicePackages)
    .where(eq(servicePackages.activo, true))
    .orderBy(asc(servicePackages.nombre));

  const serviciosMap = await attachServicios(rows.map((r) => r.id));

  return rows.map((r) => ({
    id: r.id,
    nombre: r.nombre,
    descripcion: r.descripcion,
    precioPesos: r.precioPesos,
    sesionesTotal: r.sesionesTotal,
    serviciosIncluidos: (serviciosMap.get(r.id) ?? []).map((s) => s.nombre),
  }));
}

export async function listarPaquetesAdmin(): Promise<
  PaqueteConServicios[] | { ok: false; reason: "no_db" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(servicePackages)
    .orderBy(asc(servicePackages.nombre));

  const serviciosMap = await attachServicios(rows.map((r) => r.id));

  return rows.map((r) => ({
    id: r.id,
    nombre: r.nombre,
    descripcion: r.descripcion,
    precioPesos: r.precioPesos,
    sesionesTotal: r.sesionesTotal,
    activo: r.activo,
    servicios: serviciosMap.get(r.id) ?? [],
  }));
}

export async function crearPaquete(
  input: PaqueteInput
): Promise<
  | { ok: true; id: number }
  | { ok: false; reason: "no_db" | "invalido" | "duplicado" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const data = normalizePaqueteInput(input);
  if (!data.nombre) return { ok: false, reason: "invalido" };
  if (data.serviceIds.length === 0) return { ok: false, reason: "invalido" };

  const svcs = await db
    .select({ id: services.id, esGrupo: services.esGrupo })
    .from(services)
    .where(inArray(services.id, data.serviceIds));
  if (
    svcs.length !== data.serviceIds.length ||
    svcs.some((s) => s.esGrupo)
  ) {
    return { ok: false, reason: "invalido" };
  }

  const todos = await db.select({ id: servicePackages.id, nombre: servicePackages.nombre }).from(servicePackages);
  if (todos.some((p) => p.nombre.trim().toLowerCase() === data.nombre.toLowerCase())) {
    return { ok: false, reason: "duplicado" };
  }

  const [pkg] = await db
    .insert(servicePackages)
    .values({
      nombre: data.nombre,
      descripcion: data.descripcion,
      precioPesos: data.precioPesos,
      sesionesTotal: data.sesionesTotal,
      activo: data.activo,
    })
    .returning({ id: servicePackages.id });

  await db.insert(packageServices).values(
    data.serviceIds.map((serviceId) => ({
      packageId: pkg.id,
      serviceId,
    }))
  );

  return { ok: true, id: pkg.id };
}

export async function actualizarPaquete(
  id: number,
  input: PaqueteInput
): Promise<
  | { ok: true }
  | { ok: false; reason: "no_db" | "not_found" | "invalido" | "duplicado" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const data = normalizePaqueteInput(input);
  if (!data.nombre) return { ok: false, reason: "invalido" };
  if (data.serviceIds.length === 0) return { ok: false, reason: "invalido" };

  const svcs = await db
    .select({ id: services.id, esGrupo: services.esGrupo })
    .from(services)
    .where(inArray(services.id, data.serviceIds));
  if (
    svcs.length !== data.serviceIds.length ||
    svcs.some((s) => s.esGrupo)
  ) {
    return { ok: false, reason: "invalido" };
  }

  const todos = await db.select({ id: servicePackages.id, nombre: servicePackages.nombre }).from(servicePackages);
  if (
    todos.some(
      (p) => p.id !== id && p.nombre.trim().toLowerCase() === data.nombre.toLowerCase()
    )
  ) {
    return { ok: false, reason: "duplicado" };
  }

  const updated = await db
    .update(servicePackages)
    .set({
      nombre: data.nombre,
      descripcion: data.descripcion,
      precioPesos: data.precioPesos,
      sesionesTotal: data.sesionesTotal,
      activo: data.activo,
    })
    .where(eq(servicePackages.id, id))
    .returning({ id: servicePackages.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };

  await db.delete(packageServices).where(eq(packageServices.packageId, id));
  await db.insert(packageServices).values(
    data.serviceIds.map((serviceId) => ({ packageId: id, serviceId }))
  );

  return { ok: true };
}

export async function actualizarPrecioPaquete(
  id: number,
  precioPesos: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (!Number.isFinite(id) || id < 1) return { ok: false, reason: "not_found" };

  const updated = await db
    .update(servicePackages)
    .set({ precioPesos: Math.max(0, Math.round(precioPesos)) })
    .where(eq(servicePackages.id, id))
    .returning({ id: servicePackages.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

export async function eliminarPaquete(
  id: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" | "en_uso" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const activos = await db
    .select({ id: clientPackages.id })
    .from(clientPackages)
    .where(
      and(eq(clientPackages.packageId, id), eq(clientPackages.estado, "activo"))
    )
    .limit(1);
  if (activos.length > 0) return { ok: false, reason: "en_uso" };

  const deleted = await db
    .delete(servicePackages)
    .where(eq(servicePackages.id, id))
    .returning({ id: servicePackages.id });

  if (deleted.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}

export async function listarAsignacionesPaquete(
  soloActivas = true
): Promise<AsignacionPaquete[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const q = db
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
    .innerJoin(servicePackages, eq(clientPackages.packageId, servicePackages.id));

  const rows = await (soloActivas
    ? q.where(eq(clientPackages.estado, "activo"))
    : q
  ).orderBy(desc(clientPackages.fechaCompra), desc(clientPackages.id));

  return rows;
}

export async function asignarPaqueteCliente(params: {
  packageId: number;
  nombreCliente: string;
  telefono: string;
  fechaCompra: string;
  precioCobradoPesos?: number | null;
  notas?: string | null;
  sesiones?: number;
}): Promise<
  | { ok: true; id: number }
  | { ok: false; reason: "no_db" | "not_found" | "invalido" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const nombre = params.nombreCliente.trim();
  const telefono = params.telefono.trim();
  if (!nombre || !telefono || !params.fechaCompra) {
    return { ok: false, reason: "invalido" };
  }

  const [pkg] = await db
    .select()
    .from(servicePackages)
    .where(eq(servicePackages.id, params.packageId))
    .limit(1);
  if (!pkg || !pkg.activo) return { ok: false, reason: "not_found" };

  const sesiones = Math.max(
    1,
    Math.round(params.sesiones ?? pkg.sesionesTotal)
  );

  const [row] = await db
    .insert(clientPackages)
    .values({
      packageId: pkg.id,
      nombreCliente: nombre,
      telefono,
      sesionesIniciales: sesiones,
      sesionesRestantes: sesiones,
      precioCobradoPesos:
        params.precioCobradoPesos != null
          ? Math.max(0, Math.round(params.precioCobradoPesos))
          : pkg.precioPesos,
      notas: params.notas?.trim() || null,
      estado: "activo",
      fechaCompra: params.fechaCompra,
    })
    .returning({ id: clientPackages.id });

  return { ok: true, id: row.id };
}

export async function consumirSesionPaquete(
  asignacionId: number
): Promise<
  | { ok: true; sesionesRestantes: number }
  | { ok: false; reason: "no_db" | "not_found" | "sin_sesiones" }
> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [row] = await db
    .select()
    .from(clientPackages)
    .where(eq(clientPackages.id, asignacionId))
    .limit(1);

  if (!row || row.estado !== "activo") {
    return { ok: false, reason: "not_found" };
  }
  if (row.sesionesRestantes <= 0) {
    return { ok: false, reason: "sin_sesiones" };
  }

  const restantes = row.sesionesRestantes - 1;
  await db
    .update(clientPackages)
    .set({
      sesionesRestantes: restantes,
      estado: restantes === 0 ? "agotado" : "activo",
    })
    .where(eq(clientPackages.id, asignacionId));

  return { ok: true, sesionesRestantes: restantes };
}

export async function cancelarAsignacionPaquete(
  asignacionId: number
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const updated = await db
    .update(clientPackages)
    .set({ estado: "cancelado" })
    .where(
      and(
        eq(clientPackages.id, asignacionId),
        eq(clientPackages.estado, "activo")
      )
    )
    .returning({ id: clientPackages.id });

  if (updated.length === 0) return { ok: false, reason: "not_found" };
  return { ok: true };
}
