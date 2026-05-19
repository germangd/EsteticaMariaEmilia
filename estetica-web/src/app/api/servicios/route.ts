import { asc } from "drizzle-orm";
import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services, type ServiceRow } from "@/db/schema";
import {
  dedupeServiciosPorNombre,
  rowToServicioApi,
} from "@/lib/servicio-format";
import { nombreCategoria } from "@/lib/servicio-tree";

export const dynamic = "force-dynamic";

/** Lista servicios (misma forma que `obtenerServicios()` en `Código.gs`). */
export async function GET() {
  const db = getDb();
  if (!db) {
    return NextResponse.json(
      {
        ok: false,
        error: "database_not_configured",
        mensaje: "Falta DATABASE_URL (Neon).",
      },
      { status: 503 }
    );
  }

  try {
    const rows = await db
      .select()
      .from(services)
      .orderBy(asc(services.id));

    const reservables = rows.filter(
      (r): r is ServiceRow => Boolean(r.nombre?.trim()) && !r.esGrupo
    );
    const byId = new Map(
      rows
        .filter((r) => Boolean(r.nombre?.trim()))
        .map((x) => [
          x.id,
          {
            id: x.id,
            nombre: x.nombre,
            parentId: x.parentId,
            esGrupo: x.esGrupo,
          },
        ])
    );
    const list = dedupeServiciosPorNombre(
      reservables.map((r) => ({
        id: r.id,
        ...rowToServicioApi(r),
        categoriaNombre: nombreCategoria(
          { id: r.id, nombre: r.nombre, parentId: r.parentId, esGrupo: r.esGrupo },
          byId
        ),
      }))
    );

    return NextResponse.json({ ok: true, servicios: list });
  } catch {
    return NextResponse.json(
      { ok: false, error: "query_failed", servicios: [] },
      { status: 500 }
    );
  }
}
