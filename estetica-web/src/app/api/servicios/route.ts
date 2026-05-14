import { asc } from "drizzle-orm";
import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services } from "@/db/schema";
import { rowToServicioApi } from "@/lib/servicio-format";

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

    const list = rows
      .filter((r) => r.nombre?.trim())
      .map((r) => rowToServicioApi(r));

    return NextResponse.json({ ok: true, servicios: list });
  } catch {
    return NextResponse.json(
      { ok: false, error: "query_failed", servicios: [] },
      { status: 500 }
    );
  }
}
