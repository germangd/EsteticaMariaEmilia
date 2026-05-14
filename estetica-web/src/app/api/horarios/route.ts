import { eq } from "drizzle-orm";
import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services } from "@/db/schema";
import {
  generarHorarios,
  getAppTimeZone,
  hoyIsoEnZona,
  horaActualEnZona,
} from "@/lib/agenda";
import { contarActivosPorHora } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

/**
 * Horarios disponibles (misma idea que `obtenerHorariosDisponibles` en `Código.gs`).
 * Query: `?servicio=Nombre exacto&fecha=YYYY-MM-DD`
 */
export async function GET(request: Request) {
  const db = getDb();
  if (!db) {
    return NextResponse.json(
      { ok: false, error: "database_not_configured", horarios: [] },
      { status: 503 }
    );
  }

  const { searchParams } = new URL(request.url);
  const servicioNombre = searchParams.get("servicio")?.trim() ?? "";
  const fecha = searchParams.get("fecha")?.trim() ?? "";

  if (!servicioNombre || !fecha || !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
    return NextResponse.json(
      {
        ok: false,
        error: "invalid_params",
        mensaje: "Usá ?servicio=Nombre&fecha=YYYY-MM-DD",
        horarios: [],
      },
      { status: 400 }
    );
  }

  try {
    const [servicio] = await db
      .select()
      .from(services)
      .where(eq(services.nombre, servicioNombre))
      .limit(1);

    if (!servicio) {
      return NextResponse.json({ ok: true, horarios: [] });
    }

    let horariosPosibles = generarHorarios(
      servicio.horarioInicio,
      servicio.horarioFin
    );
    if (horariosPosibles.length === 0) {
      return NextResponse.json({ ok: true, horarios: [] });
    }

    const tz = getAppTimeZone();
    const hoyStr = hoyIsoEnZona(tz);
    if (fecha === hoyStr) {
      const horaActual = horaActualEnZona(tz);
      horariosPosibles = horariosPosibles.filter((h) => h >= horaActual);
    }

    const ocupados = await contarActivosPorHora(
      fecha,
      servicioNombre,
      servicio.responsable
    );

    const capacidad = servicio.capacidad;
    const disponibles = horariosPosibles.filter(
      (hora) => (ocupados.get(hora) ?? 0) < capacidad
    );

    return NextResponse.json({ ok: true, horarios: disponibles });
  } catch {
    return NextResponse.json(
      { ok: false, error: "query_failed", horarios: [] },
      { status: 500 }
    );
  }
}
