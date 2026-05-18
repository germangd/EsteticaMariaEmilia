import { eq } from "drizzle-orm";
import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services } from "@/db/schema";
import {
  contarSolapamiento,
  generarHorarios,
  getAppTimeZone,
  hoyIsoEnZona,
  horaActualEnZona,
  horaAMinutos,
} from "@/lib/agenda";
import {
  evaluarReservaEnEvento,
  listarEventosActivosEnFecha,
  resolverVentanaReserva,
} from "@/lib/disponibilidad-repo";
import { listarActivosConDuracion } from "@/lib/turnos-repo";

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

    const ventana = await resolverVentanaReserva(servicio.id, servicio, fecha);
    if (!ventana || ventana.bloqueado) {
      return NextResponse.json({
        ok: true,
        horarios: [],
        mensaje: ventana?.bloqueado ?? "Fecha no disponible.",
      });
    }

    const eventos = await listarEventosActivosEnFecha(fecha);
    const duracionMin = Math.max(5, servicio.duracionMin);
    let horariosPosibles = generarHorarios(
      ventana.horarioInicio,
      ventana.horarioFin,
      duracionMin
    );

    if (eventos.length > 0) {
      horariosPosibles = horariosPosibles.filter((h) => {
        const ev = evaluarReservaEnEvento(servicio.id, h, eventos);
        if (!ev.permitido) return false;
        if (ev.ventana) {
          const finEv = horaAMinutos(ev.ventana.horarioFin);
          const inicio = horaAMinutos(h);
          return inicio >= 0 && inicio + duracionMin <= finEv;
        }
        return true;
      });
    }

    if (horariosPosibles.length === 0) {
      const mensaje =
        eventos.length > 0
          ? "No hay horarios disponibles para este servicio en esa fecha (revisá franjas de eventos)."
          : undefined;
      return NextResponse.json({ ok: true, horarios: [], mensaje });
    }

    const tz = getAppTimeZone();
    const hoyStr = hoyIsoEnZona(tz);
    if (fecha === hoyStr) {
      const horaActual = horaActualEnZona(tz);
      horariosPosibles = horariosPosibles.filter((h) => h >= horaActual);
    }

    const ocupados = await listarActivosConDuracion(
      fecha,
      servicioNombre,
      servicio.responsable
    );

    const capacidad = servicio.capacidad;
    const disponibles = horariosPosibles.filter((hora) => {
      const inicio = horaAMinutos(hora);
      if (inicio < 0) return false;
      return contarSolapamiento(inicio, duracionMin, ocupados) < capacidad;
    });

    return NextResponse.json({ ok: true, horarios: disponibles });
  } catch {
    return NextResponse.json(
      { ok: false, error: "query_failed", horarios: [] },
      { status: 500 }
    );
  }
}
