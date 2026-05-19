import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import {
  calcularHorariosDisponibles,
  parsearClaveReserva,
  resolverItemReserva,
} from "@/lib/reserva-catalogo";

export const dynamic = "force-dynamic";

/**
 * Horarios disponibles.
 * Query: `?clave=s:Nombre` o `?clave=p:ID` (o legacy `?servicio=Nombre`) + `fecha` + `sedeId`
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
  const clave =
    searchParams.get("clave")?.trim() ??
    (searchParams.get("servicio")?.trim()
      ? `s:${searchParams.get("servicio")!.trim()}`
      : "");
  const fecha = searchParams.get("fecha")?.trim() ?? "";
  const sedeId = Number(searchParams.get("sedeId"));

  if (
    !clave ||
    !fecha ||
    !/^\d{4}-\d{2}-\d{2}$/.test(fecha) ||
    !Number.isFinite(sedeId) ||
    sedeId < 1
  ) {
    return NextResponse.json(
      {
        ok: false,
        error: "invalid_params",
        mensaje: "Usá ?clave=s:Nombre o p:ID&fecha=YYYY-MM-DD&sedeId=1",
        horarios: [],
      },
      { status: 400 }
    );
  }

  const parsed = parsearClaveReserva(clave);
  if (!parsed) {
    return NextResponse.json(
      {
        ok: false,
        error: "invalid_params",
        mensaje: "Clave de servicio o combo inválida.",
        horarios: [],
      },
      { status: 400 }
    );
  }

  try {
    const item = await resolverItemReserva(parsed);
    if (!item) {
      return NextResponse.json({ ok: true, horarios: [] });
    }

    const { horarios, mensaje } = await calcularHorariosDisponibles(
      item,
      fecha,
      sedeId
    );

    return NextResponse.json({ ok: true, horarios, mensaje });
  } catch {
    return NextResponse.json(
      { ok: false, error: "query_failed", horarios: [] },
      { status: 500 }
    );
  }
}
