import { NextRequest, NextResponse } from "next/server";
import { DateTime } from "luxon";
import { ADMIN_SESSION_COOKIE, verifySignedSessionValue } from "@/lib/admin-session";
import { getAppTimeZone, hoyIsoEnZona } from "@/lib/agenda";
import { listarTurnosActivosFiltrados } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

function csvEscape(s: string): string {
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

export async function GET(request: NextRequest) {
  if (!verifySignedSessionValue(request.cookies.get(ADMIN_SESSION_COOKIE)?.value)) {
    return NextResponse.json({ ok: false, error: "no_autorizado" }, { status: 401 });
  }

  const tz = getAppTimeZone();
  const hoy = hoyIsoEnZona(tz);
  const { searchParams } = new URL(request.url);
  const servicio = searchParams.get("servicio")?.trim() || null;
  let desde = searchParams.get("desde")?.trim() || hoy;
  const hasta = searchParams.get("hasta")?.trim() || null;

  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde)) desde = hoy;
  if (hasta && !/^\d{4}-\d{2}-\d{2}$/.test(hasta)) {
    return NextResponse.json({ ok: false, error: "fecha" }, { status: 400 });
  }
  if (hasta && /^\d{4}-\d{2}-\d{2}$/.test(hasta) && hasta < desde) {
    return NextResponse.json(
      { ok: false, error: "rango", mensaje: "La fecha hasta debe ser ≥ desde." },
      { status: 400 }
    );
  }

  const rows = await listarTurnosActivosFiltrados({
    fechaDesde: desde,
    fechaHasta: hasta,
    servicioNombre: servicio,
  });

  const header = [
    "fecha",
    "hora",
    "servicio",
    "cliente",
    "telefono",
    "email",
    "responsable",
    "codigo_cancelacion",
    "estado",
  ];

  const lines = [
    header.join(","),
    ...rows.map((r) =>
      [
        r.fecha,
        r.hora,
        r.servicioNombre,
        r.nombreCliente,
        r.telefono,
        r.email ?? "",
        r.responsable,
        r.codigoCancelacion,
        r.estado,
      ]
        .map((c) => csvEscape(String(c)))
        .join(",")
    ),
  ];

  const stamp = DateTime.now().setZone(tz).toFormat("yyyy-MM-dd_HHmm");
  const body = `\uFEFF${lines.join("\r\n")}\r\n`;
  return new NextResponse(body, {
    status: 200,
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="turnos_${stamp}.csv"`,
    },
  });
}
