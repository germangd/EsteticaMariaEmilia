import { NextRequest, NextResponse } from "next/server";
import { DateTime } from "luxon";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { getAppTimeZone } from "@/lib/agenda";
import { exportarHistorialVentasCsv } from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sp = request.nextUrl.searchParams;
  const fechaDesde = sp.get("desde") ?? undefined;
  const fechaHasta = sp.get("hasta") ?? undefined;
  const cliente = sp.get("cliente") ?? undefined;

  const res = await exportarHistorialVentasCsv({
    fechaDesde,
    fechaHasta,
    cliente,
  });

  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  const tz = getAppTimeZone();
  const stamp = DateTime.now().setZone(tz).toFormat("yyyy-MM-dd_HHmm");
  const rango =
    fechaDesde && fechaHasta
      ? `${fechaDesde}_${fechaHasta}`
      : fechaDesde || fechaHasta || "filtro";

  return new NextResponse(res.csv, {
    status: 200,
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="historial_caja_${rango}_${stamp}.csv"`,
      "X-Export-Total": String(res.total),
      "X-Export-Rows": String(res.exportados),
    },
  });
}
