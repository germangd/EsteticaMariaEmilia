import { NextRequest, NextResponse } from "next/server";
import { DateTime } from "luxon";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { getAppTimeZone } from "@/lib/agenda";
import {
  generarCsvCierreSesion,
  obtenerDatosCierreSesion,
} from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sessionId = Number(request.nextUrl.searchParams.get("sessionId"));
  if (!Number.isFinite(sessionId) || sessionId < 1) {
    return NextResponse.json(
      { ok: false, mensaje: "sessionId requerido." },
      { status: 400 }
    );
  }

  const data = await obtenerDatosCierreSesion(sessionId);
  if (data && typeof data === "object" && "ok" in data) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  if (!data) {
    return NextResponse.json(
      { ok: false, mensaje: "Sesi\u00f3n no encontrada." },
      { status: 404 }
    );
  }

  const tz = getAppTimeZone();
  const stamp = DateTime.now().setZone(tz).toFormat("yyyy-MM-dd_HHmm");
  const body = generarCsvCierreSesion(data);

  return new NextResponse(body, {
    status: 200,
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="cierre_caja_${sessionId}_${stamp}.csv"`,
    },
  });
}
