import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { obtenerResumenVentasPeriodo } from "@/lib/caja-repo";
import { etiquetaBucket } from "@/lib/reportes-fechas";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sp = request.nextUrl.searchParams;
  const fechaDesde = sp.get("desde")?.trim() ?? "";
  const fechaHasta = sp.get("hasta")?.trim() ?? "";
  const agrupacion = sp.get("agrupacion")?.trim() ?? "dia";

  if (!fechaDesde || !fechaHasta) {
    return NextResponse.json(
      { ok: false, mensaje: "Indicá fecha desde y hasta." },
      { status: 400 }
    );
  }

  const agr =
    agrupacion === "semana" || agrupacion === "mes" ? agrupacion : "dia";

  const res = await obtenerResumenVentasPeriodo({
    fechaDesde,
    fechaHasta,
    agrupacion: agr,
    etiquetaBucket,
  });

  if (!res || (typeof res === "object" && "ok" in res)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, resumen: res });
}
