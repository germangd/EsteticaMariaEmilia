import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { buscarVentasHistorial } from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sp = request.nextUrl.searchParams;
  const fechaDesde = sp.get("desde") ?? undefined;
  const fechaHasta = sp.get("hasta") ?? undefined;
  const cliente = sp.get("cliente") ?? undefined;
  const limite = sp.get("limite") != null ? Number(sp.get("limite")) : undefined;
  const offset = sp.get("offset") != null ? Number(sp.get("offset")) : undefined;

  const res = await buscarVentasHistorial({
    fechaDesde,
    fechaHasta,
    cliente,
    limite: Number.isFinite(limite) ? limite : undefined,
    offset: Number.isFinite(offset) ? offset : undefined,
  });

  if (!("ventas" in res)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({
    ok: true,
    ventas: res.ventas,
    total: res.total,
  });
}
