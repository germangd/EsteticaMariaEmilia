import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { listarSesionesCaja } from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sp = request.nextUrl.searchParams;
  const fechaDesde = sp.get("desde")?.trim() || undefined;
  const fechaHasta = sp.get("hasta")?.trim() || undefined;
  const limite = Number(sp.get("limite") ?? 40);

  const sesiones = await listarSesionesCaja({
    fechaDesde,
    fechaHasta,
    limite: Number.isFinite(limite) ? limite : 40,
  });

  if (!Array.isArray(sesiones)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, sesiones });
}
