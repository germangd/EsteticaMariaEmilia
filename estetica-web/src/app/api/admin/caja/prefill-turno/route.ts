import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { obtenerPrefillCobroTurno } from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const appointmentId = Number(
    request.nextUrl.searchParams.get("appointmentId")
  );
  if (!Number.isFinite(appointmentId) || appointmentId < 1) {
    return NextResponse.json(
      { ok: false, mensaje: "Turno inválido." },
      { status: 400 }
    );
  }

  const prefill = await obtenerPrefillCobroTurno(appointmentId);
  if (prefill && typeof prefill === "object" && "ok" in prefill) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  if (!prefill) {
    return NextResponse.json(
      { ok: false, mensaje: "Turno no encontrado o no disponible para cobro." },
      { status: 404 }
    );
  }

  return NextResponse.json({ ok: true, prefill });
}
