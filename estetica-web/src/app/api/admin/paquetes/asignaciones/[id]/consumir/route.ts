import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { consumirSesionPaquete } from "@/lib/paquetes-repo";

export const dynamic = "force-dynamic";

export async function POST(
  request: NextRequest,
  context: { params: Promise<{ id: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const id = Number((await context.params).id);
  if (!Number.isFinite(id) || id < 1) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  const res = await consumirSesionPaquete(id);
  if (!res.ok) {
    const msg =
      res.reason === "sin_sesiones"
        ? "No quedan sesiones en este paquete."
        : "Asignación no encontrada o inactiva.";
    return NextResponse.json(
      { ok: false, mensaje: msg },
      { status: res.reason === "not_found" ? 404 : 400 }
    );
  }

  return NextResponse.json({
    ok: true,
    sesionesRestantes: res.sesionesRestantes,
  });
}
