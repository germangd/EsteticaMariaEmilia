import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { cancelarAsignacionPaquete } from "@/lib/paquetes-repo";

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

  const res = await cancelarAsignacionPaquete(id);
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se encontró una asignación activa." },
      { status: 404 }
    );
  }

  return NextResponse.json({ ok: true });
}
