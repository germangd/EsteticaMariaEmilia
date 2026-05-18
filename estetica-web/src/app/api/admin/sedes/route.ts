import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { crearSede, listarSedesAdmin } from "@/lib/sedes-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const rows = await listarSedesAdmin();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, sedes: rows });
}

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const b = body as Record<string, unknown>;
  const nombre = typeof b.nombre === "string" ? b.nombre : "";
  const orden = Number(b.orden) || 0;
  const activo = b.activo !== false;

  const res = await crearSede({ nombre, orden, activo });
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo crear la sede." },
      { status: res.reason === "invalido" ? 400 : 503 }
    );
  }
  return NextResponse.json({ ok: true, id: res.id });
}
