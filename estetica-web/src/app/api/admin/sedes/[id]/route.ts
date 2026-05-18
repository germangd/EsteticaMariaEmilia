import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { actualizarSede, eliminarSede } from "@/lib/sedes-repo";

export const dynamic = "force-dynamic";

function parseId(raw: string): number | null {
  const id = Number(raw);
  return Number.isFinite(id) && id > 0 ? id : null;
}

export async function PATCH(
  request: NextRequest,
  context: { params: Promise<{ id: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const id = parseId((await context.params).id);
  if (!id) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

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

  const res = await actualizarSede(id, { nombre, orden, activo });
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo actualizar." },
      { status: res.reason === "not_found" ? 404 : 400 }
    );
  }
  return NextResponse.json({ ok: true });
}

export async function DELETE(
  request: NextRequest,
  context: { params: Promise<{ id: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const id = parseId((await context.params).id);
  if (!id) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  const res = await eliminarSede(id);
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo desactivar." },
      { status: res.reason === "not_found" ? 404 : 503 }
    );
  }
  return NextResponse.json({ ok: true });
}
