import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { actualizarPrecioPaquete } from "@/lib/paquetes-repo";

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
  const precioPesos =
    b.precioPesos != null
      ? Number(b.precioPesos)
      : b.precio != null
        ? Number(b.precio)
        : NaN;

  if (!Number.isFinite(precioPesos) || precioPesos < 0) {
    return NextResponse.json(
      { ok: false, mensaje: "Precio inválido." },
      { status: 400 }
    );
  }

  const res = await actualizarPrecioPaquete(id, precioPesos);
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo actualizar el precio." },
      { status: res.reason === "not_found" ? 404 : 503 }
    );
  }

  return NextResponse.json({ ok: true, precioPesos: Math.round(precioPesos) });
}
