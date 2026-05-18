import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { actualizarPaquete, eliminarPaquete } from "@/lib/paquetes-repo";

export const dynamic = "force-dynamic";

function parseId(raw: string): number | null {
  const id = Number(raw);
  return Number.isFinite(id) && id > 0 ? id : null;
}

function parsePaqueteBody(body: unknown) {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;
  const nombre = typeof b.nombre === "string" ? b.nombre : "";
  const descripcion =
    typeof b.descripcion === "string" ? b.descripcion : null;
  const precioPesos = Number(b.precioPesos);
  const sesionesTotal = Number(b.sesionesTotal);
  const activo = b.activo !== false;
  const serviceIds = Array.isArray(b.serviceIds)
    ? b.serviceIds.map(Number).filter((n) => Number.isFinite(n) && n > 0)
    : [];
  if (!nombre.trim() || !Number.isFinite(precioPesos) || !Number.isFinite(sesionesTotal)) {
    return null;
  }
  return { nombre, descripcion, precioPesos, sesionesTotal, activo, serviceIds };
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

  const input = parsePaqueteBody(body);
  if (!input || input.serviceIds.length === 0) {
    return NextResponse.json(
      { ok: false, mensaje: "Datos incompletos." },
      { status: 400 }
    );
  }

  const res = await actualizarPaquete(id, input);
  if (!res.ok) {
    const status =
      res.reason === "not_found" ? 404 : res.reason === "duplicado" ? 409 : 503;
    return NextResponse.json({ ok: false, mensaje: "No se pudo actualizar." }, { status });
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

  const res = await eliminarPaquete(id);
  if (!res.ok) {
    if (res.reason === "en_uso") {
      return NextResponse.json(
        {
          ok: false,
          mensaje: "Hay clientes con este paquete activo. Cancelá esas asignaciones primero.",
        },
        { status: 409 }
      );
    }
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo eliminar." },
      { status: res.reason === "not_found" ? 404 : 503 }
    );
  }
  return NextResponse.json({ ok: true });
}
