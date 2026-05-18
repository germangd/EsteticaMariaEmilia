import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { actualizarServicio, eliminarServicio } from "@/lib/servicios-repo";
import { mensajeErrorServicio, parseServicioBody } from "@/lib/parse-servicio-body";

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

  const { id: idRaw } = await context.params;
  const id = parseId(idRaw);
  if (!id) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const input = parseServicioBody(body);
  if (!input) {
    return NextResponse.json(
      { ok: false, mensaje: "Datos incompletos o inválidos." },
      { status: 400 }
    );
  }

  const res = await actualizarServicio(id, input);
  if (!res.ok) {
    const status =
      res.reason === "not_found"
        ? 404
        : res.reason === "duplicado"
          ? 409
          : res.reason === "invalido" ||
              res.reason === "parent_invalido" ||
              res.reason === "tiene_hijos"
            ? 400
            : 503;
    return NextResponse.json(
      { ok: false, mensaje: mensajeErrorServicio(res.reason) },
      { status }
    );
  }

  return NextResponse.json({ ok: true });
}

export async function DELETE(
  request: NextRequest,
  context: { params: Promise<{ id: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const { id: idRaw } = await context.params;
  const id = parseId(idRaw);
  if (!id) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  const res = await eliminarServicio(id);
  if (!res.ok) {
    if (res.reason === "not_found") {
      return NextResponse.json({ ok: false, mensaje: "Servicio no encontrado." }, { status: 404 });
    }
    if (res.reason === "tiene_hijos") {
      return NextResponse.json(
        {
          ok: false,
          mensaje:
            "Esta categoría tiene sub-servicios. Eliminalos o reasignalos antes.",
        },
        { status: 409 }
      );
    }
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo eliminar." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true });
}
