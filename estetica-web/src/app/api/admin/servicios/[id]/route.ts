import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { actualizarServicio, eliminarServicio } from "@/lib/servicios-repo";

export const dynamic = "force-dynamic";

function parseId(raw: string): number | null {
  const id = Number(raw);
  if (!Number.isFinite(id) || id < 1) return null;
  return id;
}

function parseBody(body: unknown) {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;
  const nombre = typeof b.nombre === "string" ? b.nombre : "";
  const duracionMin = Number(b.duracionMin ?? b.duracion);
  const responsable = typeof b.responsable === "string" ? b.responsable : "";
  const capacidad = Number(b.capacidad);
  const horarioInicio =
    typeof b.horarioInicio === "string" ? b.horarioInicio : "09:00";
  const horarioFin = typeof b.horarioFin === "string" ? b.horarioFin : "18:00";
  if (!nombre.trim() || !Number.isFinite(duracionMin) || !Number.isFinite(capacidad)) {
    return null;
  }
  return {
    nombre,
    duracionMin,
    responsable,
    capacidad,
    horarioInicio,
    horarioFin,
  };
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

  const input = parseBody(body);
  if (!input) {
    return NextResponse.json(
      { ok: false, mensaje: "Datos incompletos o inválidos." },
      { status: 400 }
    );
  }

  const res = await actualizarServicio(id, input);
  if (!res.ok) {
    if (res.reason === "not_found") {
      return NextResponse.json({ ok: false, mensaje: "Servicio no encontrado." }, { status: 404 });
    }
    if (res.reason === "duplicado") {
      return NextResponse.json(
        { ok: false, mensaje: "Ya existe otro servicio con ese nombre." },
        { status: 409 }
      );
    }
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo actualizar." },
      { status: 503 }
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
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo eliminar." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true });
}
