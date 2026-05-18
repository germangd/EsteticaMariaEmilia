import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  actualizarEvento,
  eliminarEvento,
} from "@/lib/eventos-repo";

export const dynamic = "force-dynamic";

function parseId(raw: string): number | null {
  const id = Number(raw);
  return Number.isFinite(id) && id > 0 ? id : null;
}

function parseEventoBody(body: unknown) {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;
  const nombre = typeof b.nombre === "string" ? b.nombre : "";
  const descripcion =
    typeof b.descripcion === "string" ? b.descripcion : null;
  const fecha = typeof b.fecha === "string" ? b.fecha : "";
  const horarioInicio =
    typeof b.horarioInicio === "string" ? b.horarioInicio : "09:00";
  const horarioFin = typeof b.horarioFin === "string" ? b.horarioFin : "18:00";
  const precioPesos =
    typeof b.precioPesos === "number"
      ? b.precioPesos
      : Number(b.precioPesos) || 0;
  const clienteTelefono =
    typeof b.clienteTelefono === "string" ? b.clienteTelefono : null;
  const clienteNombre =
    typeof b.clienteNombre === "string" ? b.clienteNombre : null;
  const activo = b.activo !== false;
  const serviceIds = Array.isArray(b.serviceIds)
    ? b.serviceIds.map(Number).filter((n) => Number.isFinite(n) && n > 0)
    : [];
  if (!nombre.trim() || !fecha.trim() || serviceIds.length === 0) return null;
  return {
    nombre,
    descripcion,
    fecha,
    horarioInicio,
    horarioFin,
    precioPesos,
    clienteTelefono,
    clienteNombre,
    activo,
    serviceIds,
  };
}

function mensajeError(reason: string): string {
  switch (reason) {
    case "franja_solapada":
      return "La franja horaria se superpone con otro evento activo ese día.";
    case "invalido":
      return "Revisá nombre, fecha, horario (desde < hasta) y servicios.";
    case "not_found":
      return "Evento no encontrado.";
    default:
      return "No se pudo actualizar.";
  }
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

  const input = parseEventoBody(body);
  if (!input) {
    return NextResponse.json(
      { ok: false, mensaje: "Datos incompletos." },
      { status: 400 }
    );
  }

  const res = await actualizarEvento(id, input);
  if (!res.ok) {
    const status =
      res.reason === "not_found"
        ? 404
        : res.reason === "franja_solapada"
          ? 409
          : res.reason === "invalido"
            ? 400
            : 503;
    return NextResponse.json(
      { ok: false, mensaje: mensajeError(res.reason) },
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

  const id = parseId((await context.params).id);
  if (!id) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  const res = await eliminarEvento(id);
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se pudo eliminar." },
      { status: res.reason === "not_found" ? 404 : 503 }
    );
  }

  return NextResponse.json({ ok: true });
}
