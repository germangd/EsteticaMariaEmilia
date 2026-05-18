import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { crearEvento, listarEventosAdmin } from "@/lib/eventos-repo";

export const dynamic = "force-dynamic";

function parseEventoBody(body: unknown) {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;
  const nombre = typeof b.nombre === "string" ? b.nombre : "";
  const descripcion =
    typeof b.descripcion === "string" ? b.descripcion : null;
  const fecha = typeof b.fecha === "string" ? b.fecha : "";
  const sedeId = Number(b.sedeId);
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
  if (
    !nombre.trim() ||
    !fecha.trim() ||
    serviceIds.length === 0 ||
    !Number.isFinite(sedeId) ||
    sedeId < 1
  ) {
    return null;
  }
  return {
    nombre,
    descripcion,
    fecha,
    sedeId,
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
      return "Revisá nombre, fecha, horario (desde < hasta) y servicios (solo sub-servicios o sueltos).";
    default:
      return "No se pudo guardar el evento.";
  }
}

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const rows = await listarEventosAdmin();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, eventos: rows });
}

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const input = parseEventoBody(body);
  if (!input) {
    return NextResponse.json(
      { ok: false, mensaje: "Datos incompletos (nombre, fecha y servicios)." },
      { status: 400 }
    );
  }

  const res = await crearEvento(input);
  if (!res.ok) {
    const status =
      res.reason === "franja_solapada"
        ? 409
        : res.reason === "invalido"
          ? 400
          : 503;
    return NextResponse.json(
      { ok: false, mensaje: mensajeError(res.reason) },
      { status }
    );
  }

  return NextResponse.json({ ok: true, id: res.id });
}
