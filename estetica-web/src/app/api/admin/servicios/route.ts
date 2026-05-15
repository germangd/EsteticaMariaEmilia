import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  crearServicio,
  listarServiciosAdmin,
  type ServicioInput,
} from "@/lib/servicios-repo";
import { rowToServicioApi } from "@/lib/servicio-format";

export const dynamic = "force-dynamic";

function parseBody(body: unknown): ServicioInput | null {
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

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const rows = await listarServiciosAdmin();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({
    ok: true,
    servicios: rows.map((r) => ({
      id: r.id,
      ...rowToServicioApi(r),
    })),
  });
}

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

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

  const res = await crearServicio(input);
  if (!res.ok) {
    if (res.reason === "duplicado") {
      return NextResponse.json(
        { ok: false, mensaje: "Ya existe un servicio con ese nombre." },
        { status: 409 }
      );
    }
    if (res.reason === "invalido") {
      return NextResponse.json(
        { ok: false, mensaje: "Nombre de servicio obligatorio." },
        { status: 400 }
      );
    }
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, id: res.id });
}
