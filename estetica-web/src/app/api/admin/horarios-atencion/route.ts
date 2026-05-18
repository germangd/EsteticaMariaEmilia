import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  listarHorarioAtencionAdmin,
  reemplazarHorarioAtencion,
  type FranjaAtencionInput,
} from "@/lib/horario-atencion-repo";

export const dynamic = "force-dynamic";

function parseSedeId(request: NextRequest, body?: unknown): number | null {
  const fromQuery = Number(request.nextUrl.searchParams.get("sedeId"));
  if (Number.isFinite(fromQuery) && fromQuery > 0) return fromQuery;
  if (body && typeof body === "object") {
    const b = body as Record<string, unknown>;
    const id = Number(b.sedeId);
    if (Number.isFinite(id) && id > 0) return id;
  }
  return null;
}

function parseBody(body: unknown, sedeId: number): FranjaAtencionInput[] | null {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;
  if (!Array.isArray(b.franjas)) return null;
  const out: FranjaAtencionInput[] = [];
  for (const item of b.franjas) {
    if (!item || typeof item !== "object") continue;
    const row = item as Record<string, unknown>;
    const diaSemana = Number(row.diaSemana);
    const horarioInicio =
      typeof row.horarioInicio === "string" ? row.horarioInicio : "";
    const horarioFin =
      typeof row.horarioFin === "string" ? row.horarioFin : "";
    if (!Number.isFinite(diaSemana) || !horarioInicio || !horarioFin) continue;
    out.push({ sedeId, diaSemana, horarioInicio, horarioFin });
  }
  return out;
}

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sedeId = parseSedeId(request);
  if (!sedeId) {
    return NextResponse.json(
      { ok: false, mensaje: "Falta sedeId." },
      { status: 400 }
    );
  }

  const rows = await listarHorarioAtencionAdmin(sedeId);
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, sedeId, dias: rows });
}

export async function PUT(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const sedeId = parseSedeId(request, body);
  if (!sedeId) {
    return NextResponse.json(
      { ok: false, mensaje: "Falta sedeId." },
      { status: 400 }
    );
  }

  const franjas = parseBody(body, sedeId);
  if (franjas === null) {
    return NextResponse.json(
      { ok: false, mensaje: "Formato inválido (franjas)." },
      { status: 400 }
    );
  }

  const res = await reemplazarHorarioAtencion(sedeId, franjas);
  if (!res.ok) {
    return NextResponse.json(
      {
        ok: false,
        mensaje:
          res.reason === "invalido"
            ? "Cada franja debe tener desde < hasta (ej. 09:00 y 13:00)."
            : "No se pudo guardar.",
      },
      { status: res.reason === "invalido" ? 400 : 503 }
    );
  }

  return NextResponse.json({ ok: true });
}
