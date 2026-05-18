import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  abrirSesionCaja,
  cerrarSesionCaja,
  obtenerSesionAbierta,
} from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sesion = await obtenerSesionAbierta();
  if (sesion && typeof sesion === "object" && "ok" in sesion) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, sesion });
}

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inv\u00e1lido." }, {
      status: 400,
    });
  }

  const b = (body && typeof body === "object" ? body : {}) as Record<
    string,
    unknown
  >;
  const openingAmountPesos =
    b.openingAmountPesos != null ? Number(b.openingAmountPesos) : 0;
  const notes = typeof b.notes === "string" ? b.notes : null;

  const res = await abrirSesionCaja({ openingAmountPesos, notes });
  if (!res.ok) {
    const mensaje =
      res.reason === "ya_abierta"
        ? "Ya hay una caja abierta."
        : res.reason === "no_db"
          ? "Base de datos no disponible."
          : "No se pudo abrir la caja.";
    return NextResponse.json({ ok: false, mensaje }, { status: 400 });
  }

  return NextResponse.json({ ok: true, sesion: res.sesion });
}

export async function PATCH(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inv\u00e1lido." }, {
      status: 400,
    });
  }

  const b = (body && typeof body === "object" ? body : {}) as Record<
    string,
    unknown
  >;
  const sessionId = Number(b.sessionId);
  const closingAmountPesos = Number(b.closingAmountPesos);
  const notes = typeof b.notes === "string" ? b.notes : null;

  const res = await cerrarSesionCaja({
    sessionId,
    closingAmountPesos,
    notes,
  });

  if (!res.ok) {
    const mensaje =
      res.reason === "not_found"
        ? "Sesi\u00f3n no encontrada."
        : res.reason === "cerrada"
          ? "La caja ya est\u00e1 cerrada."
          : res.reason === "no_db"
            ? "Base de datos no disponible."
            : "Complet\u00e1 el monto de cierre.";
    return NextResponse.json(
      { ok: false, mensaje },
      { status: res.reason === "not_found" ? 404 : 400 }
    );
  }

  return NextResponse.json({ ok: true, sesion: res.sesion });
}
