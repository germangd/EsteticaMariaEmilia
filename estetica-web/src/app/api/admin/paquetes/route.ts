import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { crearPaquete, listarPaquetesAdmin } from "@/lib/paquetes-repo";

export const dynamic = "force-dynamic";

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

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const rows = await listarPaquetesAdmin();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, paquetes: rows });
}

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const input = parsePaqueteBody(body);
  if (!input || input.serviceIds.length === 0) {
    return NextResponse.json(
      { ok: false, mensaje: "Nombre, precio, sesiones y al menos un servicio." },
      { status: 400 }
    );
  }

  const res = await crearPaquete(input);
  if (!res.ok) {
    const msg =
      res.reason === "duplicado"
        ? "Ya existe un paquete con ese nombre."
        : res.reason === "invalido"
          ? "Datos inválidos."
          : "No se pudo crear.";
    return NextResponse.json(
      { ok: false, mensaje: msg },
      { status: res.reason === "duplicado" ? 409 : 503 }
    );
  }
  return NextResponse.json({ ok: true, id: res.id });
}
