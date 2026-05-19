import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  crearServicio,
  listarServiciosAdmin,
} from "@/lib/servicios-repo";
import { rowToServicioApi } from "@/lib/servicio-format";
import { mensajeErrorServicio, parseServicioBody } from "@/lib/parse-servicio-body";
import { nombreCategoria } from "@/lib/servicio-tree";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const rows = await listarServiciosAdmin();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  const byId = new Map(
    rows.map((x) => [
      x.id,
      { id: x.id, nombre: x.nombre, parentId: x.parentId, esGrupo: x.esGrupo },
    ])
  );
  return NextResponse.json({
    ok: true,
    servicios: rows.map((r) => ({
      id: r.id,
      ...rowToServicioApi(r),
      categoriaNombre: nombreCategoria(
        { id: r.id, nombre: r.nombre, parentId: r.parentId, esGrupo: r.esGrupo },
        byId
      ),
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

  const input = parseServicioBody(body);
  if (!input) {
    return NextResponse.json(
      { ok: false, mensaje: "Datos incompletos o inválidos." },
      { status: 400 }
    );
  }

  const res = await crearServicio(input);
  if (!res.ok) {
    const status =
      res.reason === "duplicado"
        ? 409
        : res.reason === "invalido" || res.reason === "parent_invalido"
          ? 400
          : 503;
    return NextResponse.json(
      {
        ok: false,
        mensaje: mensajeErrorServicio(res.reason, res.conflicto),
        conflicto: res.conflicto ?? null,
      },
      { status }
    );
  }

  return NextResponse.json({ ok: true, id: res.id });
}
