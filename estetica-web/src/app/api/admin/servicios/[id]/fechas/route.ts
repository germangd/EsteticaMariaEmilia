import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  agregarFechaServicio,
  listarFechasServicio,
  quitarFechaServicio,
  reemplazarFechasServicio,
} from "@/lib/disponibilidad-repo";

export const dynamic = "force-dynamic";

function parseId(raw: string): number | null {
  const id = Number(raw);
  return Number.isFinite(id) && id > 0 ? id : null;
}

export async function GET(
  request: NextRequest,
  context: { params: Promise<{ id: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const id = parseId((await context.params).id);
  if (!id) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  const fechas = await listarFechasServicio(id);
  if (!Array.isArray(fechas)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({
    ok: true,
    fechas,
    usaCalendario: fechas.length > 0,
  });
}

export async function PUT(
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

  const fechas = Array.isArray((body as { fechas?: unknown }).fechas)
    ? ((body as { fechas: unknown[] }).fechas as unknown[])
        .map(String)
        .filter(Boolean)
    : [];

  const res = await reemplazarFechasServicio(id, fechas);
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "No se pudieron guardar las fechas." },
      { status: res.reason === "not_found" ? 404 : 503 }
    );
  }

  return NextResponse.json({ ok: true, fechas });
}

export async function POST(
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

  const fecha =
    typeof (body as { fecha?: string }).fecha === "string"
      ? (body as { fecha: string }).fecha.trim()
      : "";

  const res = await agregarFechaServicio(id, fecha);
  if (!res.ok) {
    return NextResponse.json(
      { ok: false, mensaje: "Fecha inválida o no guardada." },
      { status: 400 }
    );
  }

  const fechas = await listarFechasServicio(id);
  return NextResponse.json({
    ok: true,
    fechas: Array.isArray(fechas) ? fechas : [],
  });
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

  const { searchParams } = new URL(request.url);
  const fecha = searchParams.get("fecha")?.trim() ?? "";
  if (!fecha) {
    return NextResponse.json(
      { ok: false, mensaje: "Usá ?fecha=YYYY-MM-DD" },
      { status: 400 }
    );
  }

  await quitarFechaServicio(id, fecha);
  const fechas = await listarFechasServicio(id);
  return NextResponse.json({
    ok: true,
    fechas: Array.isArray(fechas) ? fechas : [],
  });
}
