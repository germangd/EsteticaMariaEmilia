import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { listarClientesResumen } from "@/lib/clientes-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const q = request.nextUrl.searchParams.get("q") ?? undefined;
  const rows = await listarClientesResumen(q);

  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, clientes: rows });
}
