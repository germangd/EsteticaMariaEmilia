import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { obtenerCatalogoCaja } from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const catalogo = await obtenerCatalogoCaja();
  if ("ok" in catalogo && catalogo.ok === false) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, catalogo });
}
