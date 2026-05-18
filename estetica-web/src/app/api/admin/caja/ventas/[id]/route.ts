import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { anularVenta, obtenerVentaDetalle } from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

type Ctx = { params: Promise<{ id: string }> };

export async function GET(request: NextRequest, ctx: Ctx) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const { id: idStr } = await ctx.params;
  const id = Number(idStr);
  if (!Number.isFinite(id) || id < 1) {
    return NextResponse.json({ ok: false, mensaje: "ID inv\u00e1lido." }, {
      status: 400,
    });
  }

  const venta = await obtenerVentaDetalle(id);
  if (venta && typeof venta === "object" && "ok" in venta) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  if (!venta) {
    return NextResponse.json({ ok: false, mensaje: "Venta no encontrada." }, {
      status: 404,
    });
  }

  return NextResponse.json({ ok: true, venta });
}

export async function DELETE(request: NextRequest, ctx: Ctx) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const { id: idStr } = await ctx.params;
  const id = Number(idStr);
  if (!Number.isFinite(id) || id < 1) {
    return NextResponse.json({ ok: false, mensaje: "ID inv\u00e1lido." }, {
      status: 400,
    });
  }

  const res = await anularVenta(id);
  if (!res.ok) {
    const mensaje =
      res.reason === "not_found"
        ? "Venta no encontrada."
        : res.reason === "ya_anulada"
          ? "La venta ya est\u00e1 anulada."
          : "Base de datos no disponible.";
    return NextResponse.json(
      { ok: false, mensaje },
      { status: res.reason === "not_found" ? 404 : 400 }
    );
  }

  return NextResponse.json({ ok: true });
}
