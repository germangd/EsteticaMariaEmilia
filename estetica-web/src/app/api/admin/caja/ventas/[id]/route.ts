import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  anularVenta,
  modificarVenta,
  obtenerVentaDetalle,
  type LineaVentaInput,
} from "@/lib/caja-repo";

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

export async function PATCH(request: NextRequest, ctx: Ctx) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const { id: idStr } = await ctx.params;
  const id = Number(idStr);
  if (!Number.isFinite(id) || id < 1) {
    return NextResponse.json({ ok: false, mensaje: "ID inv\u00e1lido." }, {
      status: 400,
    });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inv\u00e1lido." }, {
      status: 400,
    });
  }

  if (!body || typeof body !== "object") {
    return NextResponse.json({ ok: false, mensaje: "Datos inv\u00e1lidos." }, {
      status: 400,
    });
  }

  const b = body as Record<string, unknown>;
  const lineasRaw = Array.isArray(b.lineas) ? b.lineas : [];
  const lineas: LineaVentaInput[] = lineasRaw
    .filter((x) => x && typeof x === "object")
    .map((x) => {
      const l = x as Record<string, unknown>;
      return {
        tipo:
          l.tipo === "servicio" || l.tipo === "paquete" ? l.tipo : "otro",
        descripcion: typeof l.descripcion === "string" ? l.descripcion : "",
        cantidad: Number(l.cantidad),
        precioUnitarioPesos: Number(l.precioUnitarioPesos),
        serviceId: l.serviceId != null ? Number(l.serviceId) : undefined,
        servicePackageId:
          l.servicePackageId != null
            ? Number(l.servicePackageId)
            : undefined,
      };
    });

  const res = await modificarVenta(id, {
    clienteTelefono:
      typeof b.clienteTelefono === "string" ? b.clienteTelefono : undefined,
    clienteNombre:
      typeof b.clienteNombre === "string" ? b.clienteNombre : undefined,
    descuentoPesos:
      b.descuentoPesos != null ? Number(b.descuentoPesos) : undefined,
    metodoPago: typeof b.metodoPago === "string" ? b.metodoPago : "efectivo",
    notas: typeof b.notas === "string" ? b.notas : null,
    motivo: typeof b.motivo === "string" ? b.motivo : null,
    lineas,
  });

  if (!res.ok) {
    const mensaje =
      res.reason === "not_found"
        ? "Venta no encontrada."
        : res.reason === "anulada"
          ? "No se puede modificar un comprobante anulado."
          : res.reason === "sin_lineas"
            ? "Agreg\u00e1 al menos un \u00edtem."
            : "Base de datos no disponible.";
    return NextResponse.json(
      { ok: false, mensaje },
      { status: res.reason === "not_found" ? 404 : 400 }
    );
  }

  const venta = await obtenerVentaDetalle(id);
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

  let motivo: string | null = null;
  try {
    const body = await request.json();
    if (body && typeof body === "object" && "motivo" in body) {
      motivo =
        typeof (body as { motivo?: unknown }).motivo === "string"
          ? (body as { motivo: string }).motivo
          : null;
    }
  } catch {
    /* DELETE sin cuerpo */
  }

  const res = await anularVenta(id, motivo);
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
