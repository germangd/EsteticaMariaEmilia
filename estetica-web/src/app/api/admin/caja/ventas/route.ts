import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  crearVenta,
  listarVentasSesion,
  type LineaVentaInput,
} from "@/lib/caja-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const sessionId = Number(request.nextUrl.searchParams.get("sessionId"));
  if (!Number.isFinite(sessionId) || sessionId < 1) {
    return NextResponse.json(
      { ok: false, mensaje: "sessionId requerido." },
      { status: 400 }
    );
  }

  const ventas = await listarVentasSesion(sessionId);
  if (!Array.isArray(ventas)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, ventas });
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

  const res = await crearVenta({
    sessionId: b.sessionId != null ? Number(b.sessionId) : undefined,
    clienteTelefono:
      typeof b.clienteTelefono === "string" ? b.clienteTelefono : undefined,
    clienteNombre:
      typeof b.clienteNombre === "string" ? b.clienteNombre : undefined,
    descuentoPesos:
      b.descuentoPesos != null ? Number(b.descuentoPesos) : undefined,
    metodoPago: typeof b.metodoPago === "string" ? b.metodoPago : "efectivo",
    notas: typeof b.notas === "string" ? b.notas : null,
    appointmentId:
      b.appointmentId != null ? Number(b.appointmentId) : undefined,
    clientPackageId:
      b.clientPackageId != null ? Number(b.clientPackageId) : undefined,
    lineas,
  });

  if (!res.ok) {
    const mensaje =
      res.reason === "sin_sesion"
        ? "Abr\u00ed la caja antes de registrar ventas."
        : res.reason === "sesion_cerrada"
          ? "La caja est\u00e1 cerrada."
          : res.reason === "sin_lineas"
            ? "Agreg\u00e1 al menos un \u00edtem."
            : res.reason === "no_db"
              ? "Base de datos no disponible."
              : "No se pudo registrar la venta.";
    return NextResponse.json({ ok: false, mensaje }, { status: 400 });
  }

  return NextResponse.json({ ok: true, id: res.id, numero: res.numero });
}
