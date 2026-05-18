import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { crearVentaDesdePaquete } from "@/lib/caja-repo";
import {
  asignarPaqueteCliente,
  listarAsignacionesPaquete,
} from "@/lib/paquetes-repo";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const todas = request.nextUrl.searchParams.get("todas") === "1";
  const rows = await listarAsignacionesPaquete(!todas);
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, asignaciones: rows });
}

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  if (!body || typeof body !== "object") {
    return NextResponse.json({ ok: false, mensaje: "Datos inválidos." }, { status: 400 });
  }

  const b = body as Record<string, unknown>;
  const packageId = Number(b.packageId);
  const nombreCliente = typeof b.nombreCliente === "string" ? b.nombreCliente : "";
  const telefono = typeof b.telefono === "string" ? b.telefono : "";
  const fechaCompra = typeof b.fechaCompra === "string" ? b.fechaCompra : "";
  const precioCobradoPesos =
    b.precioCobradoPesos != null ? Number(b.precioCobradoPesos) : undefined;
  const notas = typeof b.notas === "string" ? b.notas : null;
  const sesiones = b.sesiones != null ? Number(b.sesiones) : undefined;
  const registrarEnCaja = b.registrarEnCaja !== false;
  const metodoPago =
    typeof b.metodoPago === "string" ? b.metodoPago : "efectivo";

  const res = await asignarPaqueteCliente({
    packageId,
    nombreCliente,
    telefono,
    fechaCompra,
    precioCobradoPesos,
    notas,
    sesiones: Number.isFinite(sesiones) ? sesiones : undefined,
  });

  if (!res.ok) {
    return NextResponse.json(
      {
        ok: false,
        mensaje:
          res.reason === "not_found"
            ? "Paquete no encontrado o inactivo."
            : "Completá cliente, teléfono y fecha.",
      },
      { status: res.reason === "not_found" ? 404 : 400 }
    );
  }

  let ventaCaja: { id: number; numero: number } | null = null;
  let avisoCaja: string | undefined;

  if (registrarEnCaja) {
    const v = await crearVentaDesdePaquete({
      clientPackageId: res.id,
      metodoPago,
      notas: notas ?? undefined,
    });
    if (v.ok) {
      ventaCaja = { id: v.id, numero: v.numero };
    } else if (v.reason === "ya_cobrado") {
      avisoCaja = "Paquete asignado; el cobro en caja ya estaba registrado.";
    } else if (v.reason === "sin_sesion") {
      avisoCaja =
        "Paquete asignado. Abr\u00ed la caja para emitir el ticket de cobro.";
    } else if (v.reason === "sesion_cerrada") {
      avisoCaja = "Paquete asignado. La caja est\u00e1 cerrada; no se emiti\u00f3 ticket.";
    } else if (v.reason === "sin_monto") {
      avisoCaja = "Paquete asignado. Indic\u00e1 un monto mayor a cero para ticket.";
    }
  }

  return NextResponse.json({
    ok: true,
    id: res.id,
    ventaCaja,
    avisoCaja,
  });
}
