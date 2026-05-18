import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  actualizarCliente,
  eliminarCliente,
  normalizarTelefono,
  obtenerClienteDetalle,
} from "@/lib/clientes-repo";

export const dynamic = "force-dynamic";

function decodeTelefonoParam(raw: string): string {
  try {
    return normalizarTelefono(decodeURIComponent(raw));
  } catch {
    return normalizarTelefono(raw);
  }
}

export async function GET(
  request: NextRequest,
  context: { params: Promise<{ telefono: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const telefono = decodeTelefonoParam((await context.params).telefono);
  const detalle = await obtenerClienteDetalle(telefono);

  if ("reason" in detalle) {
    if (detalle.reason === "not_found") {
      return NextResponse.json(
        { ok: false, mensaje: "Cliente no encontrado." },
        { status: 404 }
      );
    }
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }

  return NextResponse.json({ ok: true, cliente: detalle });
}

export async function PATCH(
  request: NextRequest,
  context: { params: Promise<{ telefono: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const telefono = decodeTelefonoParam((await context.params).telefono);

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const b = (body && typeof body === "object" ? body : {}) as Record<
    string,
    unknown
  >;

  const res = await actualizarCliente({
    telefono,
    telefonoNuevo:
      typeof b.telefonoNuevo === "string" ? b.telefonoNuevo : undefined,
    nombre: typeof b.nombre === "string" ? b.nombre : null,
    email: typeof b.email === "string" ? b.email : null,
    notas: typeof b.notas === "string" ? b.notas : null,
  });

  if (!res.ok) {
    const mensajes: Record<string, string> = {
      not_found: "Cliente no encontrado.",
      invalido: "Teléfono inválido.",
      telefono_ocupado: "Ya existe otro cliente con ese teléfono.",
      no_db: "Base de datos no disponible.",
    };
    const status =
      res.reason === "not_found"
        ? 404
        : res.reason === "telefono_ocupado" || res.reason === "invalido"
          ? 400
          : 503;
    return NextResponse.json(
      { ok: false, mensaje: mensajes[res.reason] ?? "No se pudo guardar." },
      { status }
    );
  }

  return NextResponse.json({ ok: true, telefono: res.telefono });
}

export async function DELETE(
  request: NextRequest,
  context: { params: Promise<{ telefono: string }> }
) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const telefono = decodeTelefonoParam((await context.params).telefono);
  const res = await eliminarCliente(telefono);

  if (!res.ok) {
    const mensajes: Record<string, string> = {
      not_found: "Cliente no encontrado.",
      invalido: "Teléfono inválido.",
      no_db: "Base de datos no disponible.",
    };
    const status = res.reason === "not_found" ? 404 : res.reason === "invalido" ? 400 : 503;
    return NextResponse.json(
      { ok: false, mensaje: mensajes[res.reason] ?? "No se pudo eliminar." },
      { status }
    );
  }

  return NextResponse.json({ ok: true, eliminados: res.eliminados });
}
