import { NextRequest, NextResponse } from "next/server";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { enviarMailsTurnoConfirmado } from "@/lib/mail-turno";
import { confirmarTurnoAnticipo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const id = Number((body as { id?: number }).id);
  if (!Number.isFinite(id) || id < 1) {
    return NextResponse.json({ ok: false, mensaje: "ID inválido." }, { status: 400 });
  }

  const res = await confirmarTurnoAnticipo(id);
  if (!res.ok) {
    const mensaje =
      res.reason === "no_pendiente"
        ? "El turno no está pendiente de anticipo o ya fue confirmado."
        : res.reason === "not_found"
          ? "Turno no encontrado."
          : "No se pudo confirmar.";
    const status =
      res.reason === "no_pendiente" || res.reason === "not_found" ? 404 : 503;
    return NextResponse.json({ ok: false, mensaje }, { status });
  }

  const t = res.turno;
  try {
    await enviarMailsTurnoConfirmado({
      nombre: t.nombreCliente,
      telefono: t.telefono,
      emailCliente: t.email,
      servicio: t.servicioNombre,
      sede: t.sedeNombre,
      responsable: t.responsable,
      fecha: t.fecha,
      hora: t.hora,
      codigoCancelacion: t.codigoCancelacion,
    });
  } catch (err) {
    console.error("[mail] confirmar anticipo:", err);
  }

  return NextResponse.json({
    ok: true,
    mensaje: "Turno confirmado. Se envió el aviso al cliente si tiene email.",
  });
}
