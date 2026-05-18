import { eq } from "drizzle-orm";
import { NextRequest, NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services } from "@/db/schema";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import {
  esFechaHoraValida,
  getAppTimeZone,
  normalizarHora,
  turnoCabeEnHorario,
} from "@/lib/agenda";
import { enviarMailsTurnoConfirmado } from "@/lib/mail-turno";
import { resolverVentanaReserva } from "@/lib/disponibilidad-repo";
import { insertarTurnoSiHayCupo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

type Body = {
  servicioId?: number;
  servicio?: string;
  fecha?: string;
  hora?: string;
  nombre?: string;
  telefono?: string;
  email?: string;
  enviarMail?: boolean;
};

export async function POST(request: NextRequest) {
  if (!isAdminRequest(request)) return adminUnauthorizedResponse();

  const db = getDb();
  if (!db) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no configurada." },
      { status: 503 }
    );
  }

  let body: Body;
  try {
    body = (await request.json()) as Body;
  } catch {
    return NextResponse.json({ ok: false, mensaje: "JSON inválido." }, { status: 400 });
  }

  const fecha = body.fecha?.trim() ?? "";
  const horaRaw = body.hora?.trim() ?? "";
  const hora = normalizarHora(horaRaw);
  const nombre = body.nombre?.trim() ?? "";
  const telefono = body.telefono?.trim() ?? "";
  const email = body.email?.trim() || "";

  if (!fecha || !horaRaw || !nombre || !telefono) {
    return NextResponse.json(
      { ok: false, mensaje: "Faltan fecha, hora, nombre o teléfono." },
      { status: 400 }
    );
  }

  if (!hora) {
    return NextResponse.json({ ok: false, mensaje: "Hora inválida." }, { status: 400 });
  }

  const tz = getAppTimeZone();
  if (!esFechaHoraValida(fecha, hora, tz)) {
    return NextResponse.json({
      ok: false,
      mensaje: "Fecha u hora inválida (pasada, domingo o fuera de horario).",
    });
  }

  let servicioRow: (typeof services.$inferSelect) | undefined;

  if (body.servicioId && Number.isFinite(body.servicioId)) {
    [servicioRow] = await db
      .select()
      .from(services)
      .where(eq(services.id, body.servicioId))
      .limit(1);
  } else {
    const nombreServ = body.servicio?.trim() ?? "";
    if (!nombreServ) {
      return NextResponse.json(
        { ok: false, mensaje: "Elegí un servicio." },
        { status: 400 }
      );
    }
    [servicioRow] = await db
      .select()
      .from(services)
      .where(eq(services.nombre, nombreServ))
      .limit(1);
  }

  if (!servicioRow) {
    return NextResponse.json(
      { ok: false, mensaje: "Servicio no encontrado." },
      { status: 404 }
    );
  }

  const ventana = await resolverVentanaReserva(
    servicioRow.id,
    servicioRow,
    fecha,
    hora
  );
  if (!ventana || ventana.bloqueado) {
    return NextResponse.json({
      ok: false,
      mensaje:
        ventana?.bloqueado ??
        "Esta fecha no está habilitada para este servicio.",
    });
  }

  if (
    !turnoCabeEnHorario(hora, servicioRow.duracionMin, ventana.horarioFin)
  ) {
    return NextResponse.json({
      ok: false,
      mensaje:
        "Ese horario no alcanza para la duración del servicio antes del cierre.",
    });
  }

  const ins = await insertarTurnoSiHayCupo({
    fecha,
    hora,
    nombre,
    telefono,
    email: email || null,
    servicioNombre: servicioRow.nombre,
    responsable: servicioRow.responsable,
    capacidad: servicioRow.capacidad,
    duracionMin: servicioRow.duracionMin,
  });

  if (ins.ok === false) {
    if (ins.reason === "cupo") {
      return NextResponse.json({
        ok: false,
        mensaje: `No hay cupo (${servicioRow.capacidad} cliente(s) por turno en ese horario).`,
      });
    }
    return NextResponse.json({
      ok: false,
      mensaje: "No se pudo guardar el turno.",
    });
  }

  if (body.enviarMail !== false) {
    try {
      await enviarMailsTurnoConfirmado({
        nombre,
        telefono,
        emailCliente: email || null,
        servicio: servicioRow.nombre,
        responsable: servicioRow.responsable,
        fecha,
        hora,
        codigoCancelacion: ins.codigo,
      });
    } catch (err) {
      console.error("[mail] admin crear turno:", err);
    }
  }

  return NextResponse.json({
    ok: true,
    mensaje: "Turno cargado.",
    codigo: ins.codigo,
  });
}
