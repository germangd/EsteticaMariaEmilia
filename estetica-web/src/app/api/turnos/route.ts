import { eq } from "drizzle-orm";
import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services } from "@/db/schema";
import {
  esFechaHoraValida,
  getAppTimeZone,
  normalizarHora,
  turnoCabeEnHorario,
} from "@/lib/agenda";
import { enviarMailsTurnoConfirmado } from "@/lib/mail-turno";
import { insertarTurnoSiHayCupo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

type Body = {
  servicio?: string;
  fecha?: string;
  hora?: string;
  nombre?: string;
  telefono?: string;
  email?: string;
};

/**
 * Crear turno (paridad con `guardarTurno` en `Código.gs`).
 * Mails vía Resend si están `RESEND_API_KEY`, `EMAIL_FROM` y (opcional cliente) email en el body; `OWNER_EMAIL` para el aviso al dueño.
 */
export async function POST(request: Request) {
  const db = getDb();
  if (!db) {
    return NextResponse.json(
      { exito: false, mensaje: "Base de datos no configurada." },
      { status: 503 }
    );
  }

  let body: Body;
  try {
    body = (await request.json()) as Body;
  } catch {
    return NextResponse.json(
      { exito: false, mensaje: "JSON inválido." },
      { status: 400 }
    );
  }

  const servicioNombre = body.servicio?.trim() ?? "";
  const fecha = body.fecha?.trim() ?? "";
  const horaRaw = body.hora?.trim() ?? "";
  const hora = normalizarHora(horaRaw);
  const nombre = body.nombre?.trim() ?? "";
  const telefono = body.telefono?.trim() ?? "";
  const email = body.email?.trim() || "";

  if (!servicioNombre || !fecha || !horaRaw || !nombre || !telefono) {
    return NextResponse.json(
      {
        exito: false,
        mensaje: "Faltan servicio, fecha, hora, nombre o teléfono.",
      },
      { status: 400 }
    );
  }

  if (!hora) {
    return NextResponse.json(
      { exito: false, mensaje: "Hora inválida." },
      { status: 400 }
    );
  }

  const tz = getAppTimeZone();
  if (!esFechaHoraValida(fecha, hora, tz)) {
    return NextResponse.json({
      exito: false,
      mensaje:
        "Fecha u hora inválida (pasada, domingo o ya transcurrida en la zona horaria del negocio).",
    });
  }

  try {
    const [servicio] = await db
      .select()
      .from(services)
      .where(eq(services.nombre, servicioNombre))
      .limit(1);

    if (!servicio) {
      return NextResponse.json({
        exito: false,
        mensaje: "Servicio no encontrado.",
      });
    }

    if (!turnoCabeEnHorario(hora, servicio.duracionMin, servicio.horarioFin)) {
      return NextResponse.json({
        exito: false,
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
      servicioNombre,
      responsable: servicio.responsable,
      capacidad: servicio.capacidad,
      duracionMin: servicio.duracionMin,
    });

    if (ins.ok === false) {
      if (ins.reason === "no_db") {
        return NextResponse.json(
          { exito: false, mensaje: "Base de datos no disponible." },
          { status: 503 }
        );
      }
      if (ins.reason === "cupo") {
        return NextResponse.json({
          exito: false,
          mensaje: `No hay cupo para ${servicioNombre} a las ${hora}. Capacidad: ${servicio.capacidad}.`,
        });
      }
      return NextResponse.json({
        exito: false,
        mensaje: "No se pudo generar código de cancelación, intentá de nuevo.",
      });
    }

    try {
      await enviarMailsTurnoConfirmado({
        nombre,
        telefono,
        emailCliente: email || null,
        servicio: servicioNombre,
        responsable: servicio.responsable,
        fecha,
        hora,
        codigoCancelacion: ins.codigo,
      });
    } catch (err) {
      console.error("[mail] turno:", err);
    }

    return NextResponse.json({
      exito: true,
      mensaje: `Turno guardado con ${servicio.responsable}. Código: ${ins.codigo}`,
      codigo: ins.codigo,
    });
  } catch (e) {
    return NextResponse.json(
      { exito: false, mensaje: String(e) },
      { status: 500 }
    );
  }
}
