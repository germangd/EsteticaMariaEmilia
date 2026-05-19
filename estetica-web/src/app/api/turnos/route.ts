import { NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { esFechaHoraValida, getAppTimeZone, normalizarHora } from "@/lib/agenda";
import { enviarMailsTurnoConfirmado } from "@/lib/mail-turno";
import {
  parsearClaveReserva,
  resolverItemReserva,
  validarReservaItem,
} from "@/lib/reserva-catalogo";
import { obtenerSedePorId } from "@/lib/sedes-repo";
import { insertarTurnoSiHayCupo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

type Body = {
  /** Clave `s:nombre` o `p:id` (preferido). */
  clave?: string;
  servicio?: string;
  fecha?: string;
  hora?: string;
  nombre?: string;
  telefono?: string;
  email?: string;
  sedeId?: number;
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

  const claveRaw =
    body.clave?.trim() ??
    (body.servicio?.trim() ? `s:${body.servicio.trim()}` : "");
  const fecha = body.fecha?.trim() ?? "";
  const horaRaw = body.hora?.trim() ?? "";
  const hora = normalizarHora(horaRaw);
  const nombre = body.nombre?.trim() ?? "";
  const telefono = body.telefono?.trim() ?? "";
  const email = body.email?.trim() || "";
  const sedeId = Number(body.sedeId);

  const parsed = parsearClaveReserva(claveRaw);

  if (
    !parsed ||
    !fecha ||
    !horaRaw ||
    !nombre ||
    !telefono ||
    !Number.isFinite(sedeId) ||
    sedeId < 1
  ) {
    return NextResponse.json(
      {
        exito: false,
        mensaje: "Faltan servicio o combo, sede, fecha, hora, nombre o teléfono.",
      },
      { status: 400 }
    );
  }

  const sede = await obtenerSedePorId(sedeId);
  if (!sede?.activo) {
    return NextResponse.json({
      exito: false,
      mensaje: "Sede no válida.",
    });
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
    const item = await resolverItemReserva(parsed);
    if (!item) {
      return NextResponse.json({
        exito: false,
        mensaje: "Servicio o combo no encontrado.",
      });
    }

    const validacion = await validarReservaItem(item, fecha, hora, sedeId);
    if (!validacion.ok) {
      return NextResponse.json({
        exito: false,
        mensaje: validacion.mensaje,
      });
    }

    const ins = await insertarTurnoSiHayCupo({
      fecha,
      hora,
      nombre,
      telefono,
      email: email || null,
      servicioNombre: item.nombre,
      responsable: item.responsable,
      sedeId,
      capacidad: item.capacidad,
      duracionMin: item.duracionMin,
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
          mensaje: `No hay cupo para ${item.nombre} a las ${hora}.`,
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
        servicio: item.nombre,
        sede: sede.nombre,
        responsable: item.responsable,
        fecha,
        hora,
        codigoCancelacion: ins.codigo,
      });
    } catch (err) {
      console.error("[mail] turno:", err);
    }

    const tipoEtiqueta = item.tipo === "paquete" ? "Combo" : "Servicio";
    return NextResponse.json({
      exito: true,
      mensaje: `${tipoEtiqueta} reservado con ${item.responsable}. Código: ${ins.codigo}`,
      codigo: ins.codigo,
    });
  } catch (e) {
    return NextResponse.json(
      { exito: false, mensaje: String(e) },
      { status: 500 }
    );
  }
}
