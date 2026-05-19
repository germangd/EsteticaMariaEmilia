import { eq } from "drizzle-orm";
import { NextRequest, NextResponse } from "next/server";
import { getDb } from "@/db/client";
import { services } from "@/db/schema";
import {
  adminUnauthorizedResponse,
  isAdminRequest,
} from "@/lib/admin-api-auth";
import { esFechaHoraValida, getAppTimeZone, normalizarHora } from "@/lib/agenda";
import { enviarMailsTurnoConfirmado } from "@/lib/mail-turno";
import {
  resolverItemReserva,
  validarReservaItem,
} from "@/lib/reserva-catalogo";
import { obtenerSedePorId } from "@/lib/sedes-repo";
import { insertarTurnoSiHayCupo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

type Body = {
  servicioId?: number;
  paqueteId?: number;
  sedeId?: number;
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
  const sedeId = Number(body.sedeId);
  const paqueteId = Number(body.paqueteId);
  const servicioId = Number(body.servicioId);

  if (!fecha || !horaRaw || !nombre || !telefono || !Number.isFinite(sedeId) || sedeId < 1) {
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

  let item = null;
  if (Number.isFinite(paqueteId) && paqueteId > 0) {
    item = await resolverItemReserva({ tipo: "paquete", id: paqueteId });
  } else if (Number.isFinite(servicioId) && servicioId > 0) {
    const db = getDb();
    if (db) {
      const [row] = await db
        .select({ nombre: services.nombre, esGrupo: services.esGrupo })
        .from(services)
        .where(eq(services.id, servicioId))
        .limit(1);
      if (row && !row.esGrupo) {
        item = await resolverItemReserva({
          tipo: "servicio",
          nombre: row.nombre,
        });
      }
    }
  } else {
    const nombreServ = body.servicio?.trim() ?? "";
    if (nombreServ) {
      item = await resolverItemReserva({ tipo: "servicio", nombre: nombreServ });
    }
  }

  if (!item) {
    return NextResponse.json(
      { ok: false, mensaje: "Elegí un servicio o combo válido." },
      { status: 400 }
    );
  }

  const sede = await obtenerSedePorId(sedeId);
  if (!sede?.activo) {
    return NextResponse.json({ ok: false, mensaje: "Sede no válida." }, { status: 400 });
  }

  const validacion = await validarReservaItem(item, fecha, hora, sedeId);
  if (!validacion.ok) {
    return NextResponse.json({ ok: false, mensaje: validacion.mensaje });
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
    if (ins.reason === "cupo") {
      return NextResponse.json({
        ok: false,
        mensaje: `No hay cupo para ${item.nombre} a las ${hora}.`,
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
        servicio: item.nombre,
        sede: sede.nombre,
        responsable: item.responsable,
        fecha,
        hora,
        codigoCancelacion: ins.codigo,
      });
    } catch (err) {
      console.error("[mail] admin crear turno:", err);
    }
  }

  const tipoEtiqueta = item.tipo === "paquete" ? "Combo" : "Servicio";
  return NextResponse.json({
    ok: true,
    mensaje: `${tipoEtiqueta} cargado.`,
    codigo: ins.codigo,
  });
}
