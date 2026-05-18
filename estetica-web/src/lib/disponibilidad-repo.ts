import { and, asc, eq, inArray } from "drizzle-orm";
import { getDb } from "@/db/client";
import {
  agendaEventServices,
  agendaEvents,
  serviceAvailabilityDates,
  services,
} from "@/db/schema";
import {
  type FranjaHoraria,
  horaAMinutos,
  intersectarFranjasConVentanaServicio,
} from "@/lib/agenda";
import { listarFranjasAtencionParaFecha } from "@/lib/horario-atencion-repo";
import { padHoraHHmm } from "@/lib/servicio-format";

export type EventoFranja = {
  id: number;
  nombre: string;
  descripcion: string | null;
  horarioInicio: string;
  horarioFin: string;
  serviceIds: number[];
};

export type VentanaHorario = {
  franjas: FranjaHoraria[];
  /** Motivo si no hay turnos (evento, calendario, etc.). */
  bloqueado?: string;
};

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

export function esFechaIsoValida(fecha: string): boolean {
  return ISO_DATE.test(fecha);
}

/** True si el inicio del turno (hora) cae dentro de [inicio, fin). */
export function horaDentroDeFranja(
  hora: string,
  inicio: string,
  fin: string
): boolean {
  const h = horaAMinutos(hora);
  const i = horaAMinutos(inicio);
  const f = horaAMinutos(fin);
  if (h < 0 || i < 0 || f < 0) return false;
  return h >= i && h < f;
}

/** Solapamiento de franjas horarias (mismo día). */
export function franjasHorariasSeSolapan(
  inicioA: string,
  finA: string,
  inicioB: string,
  finB: string
): boolean {
  const a0 = horaAMinutos(inicioA);
  const a1 = horaAMinutos(finA);
  const b0 = horaAMinutos(inicioB);
  const b1 = horaAMinutos(finB);
  if (a0 < 0 || a1 < 0 || b0 < 0 || b1 < 0 || a1 <= a0 || b1 <= b0) return false;
  return a0 < b1 && b0 < a1;
}

/** Reglas de eventos para un servicio en un horario concreto. */
export function evaluarReservaEnEvento(
  serviceId: number,
  hora: string,
  eventos: EventoFranja[]
): {
  permitido: boolean;
  mensaje?: string;
  ventana?: { horarioInicio: string; horarioFin: string };
} {
  if (eventos.length === 0) return { permitido: true };

  const eventosEnHora = eventos.filter((e) =>
    horaDentroDeFranja(hora, e.horarioInicio, e.horarioFin)
  );

  const participaEnDia = eventos.some((e) => e.serviceIds.includes(serviceId));

  if (eventosEnHora.length > 0) {
    const eventoDelServicio = eventosEnHora.find((e) =>
      e.serviceIds.includes(serviceId)
    );
    if (!eventoDelServicio) {
      const ev = eventosEnHora[0]!;
      return {
        permitido: false,
        mensaje: `De ${ev.horarioInicio} a ${ev.horarioFin} hay evento (“${ev.nombre}”). Este servicio no participa.`,
      };
    }
    return {
      permitido: true,
      ventana: {
        horarioInicio: eventoDelServicio.horarioInicio,
        horarioFin: eventoDelServicio.horarioFin,
      },
    };
  }

  if (participaEnDia) {
    const misEventos = eventos.filter((e) => e.serviceIds.includes(serviceId));
    const franjas = misEventos
      .map((e) => `${e.horarioInicio}–${e.horarioFin}`)
      .join(", ");
    return {
      permitido: false,
      mensaje: `Este servicio solo se reserva en el horario del evento (${franjas}).`,
    };
  }

  for (const ev of eventos) {
    if (horaDentroDeFranja(hora, ev.horarioInicio, ev.horarioFin)) {
      return {
        permitido: false,
        mensaje: `De ${ev.horarioInicio} a ${ev.horarioFin} hay evento (“${ev.nombre}”). En esa franja no hay turnos para otros servicios.`,
      };
    }
  }

  return { permitido: true };
}

/** Eventos activos en una fecha y sede con servicios vinculados. */
export async function listarEventosActivosEnFecha(
  fecha: string,
  sedeId: number
): Promise<EventoFranja[]> {
  if (!esFechaIsoValida(fecha) || sedeId < 1) return [];
  const db = getDb();
  if (!db) return [];

  const rows = await db
    .select()
    .from(agendaEvents)
    .where(
      and(
        eq(agendaEvents.fecha, fecha),
        eq(agendaEvents.sedeId, sedeId),
        eq(agendaEvents.activo, true)
      )
    )
    .orderBy(asc(agendaEvents.horarioInicio), asc(agendaEvents.id));

  if (rows.length === 0) return [];

  const eventIds = rows.map((r) => r.id);
  const links = await db
    .select({
      eventId: agendaEventServices.eventId,
      serviceId: agendaEventServices.serviceId,
    })
    .from(agendaEventServices)
    .where(inArray(agendaEventServices.eventId, eventIds));

  const svcByEvent = new Map<number, number[]>();
  for (const link of links) {
    const list = svcByEvent.get(link.eventId) ?? [];
    list.push(link.serviceId);
    svcByEvent.set(link.eventId, list);
  }

  return rows.map((r) => ({
    id: r.id,
    nombre: r.nombre,
    descripcion: r.descripcion,
    horarioInicio: padHoraHHmm(r.horarioInicio),
    horarioFin: padHoraHHmm(r.horarioFin),
    serviceIds: svcByEvent.get(r.id) ?? [],
  }));
}

/** Fechas habilitadas para un servicio (vacío = sin restricción por calendario). */
export async function listarFechasServicio(
  serviceId: number
): Promise<string[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const rows = await db
    .select({ fecha: serviceAvailabilityDates.fecha })
    .from(serviceAvailabilityDates)
    .where(eq(serviceAvailabilityDates.serviceId, serviceId))
    .orderBy(asc(serviceAvailabilityDates.fecha));

  return rows.map((r) => r.fecha);
}

export async function servicioUsaCalendarioFechas(
  serviceId: number
): Promise<boolean> {
  const db = getDb();
  if (!db) return false;
  const rows = await db
    .select({ fecha: serviceAvailabilityDates.fecha })
    .from(serviceAvailabilityDates)
    .where(eq(serviceAvailabilityDates.serviceId, serviceId))
    .limit(1);
  return rows.length > 0;
}

export async function fechaPermitidaParaServicio(
  serviceId: number,
  fecha: string
): Promise<boolean> {
  if (!esFechaIsoValida(fecha)) return false;
  const fechas = await listarFechasServicio(serviceId);
  if (Array.isArray(fechas) === false) return true;
  if (fechas.length === 0) return true;
  return fechas.includes(fecha);
}

export async function reemplazarFechasServicio(
  serviceId: number,
  fechas: string[]
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "not_found" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  const [svc] = await db
    .select({ id: services.id })
    .from(services)
    .where(eq(services.id, serviceId))
    .limit(1);
  if (!svc) return { ok: false, reason: "not_found" };

  const limpias = [
    ...new Set(
      fechas.filter((f) => esFechaIsoValida(f)).sort((a, b) => a.localeCompare(b))
    ),
  ];

  await db
    .delete(serviceAvailabilityDates)
    .where(eq(serviceAvailabilityDates.serviceId, serviceId));

  if (limpias.length > 0) {
    await db.insert(serviceAvailabilityDates).values(
      limpias.map((fecha) => ({ serviceId, fecha }))
    );
  }

  return { ok: true };
}

export async function agregarFechaServicio(
  serviceId: number,
  fecha: string
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "invalido" }> {
  if (!esFechaIsoValida(fecha)) return { ok: false, reason: "invalido" };
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  await db
    .insert(serviceAvailabilityDates)
    .values({ serviceId, fecha })
    .onConflictDoNothing();

  return { ok: true };
}

export async function quitarFechaServicio(
  serviceId: number,
  fecha: string
): Promise<{ ok: true } | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };

  await db
    .delete(serviceAvailabilityDates)
    .where(
      and(
        eq(serviceAvailabilityDates.serviceId, serviceId),
        eq(serviceAvailabilityDates.fecha, fecha)
      )
    );

  return { ok: true };
}

export async function servicioPermitidoEnEvento(
  eventId: number,
  serviceId: number
): Promise<boolean> {
  const db = getDb();
  if (!db) return false;

  const [link] = await db
    .select({ serviceId: agendaEventServices.serviceId })
    .from(agendaEventServices)
    .where(
      and(
        eq(agendaEventServices.eventId, eventId),
        eq(agendaEventServices.serviceId, serviceId)
      )
    )
    .limit(1);

  return Boolean(link);
}

/** Horario y permiso de reserva para servicio en una fecha (y hora si se indica). */
export async function resolverVentanaReserva(
  serviceId: number,
  servicio: {
    horarioInicio: string;
    horarioFin: string;
    nombre: string;
  },
  fecha: string,
  sedeId: number,
  hora?: string
): Promise<VentanaHorario | null> {
  const ventanaServicio: FranjaHoraria = {
    horarioInicio: padHoraHHmm(servicio.horarioInicio),
    horarioFin: padHoraHHmm(servicio.horarioFin),
  };

  const permitida = await fechaPermitidaParaServicio(serviceId, fecha);
  if (!permitida) {
    return {
      franjas: [],
      bloqueado:
        "Este servicio solo admite turnos en fechas habilitadas en el calendario (configuración admin).",
    };
  }

  if (sedeId < 1) {
    return { franjas: [], bloqueado: "Elegí una sede válida." };
  }

  const franjasLocal = await listarFranjasAtencionParaFecha(fecha, sedeId);
  if (franjasLocal.length === 0) {
    return {
      franjas: [],
      bloqueado:
        "No hay atención ese día en esta sede (domingo o sin franjas en Horarios).",
    };
  }

  let franjas = intersectarFranjasConVentanaServicio(
    franjasLocal,
    ventanaServicio
  );
  if (franjas.length === 0) {
    return {
      franjas: [],
      bloqueado:
        "El horario del servicio no coincide con las franjas de atención del local ese día.",
    };
  }

  const eventos = await listarEventosActivosEnFecha(fecha, sedeId);

  if (hora && eventos.length > 0) {
    const ev = evaluarReservaEnEvento(serviceId, hora, eventos);
    if (!ev.permitido) {
      return { franjas: [], bloqueado: ev.mensaje };
    }
    if (ev.ventana) {
      franjas = intersectarFranjasConVentanaServicio([ev.ventana], ventanaServicio);
      if (franjas.length === 0) {
        return {
          franjas: [],
          bloqueado: "El horario del evento no es compatible con este servicio.",
        };
      }
    }
  }

  return { franjas };
}
