import { and, eq, inArray } from "drizzle-orm";
import { getDb } from "@/db/client";
import {
  packageServices,
  servicePackages,
  services,
  type ServiceRow,
} from "@/db/schema";
import {
  contarSolapamiento,
  generarHorariosDesdeFranjas,
  getAppTimeZone,
  hoyIsoEnZona,
  horaActualEnZona,
  horaAMinutos,
  intersectarListasFranjas,
  turnoCabeEnFranjas,
  type FranjaHoraria,
} from "@/lib/agenda";
import {
  evaluarReservaEnEvento,
  listarEventosActivosEnFecha,
  resolverVentanaReserva,
  type VentanaHorario,
} from "@/lib/disponibilidad-repo";
import {
  claveReservaPaquete,
  claveReservaServicio,
  parsearClaveReserva,
} from "@/lib/reserva-claves";
import { listarActivosConDuracion } from "@/lib/turnos-repo";

export { claveReservaPaquete, claveReservaServicio, parsearClaveReserva };

export type ItemReserva = {
  tipo: "servicio" | "paquete";
  nombre: string;
  duracionMin: number;
  capacidad: number;
  responsable: string;
  serviceIds: number[];
  paqueteId?: number;
  sesionesTotal?: number;
};

async function listarServiciosDePaquete(
  packageId: number
): Promise<ServiceRow[]> {
  const db = getDb();
  if (!db) return [];

  const links = await db
    .select({ serviceId: packageServices.serviceId })
    .from(packageServices)
    .where(eq(packageServices.packageId, packageId));

  const ids = links.map((l) => l.serviceId);
  if (ids.length === 0) return [];

  return db
    .select()
    .from(services)
    .where(and(inArray(services.id, ids), eq(services.esGrupo, false)));
}

function agregarDatosScheduling(
  svcs: ServiceRow[]
): Pick<ItemReserva, "duracionMin" | "capacidad" | "responsable" | "serviceIds"> {
  if (svcs.length === 0) {
    return {
      duracionMin: 30,
      capacidad: 1,
      responsable: "No asignado",
      serviceIds: [],
    };
  }
  return {
    duracionMin: svcs.reduce(
      (acc, s) => acc + Math.max(5, s.duracionMin),
      0
    ),
    capacidad: Math.min(...svcs.map((s) => Math.max(1, s.capacidad))),
    responsable: svcs[0]!.responsable.trim() || "No asignado",
    serviceIds: svcs.map((s) => s.id),
  };
}

export async function resolverItemReserva(
  parsed:
    | { tipo: "servicio"; nombre: string }
    | { tipo: "paquete"; id: number }
): Promise<ItemReserva | null> {
  const db = getDb();
  if (!db) return null;

  if (parsed.tipo === "servicio") {
    const [servicio] = await db
      .select()
      .from(services)
      .where(eq(services.nombre, parsed.nombre))
      .limit(1);
    if (!servicio || servicio.esGrupo) return null;
    const sched = agregarDatosScheduling([servicio]);
    return {
      tipo: "servicio",
      nombre: servicio.nombre,
      ...sched,
    };
  }

  const [paquete] = await db
    .select()
    .from(servicePackages)
    .where(
      and(eq(servicePackages.id, parsed.id), eq(servicePackages.activo, true))
    )
    .limit(1);
  if (!paquete) return null;

  const svcs = await listarServiciosDePaquete(paquete.id);
  if (svcs.length === 0) return null;

  const sched = agregarDatosScheduling(svcs);
  return {
    tipo: "paquete",
    nombre: paquete.nombre,
    paqueteId: paquete.id,
    sesionesTotal: paquete.sesionesTotal,
    ...sched,
  };
}

async function resolverVentanaReservaPaquete(
  svcs: ServiceRow[],
  fecha: string,
  sedeId: number,
  hora?: string
): Promise<VentanaHorario | null> {
  let franjas: FranjaHoraria[] | null = null;
  let bloqueado: string | undefined;

  for (const svc of svcs) {
    const ventana = await resolverVentanaReserva(
      svc.id,
      svc,
      fecha,
      sedeId,
      hora
    );
    if (!ventana) return null;
    if (ventana.bloqueado) {
      bloqueado = ventana.bloqueado;
      break;
    }
    franjas =
      franjas === null
        ? ventana.franjas
        : intersectarListasFranjas(franjas, ventana.franjas);
    if (franjas.length === 0) {
      bloqueado =
        "El combo no tiene horario compatible ese día (revisá servicios incluidos y sede).";
      break;
    }
  }

  if (bloqueado || !franjas?.length) {
    return { franjas: [], bloqueado };
  }
  return { franjas };
}

export async function calcularHorariosDisponibles(
  item: ItemReserva,
  fecha: string,
  sedeId: number
): Promise<{ horarios: string[]; mensaje?: string }> {
  const db = getDb();
  if (!db) return { horarios: [], mensaje: "Base de datos no disponible." };

  let ventana: VentanaHorario | null;
  if (item.tipo === "servicio") {
    const [servicio] = await db
      .select()
      .from(services)
      .where(eq(services.nombre, item.nombre))
      .limit(1);
    if (!servicio) return { horarios: [] };
    ventana = await resolverVentanaReserva(
      servicio.id,
      servicio,
      fecha,
      sedeId
    );
  } else {
    const svcs = await listarServiciosDePaquete(item.paqueteId!);
    if (svcs.length === 0) {
      return { horarios: [], mensaje: "El combo no tiene servicios configurados." };
    }
    ventana = await resolverVentanaReservaPaquete(svcs, fecha, sedeId);
  }

  if (!ventana || ventana.bloqueado) {
    return {
      horarios: [],
      mensaje: ventana?.bloqueado ?? "Fecha no disponible.",
    };
  }

  const eventos = await listarEventosActivosEnFecha(fecha, sedeId);
  const duracionMin = Math.max(5, item.duracionMin);
  let horariosPosibles = generarHorariosDesdeFranjas(
    ventana.franjas,
    duracionMin
  );

  if (eventos.length > 0) {
    horariosPosibles = horariosPosibles.filter((h) => {
      for (const serviceId of item.serviceIds) {
        const ev = evaluarReservaEnEvento(serviceId, h, eventos);
        if (!ev.permitido) return false;
        if (ev.ventana) {
          const finEv = horaAMinutos(ev.ventana.horarioFin);
          const inicio = horaAMinutos(h);
          if (inicio < 0 || inicio + duracionMin > finEv) return false;
        }
      }
      return true;
    });
  }

  if (horariosPosibles.length === 0) {
    return {
      horarios: [],
      mensaje:
        eventos.length > 0
          ? "No hay horarios para este servicio o combo en esa fecha."
          : undefined,
    };
  }

  const tz = getAppTimeZone();
  const hoyStr = hoyIsoEnZona(tz);
  if (fecha === hoyStr) {
    const horaActual = horaActualEnZona(tz);
    horariosPosibles = horariosPosibles.filter((h) => h >= horaActual);
  }

  const ocupados = await listarActivosConDuracion(fecha, item.responsable, sedeId);
  const disponibles = horariosPosibles.filter((hora) => {
    const inicio = horaAMinutos(hora);
    if (inicio < 0) return false;
    return contarSolapamiento(inicio, duracionMin, ocupados) < item.capacidad;
  });

  return { horarios: disponibles };
}

export async function validarReservaItem(
  item: ItemReserva,
  fecha: string,
  hora: string,
  sedeId: number
): Promise<{ ok: true; franjas: FranjaHoraria[] } | { ok: false; mensaje: string }> {
  let ventana: VentanaHorario | null;

  if (item.tipo === "servicio") {
    const db = getDb();
    if (!db) return { ok: false, mensaje: "Base de datos no disponible." };
    const [servicio] = await db
      .select()
      .from(services)
      .where(eq(services.nombre, item.nombre))
      .limit(1);
    if (!servicio) return { ok: false, mensaje: "Servicio no encontrado." };
    ventana = await resolverVentanaReserva(
      servicio.id,
      servicio,
      fecha,
      sedeId,
      hora
    );
  } else {
    const svcs = await listarServiciosDePaquete(item.paqueteId!);
    if (svcs.length === 0) {
      return { ok: false, mensaje: "El combo no tiene servicios configurados." };
    }
    ventana = await resolverVentanaReservaPaquete(svcs, fecha, sedeId, hora);
  }

  if (!ventana || ventana.bloqueado) {
    return {
      ok: false,
      mensaje:
        ventana?.bloqueado ??
        "Esta fecha no está habilitada para reservar.",
    };
  }

  if (!turnoCabeEnFranjas(hora, item.duracionMin, ventana.franjas)) {
    return {
      ok: false,
      mensaje:
        "Ese horario no alcanza para la duración del servicio o combo antes del cierre.",
    };
  }

  const eventos = await listarEventosActivosEnFecha(fecha, sedeId);
  if (eventos.length > 0) {
    for (const serviceId of item.serviceIds) {
      const ev = evaluarReservaEnEvento(serviceId, hora, eventos);
      if (!ev.permitido) {
        return { ok: false, mensaje: ev.mensaje ?? "Horario no disponible." };
      }
    }
  }

  return { ok: true, franjas: ventana.franjas };
}
