import { DateTime } from "luxon";

export type VistaAgenda = "dia" | "semana" | "mes" | "fecha";

/** Ancla para volver al listado tras aplicar filtros en /admin/turnos */
export const TURNOS_ASIGNADOS_ANCHOR = "turnos-asignados";

export function parseVistaAgenda(raw: string | undefined): VistaAgenda {
  if (raw === "semana" || raw === "mes" || raw === "fecha") return raw;
  return "dia";
}

export function rangoAgenda(params: {
  vista: VistaAgenda;
  refIso: string;
  tz: string;
  hoyIso: string;
}): { desde: string; hasta: string; etiqueta: string } {
  const ref = DateTime.fromISO(params.refIso, { zone: params.tz });
  if (!ref.isValid) {
    return {
      desde: params.hoyIso,
      hasta: params.hoyIso,
      etiqueta: params.hoyIso,
    };
  }

  const fmtCorto = (iso: string) => {
    const dt = DateTime.fromISO(iso, { zone: params.tz });
    return dt.isValid
      ? dt.setLocale("es").toFormat("ccc d MMM yyyy")
      : iso;
  };

  if (params.vista === "dia" || params.vista === "fecha") {
    const d = ref.toISODate()!;
    const esHoy = d === params.hoyIso && params.vista === "dia";
    return {
      desde: d,
      hasta: d,
      etiqueta: esHoy ? `Hoy · ${fmtCorto(d)}` : fmtCorto(d),
    };
  }

  if (params.vista === "semana") {
    const inicio = ref.startOf("week");
    const fin = ref.endOf("week");
    const desde = inicio.toISODate()!;
    const hasta = fin.toISODate()!;
    return {
      desde,
      hasta,
      etiqueta: `Semana · ${fmtCorto(desde)} – ${fmtCorto(hasta)}`,
    };
  }

  const inicio = ref.startOf("month");
  const fin = ref.endOf("month");
  const desde = inicio.toISODate()!;
  const hasta = fin.toISODate()!;
  const mesNombre = inicio.setLocale("es").toFormat("MMMM yyyy");
  return {
    desde,
    hasta,
    etiqueta: `Mes · ${mesNombre}`,
  };
}

export function buildTurnosQuery(q: {
  vista: VistaAgenda;
  ref: string;
  servicio?: string;
}): string {
  const u = new URLSearchParams();
  u.set("vista", q.vista);
  u.set("ref", q.ref);
  if (q.servicio?.trim()) u.set("servicio", q.servicio.trim());
  return `?${u.toString()}`;
}

export function buildTurnosHref(q: {
  vista: VistaAgenda;
  ref: string;
  servicio?: string;
}): string {
  return `/admin/turnos${buildTurnosQuery(q)}#${TURNOS_ASIGNADOS_ANCHOR}`;
}
