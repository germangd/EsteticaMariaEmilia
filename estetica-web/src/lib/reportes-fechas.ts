import { DateTime } from "luxon";

export type AgrupacionVentas = "dia" | "semana" | "mes";

export type RangoFechas = { desde: string; hasta: string };

export function hoyIso(): string {
  return DateTime.now().toISODate() ?? "";
}

export function rangoPreset(preset: "hoy" | "semana" | "mes"): RangoFechas {
  const now = DateTime.now();
  const hasta = now.toISODate() ?? "";
  if (preset === "hoy") {
    return { desde: hasta, hasta };
  }
  if (preset === "semana") {
    const desde = now.startOf("week").toISODate() ?? hasta;
    return { desde, hasta };
  }
  const desde = now.startOf("month").toISODate() ?? hasta;
  return { desde, hasta };
}

export function etiquetaBucket(
  periodoIso: string,
  agrupacion: AgrupacionVentas
): string {
  const dt = DateTime.fromISO(periodoIso);
  if (!dt.isValid) return periodoIso;
  if (agrupacion === "dia") {
    return dt.setLocale("es").toFormat("ccc d MMM yyyy");
  }
  if (agrupacion === "semana") {
    const fin = dt.endOf("week");
    return `${dt.setLocale("es").toFormat("d MMM")} – ${fin.setLocale("es").toFormat("d MMM yyyy")}`;
  }
  return dt.setLocale("es").toFormat("MMMM yyyy");
}
