import { and, asc, eq } from "drizzle-orm";
import { DateTime } from "luxon";
import { getDb } from "@/db/client";
import { businessHourSlots } from "@/db/schema";
import {
  type FranjaHoraria,
  getAppTimeZone,
  horaAMinutos,
  normalizarHora,
} from "@/lib/agenda";
import { padHoraHHmm } from "@/lib/servicio-format";

/** Lun–sáb por defecto si la tabla está vacía. */
export const FRANJAS_ATENCION_DEFAULT: FranjaHoraria[] = [
  { horarioInicio: "09:00", horarioFin: "13:00" },
  { horarioInicio: "17:00", horarioFin: "21:00" },
];

export const DIAS_SEMANA_ADMIN = [
  { diaSemana: 1, label: "Lunes" },
  { diaSemana: 2, label: "Martes" },
  { diaSemana: 3, label: "Mi\u00e9rcoles" },
  { diaSemana: 4, label: "Jueves" },
  { diaSemana: 5, label: "Viernes" },
  { diaSemana: 6, label: "S\u00e1bado" },
] as const;

export type FranjaAtencionInput = {
  sedeId: number;
  diaSemana: number;
  horarioInicio: string;
  horarioFin: string;
};

export type HorarioAtencionPorDia = {
  diaSemana: number;
  label: string;
  franjas: FranjaHoraria[];
};

function franjaValida(inicio: string, fin: string): boolean {
  const i = horaAMinutos(inicio);
  const f = horaAMinutos(fin);
  return i >= 0 && f > i;
}

export function diaSemanaDesdeFecha(fecha: string, tz: string): number | null {
  const dt = DateTime.fromISO(fecha, { zone: tz });
  if (!dt.isValid) return null;
  return dt.weekday;
}

/** Franjas activas del local para un día de la semana (1–7) en una sede. */
export async function listarFranjasPorDiaSemana(
  sedeId: number,
  diaSemana: number
): Promise<FranjaHoraria[]> {
  const db = getDb();
  if (!db || sedeId < 1) return [...FRANJAS_ATENCION_DEFAULT];

  const rows = await db
    .select()
    .from(businessHourSlots)
    .where(
      and(
        eq(businessHourSlots.sedeId, sedeId),
        eq(businessHourSlots.diaSemana, diaSemana),
        eq(businessHourSlots.activo, true)
      )
    )
    .orderBy(asc(businessHourSlots.horarioInicio));

  if (rows.length === 0) {
    const any = await db
      .select({ id: businessHourSlots.id })
      .from(businessHourSlots)
      .where(eq(businessHourSlots.sedeId, sedeId))
      .limit(1);
    if (any.length === 0) return [...FRANJAS_ATENCION_DEFAULT];
    return [];
  }

  return rows.map((r) => ({
    horarioInicio: padHoraHHmm(r.horarioInicio),
    horarioFin: padHoraHHmm(r.horarioFin),
  }));
}

/** Franjas del local para una fecha ISO en una sede (domingo → []). */
export async function listarFranjasAtencionParaFecha(
  fecha: string,
  sedeId: number
): Promise<FranjaHoraria[]> {
  const tz = getAppTimeZone();
  const dow = diaSemanaDesdeFecha(fecha, tz);
  if (dow == null) return [];
  if (dow === 7) return [];
  return listarFranjasPorDiaSemana(sedeId, dow);
}

export async function listarHorarioAtencionAdmin(
  sedeId: number
): Promise<HorarioAtencionPorDia[] | { ok: false; reason: "no_db" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (sedeId < 1) return { ok: false, reason: "no_db" };

  const rows = await db
    .select()
    .from(businessHourSlots)
    .where(
      and(
        eq(businessHourSlots.sedeId, sedeId),
        eq(businessHourSlots.activo, true)
      )
    )
    .orderBy(
      asc(businessHourSlots.diaSemana),
      asc(businessHourSlots.horarioInicio)
    );

  const hasRows = rows.length > 0;
  const porDia = new Map<number, FranjaHoraria[]>();

  for (const d of DIAS_SEMANA_ADMIN) {
    porDia.set(
      d.diaSemana,
      hasRows
        ? []
        : d.diaSemana <= 6
          ? [...FRANJAS_ATENCION_DEFAULT]
          : []
    );
  }

  for (const r of rows) {
    if (r.diaSemana < 1 || r.diaSemana > 7) continue;
    const list = porDia.get(r.diaSemana) ?? [];
    list.push({
      horarioInicio: padHoraHHmm(r.horarioInicio),
      horarioFin: padHoraHHmm(r.horarioFin),
    });
    porDia.set(r.diaSemana, list);
  }

  return DIAS_SEMANA_ADMIN.map((d) => ({
    diaSemana: d.diaSemana,
    label: d.label,
    franjas: porDia.get(d.diaSemana) ?? [],
  }));
}

export async function reemplazarHorarioAtencion(
  sedeId: number,
  franjas: FranjaAtencionInput[]
): Promise<{ ok: true } | { ok: false; reason: "no_db" | "invalido" }> {
  const db = getDb();
  if (!db) return { ok: false, reason: "no_db" };
  if (sedeId < 1) return { ok: false, reason: "invalido" };

  const limpias: FranjaAtencionInput[] = [];
  for (const f of franjas) {
    const dia = Math.round(f.diaSemana);
    if (dia < 1 || dia > 6) continue;
    const inicio = padHoraHHmm(normalizarHora(f.horarioInicio));
    const fin = padHoraHHmm(normalizarHora(f.horarioFin));
    if (!franjaValida(inicio, fin)) return { ok: false, reason: "invalido" };
    limpias.push({
      sedeId,
      diaSemana: dia,
      horarioInicio: inicio,
      horarioFin: fin,
    });
  }

  await db
    .delete(businessHourSlots)
    .where(eq(businessHourSlots.sedeId, sedeId));

  if (limpias.length > 0) {
    await db.insert(businessHourSlots).values(
      limpias.map((f) => ({
        sedeId: f.sedeId,
        diaSemana: f.diaSemana,
        horarioInicio: f.horarioInicio,
        horarioFin: f.horarioFin,
        activo: true,
      }))
    );
  }

  return { ok: true };
}
