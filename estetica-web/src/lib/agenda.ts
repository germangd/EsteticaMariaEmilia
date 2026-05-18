import { randomInt } from "node:crypto";
import { DateTime } from "luxon";

export function getAppTimeZone(): string {
  return (
    process.env.APP_TIMEZONE?.trim() || "America/Argentina/Buenos_Aires"
  );
}

/** HH:mm con horas 1–9 rellenadas a 01–09 (paridad con `normalizarHora` en `Código.gs`). */
export function normalizarHora(hora: string): string {
  if (!hora) return "";
  const t = hora.toString().trim();
  const m2 = t.match(/^(\d{2}):(\d{2})/);
  if (m2) return `${m2[1]}:${m2[2]}`;
  const m1 = t.match(/^(\d{1,2}):(\d{2})$/);
  if (m1) return `${m1[1].padStart(2, "0")}:${m1[2]}`;
  return t;
}

/** Minutos desde medianoche (0–1439) a partir de HH:mm. */
export function horaAMinutos(horaStr: string): number {
  const h = normalizarHora(horaStr);
  const m = h.match(/^(\d{2}):(\d{2})$/);
  if (!m) return -1;
  return parseInt(m[1]!, 10) * 60 + parseInt(m[2]!, 10);
}

/** True si los intervalos [inicio, inicio+duración) se solapan. */
export function intervalosSeSolapan(
  inicioA: number,
  duracionA: number,
  inicioB: number,
  duracionB: number
): boolean {
  if (duracionA <= 0 || duracionB <= 0) return false;
  const finA = inicioA + duracionA;
  const finB = inicioB + duracionB;
  return inicioA < finB && inicioB < finA;
}

export type TurnoOcupado = { hora: string; duracionMin: number };

/** Cuántos turnos activos solapan el intervalo candidato. */
export function contarSolapamiento(
  inicioCandidato: number,
  duracionCandidato: number,
  ocupados: TurnoOcupado[]
): number {
  let n = 0;
  for (const t of ocupados) {
    const inicio = horaAMinutos(t.hora);
    if (inicio < 0) continue;
    if (
      intervalosSeSolapan(
        inicioCandidato,
        duracionCandidato,
        inicio,
        Math.max(1, t.duracionMin)
      )
    ) {
      n++;
    }
  }
  return n;
}

/**
 * Horarios de inicio posibles: desde `inicio` hasta `fin`, separados por `pasoMin`
 * (duración del servicio). Solo incluye inicios donde el servicio termina antes o al
 * cerrar el día (ej. 9:30 + 150 min → próximo 12:00).
 */
export function generarHorarios(
  inicioStr: string,
  finStr: string,
  pasoMin = 30
): string[] {
  try {
    const inicioN = normalizarHora(inicioStr);
    const finN = normalizarHora(finStr);
    if (!inicioN || !finN) return [];
    const paso = Math.max(5, Math.round(pasoMin));
    const base = DateTime.fromObject(
      { year: 2000, month: 1, day: 1 },
      { zone: "utc" }
    );
    const [hi, mi] = inicioN.split(":").map((x) => parseInt(x, 10));
    const [hf, mf] = finN.split(":").map((x) => parseInt(x, 10));
    let actual = base.set({ hour: hi, minute: mi });
    const finDt = base.set({ hour: hf, minute: mf });
    const horarios: string[] = [];
    while (actual.plus({ minutes: paso }) <= finDt) {
      horarios.push(actual.toFormat("HH:mm"));
      actual = actual.plus({ minutes: paso });
    }
    return horarios;
  } catch {
    return [];
  }
}

/** El turno que empieza en `hora` con `duracionMin` debe caber antes de `horarioFin`. */
export function turnoCabeEnHorario(
  hora: string,
  duracionMin: number,
  horarioFin: string
): boolean {
  const inicio = horaAMinutos(hora);
  const finNegocio = horaAMinutos(horarioFin);
  if (inicio < 0 || finNegocio < 0) return false;
  return inicio + Math.max(1, duracionMin) <= finNegocio;
}

export function esFechaHoraValida(
  fechaStr: string,
  horaStr: string,
  tz: string
): boolean {
  if (!fechaStr || !horaStr) return false;
  const now = DateTime.now().setZone(tz);
  const hoyStr = now.toISODate();
  if (!hoyStr) return false;
  if (fechaStr < hoyStr) return false;
  const fecha = DateTime.fromISO(fechaStr, { zone: tz });
  if (!fecha.isValid) return false;
  if (fecha.weekday === 7) return false;
  if (fechaStr === hoyStr) {
    const horaActual = now.toFormat("HH:mm");
    if (horaStr < horaActual) return false;
  }
  return true;
}

export function hoyIsoEnZona(tz: string): string {
  return DateTime.now().setZone(tz).toISODate() ?? "";
}

export function horaActualEnZona(tz: string): string {
  return DateTime.now().setZone(tz).toFormat("HH:mm");
}

const CODIGO_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

export function generarCodigo(longitud = 6): string {
  let codigo = "";
  for (let i = 0; i < longitud; i++) {
    codigo += CODIGO_CHARS[randomInt(CODIGO_CHARS.length)]!;
  }
  return codigo;
}
