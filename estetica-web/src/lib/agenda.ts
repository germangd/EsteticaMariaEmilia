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

/** Slots cada 30 min entre inicio y fin inclusive (misma lógica que `generarHorarios`). */
export function generarHorarios(inicioStr: string, finStr: string): string[] {
  try {
    const inicioN = normalizarHora(inicioStr);
    const finN = normalizarHora(finStr);
    if (!inicioN || !finN) return [];
    const base = DateTime.fromObject(
      { year: 2000, month: 1, day: 1 },
      { zone: "utc" }
    );
    const [hi, mi] = inicioN.split(":").map((x) => parseInt(x, 10));
    const [hf, mf] = finN.split(":").map((x) => parseInt(x, 10));
    let actual = base.set({ hour: hi, minute: mi });
    const finDt = base.set({ hour: hf, minute: mf });
    const horarios: string[] = [];
    while (actual <= finDt) {
      horarios.push(actual.toFormat("HH:mm"));
      actual = actual.plus({ minutes: 30 });
    }
    return horarios;
  } catch {
    return [];
  }
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
