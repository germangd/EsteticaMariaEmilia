/** Igual que en `Código.gs`: horas tipo 9:00 → 09:00 si hace falta. */
export function padHoraHHmm(hora: string): string {
  const t = hora.trim();
  if (/^\d{1,2}:\d{2}$/.test(t) && !/^\d{2}:/.test(t)) {
    return `0${t}`;
  }
  return t;
}

/** Formato JSON compatible con `obtenerServicios()` del Apps Script. */
export function rowToServicioApi(row: {
  nombre: string;
  duracionMin: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
}) {
  return {
    nombre: row.nombre.trim(),
    duracion: row.duracionMin,
    responsable: row.responsable.trim(),
    capacidad: row.capacidad,
    horarioInicio: padHoraHHmm(row.horarioInicio || "09:00"),
    horarioFin: padHoraHHmm(row.horarioFin || "18:00"),
  };
}
