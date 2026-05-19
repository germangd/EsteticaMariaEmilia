/** Igual que en `Código.gs`: horas tipo 9:00 → 09:00 si hace falta. */
export function padHoraHHmm(hora: string): string {
  const t = hora.trim();
  if (/^\d{1,2}:\d{2}$/.test(t) && !/^\d{2}:/.test(t)) {
    return `0${t}`;
  }
  return t;
}

/** Quita duplicados por `nombre` (sin distinguir mayúsculas); conserva el primero. */
export function dedupeServiciosPorNombre<T extends { nombre: string }>(
  list: T[]
): T[] {
  const seen = new Set<string>();
  const out: T[] = [];
  for (const item of list) {
    const k = item.nombre.trim().toLowerCase();
    if (!k || seen.has(k)) continue;
    seen.add(k);
    out.push(item);
  }
  return out;
}

/** Lista de nombres únicos (mismo criterio que `dedupeServiciosPorNombre`). */
export function dedupeNombresServicio(list: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const n of list) {
    const k = n.trim().toLowerCase();
    if (!k || seen.has(k)) continue;
    seen.add(k);
    out.push(n.trim());
  }
  return out;
}

/** Formato JSON compatible con `obtenerServicios()` del Apps Script. */
export function rowToServicioApi(row: {
  nombre: string;
  duracionMin: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
  precioPesos?: number;
  parentId?: number | null;
  esGrupo?: boolean;
  anticipoRequerido?: boolean;
  anticipoPorcentaje?: number;
}) {
  const esGrupo = Boolean(row.esGrupo);
  const anticipoRequerido = !esGrupo && Boolean(row.anticipoRequerido);
  const anticipoPorcentaje = anticipoRequerido
    ? Math.min(100, Math.max(1, Math.round(row.anticipoPorcentaje ?? 0)))
    : 0;
  return {
    nombre: row.nombre.trim(),
    duracion: row.duracionMin,
    responsable: row.responsable.trim(),
    capacidad: row.capacidad,
    horarioInicio: padHoraHHmm(row.horarioInicio || "09:00"),
    horarioFin: padHoraHHmm(row.horarioFin || "18:00"),
    precioPesos: Math.max(0, Math.round(row.precioPesos ?? 0)),
    parentId: row.parentId ?? null,
    esGrupo,
    anticipoRequerido,
    anticipoPorcentaje,
  };
}
