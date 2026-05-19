/** Alcance de unicidad del nombre (categoría, sub bajo misma categoría, o suelto). */
export function mismoAlcanceNombre(
  row: { parentId: number | null; esGrupo: boolean },
  data: { parentId?: number | null; esGrupo?: boolean }
): boolean {
  const esGrupo = data.esGrupo === true;
  if (esGrupo) return row.esGrupo;
  const parentId = data.parentId ?? null;
  if (parentId != null) {
    return !row.esGrupo && row.parentId === parentId;
  }
  return !row.esGrupo && row.parentId == null;
}

export function normalizarNombreServicio(nombre: string): string {
  return nombre.trim().toLowerCase();
}

export function buscarConflictoNombre(
  todos: { id: number; nombre: string; parentId: number | null; esGrupo: boolean }[],
  data: { nombre: string; parentId?: number | null; esGrupo?: boolean },
  excludeId?: number
): { id: number; nombre: string } | null {
  const key = normalizarNombreServicio(data.nombre);
  if (!key) return null;
  for (const r of todos) {
    if (excludeId != null && r.id === excludeId) continue;
    if (normalizarNombreServicio(r.nombre) !== key) continue;
    if (!mismoAlcanceNombre(r, data)) continue;
    return { id: r.id, nombre: r.nombre.trim() };
  }
  return null;
}

function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array<number>(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i]![0] = i;
  for (let j = 0; j <= n; j++) dp[0]![j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i]![j] = Math.min(
        dp[i - 1]![j]! + 1,
        dp[i]![j - 1]! + 1,
        dp[i - 1]![j - 1]! + cost
      );
    }
  }
  return dp[m]![n]!;
}

/** Detecta posibles typos/duplicados entre sub-servicios de la misma categoría. */
export function nombresMuyParecidos(a: string, b: string): boolean {
  const na = normalizarNombreServicio(a).replace(/\s+/g, " ");
  const nb = normalizarNombreServicio(b).replace(/\s+/g, " ");
  if (!na || !nb || na === nb) return false;
  if (na.includes(nb) || nb.includes(na)) return true;
  const dist = levenshtein(na, nb);
  const maxLen = Math.max(na.length, nb.length);
  return dist <= 2 && dist / maxLen < 0.15;
}

export function idsConNombreParecido(
  list: { id: number; nombre: string; parentId: number | null; esGrupo: boolean }[]
): Set<number> {
  const subs = list.filter((s) => !s.esGrupo && s.parentId != null);
  const flagged = new Set<number>();
  for (let i = 0; i < subs.length; i++) {
    for (let j = i + 1; j < subs.length; j++) {
      const a = subs[i]!;
      const b = subs[j]!;
      if (a.parentId !== b.parentId) continue;
      if (normalizarNombreServicio(a.nombre) === normalizarNombreServicio(b.nombre)) {
        flagged.add(a.id);
        flagged.add(b.id);
        continue;
      }
      if (nombresMuyParecidos(a.nombre, b.nombre)) {
        flagged.add(a.id);
        flagged.add(b.id);
      }
    }
  }
  return flagged;
}
