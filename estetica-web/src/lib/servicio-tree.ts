/** Utilidades para categorías (grupos) y sub-servicios. */

export type ServicioJerarquia = {
  id: number;
  nombre: string;
  parentId: number | null;
  esGrupo: boolean;
};

export function esServicioReservable(s: { esGrupo: boolean }): boolean {
  return !s.esGrupo;
}

export function filtrarServiciosReservables<T extends { esGrupo: boolean }>(
  list: T[]
): T[] {
  return list.filter(esServicioReservable);
}

export type GrupoServiciosUi<T extends ServicioJerarquia> = {
  grupo: { id: number; nombre: string };
  hijos: T[];
};

/** Agrupa sub-servicios bajo su categoría; el resto queda como sueltos. */
export function agruparServiciosParaUi<T extends ServicioJerarquia>(
  list: T[]
): { grupos: GrupoServiciosUi<T>[]; sueltos: T[] } {
  const grupos = list.filter((s) => s.esGrupo);
  const hijosByParent = new Map<number, T[]>();
  const sueltos: T[] = [];

  for (const s of list) {
    if (s.esGrupo) continue;
    if (s.parentId != null) {
      const arr = hijosByParent.get(s.parentId) ?? [];
      arr.push(s);
      hijosByParent.set(s.parentId, arr);
    } else {
      sueltos.push(s);
    }
  }

  const sortNombre = (a: T, b: T) =>
    a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" });

  const grupoIds = new Set(grupos.map((g) => g.id));
  for (const [parentId, hijos] of hijosByParent) {
    if (!grupoIds.has(parentId)) {
      sueltos.push(...hijos);
    }
  }

  const gruposUi = grupos
    .map((g) => ({
      grupo: { id: g.id, nombre: g.nombre },
      hijos: (hijosByParent.get(g.id) ?? []).sort(sortNombre),
    }))
    .filter((g) => g.hijos.length > 0)
    .sort((a, b) =>
      a.grupo.nombre.localeCompare(b.grupo.nombre, "es", { sensitivity: "base" })
    );

  sueltos.sort(sortNombre);
  return { grupos: gruposUi, sueltos };
}

/** Catálogo admin: todas las categorías (aunque estén vacías) y sub-servicios agrupados. */
export function agruparServiciosCatalogoAdmin<T extends ServicioJerarquia>(
  list: T[]
): { grupos: { grupo: T; hijos: T[] }[]; sueltos: T[] } {
  const grupos = list.filter((s) => s.esGrupo);
  const hijosByParent = new Map<number, T[]>();
  const sueltos: T[] = [];

  for (const s of list) {
    if (s.esGrupo) continue;
    if (s.parentId != null) {
      const arr = hijosByParent.get(s.parentId) ?? [];
      arr.push(s);
      hijosByParent.set(s.parentId, arr);
    } else {
      sueltos.push(s);
    }
  }

  const sortNombre = (a: T, b: T) =>
    a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" });

  const grupoIds = new Set(grupos.map((g) => g.id));
  for (const [parentId, hijos] of hijosByParent) {
    if (!grupoIds.has(parentId)) {
      sueltos.push(...hijos);
    }
  }

  const gruposUi = grupos
    .map((g) => ({
      grupo: g,
      hijos: (hijosByParent.get(g.id) ?? []).sort(sortNombre),
    }))
    .sort((a, b) =>
      a.grupo.nombre.localeCompare(b.grupo.nombre, "es", { sensitivity: "base" })
    );

  sueltos.sort(sortNombre);
  return { grupos: gruposUi, sueltos };
}

/** Orden para tabla admin: categoría, luego sus hijos, luego sueltos. */
export function ordenarServiciosArbol<T extends ServicioJerarquia>(
  list: T[]
): T[] {
  const { grupos, sueltos } = agruparServiciosParaUi(list);
  const byId = new Map(list.map((s) => [s.id, s]));
  const out: T[] = [];
  for (const g of grupos) {
    const parent = byId.get(g.grupo.id);
    if (parent) out.push(parent);
    out.push(...g.hijos);
  }
  out.push(...sueltos);
  return out;
}

export function nombreCategoria(
  s: ServicioJerarquia,
  byId: Map<number, ServicioJerarquia>
): string | null {
  if (!s.parentId) return null;
  return byId.get(s.parentId)?.nombre ?? null;
}
