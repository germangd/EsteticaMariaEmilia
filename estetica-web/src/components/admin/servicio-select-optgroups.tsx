import type { ServicioJerarquia } from "@/lib/servicio-tree";
import { agruparServiciosParaUi, filtrarServiciosReservables } from "@/lib/servicio-tree";

type ServicioOption = ServicioJerarquia & {
  nombre: string;
  capacidad?: number;
  categoriaNombre?: string | null;
};

type Props = {
  servicios: ServicioOption[];
  value: string;
  onChange: (value: string) => void;
  className?: string;
  required?: boolean;
  placeholder?: string;
  /** Si true, el value de cada option es el id (string); si false, el nombre. */
  valueMode?: "id" | "nombre";
  showCupo?: boolean;
};

function agruparPorCategoriaNombre<T extends ServicioOption>(list: T[]) {
  const map = new Map<string, T[]>();
  const sueltos: T[] = [];
  for (const s of list) {
    const cat = s.categoriaNombre?.trim();
    if (cat) {
      const arr = map.get(cat) ?? [];
      arr.push(s);
      map.set(cat, arr);
    } else {
      sueltos.push(s);
    }
  }
  const grupos = [...map.entries()]
    .map(([nombre, hijos]) => ({
      grupo: { id: nombre, nombre },
      hijos: hijos.sort((a, b) =>
        a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" })
      ),
    }))
    .sort((a, b) =>
      a.grupo.nombre.localeCompare(b.grupo.nombre, "es", { sensitivity: "base" })
    );
  sueltos.sort((a, b) =>
    a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" })
  );
  return { grupos, sueltos };
}

export function ServicioSelectOptgroups({
  servicios,
  value,
  onChange,
  className,
  required,
  placeholder = "Eleg\u00ed un servicio",
  valueMode = "nombre",
  showCupo = false,
}: Props) {
  const reservables = filtrarServiciosReservables(servicios);
  const usaCategoria = reservables.some((s) => s.categoriaNombre?.trim());
  const { grupos, sueltos } = usaCategoria
    ? agruparPorCategoriaNombre(reservables)
    : agruparServiciosParaUi(reservables);

  const optValue = (s: ServicioOption) =>
    valueMode === "id" ? String(s.id) : s.nombre;

  const optLabel = (s: ServicioOption) => {
    const cupo =
      showCupo && s.capacidad != null ? ` (cupo ${s.capacidad})` : "";
    return `${s.nombre}${cupo}`;
  };

  return (
    <select
      required={required}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className={className}
    >
      <option value="">{placeholder}</option>
      {grupos.map((g) => (
        <optgroup key={g.grupo.id} label={g.grupo.nombre}>
          {g.hijos.map((s) => (
            <option key={s.id} value={optValue(s)}>
              {optLabel(s)}
            </option>
          ))}
        </optgroup>
      ))}
      {sueltos.length > 0 ? (
        grupos.length > 0 ? (
          <optgroup label="Otros servicios">
            {sueltos.map((s) => (
              <option key={s.id} value={optValue(s)}>
                {optLabel(s)}
              </option>
            ))}
          </optgroup>
        ) : (
          sueltos.map((s) => (
            <option key={s.id} value={optValue(s)}>
              {optLabel(s)}
            </option>
          ))
        )
      ) : null}
    </select>
  );
}
