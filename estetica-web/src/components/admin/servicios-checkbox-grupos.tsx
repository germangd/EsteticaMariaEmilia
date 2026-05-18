"use client";

import { fmtPesos } from "@/lib/fmt-pesos";
import {
  agruparServiciosParaUi,
  filtrarServiciosReservables,
  type ServicioJerarquia,
} from "@/lib/servicio-tree";

type ServicioCheck = ServicioJerarquia & {
  nombre: string;
  precioPesos?: number;
};

type Props = {
  servicios: ServicioCheck[];
  selectedIds: number[];
  onToggle: (id: number) => void;
  disabled?: boolean;
};

export function ServiciosCheckboxGrupos({
  servicios,
  selectedIds,
  onToggle,
  disabled,
}: Props) {
  const reservables = filtrarServiciosReservables(servicios);
  const { grupos, sueltos } = agruparServiciosParaUi(reservables);

  const renderCheck = (s: ServicioCheck) => (
    <label
      key={s.id}
      className="flex cursor-pointer items-center gap-2 rounded-sm border border-gold/25 bg-cream/60 px-3 py-2 text-sm"
    >
      <input
        type="checkbox"
        disabled={disabled}
        checked={selectedIds.includes(s.id)}
        onChange={() => onToggle(s.id)}
      />
      {s.nombre}
      {(s.precioPesos ?? 0) > 0 ? ` (${fmtPesos(s.precioPesos!)})` : ""}
    </label>
  );

  if (reservables.length === 0) {
    return (
      <p className="text-sm text-ink-muted">
        {"Cre\u00e1 categor\u00edas y sub-servicios en "}
        <a href="/admin/servicios" className="text-gold-dark underline">
          Servicios
        </a>
        .
      </p>
    );
  }

  return (
    <div className="space-y-4">
      {grupos.map((g) => (
        <div key={g.grupo.id}>
          <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-gold-dark">
            {g.grupo.nombre}
          </p>
          <div className="flex flex-wrap gap-3">{g.hijos.map(renderCheck)}</div>
        </div>
      ))}
      {sueltos.length > 0 ? (
        <div>
          {grupos.length > 0 ? (
            <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-ink-muted">
              Otros servicios
            </p>
          ) : null}
          <div className="flex flex-wrap gap-3">{sueltos.map(renderCheck)}</div>
        </div>
      ) : null}
    </div>
  );
}
