"use client";

import { useState } from "react";
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
  const { grupos, sueltos } = agruparServiciosParaUi(servicios);
  const [expandedGrupoId, setExpandedGrupoId] = useState<number | "sueltos" | null>(
    null
  );

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

  const toggleGrupo = (id: number | "sueltos") => {
    if (disabled) return;
    setExpandedGrupoId((prev) => (prev === id ? null : id));
  };

  return (
    <div className="space-y-2">
      {grupos.length > 0 ? (
        <p className="text-xs text-ink-muted">
          Tocá una categoría para ver y marcar los servicios incluidos.
        </p>
      ) : null}
      <div className="overflow-hidden rounded-sm border border-gold/25 bg-white/50 divide-y divide-gold/15">
        {grupos.map((g) => {
          const isOpen = expandedGrupoId === g.grupo.id;
          const countSelected = g.hijos.filter((s) =>
            selectedIds.includes(s.id)
          ).length;
          return (
            <div key={g.grupo.id}>
              <button
                type="button"
                disabled={disabled}
                onClick={() => toggleGrupo(g.grupo.id)}
                aria-expanded={isOpen}
                className={`flex w-full items-center justify-between gap-2 px-3 py-2.5 text-left text-sm transition-colors ${
                  isOpen
                    ? "bg-cream/90 font-medium text-ink-dark"
                    : "hover:bg-cream/50"
                }`}
              >
                <span className="text-xs font-semibold uppercase tracking-wide text-gold-dark">
                  {g.grupo.nombre}
                </span>
                <span className="shrink-0 text-xs text-ink-muted">
                  {countSelected > 0 && !isOpen
                    ? `${countSelected} marcados · `
                    : ""}
                  {g.hijos.length}
                  <span className="ml-1.5 inline-block w-4 text-center">
                    {isOpen ? "▴" : "▾"}
                  </span>
                </span>
              </button>
              {isOpen ? (
                <div className="flex flex-wrap gap-2 border-t border-gold/15 bg-white/70 px-3 py-3">
                  {g.hijos.map(renderCheck)}
                </div>
              ) : null}
            </div>
          );
        })}
        {sueltos.length > 0 ? (
          <div>
            <button
              type="button"
              disabled={disabled}
              onClick={() => toggleGrupo("sueltos")}
              aria-expanded={expandedGrupoId === "sueltos"}
              className={`flex w-full items-center justify-between gap-2 px-3 py-2.5 text-left text-sm transition-colors ${
                expandedGrupoId === "sueltos"
                  ? "bg-cream/90 font-medium text-ink-dark"
                  : "hover:bg-cream/50"
              }`}
            >
              <span className="text-xs font-semibold uppercase tracking-wide text-ink-muted">
                Otros servicios
              </span>
              <span className="shrink-0 text-xs text-ink-muted">
                {sueltos.length}
                <span className="ml-1.5 inline-block w-4 text-center">
                  {expandedGrupoId === "sueltos" ? "▴" : "▾"}
                </span>
              </span>
            </button>
            {expandedGrupoId === "sueltos" ? (
              <div className="flex flex-wrap gap-2 border-t border-gold/15 bg-white/70 px-3 py-3">
                {sueltos.map(renderCheck)}
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
    </div>
  );
}
