"use client";

import { useEffect, useMemo, useState } from "react";
import type { ServicioAdmin } from "@/components/admin/admin-servicios-manager";
import { PrecioInlineEditor } from "@/components/admin/precio-inline-editor";
import {
  calcularAnticipoPesos,
  etiquetaAnticipo,
} from "@/lib/servicio-anticipo";
import { fmtPesos } from "@/lib/fmt-pesos";
import { agruparServiciosCatalogoAdmin } from "@/lib/servicio-tree";
import { uiTableHead, uiTableWrap } from "@/lib/ui-classes";

type ExpandedId = number | "sueltos" | null;

function CatalogoTabla({
  filas,
  idsParecidos,
  pending,
  onEdit,
  onDelete,
  onSavePrecio,
}: {
  filas: ServicioAdmin[];
  idsParecidos: Set<number>;
  pending: boolean;
  onEdit: (s: ServicioAdmin) => void;
  onDelete: (id: number, nombre: string) => void;
  onSavePrecio: (id: number, precio: number) => Promise<boolean>;
}) {
  if (filas.length === 0) {
    return (
      <p className="px-4 py-3 text-sm text-ink-muted">
        Sin sub-servicios en esta categoría.
      </p>
    );
  }

  return (
    <div className="overflow-x-auto border-t border-gold/15">
      <table className="min-w-[720px] w-full text-left text-sm">
        <thead className={uiTableHead}>
          <tr>
            <th className="px-3 py-2.5 pl-4">Servicio</th>
            <th className="px-3 py-2.5">Duración</th>
            <th className="px-3 py-2.5">Precio</th>
            <th className="px-3 py-2.5">Anticipo</th>
            <th className="px-3 py-2.5">Cupo</th>
            <th className="px-3 py-2.5">Responsable</th>
            <th className="px-3 py-2.5">Horario</th>
            <th className="px-3 py-2.5 pr-4 text-right">Acciones</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gold/10 bg-white/70">
          {filas.map((s) => (
            <tr
              key={s.id}
              className={`hover:bg-cream/60 ${
                idsParecidos.has(s.id)
                  ? "bg-amber-50/80 ring-1 ring-inset ring-amber-300/50"
                  : ""
              }`}
            >
              <td className="px-3 py-2.5 pl-4 font-medium text-ink-dark">
                {s.nombre}
                {idsParecidos.has(s.id) ? (
                  <span
                    className="ml-2 block text-[0.65rem] font-normal text-amber-900"
                    title="Hay otro sub-servicio con nombre muy parecido en la misma categoría."
                  >
                    Posible duplicado
                  </span>
                ) : null}
              </td>
              <td className="px-3 py-2.5 text-ink-muted">{s.duracion} min</td>
              <td className="px-3 py-2.5 text-ink-muted">
                <PrecioInlineEditor
                  value={s.precioPesos ?? 0}
                  disabled={pending}
                  onSave={(precio) => onSavePrecio(s.id, precio)}
                />
              </td>
              <td className="px-3 py-2.5 text-ink-muted">
                {s.anticipoRequerido ? (
                  <span
                    title={
                      etiquetaAnticipo(
                        s.precioPesos,
                        s.anticipoRequerido,
                        s.anticipoPorcentaje
                      ) ?? undefined
                    }
                  >
                    {s.anticipoPorcentaje}%
                    {(s.precioPesos ?? 0) > 0
                      ? ` (${fmtPesos(
                          calcularAnticipoPesos(
                            s.precioPesos,
                            true,
                            s.anticipoPorcentaje
                          )
                        )})`
                      : ""}
                  </span>
                ) : (
                  "\u2014"
                )}
              </td>
              <td className="px-3 py-2.5 text-ink-muted">{s.capacidad}</td>
              <td className="px-3 py-2.5 text-ink-muted">{s.responsable}</td>
              <td className="whitespace-nowrap px-3 py-2.5 text-ink-muted">
                {s.horarioInicio} – {s.horarioFin}
              </td>
              <td className="space-x-2 px-3 py-2 pr-4 text-right whitespace-nowrap">
                <button
                  type="button"
                  disabled={pending}
                  onClick={() => onEdit(s)}
                  className="text-[0.65rem] font-medium uppercase tracking-wide text-gold-dark underline"
                >
                  Editar
                </button>
                <button
                  type="button"
                  disabled={pending}
                  onClick={() => void onDelete(s.id, s.nombre)}
                  className="text-[0.65rem] font-medium uppercase tracking-wide text-red-800 underline"
                >
                  Eliminar
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function AdminServiciosCatalogo({
  list,
  idsParecidos,
  editingId,
  pending,
  onEdit,
  onDelete,
  onSavePrecio,
}: {
  list: ServicioAdmin[];
  idsParecidos: Set<number>;
  editingId: number | null;
  pending: boolean;
  onEdit: (s: ServicioAdmin) => void;
  onDelete: (id: number, nombre: string) => void;
  onSavePrecio: (id: number, precio: number) => Promise<boolean>;
}) {
  const { grupos, sueltos } = useMemo(
    () => agruparServiciosCatalogoAdmin(list),
    [list]
  );

  const [expandedId, setExpandedId] = useState<ExpandedId>(null);

  useEffect(() => {
    if (editingId == null) return;
    const s = list.find((x) => x.id === editingId);
    if (!s) return;
    if (s.esGrupo) {
      setExpandedId(s.id);
      return;
    }
    if (s.parentId != null) {
      setExpandedId(s.parentId);
      return;
    }
    if (sueltos.some((x) => x.id === s.id)) {
      setExpandedId("sueltos");
    }
  }, [editingId, list, sueltos]);

  const toggle = (id: ExpandedId) => {
    setExpandedId((prev) => (prev === id ? null : id));
  };

  if (list.length === 0) {
    return (
      <p className="text-sm font-medium text-ink">
        No hay servicios. Agregá el primero arriba; aparecerán en la web de
        reservas.
      </p>
    );
  }

  return (
    <div className="space-y-2">
      <p className="text-xs text-ink-muted">
        Tocá una categoría para ver y editar sus sub-servicios.
      </p>
      <div className={`${uiTableWrap} overflow-hidden divide-y divide-gold/15`}>
        {grupos.map(({ grupo, hijos }) => {
          const isOpen = expandedId === grupo.id;
          const parecidosEnGrupo = hijos.filter((h) => idsParecidos.has(h.id)).length;
          return (
            <div key={grupo.id}>
              <div
                className={`flex flex-wrap items-center justify-between gap-2 px-3 py-2.5 ${
                  isOpen ? "bg-cream/90" : "bg-white/50 hover:bg-cream/50"
                }`}
              >
                <button
                  type="button"
                  onClick={() => toggle(grupo.id)}
                  aria-expanded={isOpen}
                  className="flex min-w-0 flex-1 items-center justify-between gap-2 text-left"
                >
                  <span className="text-xs font-semibold uppercase tracking-wide text-gold-dark">
                    {grupo.nombre}
                  </span>
                  <span className="shrink-0 text-xs text-ink-muted">
                    {parecidosEnGrupo > 0 && !isOpen
                      ? `${parecidosEnGrupo} posible duplicado · `
                      : ""}
                    {hijos.length} sub-servicio{hijos.length === 1 ? "" : "s"}
                    <span className="ml-1.5 inline-block w-4 text-center">
                      {isOpen ? "▴" : "▾"}
                    </span>
                  </span>
                </button>
                <span className="flex shrink-0 gap-3">
                  <button
                    type="button"
                    disabled={pending}
                    onClick={() => onEdit(grupo)}
                    className="text-[0.65rem] font-medium uppercase tracking-wide text-gold-dark underline"
                  >
                    Editar categoría
                  </button>
                  <button
                    type="button"
                    disabled={pending}
                    onClick={() => void onDelete(grupo.id, grupo.nombre)}
                    className="text-[0.65rem] font-medium uppercase tracking-wide text-red-800 underline"
                  >
                    Eliminar
                  </button>
                </span>
              </div>
              {isOpen ? (
                <CatalogoTabla
                  filas={hijos}
                  idsParecidos={idsParecidos}
                  pending={pending}
                  onEdit={onEdit}
                  onDelete={onDelete}
                  onSavePrecio={onSavePrecio}
                />
              ) : null}
            </div>
          );
        })}
        {sueltos.length > 0 ? (
          <div>
            <button
              type="button"
              onClick={() => toggle("sueltos")}
              aria-expanded={expandedId === "sueltos"}
              className={`flex w-full items-center justify-between gap-2 px-3 py-2.5 text-left text-sm transition-colors ${
                expandedId === "sueltos"
                  ? "bg-cream/90 font-medium text-ink-dark"
                  : "bg-white/50 hover:bg-cream/50"
              }`}
            >
              <span className="text-xs font-semibold uppercase tracking-wide text-ink-muted">
                Servicios sueltos
              </span>
              <span className="shrink-0 text-xs text-ink-muted">
                {sueltos.length}
                <span className="ml-1.5 inline-block w-4 text-center">
                  {expandedId === "sueltos" ? "▴" : "▾"}
                </span>
              </span>
            </button>
            {expandedId === "sueltos" ? (
              <CatalogoTabla
                filas={sueltos}
                idsParecidos={idsParecidos}
                pending={pending}
                onEdit={onEdit}
                onDelete={onDelete}
                onSavePrecio={onSavePrecio}
              />
            ) : null}
          </div>
        ) : null}
      </div>
    </div>
  );
}
