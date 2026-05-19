"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { AdminServicioFechas } from "@/components/admin/admin-servicio-fechas";
import { PrecioInlineEditor } from "@/components/admin/precio-inline-editor";
import { ordenarServiciosArbol } from "@/lib/servicio-tree";
import {
  calcularAnticipoPesos,
  etiquetaAnticipo,
} from "@/lib/servicio-anticipo";
import { fmtPesos } from "@/lib/fmt-pesos";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiSelect,
  uiSubsectionTitle,
  uiTableHead,
  uiTableWrap,
} from "@/lib/ui-classes";

export type ServicioAdmin = {
  id: number;
  nombre: string;
  duracion: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
  precioPesos: number;
  anticipoRequerido: boolean;
  anticipoPorcentaje: number;
  parentId: number | null;
  esGrupo: boolean;
  categoriaNombre?: string | null;
};

type TipoServicio = "grupo" | "sub" | "suelto";

type FormState = Omit<ServicioAdmin, "id" | "parentId" | "esGrupo" | "categoriaNombre"> & {
  tipo: TipoServicio;
  parentId: string;
};

function tipoDeServicio(s: Pick<ServicioAdmin, "esGrupo" | "parentId">): TipoServicio {
  if (s.esGrupo) return "grupo";
  if (s.parentId != null) return "sub";
  return "suelto";
}

const emptyForm = (): FormState => ({
  tipo: "suelto",
  parentId: "",
  nombre: "",
  duracion: 30,
  responsable: "Mar\u00eda Emilia",
  capacidad: 1,
  horarioInicio: "09:00",
  horarioFin: "18:00",
  precioPesos: 0,
  anticipoRequerido: false,
  anticipoPorcentaje: 30,
});

export function AdminServiciosManager({
  initialServicios,
}: {
  initialServicios: ServicioAdmin[];
}) {
  const [list, setList] = useState(initialServicios);
  const [form, setForm] = useState(emptyForm());
  const [editingId, setEditingId] = useState<number | null>(null);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const resetForm = useCallback(() => {
    setForm(emptyForm());
    setEditingId(null);
  }, []);

  useEffect(() => {
    setList(initialServicios);
  }, [initialServicios]);

  const categorias = useMemo(
    () => list.filter((s) => s.esGrupo).sort((a, b) => a.nombre.localeCompare(b.nombre, "es")),
    [list]
  );

  const listOrdenada = useMemo(() => ordenarServiciosArbol(list), [list]);

  const esFormGrupo = form.tipo === "grupo";

  async function refresh() {
    const r = await fetch("/api/admin/servicios", { credentials: "same-origin" });
    const data = (await r.json()) as {
      ok?: boolean;
      servicios?: ServicioAdmin[];
    };
    if (data.ok && data.servicios) setList(data.servicios);
  }

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    const payload = {
      tipo: form.tipo,
      nombre: form.nombre,
      parentId: form.tipo === "sub" ? form.parentId : null,
      duracionMin: form.duracion,
      responsable: form.responsable,
      capacidad: form.capacidad,
      horarioInicio: form.horarioInicio,
      horarioFin: form.horarioFin,
      precioPesos: form.precioPesos,
      anticipoRequerido: form.anticipoRequerido,
      anticipoPorcentaje: form.anticipoRequerido ? form.anticipoPorcentaje : 0,
    };

    try {
      const url = editingId
        ? `/api/admin/servicios/${editingId}`
        : "/api/admin/servicios";
      const method = editingId ? "PATCH" : "POST";
      const r = await fetch(url, {
        method,
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo guardar.");
        return;
      }
      setMsg(editingId ? "Servicio actualizado." : "Servicio creado.");
      resetForm();
      await refresh();
    } finally {
      setPending(false);
    }
  }

  function startEdit(s: ServicioAdmin) {
    setEditingId(s.id);
    setForm({
      tipo: tipoDeServicio(s),
      parentId: s.parentId != null ? String(s.parentId) : "",
      nombre: s.nombre,
      duracion: s.duracion,
      responsable: s.responsable,
      capacidad: s.capacidad,
      horarioInicio: s.horarioInicio,
      horarioFin: s.horarioFin,
      precioPesos: s.precioPesos ?? 0,
      anticipoRequerido: s.anticipoRequerido ?? false,
      anticipoPorcentaje: s.anticipoPorcentaje ?? 30,
    });
    setMsg(null);
  }

  async function guardarPrecio(id: number, precioPesos: number): Promise<boolean> {
    const r = await fetch(`/api/admin/servicios/${id}/precio`, {
      method: "PATCH",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ precioPesos }),
    });
    const data = (await r.json()) as { ok?: boolean };
    if (!r.ok || !data.ok) {
      setMsg("No se pudo actualizar el precio.");
      return false;
    }
    setList((prev) =>
      prev.map((s) => (s.id === id ? { ...s, precioPesos } : s))
    );
    if (editingId === id) setForm((f) => ({ ...f, precioPesos }));
    return true;
  }

  async function onDelete(id: number, nombre: string) {
    if (!window.confirm(`¿Eliminar el servicio "${nombre}"?`)) return;
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/servicios/${id}`, {
        method: "DELETE",
        credentials: "same-origin",
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo eliminar.");
        return;
      }
      setMsg("Servicio eliminado.");
      if (editingId === id) resetForm();
      await refresh();
    } finally {
      setPending(false);
    }
  }

  const inputClass = uiInput;

  return (
    <div className="space-y-10">
      <section className={uiCard}>
        <h2 className={uiSubsectionTitle}>
          {editingId ? "Editar servicio" : "Nuevo servicio"}
        </h2>
        <p className="mb-4 text-sm font-medium text-ink">
          {"Cre\u00e1 una "}
          <strong>{"categor\u00eda"}</strong>
          {" (ej. Depilaci\u00f3n) y sub-servicios (cavado, axilas). Solo los sub-servicios y sueltos se reservan online y entran en paquetes."}
        </p>
        <form onSubmit={(e) => void onSubmit(e)} className="grid gap-4 md:grid-cols-2">
          <div>
            <label className={uiLabel}>Tipo</label>
            <select
              className={uiSelect}
              value={form.tipo}
              onChange={(e) => {
                const tipo = e.target.value as TipoServicio;
                setForm((f) => ({
                  ...f,
                  tipo,
                  parentId: tipo === "sub" ? f.parentId : "",
                }));
              }}
            >
              <option value="grupo">{"Categor\u00eda (agrupa otros)"}</option>
              <option value="sub">Sub-servicio</option>
              <option value="suelto">Servicio suelto</option>
            </select>
          </div>
          {form.tipo === "sub" ? (
            <div>
              <label className={uiLabel}>{"Categor\u00eda"}</label>
              <select
                required
                className={uiSelect}
                value={form.parentId}
                onChange={(e) => setForm({ ...form, parentId: e.target.value })}
              >
                <option value="">{"Eleg\u00ed..."}</option>
                {categorias.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.nombre}
                  </option>
                ))}
              </select>
            </div>
          ) : (
            <div />
          )}
          <div className="md:col-span-2">
            <label className={uiLabel}>
              Nombre
            </label>
            <input
              required
              className={inputClass}
              value={form.nombre}
              onChange={(e) => setForm({ ...form, nombre: e.target.value })}
            />
          </div>
          {!esFormGrupo ? (
          <>
          <div>
            <label className={uiLabel}>
              Duración (min)
            </label>
            <input
              type="number"
              min={5}
              step={5}
              required
              className={inputClass}
              value={form.duracion}
              onChange={(e) =>
                setForm({ ...form, duracion: Number(e.target.value) })
              }
            />
          </div>
          <div>
            <label className={uiLabel}>
              Cupo por turno
            </label>
            <input
              type="number"
              min={1}
              required
              className={inputClass}
              value={form.capacidad}
              onChange={(e) =>
                setForm({ ...form, capacidad: Number(e.target.value) })
              }
            />
          </div>
          <div>
            <label className={uiLabel}>
              Responsable
            </label>
            <input
              className={inputClass}
              value={form.responsable}
              onChange={(e) => setForm({ ...form, responsable: e.target.value })}
            />
          </div>
          <div>
            <label className={uiLabel}>
              Horario desde
            </label>
            <input
              type="time"
              required
              className={inputClass}
              value={form.horarioInicio}
              onChange={(e) =>
                setForm({ ...form, horarioInicio: e.target.value })
              }
            />
          </div>
          <div>
            <label className={uiLabel}>
              Horario hasta
            </label>
            <input
              type="time"
              required
              className={inputClass}
              value={form.horarioFin}
              onChange={(e) => setForm({ ...form, horarioFin: e.target.value })}
            />
          </div>
          <div>
            <label className={uiLabel}>Precio sugerido (ARS)</label>
            <input
              type="number"
              min={0}
              step={1}
              className={inputClass}
              value={form.precioPesos}
              onChange={(e) =>
                setForm({ ...form, precioPesos: Number(e.target.value) || 0 })
              }
            />
          </div>
          <div className="md:col-span-2">
            <label className="flex cursor-pointer items-center gap-2 text-sm text-ink-dark">
              <input
                type="checkbox"
                checked={form.anticipoRequerido}
                onChange={(e) =>
                  setForm({
                    ...form,
                    anticipoRequerido: e.target.checked,
                  })
                }
                className="size-4 rounded border-gold/40"
              />
              Solicitar anticipo al reservar / cobrar
            </label>
          </div>
          {form.anticipoRequerido ? (
            <div>
              <label className={uiLabel}>Anticipo (% del precio)</label>
              <input
                type="number"
                min={1}
                max={100}
                step={1}
                required
                className={inputClass}
                value={form.anticipoPorcentaje}
                onChange={(e) =>
                  setForm({
                    ...form,
                    anticipoPorcentaje: Number(e.target.value) || 1,
                  })
                }
              />
              {form.precioPesos > 0 ? (
                <p className="mt-1 text-xs text-ink-muted">
                  Monto de referencia:{" "}
                  {fmtPesos(
                    calcularAnticipoPesos(
                      form.precioPesos,
                      true,
                      form.anticipoPorcentaje
                    )
                  )}{" "}
                  ({form.anticipoPorcentaje}% de {fmtPesos(form.precioPesos)})
                </p>
              ) : (
                <p className="mt-1 text-xs text-ink-muted">
                  Definí un precio sugerido para calcular el monto del anticipo
                  en caja y reservas.
                </p>
              )}
            </div>
          ) : null}
          </>
          ) : null}
          <div className="flex flex-wrap gap-2 md:col-span-2">
            <button type="submit" disabled={pending} className={uiBtnPrimary}>
              {pending ? "Guardando\u2026" : editingId ? "Actualizar" : "Agregar"}
            </button>
            {editingId ? (
              <button
                type="button"
                disabled={pending}
                onClick={resetForm}
                className={uiBtnSecondary}
              >
                Cancelar edición
              </button>
            ) : null}
          </div>
        </form>
        {editingId && form.tipo !== "grupo" ? (
          <div className="mt-8 border-t border-gold/25 pt-8">
            <h3 className={uiSubsectionTitle}>Fechas habilitadas</h3>
            <p className="mb-4 text-sm text-ink-muted">
              {
                "Si agreg\u00e1s fechas, solo esos d\u00edas aceptan reservas de este servicio. Sin fechas, rige el horario del local (Admin \u2192 Horarios) y el rango de este servicio."
              }
            </p>
            <AdminServicioFechas serviceId={editingId} esGrupo={false} />
          </div>
        ) : null}
        {msg ? (
          <p className="mt-3 text-sm font-medium text-ink">{msg}</p>
        ) : null}
      </section>

      <section>
        <h2 className="mb-4 font-serif text-xl font-semibold text-ink-dark">
          Catálogo ({list.length})
        </h2>
        {list.length === 0 ? (
          <p className="text-sm font-medium text-ink">
            No hay servicios. Agregá el primero arriba; aparecerán en la web de
            reservas.
          </p>
        ) : (
          <div className={uiTableWrap}>
            <table className="min-w-[720px] w-full text-left text-sm">
              <thead className={uiTableHead}>
                <tr>
                  <th className="px-3 py-3 pl-4">Servicio</th>
                  <th className="px-3 py-3">Tipo</th>
                  <th className="px-3 py-3">Duración</th>
                  <th className="px-3 py-3">Precio</th>
                  <th className="px-3 py-3">Anticipo</th>
                  <th className="px-3 py-3">Cupo</th>
                  <th className="px-3 py-3">Responsable</th>
                  <th className="px-3 py-3">Horario</th>
                  <th className="px-3 py-3 pr-4 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gold/10">
                {listOrdenada.map((s) => (
                  <tr key={s.id} className="hover:bg-cream/60">
                    <td
                      className={`px-3 py-2.5 font-medium text-ink-dark ${
                        s.parentId ? "pl-8" : "pl-4"
                      }`}
                    >
                      {s.nombre}
                      {s.categoriaNombre ? (
                        <span className="ml-1 text-xs font-normal text-ink-muted">
                          ({s.categoriaNombre})
                        </span>
                      ) : null}
                    </td>
                    <td className="px-3 py-2.5 text-xs uppercase text-ink-muted">
                      {s.esGrupo
                        ? "Categor\u00eda"
                        : s.parentId
                          ? "Sub"
                          : "Suelto"}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">
                      {s.esGrupo ? "\u2014" : `${s.duracion} min`}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">
                      {s.esGrupo ? (
                        "\u2014"
                      ) : (
                        <PrecioInlineEditor
                          value={s.precioPesos ?? 0}
                          disabled={pending}
                          onSave={(precio) => guardarPrecio(s.id, precio)}
                        />
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">
                      {s.esGrupo ? (
                        "\u2014"
                      ) : s.anticipoRequerido ? (
                        <span title={etiquetaAnticipo(
                          s.precioPesos,
                          s.anticipoRequerido,
                          s.anticipoPorcentaje
                        ) ?? undefined}
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
                    <td className="px-3 py-2.5 text-ink-muted">
                      {s.esGrupo ? "\u2014" : s.capacidad}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">
                      {s.esGrupo ? "\u2014" : s.responsable}
                    </td>
                    <td className="whitespace-nowrap px-3 py-2.5 text-ink-muted">
                      {s.esGrupo
                        ? "\u2014"
                        : `${s.horarioInicio} \u2013 ${s.horarioFin}`}
                    </td>
                    <td className="space-x-2 px-3 py-2 pr-4 text-right">
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => startEdit(s)}
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
        )}
      </section>
    </div>
  );
}
