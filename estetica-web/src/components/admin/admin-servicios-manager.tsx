"use client";

import { useCallback, useEffect, useState } from "react";

export type ServicioAdmin = {
  id: number;
  nombre: string;
  duracion: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
};

const emptyForm = (): Omit<ServicioAdmin, "id"> => ({
  nombre: "",
  duracion: 30,
  responsable: "María Emilia",
  capacidad: 1,
  horarioInicio: "09:00",
  horarioFin: "18:00",
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
      nombre: form.nombre,
      duracionMin: form.duracion,
      responsable: form.responsable,
      capacidad: form.capacidad,
      horarioInicio: form.horarioInicio,
      horarioFin: form.horarioFin,
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
      nombre: s.nombre,
      duracion: s.duracion,
      responsable: s.responsable,
      capacidad: s.capacidad,
      horarioInicio: s.horarioInicio,
      horarioFin: s.horarioFin,
    });
    setMsg(null);
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

  const inputClass =
    "w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink";

  return (
    <div className="space-y-10">
      <section className="rounded-sm border border-gold/20 bg-white p-5 shadow-sm md:p-6">
        <h2 className="mb-1 font-serif text-lg font-normal text-ink-dark">
          {editingId ? "Editar servicio" : "Nuevo servicio"}
        </h2>
        <p className="mb-4 text-sm text-ink-muted">
          La <strong>capacidad</strong> es cuántos clientes pueden reservar el
          mismo servicio a la misma hora (cupo por turno).
        </p>
        <form onSubmit={(e) => void onSubmit(e)} className="grid gap-4 md:grid-cols-2">
          <div className="md:col-span-2">
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Nombre
            </label>
            <input
              required
              className={inputClass}
              value={form.nombre}
              onChange={(e) => setForm({ ...form, nombre: e.target.value })}
            />
          </div>
          <div>
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
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
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
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
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Responsable
            </label>
            <input
              className={inputClass}
              value={form.responsable}
              onChange={(e) => setForm({ ...form, responsable: e.target.value })}
            />
          </div>
          <div>
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
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
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
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
          <div className="flex flex-wrap gap-2 md:col-span-2">
            <button
              type="submit"
              disabled={pending}
              className="rounded-sm bg-gold px-5 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-white hover:bg-gold-dark disabled:opacity-50"
            >
              {pending ? "Guardando…" : editingId ? "Actualizar" : "Agregar"}
            </button>
            {editingId ? (
              <button
                type="button"
                disabled={pending}
                onClick={resetForm}
                className="rounded-sm border border-gold/40 px-5 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-gold-dark"
              >
                Cancelar edición
              </button>
            ) : null}
          </div>
        </form>
        {msg ? (
          <p className="mt-3 text-sm text-ink-muted">{msg}</p>
        ) : null}
      </section>

      <section>
        <h2 className="mb-4 font-serif text-xl font-normal text-ink-dark">
          Catálogo ({list.length})
        </h2>
        {list.length === 0 ? (
          <p className="text-sm text-ink-muted">
            No hay servicios. Agregá el primero arriba; aparecerán en la web de
            reservas.
          </p>
        ) : (
          <div className="overflow-x-auto rounded-sm border border-gold/25 bg-white shadow-sm">
            <table className="min-w-[720px] w-full text-left text-sm">
              <thead className="border-b border-gold/20 bg-cream text-[0.65rem] font-medium uppercase tracking-[0.12em] text-ink-muted">
                <tr>
                  <th className="px-3 py-3 pl-4">Servicio</th>
                  <th className="px-3 py-3">Duración</th>
                  <th className="px-3 py-3">Cupo</th>
                  <th className="px-3 py-3">Responsable</th>
                  <th className="px-3 py-3">Horario</th>
                  <th className="px-3 py-3 pr-4 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gold/10">
                {list.map((s) => (
                  <tr key={s.id} className="hover:bg-cream/60">
                    <td className="px-3 py-2.5 pl-4 font-medium text-ink-dark">
                      {s.nombre}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">{s.duracion} min</td>
                    <td className="px-3 py-2.5 text-ink-muted">{s.capacidad}</td>
                    <td className="px-3 py-2.5 text-ink-muted">{s.responsable}</td>
                    <td className="whitespace-nowrap px-3 py-2.5 text-ink-muted">
                      {s.horarioInicio} – {s.horarioFin}
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
