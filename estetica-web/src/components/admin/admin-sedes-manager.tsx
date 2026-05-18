"use client";

import { useCallback, useEffect, useState } from "react";
import type { SedeRow } from "@/db/schema";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiTableHead,
  uiTableWrap,
} from "@/lib/ui-classes";

export function AdminSedesManager({ initialSedes }: { initialSedes: SedeRow[] }) {
  const [sedes, setSedes] = useState(initialSedes);
  const [nombre, setNombre] = useState("");
  const [orden, setOrden] = useState(0);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  useEffect(() => {
    setSedes(initialSedes);
  }, [initialSedes]);

  const resetForm = useCallback(() => {
    setNombre("");
    setOrden(0);
    setEditingId(null);
  }, []);

  async function refresh() {
    const r = await fetch("/api/admin/sedes", { credentials: "same-origin" });
    const data = (await r.json()) as { ok?: boolean; sedes?: SedeRow[] };
    if (data.ok && data.sedes) setSedes(data.sedes);
  }

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const url = editingId ? `/api/admin/sedes/${editingId}` : "/api/admin/sedes";
      const r = await fetch(url, {
        method: editingId ? "PATCH" : "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ nombre, orden, activo: true }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg("No se pudo guardar.");
        return;
      }
      setMsg(editingId ? "Sede actualizada." : "Sede creada.");
      resetForm();
      await refresh();
    } finally {
      setPending(false);
    }
  }

  function startEdit(s: SedeRow) {
    setEditingId(s.id);
    setNombre(s.nombre);
    setOrden(s.orden);
    setMsg(null);
  }

  async function desactivar(id: number, nombreSede: string) {
    if (!window.confirm(`¿Desactivar la sede "${nombreSede}"?`)) return;
    setPending(true);
    try {
      const r = await fetch(`/api/admin/sedes/${id}`, {
        method: "DELETE",
        credentials: "same-origin",
      });
      if (r.ok) {
        await refresh();
        if (editingId === id) resetForm();
      }
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="space-y-10">
      <section className={uiCard}>
        <h2 className="mb-4 font-serif text-lg text-ink-dark">
          {editingId ? "Editar sede" : "Nueva sede"}
        </h2>
        <form onSubmit={(e) => void onSubmit(e)} className="grid gap-4 md:grid-cols-2">
          <div className="md:col-span-2">
            <label className={uiLabel}>Nombre</label>
            <input
              required
              className={uiInput}
              value={nombre}
              onChange={(e) => setNombre(e.target.value)}
              placeholder="Ej. Ensenada"
            />
          </div>
          <div>
            <label className={uiLabel}>Orden</label>
            <input
              type="number"
              min={0}
              className={uiInput}
              value={orden}
              onChange={(e) => setOrden(Number(e.target.value) || 0)}
            />
          </div>
          <div className="flex flex-wrap gap-2 md:col-span-2">
            <button type="submit" disabled={pending} className={uiBtnPrimary}>
              {pending ? "Guardando\u2026" : editingId ? "Actualizar" : "Crear"}
            </button>
            {editingId ? (
              <button
                type="button"
                onClick={resetForm}
                className={uiBtnSecondary}
                disabled={pending}
              >
                Cancelar
              </button>
            ) : null}
          </div>
        </form>
        {msg ? <p className="mt-3 text-sm font-medium text-ink">{msg}</p> : null}
      </section>

      <section>
        <h2 className="mb-4 font-serif text-xl text-ink-dark">
          Sedes ({sedes.filter((s) => s.activo).length} activas)
        </h2>
        <div className={uiTableWrap}>
          <table className="min-w-[480px] w-full text-left text-sm">
            <thead className={uiTableHead}>
              <tr>
                <th className="px-3 py-3 pl-4">Nombre</th>
                <th className="px-3 py-3">Orden</th>
                <th className="px-3 py-3">Estado</th>
                <th className="px-3 py-3 pr-4 text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gold/10">
              {sedes.map((s) => (
                <tr key={s.id} className="hover:bg-cream/60">
                  <td className="px-3 py-2.5 pl-4 font-medium">{s.nombre}</td>
                  <td className="px-3 py-2.5">{s.orden}</td>
                  <td className="px-3 py-2.5">
                    {s.activo ? "Activa" : "Inactiva"}
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
                    {s.activo ? (
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => void desactivar(s.id, s.nombre)}
                        className="text-[0.65rem] font-medium uppercase tracking-wide text-red-800 underline"
                      >
                        Desactivar
                      </button>
                    ) : null}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}
