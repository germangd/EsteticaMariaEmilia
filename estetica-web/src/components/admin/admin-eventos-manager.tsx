"use client";

import { useCallback, useEffect, useState } from "react";
import type { ServicioAdmin } from "@/components/admin/admin-servicios-manager";
import { ServiciosCheckboxGrupos } from "@/components/admin/servicios-checkbox-grupos";
import { filtrarServiciosReservables } from "@/lib/servicio-tree";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiTableHead,
  uiTableWrap,
} from "@/lib/ui-classes";

type ClienteBusqueda = {
  telefono: string;
  nombre: string;
};

export type EventoAdmin = {
  id: number;
  nombre: string;
  descripcion: string | null;
  fecha: string;
  sedeId: number;
  sedeNombre: string;
  horarioInicio: string;
  horarioFin: string;
  precioPesos: number;
  clienteTelefono: string | null;
  clienteNombre: string | null;
  activo: boolean;
  servicios: { id: number; nombre: string }[];
};

type SedeOpt = { id: number; nombre: string };

const emptyForm = (sedeIdDefault: number) => ({
  nombre: "",
  descripcion: "",
  fecha: "",
  sedeId: sedeIdDefault,
  horarioInicio: "09:00",
  horarioFin: "18:00",
  precioPesos: 0,
  clienteTelefono: "",
  clienteNombre: "",
  activo: true,
  serviceIds: [] as number[],
});

function formatPrecio(n: number): string {
  if (!n) return "—";
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  }).format(n);
}

export function AdminEventosManager({
  initialEventos,
  servicios,
  sedes,
}: {
  initialEventos: EventoAdmin[];
  servicios: ServicioAdmin[];
  sedes: SedeOpt[];
}) {
  const sedeDefault = sedes[0]?.id ?? 0;
  const [eventos, setEventos] = useState(initialEventos);
  const [form, setForm] = useState(() => emptyForm(sedeDefault));
  const [editingId, setEditingId] = useState<number | null>(null);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);
  const [clienteQ, setClienteQ] = useState("");
  const [clienteHits, setClienteHits] = useState<ClienteBusqueda[]>([]);
  const [buscandoCliente, setBuscandoCliente] = useState(false);

  const reservables = filtrarServiciosReservables(servicios);

  useEffect(() => {
    setEventos(initialEventos);
  }, [initialEventos]);

  useEffect(() => {
    const q = clienteQ.trim();
    if (q.length < 2) {
      setClienteHits([]);
      return;
    }
    const t = window.setTimeout(() => {
      void (async () => {
        setBuscandoCliente(true);
        try {
          const r = await fetch(
            `/api/admin/clientes?q=${encodeURIComponent(q)}`,
            { credentials: "same-origin" }
          );
          const data = (await r.json()) as {
            ok?: boolean;
            clientes?: ClienteBusqueda[];
          };
          if (data.ok && Array.isArray(data.clientes)) {
            setClienteHits(
              data.clientes.map((c) => ({
                telefono: c.telefono,
                nombre: c.nombre,
              }))
            );
          }
        } finally {
          setBuscandoCliente(false);
        }
      })();
    }, 300);
    return () => window.clearTimeout(t);
  }, [clienteQ]);

  const resetForm = useCallback(() => {
    setForm(emptyForm(sedeDefault));
    setEditingId(null);
    setClienteQ("");
    setClienteHits([]);
  }, [sedeDefault]);

  async function refresh() {
    const r = await fetch("/api/admin/eventos", { credentials: "same-origin" });
    const data = (await r.json()) as { ok?: boolean; eventos?: EventoAdmin[] };
    if (data.ok && data.eventos) setEventos(data.eventos);
  }

  function toggleService(id: number) {
    setForm((f) => ({
      ...f,
      serviceIds: f.serviceIds.includes(id)
        ? f.serviceIds.filter((x) => x !== id)
        : [...f.serviceIds, id],
    }));
  }

  function seleccionarTodos() {
    setForm((f) => ({
      ...f,
      serviceIds: reservables.map((s) => s.id),
    }));
  }

  function elegirCliente(c: ClienteBusqueda) {
    setForm((f) => ({
      ...f,
      clienteTelefono: c.telefono,
      clienteNombre: c.nombre,
    }));
    setClienteQ("");
    setClienteHits([]);
  }

  function limpiarCliente() {
    setForm((f) => ({ ...f, clienteTelefono: "", clienteNombre: "" }));
    setClienteQ("");
    setClienteHits([]);
  }

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const url = editingId
        ? `/api/admin/eventos/${editingId}`
        : "/api/admin/eventos";
      const payload = {
        ...form,
        precioPesos: form.precioPesos,
        clienteTelefono: form.clienteTelefono.trim() || null,
        clienteNombre: form.clienteNombre.trim() || null,
      };
      const r = await fetch(url, {
        method: editingId ? "PATCH" : "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo guardar.");
        return;
      }
      setMsg(editingId ? "Evento actualizado." : "Evento creado.");
      resetForm();
      await refresh();
    } finally {
      setPending(false);
    }
  }

  function startEdit(ev: EventoAdmin) {
    setEditingId(ev.id);
    setForm({
      nombre: ev.nombre,
      descripcion: ev.descripcion ?? "",
      fecha: ev.fecha,
      sedeId: ev.sedeId,
      horarioInicio: ev.horarioInicio,
      horarioFin: ev.horarioFin,
      precioPesos: ev.precioPesos,
      clienteTelefono: ev.clienteTelefono ?? "",
      clienteNombre: ev.clienteNombre ?? "",
      activo: ev.activo,
      serviceIds: ev.servicios.map((s) => s.id),
    });
    setClienteQ("");
    setClienteHits([]);
    setMsg(null);
  }

  async function onDelete(id: number, nombre: string) {
    if (!window.confirm(`¿Eliminar el evento "${nombre}"?`)) return;
    setPending(true);
    try {
      const r = await fetch(`/api/admin/eventos/${id}`, {
        method: "DELETE",
        credentials: "same-origin",
      });
      const data = (await r.json()) as { ok?: boolean };
      if (!r.ok || !data.ok) {
        setMsg("No se pudo eliminar.");
        return;
      }
      if (editingId === id) resetForm();
      await refresh();
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="space-y-10">
      <section className={uiCard}>
        <h2 className="mb-1 font-serif text-lg text-ink-dark">
          {editingId ? "Editar evento" : "Nuevo evento"}
        </h2>
        <p className="mb-4 text-sm text-ink-muted">
          {
            "Pod\u00e9s programar varios eventos el mismo d\u00eda en franjas distintas (ej. 09:00\u201312:00 y 14:00\u201318:00). En cada franja solo se reservan los servicios del evento; fuera de las franjas la agenda normal sigue para el resto."
          }
        </p>
        <form onSubmit={(e) => void onSubmit(e)} className="grid gap-4 md:grid-cols-2">
          <div className="md:col-span-2">
            <label className={uiLabel}>Nombre del evento</label>
            <input
              required
              className={uiInput}
              value={form.nombre}
              onChange={(e) => setForm({ ...form, nombre: e.target.value })}
              placeholder="Ej. Jornada de belleza"
            />
          </div>
          <div>
            <label className={uiLabel}>Sede</label>
            <select
              required
              className={uiInput}
              value={form.sedeId}
              onChange={(e) =>
                setForm({ ...form, sedeId: Number(e.target.value) })
              }
            >
              {sedes.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.nombre}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className={uiLabel}>Fecha del evento</label>
            <input
              type="date"
              required
              className={uiInput}
              value={form.fecha}
              onChange={(e) => setForm({ ...form, fecha: e.target.value })}
            />
          </div>
          <div>
            <label className={uiLabel}>Precio del evento (ARS)</label>
            <input
              type="number"
              min={0}
              step={1}
              className={uiInput}
              value={form.precioPesos}
              onChange={(e) =>
                setForm({
                  ...form,
                  precioPesos: Math.max(0, Number(e.target.value) || 0),
                })
              }
            />
          </div>
          <div>
            <label className={uiLabel}>Franja desde (hora)</label>
            <input
              type="time"
              required
              className={uiInput}
              value={form.horarioInicio}
              onChange={(e) =>
                setForm({ ...form, horarioInicio: e.target.value })
              }
            />
          </div>
          <div>
            <label className={uiLabel}>Franja hasta (hora)</label>
            <input
              type="time"
              required
              className={uiInput}
              value={form.horarioFin}
              onChange={(e) => setForm({ ...form, horarioFin: e.target.value })}
            />
          </div>
          <div className="md:col-span-2">
            <label className={uiLabel}>{"Descripci\u00f3n (opcional)"}</label>
            <input
              className={uiInput}
              value={form.descripcion}
              onChange={(e) => setForm({ ...form, descripcion: e.target.value })}
            />
          </div>
          <div className="relative md:col-span-2">
            <label className={uiLabel}>Cliente del evento (opcional)</label>
            <input
              className={uiInput}
              value={clienteQ}
              onChange={(e) => setClienteQ(e.target.value)}
              placeholder={"Buscar por nombre o tel\u00e9fono\u2026"}
            />
            {buscandoCliente ? (
              <p className="mt-1 text-xs text-ink-muted">{"Buscando\u2026"}</p>
            ) : null}
            {clienteHits.length > 0 ? (
              <ul className="absolute z-10 mt-1 max-h-40 w-full overflow-auto rounded border border-gold/30 bg-white shadow-md">
                {clienteHits.map((c) => (
                  <li key={c.telefono}>
                    <button
                      type="button"
                      className="block w-full px-3 py-2 text-left text-sm hover:bg-cream"
                      onClick={() => elegirCliente(c)}
                    >
                      {c.nombre}{" "}
                      <span className="text-ink-muted">({c.telefono})</span>
                    </button>
                  </li>
                ))}
              </ul>
            ) : null}
            {form.clienteTelefono || form.clienteNombre ? (
              <div className="mt-2 flex flex-wrap items-center gap-2 text-sm">
                <span className="rounded bg-cream px-2 py-1">
                  {form.clienteNombre || "Sin nombre"} — {form.clienteTelefono}
                </span>
                <button
                  type="button"
                  className="text-xs text-gold-dark underline"
                  onClick={limpiarCliente}
                >
                  Quitar cliente
                </button>
              </div>
            ) : null}
            <div className="mt-3 grid gap-3 sm:grid-cols-2">
              <div>
                <label className={uiLabel}>Nombre (manual)</label>
                <input
                  className={uiInput}
                  value={form.clienteNombre}
                  onChange={(e) =>
                    setForm({ ...form, clienteNombre: e.target.value })
                  }
                />
              </div>
              <div>
                <label className={uiLabel}>{"Tel\u00e9fono (manual)"}</label>
                <input
                  className={uiInput}
                  value={form.clienteTelefono}
                  onChange={(e) =>
                    setForm({ ...form, clienteTelefono: e.target.value })
                  }
                />
              </div>
            </div>
          </div>
          <div className="md:col-span-2">
            <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
              <p className={uiLabel}>Servicios del evento</p>
              <button
                type="button"
                className="text-xs text-gold-dark underline"
                onClick={seleccionarTodos}
              >
                Seleccionar todos
              </button>
            </div>
            <ServiciosCheckboxGrupos
              servicios={servicios}
              selectedIds={form.serviceIds}
              onToggle={toggleService}
              disabled={pending}
            />
          </div>
          <label className="flex items-center gap-2 text-sm md:col-span-2">
            <input
              type="checkbox"
              checked={form.activo}
              onChange={(e) => setForm({ ...form, activo: e.target.checked })}
            />
            Evento activo
          </label>
          <div className="flex flex-wrap gap-2 md:col-span-2">
            <button type="submit" disabled={pending} className={uiBtnPrimary}>
              {pending ? "Guardando\u2026" : editingId ? "Actualizar" : "Crear evento"}
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
          Eventos ({eventos.length})
        </h2>
        {eventos.length === 0 ? (
          <p className="text-sm text-ink-muted">No hay eventos programados.</p>
        ) : (
          <div className={uiTableWrap}>
            <table className="min-w-[900px] w-full text-left text-sm">
              <thead className={uiTableHead}>
                <tr>
                  <th className="px-3 py-3 pl-4">Evento</th>
                  <th className="px-3 py-3">Sede</th>
                  <th className="px-3 py-3">Fecha</th>
                  <th className="px-3 py-3">Franja</th>
                  <th className="px-3 py-3">Cliente</th>
                  <th className="px-3 py-3">Precio</th>
                  <th className="px-3 py-3">Servicios</th>
                  <th className="px-3 py-3 pr-4 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gold/10">
                {eventos.map((ev) => (
                  <tr key={ev.id} className="hover:bg-cream/60">
                    <td className="px-3 py-2.5 pl-4 font-medium">
                      {ev.nombre}
                      {!ev.activo ? (
                        <span className="ml-1 text-xs text-ink-muted">
                          (inactivo)
                        </span>
                      ) : null}
                    </td>
                    <td className="px-3 py-2.5">{ev.sedeNombre}</td>
                    <td className="px-3 py-2.5">{ev.fecha}</td>
                    <td className="whitespace-nowrap px-3 py-2.5">
                      {ev.horarioInicio} – {ev.horarioFin}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">
                      {ev.clienteNombre || ev.clienteTelefono
                        ? `${ev.clienteNombre ?? ""}${ev.clienteTelefono ? ` (${ev.clienteTelefono})` : ""}`.trim()
                        : "—"}
                    </td>
                    <td className="whitespace-nowrap px-3 py-2.5">
                      {formatPrecio(ev.precioPesos)}
                    </td>
                    <td className="max-w-[200px] px-3 py-2.5 text-ink-muted">
                      {ev.servicios.map((s) => s.nombre).join(", ") || "—"}
                    </td>
                    <td className="space-x-2 px-3 py-2 pr-4 text-right">
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => startEdit(ev)}
                        className="text-[0.65rem] font-medium uppercase tracking-wide text-gold-dark underline"
                      >
                        Editar
                      </button>
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => void onDelete(ev.id, ev.nombre)}
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
