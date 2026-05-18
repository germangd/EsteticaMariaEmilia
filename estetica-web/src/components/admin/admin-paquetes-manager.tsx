"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import type { ServicioAdmin } from "@/components/admin/admin-servicios-manager";

export type PaqueteAdmin = {
  id: number;
  nombre: string;
  descripcion: string | null;
  precioPesos: number;
  sesionesTotal: number;
  activo: boolean;
  servicios: { id: number; nombre: string }[];
};

export type AsignacionAdmin = {
  id: number;
  packageId: number;
  paqueteNombre: string;
  nombreCliente: string;
  telefono: string;
  sesionesIniciales: number;
  sesionesRestantes: number;
  precioCobradoPesos: number | null;
  notas: string | null;
  estado: string;
  fechaCompra: string;
};

function fmtPesos(n: number): string {
  return n.toLocaleString("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  });
}

function hoyIso(): string {
  return new Date().toISOString().slice(0, 10);
}

const emptyPaquete = () => ({
  nombre: "",
  descripcion: "",
  precioPesos: 0,
  sesionesTotal: 3,
  activo: true,
  serviceIds: [] as number[],
});

export function AdminPaquetesManager({
  initialPaquetes,
  initialAsignaciones,
  servicios,
}: {
  initialPaquetes: PaqueteAdmin[];
  initialAsignaciones: AsignacionAdmin[];
  servicios: ServicioAdmin[];
}) {
  const [paquetes, setPaquetes] = useState(initialPaquetes);
  const [asignaciones, setAsignaciones] = useState(initialAsignaciones);
  const [form, setForm] = useState(emptyPaquete());
  const [editingId, setEditingId] = useState<number | null>(null);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const [asigPackageId, setAsigPackageId] = useState(
    String(initialPaquetes[0]?.id ?? "")
  );
  const [asigNombre, setAsigNombre] = useState("");
  const [asigTel, setAsigTel] = useState("");
  const [asigFecha, setAsigFecha] = useState(hoyIso());
  const [asigPrecio, setAsigPrecio] = useState("");
  const [asigNotas, setAsigNotas] = useState("");

  const paqueteSel = useMemo(
    () => paquetes.find((p) => String(p.id) === asigPackageId),
    [paquetes, asigPackageId]
  );

  useEffect(() => {
    setPaquetes(initialPaquetes);
    setAsignaciones(initialAsignaciones);
  }, [initialPaquetes, initialAsignaciones]);

  useEffect(() => {
    if (paqueteSel && !asigPrecio) {
      setAsigPrecio(String(paqueteSel.precioPesos));
    }
  }, [paqueteSel, asigPrecio]);

  const inputClass =
    "w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink";

  const resetForm = useCallback(() => {
    setForm(emptyPaquete());
    setEditingId(null);
  }, []);

  async function refreshPaquetes() {
    const r = await fetch("/api/admin/paquetes", { credentials: "same-origin" });
    const data = (await r.json()) as { ok?: boolean; paquetes?: PaqueteAdmin[] };
    if (data.ok && data.paquetes) setPaquetes(data.paquetes);
  }

  async function refreshAsignaciones() {
    const r = await fetch("/api/admin/paquetes/asignaciones", {
      credentials: "same-origin",
    });
    const data = (await r.json()) as {
      ok?: boolean;
      asignaciones?: AsignacionAdmin[];
    };
    if (data.ok && data.asignaciones) setAsignaciones(data.asignaciones);
  }

  function toggleService(id: number) {
    setForm((f) => ({
      ...f,
      serviceIds: f.serviceIds.includes(id)
        ? f.serviceIds.filter((x) => x !== id)
        : [...f.serviceIds, id],
    }));
  }

  async function onSubmitPaquete(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const url = editingId
        ? `/api/admin/paquetes/${editingId}`
        : "/api/admin/paquetes";
      const r = await fetch(url, {
        method: editingId ? "PATCH" : "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo guardar el paquete.");
        return;
      }
      setMsg(editingId ? "Paquete actualizado." : "Paquete creado.");
      resetForm();
      await refreshPaquetes();
    } finally {
      setPending(false);
    }
  }

  function startEdit(p: PaqueteAdmin) {
    setEditingId(p.id);
    setForm({
      nombre: p.nombre,
      descripcion: p.descripcion ?? "",
      precioPesos: p.precioPesos,
      sesionesTotal: p.sesionesTotal,
      activo: p.activo,
      serviceIds: p.servicios.map((s) => s.id),
    });
    setMsg(null);
  }

  async function onDeletePaquete(id: number, nombre: string) {
    if (!window.confirm(`¿Eliminar el paquete "${nombre}"?`)) return;
    setPending(true);
    try {
      const r = await fetch(`/api/admin/paquetes/${id}`, {
        method: "DELETE",
        credentials: "same-origin",
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo eliminar.");
        return;
      }
      setMsg("Paquete eliminado.");
      if (editingId === id) resetForm();
      await refreshPaquetes();
    } finally {
      setPending(false);
    }
  }

  async function onAsignar(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch("/api/admin/paquetes/asignaciones", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          packageId: Number(asigPackageId),
          nombreCliente: asigNombre,
          telefono: asigTel,
          fechaCompra: asigFecha,
          precioCobradoPesos: Number(asigPrecio) || undefined,
          notas: asigNotas || null,
        }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo asignar.");
        return;
      }
      setMsg("Paquete asignado al cliente (control de cobro registrado).");
      setAsigNombre("");
      setAsigTel("");
      setAsigNotas("");
      await refreshAsignaciones();
    } finally {
      setPending(false);
    }
  }

  async function consumirSesion(id: number) {
    setPending(true);
    try {
      const r = await fetch(`/api/admin/paquetes/asignaciones/${id}/consumir`, {
        method: "POST",
        credentials: "same-origin",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        sesionesRestantes?: number;
      };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo descontar sesión.");
        return;
      }
      setMsg(
        data.sesionesRestantes === 0
          ? "Última sesión usada. Paquete agotado."
          : `Sesión descontada. Quedan ${data.sesionesRestantes}.`
      );
      await refreshAsignaciones();
    } finally {
      setPending(false);
    }
  }

  async function cancelarAsignacion(id: number) {
    if (!window.confirm("¿Cancelar esta asignación de paquete?")) return;
    setPending(true);
    try {
      const r = await fetch(`/api/admin/paquetes/asignaciones/${id}/cancelar`, {
        method: "POST",
        credentials: "same-origin",
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo cancelar.");
        return;
      }
      setMsg("Asignación cancelada.");
      await refreshAsignaciones();
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="space-y-12">
      <section className="rounded-sm border border-gold/20 bg-white p-5 shadow-sm md:p-6">
        <h2 className="mb-1 font-serif text-lg font-normal text-ink-dark">
          {editingId ? "Editar paquete" : "Nuevo paquete"}
        </h2>
        <p className="mb-4 text-sm text-ink-muted">
          Definí precio de referencia, cantidad de sesiones y qué servicios
          incluye. La reserva online sigue siendo por servicio suelto.
        </p>
        <form onSubmit={(e) => void onSubmitPaquete(e)} className="grid gap-4 md:grid-cols-2">
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
              Precio (ARS)
            </label>
            <input
              type="number"
              min={0}
              required
              className={inputClass}
              value={form.precioPesos}
              onChange={(e) =>
                setForm({ ...form, precioPesos: Number(e.target.value) })
              }
            />
          </div>
          <div>
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Sesiones incluidas
            </label>
            <input
              type="number"
              min={1}
              required
              className={inputClass}
              value={form.sesionesTotal}
              onChange={(e) =>
                setForm({ ...form, sesionesTotal: Number(e.target.value) })
              }
            />
          </div>
          <div className="md:col-span-2">
            <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Descripción (opcional)
            </label>
            <input
              className={inputClass}
              value={form.descripcion}
              onChange={(e) => setForm({ ...form, descripcion: e.target.value })}
            />
          </div>
          <div className="md:col-span-2">
            <p className="mb-2 text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Servicios del paquete
            </p>
            <div className="flex flex-wrap gap-3">
              {servicios.map((s) => (
                <label
                  key={s.id}
                  className="flex cursor-pointer items-center gap-2 rounded-sm border border-gold/25 bg-cream/60 px-3 py-2 text-sm"
                >
                  <input
                    type="checkbox"
                    checked={form.serviceIds.includes(s.id)}
                    onChange={() => toggleService(s.id)}
                  />
                  {s.nombre}
                </label>
              ))}
            </div>
          </div>
          <label className="flex items-center gap-2 text-sm text-ink-muted md:col-span-2">
            <input
              type="checkbox"
              checked={form.activo}
              onChange={(e) => setForm({ ...form, activo: e.target.checked })}
            />
            Activo (visible para nuevas ventas)
          </label>
          <div className="flex flex-wrap gap-2 md:col-span-2">
            <button
              type="submit"
              disabled={pending}
              className="rounded-sm bg-gold px-5 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-white hover:bg-gold-dark disabled:opacity-50"
            >
              {pending ? "Guardando…" : editingId ? "Actualizar" : "Crear paquete"}
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
      </section>

      <section>
        <h2 className="mb-4 font-serif text-xl font-normal text-ink-dark">
          Catálogo ({paquetes.length})
        </h2>
        {paquetes.length === 0 ? (
          <p className="text-sm text-ink-muted">No hay paquetes cargados.</p>
        ) : (
          <div className="overflow-x-auto rounded-sm border border-gold/25 bg-white shadow-sm">
            <table className="min-w-[720px] w-full text-left text-sm">
              <thead className="border-b border-gold/20 bg-cream text-[0.65rem] font-medium uppercase tracking-[0.12em] text-ink-muted">
                <tr>
                  <th className="px-3 py-3 pl-4">Paquete</th>
                  <th className="px-3 py-3">Precio</th>
                  <th className="px-3 py-3">Sesiones</th>
                  <th className="px-3 py-3">Servicios</th>
                  <th className="px-3 py-3 pr-4 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gold/10">
                {paquetes.map((p) => (
                  <tr key={p.id} className="hover:bg-cream/60">
                    <td className="px-3 py-2.5 pl-4 font-medium text-ink-dark">
                      {p.nombre}
                      {!p.activo ? (
                        <span className="ml-2 text-xs text-ink-muted">(inactivo)</span>
                      ) : null}
                    </td>
                    <td className="px-3 py-2.5">{fmtPesos(p.precioPesos)}</td>
                    <td className="px-3 py-2.5">{p.sesionesTotal}</td>
                    <td className="max-w-[200px] px-3 py-2.5 text-ink-muted">
                      {p.servicios.map((s) => s.nombre).join(", ") || "—"}
                    </td>
                    <td className="space-x-2 px-3 py-2 pr-4 text-right">
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => startEdit(p)}
                        className="text-[0.65rem] font-medium uppercase tracking-wide text-gold-dark underline"
                      >
                        Editar
                      </button>
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => void onDeletePaquete(p.id, p.nombre)}
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

      <section className="rounded-sm border border-gold/20 bg-white p-5 shadow-sm md:p-6">
        <h2 className="mb-1 font-serif text-lg font-normal text-ink-dark">
          Vender / asignar paquete a cliente
        </h2>
        <p className="mb-4 text-sm text-ink-muted">
          Registrá el cobro (control interno). Descontá una sesión cuando la
          clienta asista.
        </p>
        {paquetes.length === 0 ? (
          <p className="text-sm text-ink-muted">Creá un paquete primero.</p>
        ) : (
          <form onSubmit={(e) => void onAsignar(e)} className="grid gap-4 md:grid-cols-2">
            <div className="md:col-span-2">
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Paquete
              </label>
              <select
                required
                className={inputClass}
                value={asigPackageId}
                onChange={(e) => {
                  setAsigPackageId(e.target.value);
                  const p = paquetes.find((x) => String(x.id) === e.target.value);
                  if (p) setAsigPrecio(String(p.precioPesos));
                }}
              >
                {paquetes
                  .filter((p) => p.activo)
                  .map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.nombre} — {fmtPesos(p.precioPesos)} ({p.sesionesTotal}{" "}
                      ses.)
                    </option>
                  ))}
              </select>
            </div>
            <div>
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Cliente
              </label>
              <input
                required
                className={inputClass}
                value={asigNombre}
                onChange={(e) => setAsigNombre(e.target.value)}
              />
            </div>
            <div>
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Teléfono
              </label>
              <input
                required
                className={inputClass}
                value={asigTel}
                onChange={(e) => setAsigTel(e.target.value)}
              />
            </div>
            <div>
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Fecha de cobro
              </label>
              <input
                type="date"
                required
                className={inputClass}
                value={asigFecha}
                onChange={(e) => setAsigFecha(e.target.value)}
              />
            </div>
            <div>
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Monto cobrado (ARS)
              </label>
              <input
                type="number"
                min={0}
                required
                className={inputClass}
                value={asigPrecio}
                onChange={(e) => setAsigPrecio(e.target.value)}
              />
            </div>
            <div className="md:col-span-2">
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Notas (opcional)
              </label>
              <input
                className={inputClass}
                value={asigNotas}
                onChange={(e) => setAsigNotas(e.target.value)}
              />
            </div>
            <div className="md:col-span-2">
              <button
                type="submit"
                disabled={pending}
                className="rounded-sm bg-gold px-5 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-white hover:bg-gold-dark disabled:opacity-50"
              >
                Registrar venta
              </button>
            </div>
          </form>
        )}
      </section>

      <section>
        <h2 className="mb-4 font-serif text-xl font-normal text-ink-dark">
          Clientes con paquete activo ({asignaciones.length})
        </h2>
        {asignaciones.length === 0 ? (
          <p className="text-sm text-ink-muted">No hay asignaciones activas.</p>
        ) : (
          <div className="overflow-x-auto rounded-sm border border-gold/25 bg-white shadow-sm">
            <table className="min-w-[800px] w-full text-left text-sm">
              <thead className="border-b border-gold/20 bg-cream text-[0.65rem] font-medium uppercase tracking-[0.12em] text-ink-muted">
                <tr>
                  <th className="px-3 py-3 pl-4">Cliente</th>
                  <th className="px-3 py-3">Paquete</th>
                  <th className="px-3 py-3">Sesiones</th>
                  <th className="px-3 py-3">Cobrado</th>
                  <th className="px-3 py-3">Fecha</th>
                  <th className="px-3 py-3 pr-4 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gold/10">
                {asignaciones.map((a) => (
                  <tr key={a.id} className="hover:bg-cream/60">
                    <td className="px-3 py-2.5 pl-4">
                      <div className="font-medium text-ink-dark">
                        {a.nombreCliente}
                      </div>
                      <div className="text-xs text-ink-muted">{a.telefono}</div>
                    </td>
                    <td className="px-3 py-2.5">{a.paqueteNombre}</td>
                    <td className="px-3 py-2.5 font-medium">
                      {a.sesionesRestantes} / {a.sesionesIniciales}
                    </td>
                    <td className="px-3 py-2.5">
                      {a.precioCobradoPesos != null
                        ? fmtPesos(a.precioCobradoPesos)
                        : "—"}
                    </td>
                    <td className="px-3 py-2.5 text-ink-muted">{a.fechaCompra}</td>
                    <td className="space-x-2 px-3 py-2 pr-4 text-right whitespace-nowrap">
                      <button
                        type="button"
                        disabled={pending || a.sesionesRestantes <= 0}
                        onClick={() => void consumirSesion(a.id)}
                        className="text-[0.65rem] font-medium uppercase tracking-wide text-gold-dark underline disabled:opacity-40"
                      >
                        −1 sesión
                      </button>
                      <button
                        type="button"
                        disabled={pending}
                        onClick={() => void cancelarAsignacion(a.id)}
                        className="text-[0.65rem] font-medium uppercase tracking-wide text-red-800 underline"
                      >
                        Cancelar
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {msg ? <p className="text-sm text-ink-muted">{msg}</p> : null}
    </div>
  );
}
