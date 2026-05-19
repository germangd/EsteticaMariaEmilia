"use client";

import { useCallback, useEffect, useState } from "react";
import type { VentaDetalle, VentaHistorial } from "@/lib/caja-repo";
import { METODOS_PAGO } from "@/lib/caja-repo";
import { AdminCajaAuditoriaPanel } from "@/components/admin/admin-caja-auditoria-panel";
import { urlExportHistorialCaja, urlTicketVenta } from "@/lib/caja-url";
import { fmtPesos } from "@/lib/fmt-pesos";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiSelect,
  uiTableHead,
  uiTableWrap,
} from "@/lib/ui-classes";

const METODO_LABEL: Record<string, string> = {
  efectivo: "Efectivo",
  transferencia: "Transferencia",
  debito: "Débito",
  credito: "Crédito",
  otro: "Otro",
};

function hoyIso(): string {
  return new Date().toISOString().slice(0, 10);
}

function hace30DiasIso(): string {
  const d = new Date();
  d.setDate(d.getDate() - 30);
  return d.toISOString().slice(0, 10);
}

function fmtHora(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString("es-AR", {
    dateStyle: "short",
    timeStyle: "short",
  });
}

type LineaEdit = {
  key: string;
  descripcion: string;
  cantidad: number;
  precioUnitarioPesos: number;
};

export function AdminCajaHistorial({
  onVentaChanged,
}: {
  onVentaChanged?: () => void;
}) {
  const [desde, setDesde] = useState(hace30DiasIso());
  const [hasta, setHasta] = useState(hoyIso());
  const [cliente, setCliente] = useState("");
  const [ventas, setVentas] = useState<VentaHistorial[]>([]);
  const [total, setTotal] = useState(0);
  const [pending, setPending] = useState(false);
  const [msg, setMsg] = useState<string | null>(null);

  const [editId, setEditId] = useState<number | null>(null);
  const [editVenta, setEditVenta] = useState<VentaDetalle | null>(null);
  const [editNombre, setEditNombre] = useState("");
  const [editTel, setEditTel] = useState("");
  const [editDescuento, setEditDescuento] = useState("0");
  const [editMetodo, setEditMetodo] = useState("efectivo");
  const [editNotas, setEditNotas] = useState("");
  const [editMotivo, setEditMotivo] = useState("");
  const [editLineas, setEditLineas] = useState<LineaEdit[]>([]);

  const [auditVenta, setAuditVenta] = useState<{
    id: number;
    numeroTicket: number;
  } | null>(null);

  const buscar = useCallback(async () => {
    setPending(true);
    setMsg(null);
    try {
      const q = new URLSearchParams();
      if (desde) q.set("desde", desde);
      if (hasta) q.set("hasta", hasta);
      if (cliente.trim()) q.set("cliente", cliente.trim());
      q.set("limite", "80");
      const r = await fetch(`/api/admin/caja/historial?${q}`, {
        credentials: "same-origin",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        ventas?: VentaHistorial[];
        total?: number;
        mensaje?: string;
      };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo cargar el historial.");
        return;
      }
      setVentas(data.ventas ?? []);
      setTotal(data.total ?? 0);
    } finally {
      setPending(false);
    }
  }, [desde, hasta, cliente]);

  useEffect(() => {
    void buscar();
  }, [buscar]);

  async function abrirEdicion(id: number) {
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/caja/ventas/${id}`, {
        credentials: "same-origin",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        venta?: VentaDetalle;
        mensaje?: string;
      };
      if (!r.ok || !data.ok || !data.venta) {
        setMsg(data.mensaje ?? "No se pudo cargar el comprobante.");
        return;
      }
      const v = data.venta;
      setEditId(id);
      setEditVenta(v);
      setEditNombre(v.clienteNombre ?? "");
      setEditTel(v.clienteTelefono ?? "");
      setEditDescuento(String(v.descuentoPesos));
      setEditMetodo(v.metodoPago);
      setEditNotas(v.notas ?? "");
      setEditMotivo("");
      setEditLineas(
        v.lineas.map((l) => ({
          key: String(l.id),
          descripcion: l.descripcion,
          cantidad: l.cantidad,
          precioUnitarioPesos: l.precioUnitarioPesos,
        }))
      );
    } finally {
      setPending(false);
    }
  }

  function cerrarEdicion() {
    setEditId(null);
    setEditVenta(null);
  }

  async function guardarEdicion(e: React.FormEvent) {
    e.preventDefault();
    if (!editId) return;
    if (!editMotivo.trim()) {
      setMsg("Indicá el motivo de la modificación (queda en el registro).");
      return;
    }
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/caja/ventas/${editId}`, {
        method: "PATCH",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          clienteNombre: editNombre,
          clienteTelefono: editTel,
          descuentoPesos: Number(editDescuento) || 0,
          metodoPago: editMetodo,
          notas: editNotas,
          motivo: editMotivo,
          lineas: editLineas.map((l) => ({
            tipo: "otro",
            descripcion: l.descripcion,
            cantidad: l.cantidad,
            precioUnitarioPesos: l.precioUnitarioPesos,
          })),
        }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo modificar.");
        return;
      }
      setMsg(`Ticket #${editVenta?.numeroTicket} actualizado.`);
      cerrarEdicion();
      await buscar();
      onVentaChanged?.();
    } finally {
      setPending(false);
    }
  }

  async function anular(id: number, numeroTicket: number) {
    const motivo = window.prompt(
      `Motivo de anulación del ticket #${numeroTicket}:`,
      ""
    );
    if (motivo === null) return;
    if (!motivo.trim()) {
      setMsg("La anulación requiere un motivo.");
      return;
    }
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/caja/ventas/${id}`, {
        method: "DELETE",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ motivo }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo anular.");
        return;
      }
      setMsg(`Ticket #${numeroTicket} anulado.`);
      await buscar();
      onVentaChanged?.();
    } finally {
      setPending(false);
    }
  }

  const inputClass = uiInput;

  return (
    <section className={uiCard}>
      <h2 className="mb-1 font-serif text-lg text-ink-dark">
        Historial de ventas
      </h2>
      <p className="mb-4 text-sm text-ink-muted">
        Tickets con numeración secuencial. Consultá la auditoría para ver el
        detalle de cada cambio o exportá el listado filtrado a CSV.
      </p>

      <form
        onSubmit={(e) => {
          e.preventDefault();
          void buscar();
        }}
        className="mb-6 grid gap-3 sm:grid-cols-2 lg:grid-cols-4"
      >
        <div>
          <label className={uiLabel}>Desde</label>
          <input
            type="date"
            className={inputClass}
            value={desde}
            onChange={(e) => setDesde(e.target.value)}
          />
        </div>
        <div>
          <label className={uiLabel}>Hasta</label>
          <input
            type="date"
            className={inputClass}
            value={hasta}
            onChange={(e) => setHasta(e.target.value)}
          />
        </div>
        <div className="sm:col-span-2">
          <label className={uiLabel}>Cliente (nombre o teléfono)</label>
          <input
            className={inputClass}
            placeholder="Buscar…"
            value={cliente}
            onChange={(e) => setCliente(e.target.value)}
          />
        </div>
        <div className="flex items-end gap-2 sm:col-span-2 lg:col-span-4">
          <button type="submit" disabled={pending} className={uiBtnPrimary}>
            Buscar
          </button>
          <button
            type="button"
            disabled={pending}
            className={uiBtnSecondary}
            onClick={() => {
              setDesde("");
              setHasta("");
              setCliente("");
            }}
          >
            Limpiar filtros
          </button>
          <a
            href={urlExportHistorialCaja({
              desde: desde || undefined,
              hasta: hasta || undefined,
              cliente: cliente.trim() || undefined,
            })}
            className={uiBtnSecondary}
            download
          >
            Exportar CSV
          </a>
          <span className="text-sm text-ink-muted">
            {total} resultado{total === 1 ? "" : "s"}
          </span>
        </div>
      </form>

      {msg ? (
        <p className="mb-4 rounded-sm border border-gold/40 bg-cream px-3 py-2 text-sm text-ink-dark">
          {msg}
        </p>
      ) : null}

      {ventas.length === 0 ? (
        <p className="text-sm text-ink-muted">Sin ventas en este criterio.</p>
      ) : (
        <div className={uiTableWrap}>
          <table className="w-full min-w-[800px] text-sm">
            <thead>
              <tr className={uiTableHead}>
                <th className="px-3 py-2 text-left">Ticket</th>
                <th className="px-3 py-2 text-left">Fecha</th>
                <th className="px-3 py-2 text-left">Cliente</th>
                <th className="px-3 py-2 text-left">Pago</th>
                <th className="px-3 py-2 text-right">Total</th>
                <th className="px-3 py-2 text-left">Estado</th>
                <th className="px-3 py-2 text-right">Acciones</th>
              </tr>
            </thead>
            <tbody>
              {ventas.map((v) => (
                <tr
                  key={v.id}
                  className={`border-t border-gold/15 ${
                    v.estado === "anulada" ? "opacity-50" : ""
                  }`}
                >
                  <td className="px-3 py-2 font-mono font-semibold tabular-nums">
                    #{String(v.numeroTicket).padStart(6, "0")}
                  </td>
                  <td className="px-3 py-2 whitespace-nowrap">
                    {fmtHora(v.createdAt)}
                  </td>
                  <td className="px-3 py-2">
                    <div className="font-medium">
                      {v.clienteNombre ?? "—"}
                    </div>
                    {v.clienteTelefono ? (
                      <div className="text-xs text-ink-muted">
                        {v.clienteTelefono}
                      </div>
                    ) : null}
                  </td>
                  <td className="px-3 py-2 capitalize">
                    {METODO_LABEL[v.metodoPago] ?? v.metodoPago}
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {fmtPesos(v.totalPesos)}
                  </td>
                  <td className="px-3 py-2 capitalize text-xs">
                    {v.estado === "anulada" ? "Anulada" : "Vigente"}
                  </td>
                  <td className="space-x-2 px-3 py-2 text-right whitespace-nowrap">
                    <a
                      href={urlTicketVenta(v.id)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark underline"
                    >
                      Consultar
                    </a>
                    <button
                      type="button"
                      disabled={pending}
                      onClick={() =>
                        setAuditVenta({
                          id: v.id,
                          numeroTicket: v.numeroTicket,
                        })
                      }
                      className="text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark underline"
                    >
                      Auditoría
                    </button>
                    {v.estado === "completada" ? (
                      <>
                        <button
                          type="button"
                          disabled={pending}
                          onClick={() => void abrirEdicion(v.id)}
                          className="text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark underline"
                        >
                          Modificar
                        </button>
                        <button
                          type="button"
                          disabled={pending}
                          onClick={() => void anular(v.id, v.numeroTicket)}
                          className="text-[0.65rem] font-semibold uppercase tracking-wide text-red-800 underline"
                        >
                          Anular
                        </button>
                      </>
                    ) : null}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {editId && editVenta ? (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
          role="dialog"
          aria-modal="true"
        >
          <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-sm border border-gold/40 bg-surface p-6 shadow-lg">
            <h3 className="mb-1 font-serif text-lg text-ink-dark">
              Modificar ticket #{editVenta.numeroTicket}
            </h3>
            <p className="mb-4 text-xs text-ink-muted">
              El cambio queda registrado en el historial del comprobante.
            </p>
            <form onSubmit={(e) => void guardarEdicion(e)} className="space-y-4">
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <label className={uiLabel}>Cliente</label>
                  <input
                    className={inputClass}
                    value={editNombre}
                    onChange={(e) => setEditNombre(e.target.value)}
                  />
                </div>
                <div>
                  <label className={uiLabel}>Teléfono</label>
                  <input
                    className={inputClass}
                    value={editTel}
                    onChange={(e) => setEditTel(e.target.value)}
                  />
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <label className={uiLabel}>Descuento (ARS)</label>
                  <input
                    type="number"
                    min={0}
                    className={inputClass}
                    value={editDescuento}
                    onChange={(e) => setEditDescuento(e.target.value)}
                  />
                </div>
                <div>
                  <label className={uiLabel}>Forma de pago</label>
                  <select
                    className={uiSelect}
                    value={editMetodo}
                    onChange={(e) => setEditMetodo(e.target.value)}
                  >
                    {METODOS_PAGO.map((m) => (
                      <option key={m} value={m}>
                        {METODO_LABEL[m] ?? m}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              <div>
                <label className={uiLabel}>Ítems</label>
                <div className="space-y-2">
                  {editLineas.map((l, idx) => (
                    <div
                      key={l.key}
                      className="grid gap-2 rounded border border-gold/20 p-2 sm:grid-cols-[1fr_4rem_5rem]"
                    >
                      <input
                        required
                        className={inputClass}
                        value={l.descripcion}
                        onChange={(e) => {
                          const next = [...editLineas];
                          next[idx] = {
                            ...l,
                            descripcion: e.target.value,
                          };
                          setEditLineas(next);
                        }}
                      />
                      <input
                        type="number"
                        min={1}
                        className={inputClass}
                        value={l.cantidad}
                        onChange={(e) => {
                          const next = [...editLineas];
                          next[idx] = {
                            ...l,
                            cantidad: Number(e.target.value) || 1,
                          };
                          setEditLineas(next);
                        }}
                      />
                      <input
                        type="number"
                        min={0}
                        className={inputClass}
                        value={l.precioUnitarioPesos}
                        onChange={(e) => {
                          const next = [...editLineas];
                          next[idx] = {
                            ...l,
                            precioUnitarioPesos: Number(e.target.value) || 0,
                          };
                          setEditLineas(next);
                        }}
                      />
                    </div>
                  ))}
                </div>
                <button
                  type="button"
                  className="mt-2 text-xs text-gold-dark underline"
                  onClick={() =>
                    setEditLineas((prev) => [
                      ...prev,
                      {
                        key: crypto.randomUUID(),
                        descripcion: "",
                        cantidad: 1,
                        precioUnitarioPesos: 0,
                      },
                    ])
                  }
                >
                  + Agregar ítem
                </button>
              </div>
              <div>
                <label className={uiLabel}>Notas</label>
                <input
                  className={inputClass}
                  value={editNotas}
                  onChange={(e) => setEditNotas(e.target.value)}
                />
              </div>
              <div>
                <label className={uiLabel}>Motivo del cambio *</label>
                <input
                  required
                  className={inputClass}
                  placeholder="Ej. corrección de precio acordado"
                  value={editMotivo}
                  onChange={(e) => setEditMotivo(e.target.value)}
                />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  className={uiBtnSecondary}
                  onClick={cerrarEdicion}
                >
                  Cancelar
                </button>
                <button type="submit" disabled={pending} className={uiBtnPrimary}>
                  Guardar cambios
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : null}

      {auditVenta ? (
        <AdminCajaAuditoriaPanel
          ventaId={auditVenta.id}
          numeroTicket={auditVenta.numeroTicket}
          onClose={() => setAuditVenta(null)}
        />
      ) : null}

      {total > ventas.length ? (
        <p className="mt-4 text-xs text-ink-muted">
          Mostrando {ventas.length} de {total} resultados en pantalla. La
          exportación CSV incluye hasta 5000 registros del mismo filtro.
        </p>
      ) : null}
    </section>
  );
}
