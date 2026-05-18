"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import type {
  CatalogoCaja,
  PrefillCobroTurno,
  SesionCaja,
  VentaResumen,
} from "@/lib/caja-repo";
import { METODOS_PAGO } from "@/lib/caja-repo";
import { urlTicketVenta } from "@/lib/caja-url";
import { fmtPesos } from "@/lib/fmt-pesos";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiTableHead,
  uiTableWrap,
} from "@/lib/ui-classes";

type LineaForm = {
  key: string;
  tipo: "servicio" | "paquete" | "otro";
  descripcion: string;
  cantidad: number;
  precioUnitarioPesos: number;
  serviceId?: number;
  servicePackageId?: number;
};

const METODO_LABEL: Record<string, string> = {
  efectivo: "Efectivo",
  transferencia: "Transferencia",
  debito: "D\u00e9bito",
  credito: "Cr\u00e9dito",
  otro: "Otro",
};

function nuevaLinea(): LineaForm {
  return {
    key: crypto.randomUUID(),
    tipo: "otro",
    descripcion: "",
    cantidad: 1,
    precioUnitarioPesos: 0,
  };
}

function fmtHora(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString("es-AR", {
    dateStyle: "short",
    timeStyle: "short",
  });
}

export function AdminCajaManager({
  initialSesion,
  initialVentas,
  catalogo,
  initialPrefill,
}: {
  initialSesion: SesionCaja | null;
  initialVentas: VentaResumen[];
  catalogo: CatalogoCaja;
  initialPrefill?: PrefillCobroTurno | null;
}) {
  const [sesion, setSesion] = useState(initialSesion);
  const [ventas, setVentas] = useState(initialVentas);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const [aperturaMonto, setAperturaMonto] = useState("0");
  const [cierreMonto, setCierreMonto] = useState("");

  const [lineas, setLineas] = useState<LineaForm[]>([nuevaLinea()]);
  const [clienteNombre, setClienteNombre] = useState("");
  const [clienteTel, setClienteTel] = useState("");
  const [descuento, setDescuento] = useState("0");
  const [metodoPago, setMetodoPago] = useState<string>("efectivo");
  const [notasVenta, setNotasVenta] = useState("");
  const [linkAppointmentId, setLinkAppointmentId] = useState<number | undefined>(
    undefined
  );
  const [ultimaSesionCerradaId, setUltimaSesionCerradaId] = useState<number | null>(
    null
  );

  useEffect(() => {
    if (!initialPrefill) return;
    if (initialPrefill.yaCobrado) return;

    setClienteNombre(initialPrefill.clienteNombre);
    setClienteTel(initialPrefill.clienteTelefono);
    setLinkAppointmentId(initialPrefill.appointmentId);
    setLineas([
      {
        key: crypto.randomUUID(),
        tipo: "servicio",
        descripcion: `${initialPrefill.servicioNombre} (${initialPrefill.fecha} ${initialPrefill.hora})`,
        cantidad: 1,
        precioUnitarioPesos: initialPrefill.precioSugeridoPesos,
        serviceId: initialPrefill.serviceId ?? undefined,
      },
    ]);
    setMsg(
      initialPrefill.precioSugeridoPesos > 0
        ? `Cobro del turno #${initialPrefill.appointmentId}: revis\u00e1 el importe sugerido y confirm\u00e1.`
        : `Cobro del turno #${initialPrefill.appointmentId}: indic\u00e1 el importe y confirm\u00e1.`
    );
  }, [initialPrefill]);

  const subtotal = useMemo(
    () =>
      lineas.reduce(
        (s, l) => s + Math.max(1, l.cantidad) * Math.max(0, l.precioUnitarioPesos),
        0
      ),
    [lineas]
  );
  const descuentoNum = Math.min(
    subtotal,
    Math.max(0, Math.round(Number(descuento) || 0))
  );
  const total = subtotal - descuentoNum;

  const refreshSesion = useCallback(async () => {
    const r = await fetch("/api/admin/caja/sesion", {
      credentials: "same-origin",
    });
    const data = (await r.json()) as { ok?: boolean; sesion?: SesionCaja | null };
    if (data.ok) setSesion(data.sesion ?? null);
  }, []);

  const refreshVentas = useCallback(async (sessionId: number) => {
    const r = await fetch(
      `/api/admin/caja/ventas?sessionId=${sessionId}`,
      { credentials: "same-origin" }
    );
    const data = (await r.json()) as {
      ok?: boolean;
      ventas?: VentaResumen[];
    };
    if (data.ok && data.ventas) setVentas(data.ventas);
  }, []);

  async function onAbrirCaja(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch("/api/admin/caja/sesion", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          openingAmountPesos: Number(aperturaMonto) || 0,
        }),
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        sesion?: SesionCaja;
      };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo abrir la caja.");
        return;
      }
      if (data.sesion) {
        setSesion(data.sesion);
        setVentas([]);
      }
      setMsg("Caja abierta.");
    } finally {
      setPending(false);
    }
  }

  async function onCerrarCaja(e: React.FormEvent) {
    e.preventDefault();
    if (!sesion) return;
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch("/api/admin/caja/sesion", {
        method: "PATCH",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sessionId: sesion.id,
          closingAmountPesos: Number(cierreMonto) || 0,
        }),
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        sesion?: SesionCaja;
      };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo cerrar la caja.");
        return;
      }
      await refreshSesion();
      if (data.sesion?.id) {
        setUltimaSesionCerradaId(data.sesion.id);
      }
      setMsg("Caja cerrada. Pod\u00e9s descargar el reporte de cierre.");
    } finally {
      setPending(false);
    }
  }

  function updateLinea(key: string, patch: Partial<LineaForm>) {
    setLineas((prev) =>
      prev.map((l) => (l.key === key ? { ...l, ...patch } : l))
    );
  }

  function onTipoChange(key: string, tipo: LineaForm["tipo"]) {
    setLineas((prev) =>
      prev.map((l) => {
        if (l.key !== key) return l;
        return {
          ...l,
          tipo,
          descripcion: "",
          precioUnitarioPesos: 0,
          serviceId: undefined,
          servicePackageId: undefined,
        };
      })
    );
  }

  function onPickServicio(key: string, serviceId: number) {
    const s = catalogo.servicios.find((x) => x.id === serviceId);
    if (!s) return;
    updateLinea(key, {
      serviceId,
      servicePackageId: undefined,
      descripcion: s.nombre,
      precioUnitarioPesos: s.precioPesos,
    });
  }

  function onPickPaquete(key: string, packageId: number) {
    const p = catalogo.paquetes.find((x) => x.id === packageId);
    if (!p) return;
    updateLinea(key, {
      servicePackageId: packageId,
      serviceId: undefined,
      descripcion: p.nombre,
      precioUnitarioPesos: p.precioPesos,
    });
  }

  async function onRegistrarVenta(e: React.FormEvent) {
    e.preventDefault();
    if (!sesion) {
      setMsg("Abr\u00ed la caja antes de cobrar.");
      return;
    }
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch("/api/admin/caja/ventas", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sessionId: sesion.id,
          clienteNombre,
          clienteTelefono: clienteTel,
          descuentoPesos: descuentoNum,
          metodoPago,
          notas: notasVenta,
          appointmentId: linkAppointmentId,
          lineas: lineas.map((l) => ({
            tipo: l.tipo,
            descripcion: l.descripcion,
            cantidad: l.cantidad,
            precioUnitarioPesos: l.precioUnitarioPesos,
            serviceId: l.serviceId,
            servicePackageId: l.servicePackageId,
          })),
        }),
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        id?: number;
      };
      if (!r.ok || !data.ok || !data.id) {
        setMsg(data.mensaje ?? "No se pudo registrar la venta.");
        return;
      }
      await refreshVentas(sesion.id);
      await refreshSesion();
      setLineas([nuevaLinea()]);
      setClienteNombre("");
      setClienteTel("");
      setDescuento("0");
      setNotasVenta("");
      setLinkAppointmentId(undefined);
      setMsg("Venta registrada.");
      window.open(urlTicketVenta(data.id), "_blank", "noopener");
    } finally {
      setPending(false);
    }
  }

  async function onAnular(id: number) {
    if (!confirm("\u00bfAnular este comprobante?")) return;
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/caja/ventas/${id}`, {
        method: "DELETE",
        credentials: "same-origin",
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo anular.");
        return;
      }
      if (sesion) await refreshVentas(sesion.id);
      await refreshSesion();
      setMsg("Venta anulada.");
    } finally {
      setPending(false);
    }
  }

  const efectivoEnCajon = sesion?.efectivoEsperadoEnCajon ?? 0;

  return (
    <div className="space-y-10">
      {initialPrefill?.yaCobrado && initialPrefill.ventaId ? (
        <p className="rounded-sm border border-gold/50 bg-cream px-4 py-2 text-sm text-ink-dark">
          {"Este turno ya tiene cobro registrado. "}
          <a
            href={urlTicketVenta(initialPrefill.ventaId)}
            target="_blank"
            rel="noopener noreferrer"
            className="font-semibold text-gold-dark underline"
          >
            Ver ticket
          </a>
        </p>
      ) : null}
      {msg ? (
        <p className="rounded-sm border border-gold/40 bg-cream px-4 py-2 text-sm font-medium text-ink-dark">
          {msg}
        </p>
      ) : null}

      <section className={uiCard}>
        <h2 className="mb-4 font-serif text-lg text-ink-dark">Estado de caja</h2>
        {!sesion ? (
          <form onSubmit={onAbrirCaja} className="flex flex-wrap items-end gap-4">
            <div>
              <label className={uiLabel}>Monto inicial (ARS)</label>
              <input
                type="number"
                min={0}
                step={1}
                value={aperturaMonto}
                onChange={(e) => setAperturaMonto(e.target.value)}
                className={`w-40 ${uiInput}`}
              />
            </div>
            <button type="submit" disabled={pending} className={uiBtnPrimary}>
              Abrir caja
            </button>
          </form>
        ) : (
          <div className="space-y-4">
            <dl className="grid gap-2 text-sm sm:grid-cols-2 lg:grid-cols-4">
              <div>
                <dt className="text-xs uppercase text-ink-muted">Apertura</dt>
                <dd className="font-semibold">{fmtHora(sesion.openedAt)}</dd>
              </div>
              <div>
                <dt className="text-xs uppercase text-ink-muted">Fondo inicial</dt>
                <dd className="font-semibold tabular-nums">
                  {fmtPesos(sesion.openingAmountPesos)}
                </dd>
              </div>
              <div>
                <dt className="text-xs uppercase text-ink-muted">Ventas del turno</dt>
                <dd className="font-semibold tabular-nums">
                  {fmtPesos(sesion.totalVentasPesos)} ({sesion.cantidadVentas})
                </dd>
              </div>
              <div>
                <dt className="text-xs uppercase text-ink-muted">
                  {"Efectivo en caj\u00f3n"}
                </dt>
                <dd className="font-semibold tabular-nums text-gold-dark">
                  {fmtPesos(efectivoEnCajon)}
                </dd>
              </div>
            </dl>
            {sesion.arqueoPorMetodo.length > 0 ? (
              <div className="rounded-sm border border-gold/25 bg-white/50 p-4">
                <h3 className="mb-3 text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  {"Arqueo por forma de pago"}
                </h3>
                <div className={uiTableWrap}>
                  <table className="w-full text-sm">
                    <thead>
                      <tr className={uiTableHead}>
                        <th className="px-3 py-2 text-left">{"M\u00e9todo"}</th>
                        <th className="px-3 py-2 text-center">Ops.</th>
                        <th className="px-3 py-2 text-right">Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sesion.arqueoPorMetodo.map((a) => (
                        <tr
                          key={a.metodoPago}
                          className="border-t border-gold/10"
                        >
                          <td className="px-3 py-2">
                            {METODO_LABEL[a.metodoPago] ?? a.metodoPago}
                          </td>
                          <td className="px-3 py-2 text-center tabular-nums">
                            {a.cantidad}
                          </td>
                          <td className="px-3 py-2 text-right tabular-nums font-medium">
                            {fmtPesos(a.totalPesos)}
                          </td>
                        </tr>
                      ))}
                      <tr className="border-t border-gold/30 font-semibold">
                        <td className="px-3 py-2">Total ventas</td>
                        <td className="px-3 py-2 text-center tabular-nums">
                          {sesion.cantidadVentas}
                        </td>
                        <td className="px-3 py-2 text-right tabular-nums">
                          {fmtPesos(sesion.totalVentasPesos)}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                <p className="mt-2 text-xs text-ink-muted">
                  {
                    "Al cerrar, cont\u00e1 solo el efectivo f\u00edsico (fondo + cobros en efectivo). Transferencia y tarjetas no van al caj\u00f3n."
                  }
                </p>
              </div>
            ) : null}
            <form
              onSubmit={onCerrarCaja}
              className="flex flex-wrap items-end gap-4 border-t border-gold/20 pt-4"
            >
              <div>
                <label className={uiLabel}>Monto al cerrar (ARS)</label>
                <input
                  type="number"
                  min={0}
                  step={1}
                  value={cierreMonto}
                  onChange={(e) => setCierreMonto(e.target.value)}
                  placeholder={String(efectivoEnCajon)}
                  className={`w-40 ${uiInput}`}
                />
              </div>
              <button type="submit" disabled={pending} className={uiBtnSecondary}>
                Cerrar caja
              </button>
              <a
                href={`/api/admin/caja/cierre/export?sessionId=${sesion.id}`}
                className={uiBtnSecondary}
              >
                CSV cierre
              </a>
            </form>
          </div>
        )}
      </section>

      {ultimaSesionCerradaId && !sesion ? (
        <p className="text-sm font-medium text-ink">
          <a
            href={`/api/admin/caja/cierre/export?sessionId=${ultimaSesionCerradaId}`}
            className="font-semibold text-gold-dark underline"
          >
            Descargar reporte CSV
          </a>{" "}
          del turno reci\u00e9n cerrado (sesi\u00f3n #{ultimaSesionCerradaId}).
        </p>
      ) : null}

      {sesion ? (
        <>
          <section id="nueva-venta" className={`scroll-mt-6 ${uiCard}`}>
            <h2 className="mb-4 font-serif text-lg text-ink-dark">
              Nueva venta / cobro
            </h2>
            {linkAppointmentId ? (
              <p className="mb-4 text-sm font-medium text-gold-dark">
                {`Vinculado al turno #${linkAppointmentId} de la agenda.`}
              </p>
            ) : null}
            <form onSubmit={onRegistrarVenta} className="space-y-6">
              <div className="space-y-4">
                {lineas.map((l, idx) => (
                  <div
                    key={l.key}
                    className="rounded-sm border border-gold/25 bg-white/60 p-4"
                  >
                    <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
                      <span className="text-xs font-semibold uppercase tracking-wide text-ink-muted">
                        {"\u00cdtem"} {idx + 1}
                      </span>
                      {lineas.length > 1 ? (
                        <button
                          type="button"
                          className="text-xs text-red-700 underline"
                          onClick={() =>
                            setLineas((p) => p.filter((x) => x.key !== l.key))
                          }
                        >
                          Quitar
                        </button>
                      ) : null}
                    </div>
                    <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-4">
                      <div>
                        <label className={uiLabel}>Tipo</label>
                        <select
                          value={l.tipo}
                          onChange={(e) =>
                            onTipoChange(
                              l.key,
                              e.target.value as LineaForm["tipo"]
                            )
                          }
                          className={uiInput}
                        >
                          <option value="servicio">Servicio</option>
                          <option value="paquete">Paquete</option>
                          <option value="otro">Otro</option>
                        </select>
                      </div>
                      {l.tipo === "servicio" ? (
                        <div>
                          <label className={uiLabel}>Servicio</label>
                          <select
                            value={l.serviceId ?? ""}
                            onChange={(e) =>
                              onPickServicio(l.key, Number(e.target.value))
                            }
                            className={uiInput}
                          >
                            <option value="">Elegir...</option>
                            {catalogo.servicios.map((s) => (
                              <option key={s.id} value={s.id}>
                                {s.nombre}
                                {s.precioPesos > 0
                                  ? ` (${fmtPesos(s.precioPesos)})`
                                  : ""}
                              </option>
                            ))}
                          </select>
                        </div>
                      ) : l.tipo === "paquete" ? (
                        <div>
                          <label className={uiLabel}>Paquete</label>
                          <select
                            value={l.servicePackageId ?? ""}
                            onChange={(e) =>
                              onPickPaquete(l.key, Number(e.target.value))
                            }
                            className={uiInput}
                          >
                            <option value="">Elegir...</option>
                            {catalogo.paquetes.map((p) => (
                              <option key={p.id} value={p.id}>
                                {p.nombre} ({fmtPesos(p.precioPesos)})
                              </option>
                            ))}
                          </select>
                        </div>
                      ) : (
                        <div className="md:col-span-2">
                          <label className={uiLabel}>Concepto</label>
                          <input
                            value={l.descripcion}
                            onChange={(e) =>
                              updateLinea(l.key, {
                                descripcion: e.target.value,
                              })
                            }
                            className={uiInput}
                            placeholder={"Descripci\u00f3n"}
                          />
                        </div>
                      )}
                      <div>
                        <label className={uiLabel}>Cant.</label>
                        <input
                          type="number"
                          min={1}
                          value={l.cantidad}
                          onChange={(e) =>
                            updateLinea(l.key, {
                              cantidad: Number(e.target.value) || 1,
                            })
                          }
                          className={uiInput}
                        />
                      </div>
                      <div>
                        <label className={uiLabel}>Precio unit. (ARS)</label>
                        <input
                          type="number"
                          min={0}
                          value={l.precioUnitarioPesos}
                          onChange={(e) =>
                            updateLinea(l.key, {
                              precioUnitarioPesos: Number(e.target.value) || 0,
                            })
                          }
                          className={uiInput}
                        />
                      </div>
                    </div>
                    {l.tipo !== "otro" && l.descripcion ? (
                      <p className="mt-2 text-sm text-ink-muted">{l.descripcion}</p>
                    ) : null}
                  </div>
                ))}
              </div>
              <button
                type="button"
                onClick={() => setLineas((p) => [...p, nuevaLinea()])}
                className={uiBtnSecondary}
              >
                {"+ Agregar \u00edtem"}
              </button>

              <div className="grid gap-4 border-t border-gold/20 pt-4 md:grid-cols-2">
                <div>
                  <label className={uiLabel}>Cliente (opcional)</label>
                  <input
                    value={clienteNombre}
                    onChange={(e) => setClienteNombre(e.target.value)}
                    className={uiInput}
                    placeholder="Nombre"
                  />
                </div>
                <div>
                  <label className={uiLabel}>{"Tel\u00e9fono"}</label>
                  <input
                    value={clienteTel}
                    onChange={(e) => setClienteTel(e.target.value)}
                    className={uiInput}
                    placeholder="11..."
                  />
                </div>
                <div>
                  <label className={uiLabel}>Descuento (ARS)</label>
                  <input
                    type="number"
                    min={0}
                    value={descuento}
                    onChange={(e) => setDescuento(e.target.value)}
                    className={uiInput}
                  />
                </div>
                <div>
                  <label className={uiLabel}>Forma de pago</label>
                  <select
                    value={metodoPago}
                    onChange={(e) => setMetodoPago(e.target.value)}
                    className={uiInput}
                  >
                    {METODOS_PAGO.map((m) => (
                      <option key={m} value={m}>
                        {METODO_LABEL[m] ?? m}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="md:col-span-2">
                  <label className={uiLabel}>Notas</label>
                  <input
                    value={notasVenta}
                    onChange={(e) => setNotasVenta(e.target.value)}
                    className={uiInput}
                  />
                </div>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-4 border-t border-gold/20 pt-4">
                <p className="text-lg font-semibold tabular-nums">
                  Total: {fmtPesos(total)}
                  <span className="ml-2 text-sm font-normal text-ink-muted">
                    (subtotal {fmtPesos(subtotal)})
                  </span>
                </p>
                <button type="submit" disabled={pending} className={uiBtnPrimary}>
                  Cobrar e imprimir ticket
                </button>
              </div>
            </form>
          </section>

          <section>
            <h2 className="mb-4 font-serif text-lg text-ink-dark">
              Ventas de este turno
            </h2>
            {ventas.length === 0 ? (
              <p className="text-sm text-ink-muted">Sin ventas registradas.</p>
            ) : (
              <div className={uiTableWrap}>
                <table className="w-full min-w-[640px] text-sm">
                  <thead>
                    <tr className={uiTableHead}>
                      <th className="px-3 py-2 text-left">{"N\u00b0"}</th>
                      <th className="px-3 py-2 text-left">Hora</th>
                      <th className="px-3 py-2 text-left">Cliente</th>
                      <th className="px-3 py-2 text-left">Pago</th>
                      <th className="px-3 py-2 text-right">Total</th>
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
                        <td className="px-3 py-2 font-mono text-xs">
                          {v.numero}
                        </td>
                        <td className="px-3 py-2">{fmtHora(v.createdAt)}</td>
                        <td className="px-3 py-2">
                          {v.clienteNombre ?? "\u2014"}
                        </td>
                        <td className="px-3 py-2 capitalize">
                          {METODO_LABEL[v.metodoPago] ?? v.metodoPago}
                        </td>
                        <td className="px-3 py-2 text-right tabular-nums">
                          {fmtPesos(v.totalPesos)}
                        </td>
                        <td className="px-3 py-2 text-right">
                          <a
                            href={`/admin/caja/ticket/${v.id}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="mr-2 text-gold-dark underline"
                          >
                            Ticket
                          </a>
                          {v.estado === "completada" ? (
                            <button
                              type="button"
                              onClick={() => onAnular(v.id)}
                              disabled={pending}
                              className="text-red-700 underline"
                            >
                              Anular
                            </button>
                          ) : (
                            <span className="text-xs uppercase text-red-700">
                              Anulada
                            </span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </>
      ) : null}
    </div>
  );
}
