"use client";

import { useCallback, useState } from "react";
import { AdminCajaNuevaVentaForm } from "@/components/admin/admin-caja-nueva-venta-form";
import type {
  CatalogoCaja,
  PrefillCobroPaquete,
  PrefillCobroTurno,
  SesionCaja,
  VentaResumen,
} from "@/lib/caja-repo";
import { AdminCajaHistorial } from "@/components/admin/admin-caja-historial";
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

const METODO_LABEL: Record<string, string> = {
  efectivo: "Efectivo",
  transferencia: "Transferencia",
  debito: "D\u00e9bito",
  credito: "Cr\u00e9dito",
  otro: "Otro",
};

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
  initialPrefillTurno,
  initialPrefillPaquete,
}: {
  initialSesion: SesionCaja | null;
  initialVentas: VentaResumen[];
  catalogo: CatalogoCaja;
  initialPrefillTurno?: PrefillCobroTurno | null;
  initialPrefillPaquete?: PrefillCobroPaquete | null;
}) {
  const [sesion, setSesion] = useState(initialSesion);
  const [ventas, setVentas] = useState(initialVentas);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const [aperturaMonto, setAperturaMonto] = useState("0");
  const [cierreMonto, setCierreMonto] = useState("");
  const [ultimaSesionCerradaId, setUltimaSesionCerradaId] = useState<number | null>(
    null
  );

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

  async function onAnular(id: number, numeroTicket: number) {
    const motivo = window.prompt(
      `Motivo de anulaci\u00f3n del ticket #${numeroTicket}:`,
      ""
    );
    if (motivo === null) return;
    if (!motivo.trim()) {
      setMsg("La anulaci\u00f3n requiere un motivo.");
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
      {initialPrefillTurno?.yaCobrado && initialPrefillTurno.ventaId ? (
        <p className="rounded-sm border border-gold/50 bg-cream px-4 py-2 text-sm text-ink-dark">
          {"Este turno ya tiene cobro registrado. "}
          <a
            href={urlTicketVenta(initialPrefillTurno.ventaId)}
            target="_blank"
            rel="noopener noreferrer"
            className="font-semibold text-gold-dark underline"
          >
            Ver ticket
          </a>
        </p>
      ) : null}
      {initialPrefillPaquete?.yaCobrado && initialPrefillPaquete.ventaId ? (
        <p className="rounded-sm border border-gold/50 bg-cream px-4 py-2 text-sm text-ink-dark">
          {"Este paquete ya tiene cobro registrado en caja. "}
          <a
            href={urlTicketVenta(initialPrefillPaquete.ventaId)}
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
            <AdminCajaNuevaVentaForm
              sessionId={sesion.id}
              catalogo={catalogo}
              initialPrefillTurno={initialPrefillTurno}
              initialPrefillPaquete={initialPrefillPaquete}
              onMensaje={setMsg}
              onVentaRegistrada={async () => {
                await refreshVentas(sesion.id);
                await refreshSesion();
              }}
              pending={pending}
              setPending={setPending}
            />
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
                      <th className="px-3 py-2 text-left">Ticket</th>
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
                        <td className="px-3 py-2 font-mono text-xs font-semibold">
                          #{String(v.numeroTicket).padStart(6, "0")}
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
                              onClick={() => onAnular(v.id, v.numeroTicket)}
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

      <AdminCajaHistorial
        onVentaChanged={async () => {
          if (sesion) await refreshVentas(sesion.id);
          await refreshSesion();
        }}
      />
    </div>
  );
}
