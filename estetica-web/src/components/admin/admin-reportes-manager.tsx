"use client";

import { useCallback, useEffect, useState } from "react";
import type { ResumenVentasPeriodo, SesionCaja } from "@/lib/caja-repo";
import { fmtPesos } from "@/lib/fmt-pesos";
import { hoyIso, rangoPreset, type AgrupacionVentas } from "@/lib/reportes-fechas";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiSelect,
  uiTabActive,
  uiTabBar,
  uiTabInactive,
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

type Tab = "ventas" | "cajas";

function fmtHora(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString("es-AR", {
    dateStyle: "short",
    timeStyle: "short",
  });
}

export function AdminReportesManager() {
  const [tab, setTab] = useState<Tab>("ventas");
  const [pending, setPending] = useState(false);
  const [msg, setMsg] = useState<string | null>(null);

  const mesActual = rangoPreset("mes");
  const [desde, setDesde] = useState(mesActual.desde);
  const [hasta, setHasta] = useState(mesActual.hasta);
  const [agrupacion, setAgrupacion] = useState<AgrupacionVentas>("dia");

  const [resumen, setResumen] = useState<ResumenVentasPeriodo | null>(null);
  const [sesiones, setSesiones] = useState<SesionCaja[]>([]);
  const [expandedSesionId, setExpandedSesionId] = useState<number | null>(null);

  const cargarVentas = useCallback(async () => {
    if (!desde || !hasta) {
      setMsg("Completá el rango de fechas.");
      return;
    }
    setPending(true);
    setMsg(null);
    try {
      const q = new URLSearchParams({
        desde,
        hasta,
        agrupacion,
      });
      const r = await fetch(`/api/admin/reportes/ventas?${q}`, {
        credentials: "same-origin",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        resumen?: ResumenVentasPeriodo;
      };
      if (!r.ok || !data.ok || !data.resumen) {
        setMsg(data.mensaje ?? "No se pudo cargar el resumen.");
        return;
      }
      setResumen(data.resumen);
    } finally {
      setPending(false);
    }
  }, [desde, hasta, agrupacion]);

  const cargarSesiones = useCallback(async () => {
    setPending(true);
    setMsg(null);
    try {
      const q = new URLSearchParams();
      if (desde) q.set("desde", desde);
      if (hasta) q.set("hasta", hasta);
      q.set("limite", "50");
      const r = await fetch(`/api/admin/reportes/sesiones?${q}`, {
        credentials: "same-origin",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        sesiones?: SesionCaja[];
      };
      if (!r.ok || !data.ok || !data.sesiones) {
        setMsg(data.mensaje ?? "No se pudo cargar el historial de cajas.");
        return;
      }
      setSesiones(data.sesiones);
    } finally {
      setPending(false);
    }
  }, [desde, hasta]);

  function aplicarPreset(preset: "hoy" | "semana" | "mes") {
    const r = rangoPreset(preset);
    setDesde(r.desde);
    setHasta(r.hasta);
  }

  function onBuscar() {
    if (tab === "ventas") void cargarVentas();
    else void cargarSesiones();
  }

  useEffect(() => {
    if (tab === "ventas") void cargarVentas();
    else void cargarSesiones();
  }, [tab, cargarVentas, cargarSesiones]);

  return (
    <div className="space-y-8">
      <div className={uiTabBar}>
        <button
          type="button"
          onClick={() => setTab("ventas")}
          className={tab === "ventas" ? uiTabActive : uiTabInactive}
        >
          Resumen de ventas
        </button>
        <button
          type="button"
          onClick={() => setTab("cajas")}
          className={tab === "cajas" ? uiTabActive : uiTabInactive}
        >
          Historial de cajas
        </button>
      </div>

      <section className={uiCard}>
        <h2 className="mb-4 font-serif text-xl font-semibold text-ink-dark">
          Filtros
        </h2>
        <div className="mb-4 flex flex-wrap gap-2">
          <button
            type="button"
            className={uiBtnSecondary}
            onClick={() => aplicarPreset("hoy")}
          >
            Hoy
          </button>
          <button
            type="button"
            className={uiBtnSecondary}
            onClick={() => aplicarPreset("semana")}
          >
            Esta semana
          </button>
          <button
            type="button"
            className={uiBtnSecondary}
            onClick={() => aplicarPreset("mes")}
          >
            Este mes
          </button>
        </div>
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <div>
            <label className={uiLabel}>Desde</label>
            <input
              type="date"
              className={uiInput}
              value={desde}
              max={hasta || hoyIso()}
              onChange={(e) => setDesde(e.target.value)}
            />
          </div>
          <div>
            <label className={uiLabel}>Hasta</label>
            <input
              type="date"
              className={uiInput}
              value={hasta}
              min={desde}
              max={hoyIso()}
              onChange={(e) => setHasta(e.target.value)}
            />
          </div>
          {tab === "ventas" ? (
            <div>
              <label className={uiLabel}>Agrupar por</label>
              <select
                className={uiSelect}
                value={agrupacion}
                onChange={(e) =>
                  setAgrupacion(e.target.value as AgrupacionVentas)
                }
              >
                <option value="dia">Día</option>
                <option value="semana">Semana</option>
                <option value="mes">Mes</option>
              </select>
            </div>
          ) : null}
          <div className="flex items-end">
            <button
              type="button"
              disabled={pending}
              className={`${uiBtnPrimary} w-full`}
              onClick={onBuscar}
            >
              {pending ? "Cargando…" : "Actualizar"}
            </button>
          </div>
        </div>
        {msg ? (
          <p className="mt-3 rounded bg-red-50 px-3 py-2 text-sm text-red-800">
            {msg}
          </p>
        ) : null}
      </section>

      {tab === "ventas" ? (
        <section className="space-y-6">
          {!resumen ? (
            <p className="text-sm text-ink-muted">
              Elegí el período y tocá Actualizar para ver el resumen de ventas
              vigentes.
            </p>
          ) : (
            <>
              <div className={`${uiCard} grid gap-4 sm:grid-cols-3`}>
                <div>
                  <p className="text-xs font-semibold uppercase text-ink-muted">
                    Total vendido
                  </p>
                  <p className="font-serif text-2xl text-ink-dark">
                    {fmtPesos(resumen.totalPesos)}
                  </p>
                  <p className="text-sm text-ink-muted">
                    {resumen.cantidad} operación
                    {resumen.cantidad === 1 ? "" : "es"}
                  </p>
                </div>
                <div>
                  <p className="text-xs font-semibold uppercase text-ink-muted">
                    Período
                  </p>
                  <p className="text-sm font-medium text-ink-dark">
                    {resumen.fechaDesde} — {resumen.fechaHasta}
                  </p>
                  <p className="text-xs text-ink-muted">
                    Agrupado por{" "}
                    {resumen.agrupacion === "dia"
                      ? "día"
                      : resumen.agrupacion === "semana"
                        ? "semana"
                        : "mes"}
                  </p>
                </div>
                {resumen.anuladas > 0 ? (
                  <div>
                    <p className="text-xs font-semibold uppercase text-ink-muted">
                      Anuladas (no suman)
                    </p>
                    <p className="text-sm font-medium text-red-800">
                      {resumen.anuladas}
                    </p>
                  </div>
                ) : null}
              </div>

              {resumen.porMetodo.length > 0 ? (
                <div className={uiCard}>
                  <h3 className="mb-3 font-serif text-lg font-semibold text-ink-dark">
                    Por forma de pago
                  </h3>
                  <div className={uiTableWrap}>
                    <table className="min-w-[400px] w-full text-left text-sm">
                      <thead className={uiTableHead}>
                        <tr>
                          <th className="px-3 py-2 pl-4">Método</th>
                          <th className="px-3 py-2">Ops.</th>
                          <th className="px-3 py-2 pr-4 text-right">Total</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gold/10">
                        {resumen.porMetodo.map((m) => (
                          <tr key={m.metodoPago}>
                            <td className="px-3 py-2 pl-4">
                              {METODO_LABEL[m.metodoPago] ?? m.metodoPago}
                            </td>
                            <td className="px-3 py-2">{m.cantidad}</td>
                            <td className="px-3 py-2 pr-4 text-right font-medium">
                              {fmtPesos(m.totalPesos)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : null}

              <div className={uiCard}>
                <h3 className="mb-3 font-serif text-lg font-semibold text-ink-dark">
                  Detalle por período
                </h3>
                {resumen.buckets.length === 0 ? (
                  <p className="text-sm text-ink-muted">
                    Sin ventas en este rango.
                  </p>
                ) : (
                  <div className={uiTableWrap}>
                    <table className="min-w-[480px] w-full text-left text-sm">
                      <thead className={uiTableHead}>
                        <tr>
                          <th className="px-3 py-2 pl-4">Período</th>
                          <th className="px-3 py-2">Ventas</th>
                          <th className="px-3 py-2 pr-4 text-right">Total</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gold/10">
                        {resumen.buckets.map((b) => (
                          <tr key={b.periodo}>
                            <td className="px-3 py-2 pl-4 font-medium">
                              {b.etiqueta}
                            </td>
                            <td className="px-3 py-2">{b.cantidad}</td>
                            <td className="px-3 py-2 pr-4 text-right font-medium">
                              {fmtPesos(b.totalPesos)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </>
          )}
        </section>
      ) : (
        <section className="space-y-4">
          <p className="text-sm text-ink-muted">
            Turnos de caja (apertura/cierre). Tocá un turno para ver el arqueo por
            método de pago.
          </p>
          {sesiones.length === 0 ? (
            <p className="text-sm text-ink-muted">
              Sin datos. Ajustá las fechas y tocá Actualizar.
            </p>
          ) : (
            <div className="overflow-hidden rounded-sm border border-gold/25 bg-white/50 divide-y divide-gold/15">
              {sesiones.map((s) => {
                const isOpen = expandedSesionId === s.id;
                return (
                  <div key={s.id}>
                    <button
                      type="button"
                      onClick={() =>
                        setExpandedSesionId(isOpen ? null : s.id)
                      }
                      aria-expanded={isOpen}
                      className={`flex w-full flex-wrap items-center justify-between gap-2 px-3 py-2.5 text-left text-sm transition-colors ${
                        isOpen
                          ? "bg-cream/90 font-medium text-ink-dark"
                          : "hover:bg-cream/50"
                      }`}
                    >
                      <span>
                        <span className="font-semibold text-gold-dark">
                          Turno #{s.id}
                        </span>
                        <span
                          className={`ml-2 text-xs uppercase ${
                            s.status === "abierta"
                              ? "text-emerald-700"
                              : "text-ink-muted"
                          }`}
                        >
                          {s.status === "abierta" ? "Abierta" : "Cerrada"}
                        </span>
                      </span>
                      <span className="shrink-0 text-xs text-ink-muted">
                        {fmtPesos(s.totalVentasPesos)} · {s.cantidadVentas}{" "}
                        ventas
                        <span className="ml-1.5 inline-block w-4 text-center">
                          {isOpen ? "▴" : "▾"}
                        </span>
                      </span>
                    </button>
                    {isOpen ? (
                      <div className="border-t border-gold/15 bg-white/70 px-4 py-3 text-sm">
                        <p className="mb-2 text-ink-muted">
                          Apertura: {fmtHora(s.openedAt)}
                          {s.closedAt
                            ? ` · Cierre: ${fmtHora(s.closedAt)}`
                            : " · En curso"}
                        </p>
                        <p className="mb-2">
                          Fondo inicial: {fmtPesos(s.openingAmountPesos)} ·
                          Efectivo esperado en cajón:{" "}
                          {fmtPesos(s.efectivoEsperadoEnCajon)}
                          {s.closingAmountPesos != null
                            ? ` · Contado al cerrar: ${fmtPesos(s.closingAmountPesos)}`
                            : ""}
                        </p>
                        {s.arqueoPorMetodo.length > 0 ? (
                          <table className="mb-3 w-full text-left text-xs">
                            <thead>
                              <tr className="border-b border-gold/20">
                                <th className="py-1">Método</th>
                                <th className="py-1">Ops.</th>
                                <th className="py-1 text-right">Total</th>
                              </tr>
                            </thead>
                            <tbody>
                              {s.arqueoPorMetodo.map((m) => (
                                <tr key={m.metodoPago}>
                                  <td className="py-1">
                                    {METODO_LABEL[m.metodoPago] ?? m.metodoPago}
                                  </td>
                                  <td className="py-1">{m.cantidad}</td>
                                  <td className="py-1 text-right">
                                    {fmtPesos(m.totalPesos)}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        ) : (
                          <p className="mb-2 text-ink-muted">
                            Sin ventas en este turno.
                          </p>
                        )}
                        <a
                          href={`/api/admin/caja/cierre/export?sessionId=${s.id}`}
                          className="text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark underline"
                        >
                          Descargar reporte CSV
                        </a>
                      </div>
                    ) : null}
                  </div>
                );
              })}
            </div>
          )}
        </section>
      )}
    </div>
  );
}
