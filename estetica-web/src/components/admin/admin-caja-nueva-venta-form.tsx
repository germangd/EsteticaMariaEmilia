"use client";

import { useEffect, useMemo, useState } from "react";
import { useRouter } from "next/navigation";
import {
  ServicioPasosSelect,
  type ServicioPasosOpt,
} from "@/components/admin/servicio-pasos-select";
import type {
  CatalogoCaja,
  PrefillCobroPaquete,
  PrefillCobroTurno,
} from "@/lib/caja-repo";
import { METODOS_PAGO } from "@/lib/caja-repo";
import { urlTicketVenta } from "@/lib/caja-url";
import { fmtPesos } from "@/lib/fmt-pesos";
import { calcularAnticipoPesos } from "@/lib/servicio-anticipo";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiInput,
  uiLabel,
  uiSelect,
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
  debito: "Débito",
  credito: "Crédito",
  otro: "Otro",
};

function nuevaLinea(): LineaForm {
  return {
    key: crypto.randomUUID(),
    tipo: "servicio",
    descripcion: "",
    cantidad: 1,
    precioUnitarioPesos: 0,
  };
}

function totalLinea(l: LineaForm): number {
  return Math.max(1, l.cantidad) * Math.max(0, l.precioUnitarioPesos);
}

function etiquetaTipo(tipo: LineaForm["tipo"]): string {
  if (tipo === "servicio") return "Servicio";
  if (tipo === "paquete") return "Paquete";
  return "Otro";
}

export function AdminCajaNuevaVentaForm({
  sessionId,
  catalogo,
  initialPrefillTurno,
  initialPrefillPaquete,
  onMensaje,
  onVentaRegistrada,
  pending,
  setPending,
}: {
  sessionId: number;
  catalogo: CatalogoCaja;
  initialPrefillTurno?: PrefillCobroTurno | null;
  initialPrefillPaquete?: PrefillCobroPaquete | null;
  onMensaje: (msg: string | null) => void;
  onVentaRegistrada?: () => void | Promise<void>;
  pending: boolean;
  setPending: (v: boolean) => void;
}) {
  const router = useRouter();
  const [catalogoState, setCatalogoState] = useState(catalogo);
  const [borrador, setBorrador] = useState<LineaForm>(() => nuevaLinea());
  const [cola, setCola] = useState<LineaForm[]>([]);
  const [modalCobroAbierto, setModalCobroAbierto] = useState(false);

  const [clienteNombre, setClienteNombre] = useState("");
  const [clienteTel, setClienteTel] = useState("");
  const [descuento, setDescuento] = useState("0");
  const [metodoPago, setMetodoPago] = useState<string>("efectivo");
  const [notasVenta, setNotasVenta] = useState("");
  const [linkAppointmentId, setLinkAppointmentId] = useState<number | undefined>();
  const [linkClientPackageId, setLinkClientPackageId] = useState<
    number | undefined
  >();

  useEffect(() => {
    setCatalogoState(catalogo);
  }, [catalogo]);

  const serviciosParaPasos: ServicioPasosOpt[] = useMemo(
    () =>
      catalogoState.servicios.map((s) => ({
        id: s.id,
        nombre: s.nombre,
        categoriaNombre: s.categoriaNombre,
        precioPesos: s.precioPesos,
        anticipoRequerido: s.anticipoRequerido,
        anticipoPorcentaje: s.anticipoPorcentaje,
      })),
    [catalogoState.servicios]
  );

  const anticipoBorrador = useMemo(() => {
    if (borrador.tipo !== "servicio" || !borrador.serviceId) return null;
    const s = catalogoState.servicios.find((x) => x.id === borrador.serviceId);
    if (!s?.anticipoRequerido) return null;
    const monto = calcularAnticipoPesos(
      s.precioPesos,
      true,
      s.anticipoPorcentaje
    );
    return { porcentaje: s.anticipoPorcentaje, monto };
  }, [borrador.serviceId, borrador.tipo, catalogoState.servicios]);

  useEffect(() => {
    if (!initialPrefillTurno || initialPrefillTurno.yaCobrado) return;
    setClienteNombre(initialPrefillTurno.clienteNombre);
    setClienteTel(initialPrefillTurno.clienteTelefono);
    setLinkAppointmentId(initialPrefillTurno.appointmentId);
    setLinkClientPackageId(undefined);
    setCola([
      {
        key: crypto.randomUUID(),
        tipo: "servicio",
        descripcion: `${initialPrefillTurno.servicioNombre} (${initialPrefillTurno.fecha} ${initialPrefillTurno.hora})`,
        cantidad: 1,
        precioUnitarioPesos: initialPrefillTurno.precioSugeridoPesos,
        serviceId: initialPrefillTurno.serviceId ?? undefined,
      },
    ]);
    const ant = initialPrefillTurno.anticipoRequerido
      ? initialPrefillTurno.anticipoSugeridoPesos > 0
        ? ` Anticipo configurado: ${initialPrefillTurno.anticipoPorcentaje}% (${fmtPesos(initialPrefillTurno.anticipoSugeridoPesos)}).`
        : ` Anticipo configurado: ${initialPrefillTurno.anticipoPorcentaje}% (sin precio de referencia).`
      : "";
    onMensaje(
      (initialPrefillTurno.precioSugeridoPesos > 0
        ? `Cobro del turno #${initialPrefillTurno.appointmentId}: revisá el importe sugerido y confirmá.`
        : `Cobro del turno #${initialPrefillTurno.appointmentId}: indicá el importe y confirmá.`) + ant
    );
  }, [initialPrefillTurno, onMensaje]);

  useEffect(() => {
    if (!initialPrefillPaquete || initialPrefillPaquete.yaCobrado) return;
    setClienteNombre(initialPrefillPaquete.clienteNombre);
    setClienteTel(initialPrefillPaquete.clienteTelefono);
    setLinkAppointmentId(undefined);
    setLinkClientPackageId(initialPrefillPaquete.clientPackageId);
    setCola([
      {
        key: crypto.randomUUID(),
        tipo: "paquete",
        descripcion: `Paquete: ${initialPrefillPaquete.paqueteNombre} (desde ${initialPrefillPaquete.fechaCompra})`,
        cantidad: 1,
        precioUnitarioPesos: initialPrefillPaquete.precioSugeridoPesos,
        servicePackageId: initialPrefillPaquete.packageId,
      },
    ]);
    onMensaje(
      initialPrefillPaquete.precioSugeridoPesos > 0
        ? `Cobro del paquete asignado #${initialPrefillPaquete.clientPackageId}: revisá el importe y confirmá.`
        : `Cobro del paquete #${initialPrefillPaquete.clientPackageId}: indicá el importe.`
    );
  }, [initialPrefillPaquete, onMensaje]);

  const subtotal = useMemo(
    () => cola.reduce((s, l) => s + totalLinea(l), 0),
    [cola]
  );
  const descuentoNum = Math.min(
    subtotal,
    Math.max(0, Math.round(Number(descuento) || 0))
  );
  const total = subtotal - descuentoNum;

  const inputClass = uiInput;

  function patchBorrador(patch: Partial<LineaForm>) {
    setBorrador((b) => ({ ...b, ...patch }));
  }

  function onTipoChangeBorrador(tipo: LineaForm["tipo"]) {
    setBorrador((b) => ({
      ...b,
      tipo,
      descripcion: "",
      precioUnitarioPesos: 0,
      serviceId: undefined,
      servicePackageId: undefined,
    }));
  }

  function onPickServicioBorrador(serviceId: number) {
    const s = catalogoState.servicios.find((x) => x.id === serviceId);
    if (!s) return;
    patchBorrador({
      serviceId,
      servicePackageId: undefined,
      descripcion: s.nombre,
      precioUnitarioPesos: s.precioPesos,
    });
  }

  function onPickPaqueteBorrador(packageId: number) {
    const p = catalogoState.paquetes.find((x) => x.id === packageId);
    if (!p) return;
    patchBorrador({
      servicePackageId: packageId,
      serviceId: undefined,
      descripcion: p.nombre,
      precioUnitarioPesos: p.precioPesos,
    });
  }

  function precioCatalogoLinea(l: LineaForm): number | null {
    if (l.tipo === "servicio" && l.serviceId) {
      return (
        catalogoState.servicios.find((s) => s.id === l.serviceId)?.precioPesos ??
        null
      );
    }
    if (l.tipo === "paquete" && l.servicePackageId) {
      return (
        catalogoState.paquetes.find((p) => p.id === l.servicePackageId)
          ?.precioPesos ?? null
      );
    }
    return null;
  }

  async function guardarPrecioEnCatalogo(l: LineaForm) {
    const precio = Math.max(0, Math.round(l.precioUnitarioPesos));
    if (l.tipo === "servicio" && l.serviceId) {
      const r = await fetch(`/api/admin/servicios/${l.serviceId}/precio`, {
        method: "PATCH",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ precioPesos: precio }),
      });
      const data = (await r.json()) as { ok?: boolean };
      if (!r.ok || !data.ok) {
        onMensaje("No se pudo guardar el precio del servicio en el catálogo.");
        return;
      }
      setCatalogoState((c) => ({
        ...c,
        servicios: c.servicios.map((s) =>
          s.id === l.serviceId ? { ...s, precioPesos: precio } : s
        ),
      }));
      onMensaje(`Precio del servicio actualizado a ${fmtPesos(precio)}.`);
      return;
    }
    if (l.tipo === "paquete" && l.servicePackageId) {
      const r = await fetch(`/api/admin/paquetes/${l.servicePackageId}/precio`, {
        method: "PATCH",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ precioPesos: precio }),
      });
      const data = (await r.json()) as { ok?: boolean };
      if (!r.ok || !data.ok) {
        onMensaje("No se pudo guardar el precio del paquete en el catálogo.");
        return;
      }
      setCatalogoState((c) => ({
        ...c,
        paquetes: c.paquetes.map((p) =>
          p.id === l.servicePackageId ? { ...p, precioPesos: precio } : p
        ),
      }));
      onMensaje(`Precio del paquete actualizado a ${fmtPesos(precio)}.`);
    }
  }

  function validarBorrador(): string | null {
    if (borrador.tipo === "servicio" && !borrador.serviceId) {
      return "Elegí un servicio.";
    }
    if (borrador.tipo === "paquete" && !borrador.servicePackageId) {
      return "Elegí un paquete.";
    }
    if (borrador.tipo === "otro" && !borrador.descripcion.trim()) {
      return "Indicá el concepto del ítem.";
    }
    if (totalLinea(borrador) <= 0) {
      return "El precio unitario debe ser mayor a cero.";
    }
    return null;
  }

  function agregarItemACola() {
    const err = validarBorrador();
    if (err) {
      onMensaje(err);
      return;
    }
    setCola((prev) => [...prev, { ...borrador, key: crypto.randomUUID() }]);
    setBorrador(nuevaLinea());
    onMensaje(null);
  }

  function abrirModalCobro() {
    if (cola.length === 0) {
      onMensaje("Agregá al menos un ítem a la venta.");
      return;
    }
    setModalCobroAbierto(true);
  }

  async function onRegistrarVenta(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    onMensaje(null);
    try {
      const r = await fetch("/api/admin/caja/ventas", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sessionId,
          clienteNombre,
          clienteTelefono: clienteTel,
          descuentoPesos: descuentoNum,
          metodoPago,
          notas: notasVenta,
          appointmentId: linkAppointmentId,
          clientPackageId: linkClientPackageId,
          lineas: cola.map((l) => ({
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
        onMensaje(data.mensaje ?? "No se pudo registrar la venta.");
        return;
      }
      setBorrador(nuevaLinea());
      setCola([]);
      setModalCobroAbierto(false);
      setClienteNombre("");
      setClienteTel("");
      setDescuento("0");
      setNotasVenta("");
      setLinkAppointmentId(undefined);
      setLinkClientPackageId(undefined);
      await onVentaRegistrada?.();
      router.push(urlTicketVenta(data.id));
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="space-y-6">
      {linkAppointmentId ? (
        <p className="text-sm font-medium text-gold-dark">
          {`Vinculado al turno #${linkAppointmentId} de la agenda.`}
        </p>
      ) : null}
      {linkClientPackageId ? (
        <p className="text-sm font-medium text-gold-dark">
          {`Vinculado al paquete asignado #${linkClientPackageId}.`}
        </p>
      ) : null}

      <div>
        <h3 className="mb-3 text-xs font-semibold uppercase tracking-wide text-ink-muted">
          Agregar ítem
        </h3>
        <div className="rounded-sm border border-gold/25 bg-white/60 p-4">
          <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-4">
            <div>
              <label className={uiLabel}>Tipo</label>
              <select
                value={borrador.tipo}
                onChange={(e) =>
                  onTipoChangeBorrador(e.target.value as LineaForm["tipo"])
                }
                className={inputClass}
              >
                <option value="servicio">Servicio</option>
                <option value="paquete">Paquete</option>
                <option value="otro">Otro</option>
              </select>
            </div>
            {borrador.tipo === "servicio" ? (
              <div className="md:col-span-2 lg:col-span-2">
                <ServicioPasosSelect
                  servicios={serviciosParaPasos}
                  value={borrador.serviceId ? String(borrador.serviceId) : ""}
                  onChange={(id) => {
                    if (!id) {
                      patchBorrador({
                        serviceId: undefined,
                        descripcion: "",
                        precioUnitarioPesos: 0,
                      });
                      return;
                    }
                    onPickServicioBorrador(Number(id));
                  }}
                  showPrecio
                  showAnticipo
                  selectClassName={inputClass}
                  className="space-y-3"
                />
              </div>
            ) : borrador.tipo === "paquete" ? (
              <div>
                <label className={uiLabel}>Paquete</label>
                <select
                  value={borrador.servicePackageId ?? ""}
                  onChange={(e) =>
                    onPickPaqueteBorrador(Number(e.target.value))
                  }
                  className={inputClass}
                >
                  <option value="">Elegir...</option>
                  {catalogoState.paquetes.map((p) => (
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
                  value={borrador.descripcion}
                  onChange={(e) =>
                    patchBorrador({ descripcion: e.target.value })
                  }
                  className={inputClass}
                  placeholder="Descripción"
                />
              </div>
            )}
            <div>
              <label className={uiLabel}>Cant.</label>
              <input
                type="number"
                min={1}
                value={borrador.cantidad}
                onChange={(e) =>
                  patchBorrador({ cantidad: Number(e.target.value) || 1 })
                }
                className={inputClass}
              />
            </div>
            <div>
              <label className={uiLabel}>Precio unit. (ARS)</label>
              <input
                type="number"
                min={0}
                value={borrador.precioUnitarioPesos}
                onChange={(e) =>
                  patchBorrador({
                    precioUnitarioPesos: Number(e.target.value) || 0,
                  })
                }
                className={inputClass}
              />
              {anticipoBorrador ? (
                <p className="mt-1 text-xs font-medium text-gold-dark">
                  Anticipo sugerido: {anticipoBorrador.porcentaje}%
                  {anticipoBorrador.monto > 0
                    ? ` (${fmtPesos(anticipoBorrador.monto)})`
                    : " — definí precio en catálogo"}
                  {anticipoBorrador.monto > 0 ? (
                    <>
                      {" · "}
                      <button
                        type="button"
                        className="underline"
                        onClick={() =>
                          patchBorrador({
                            precioUnitarioPesos: anticipoBorrador.monto,
                          })
                        }
                      >
                        Usar monto de anticipo
                      </button>
                    </>
                  ) : null}
                </p>
              ) : null}
              {borrador.tipo === "servicio" || borrador.tipo === "paquete" ? (
                <p className="mt-1 text-xs text-ink-muted">
                  {(() => {
                    const ref = precioCatalogoLinea(borrador);
                    if (ref == null) return null;
                    if (ref === borrador.precioUnitarioPesos) {
                      return ref > 0
                        ? "Coincide con el precio del catálogo."
                        : "Sin precio en catálogo (0).";
                    }
                    if (ref === 0 && borrador.precioUnitarioPesos > 0) {
                      return (
                        <button
                          type="button"
                          disabled={pending}
                          className="underline"
                          onClick={() => void guardarPrecioEnCatalogo(borrador)}
                        >
                          Guardar {fmtPesos(borrador.precioUnitarioPesos)} como
                          precio del catálogo
                        </button>
                      );
                    }
                    return (
                      <>
                        Catálogo: {fmtPesos(ref)}.{" "}
                        <button
                          type="button"
                          disabled={pending}
                          className="underline"
                          onClick={() => void guardarPrecioEnCatalogo(borrador)}
                        >
                          Guardar {fmtPesos(borrador.precioUnitarioPesos)} en
                          catálogo
                        </button>
                      </>
                    );
                  })()}
                </p>
              ) : null}
            </div>
          </div>
          {borrador.tipo !== "otro" && borrador.descripcion ? (
            <p className="mt-2 text-sm text-ink-muted">{borrador.descripcion}</p>
          ) : null}
        </div>
        <button
          type="button"
          onClick={agregarItemACola}
          disabled={pending}
          className={`mt-3 ${uiBtnSecondary}`}
        >
          + Agregar a la venta
        </button>
      </div>

      <div className="border-t border-gold/20 pt-4">
        <h3 className="mb-3 font-serif text-base text-ink-dark">
          Servicios a cobrar
          {cola.length > 0 ? ` (${cola.length})` : ""}
        </h3>
        {cola.length === 0 ? (
          <p className="text-sm text-ink-muted">
            La cola está vacía. Cargá ítems arriba y usá &quot;Agregar a la
            venta&quot;.
          </p>
        ) : (
          <div className={uiTableWrap}>
            <table className="w-full min-w-[560px] text-sm">
              <thead>
                <tr className={uiTableHead}>
                  <th className="px-3 py-2 text-left">#</th>
                  <th className="px-3 py-2 text-left">Tipo</th>
                  <th className="px-3 py-2 text-left">Detalle</th>
                  <th className="px-3 py-2 text-center">Cant.</th>
                  <th className="px-3 py-2 text-right">P. unit.</th>
                  <th className="px-3 py-2 text-right">Subtotal</th>
                  <th className="px-3 py-2 text-right" />
                </tr>
              </thead>
              <tbody>
                {cola.map((l, idx) => (
                  <tr
                    key={l.key}
                    className="border-t border-gold/15 hover:bg-cream/40"
                  >
                    <td className="px-3 py-2 tabular-nums">{idx + 1}</td>
                    <td className="px-3 py-2 text-xs uppercase text-ink-muted">
                      {etiquetaTipo(l.tipo)}
                    </td>
                    <td className="px-3 py-2">{l.descripcion}</td>
                    <td className="px-3 py-2 text-center tabular-nums">
                      {l.cantidad}
                    </td>
                    <td className="px-3 py-2 text-right tabular-nums">
                      {fmtPesos(l.precioUnitarioPesos)}
                    </td>
                    <td className="px-3 py-2 text-right font-medium tabular-nums">
                      {fmtPesos(totalLinea(l))}
                    </td>
                    <td className="px-3 py-2 text-right">
                      <button
                        type="button"
                        className="text-xs text-red-700 underline"
                        onClick={() =>
                          setCola((p) => p.filter((x) => x.key !== l.key))
                        }
                      >
                        Quitar
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="flex flex-wrap items-center justify-between gap-4 border-t border-gold/20 pt-4">
        <p className="text-lg font-semibold tabular-nums">
          Subtotal: {fmtPesos(subtotal)}
          {cola.length > 0 ? (
            <span className="ml-2 text-sm font-normal text-ink-muted">
              ({cola.length} {cola.length === 1 ? "ítem" : "ítems"})
            </span>
          ) : null}
        </p>
        <button
          type="button"
          disabled={pending || cola.length === 0}
          onClick={abrirModalCobro}
          className={uiBtnPrimary}
        >
          Cobrar e imprimir ticket
        </button>
      </div>

      {modalCobroAbierto ? (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="modal-cobro-titulo"
        >
          <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-sm border border-gold/40 bg-surface p-6 shadow-lg">
            <h3
              id="modal-cobro-titulo"
              className="mb-1 font-serif text-lg text-ink-dark"
            >
              Confirmar cobro
            </h3>
            <p className="mb-4 text-sm text-ink-muted">
              {cola.length} {cola.length === 1 ? "ítem" : "ítems"} — subtotal{" "}
              {fmtPesos(subtotal)}
            </p>
            <form onSubmit={onRegistrarVenta} className="space-y-4">
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <label className={uiLabel}>Cliente (opcional)</label>
                  <input
                    value={clienteNombre}
                    onChange={(e) => setClienteNombre(e.target.value)}
                    className={inputClass}
                    placeholder="Nombre"
                  />
                </div>
                <div>
                  <label className={uiLabel}>Teléfono</label>
                  <input
                    value={clienteTel}
                    onChange={(e) => setClienteTel(e.target.value)}
                    className={inputClass}
                    placeholder="11..."
                  />
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <label className={uiLabel}>Descuento (ARS)</label>
                  <input
                    type="number"
                    min={0}
                    value={descuento}
                    onChange={(e) => setDescuento(e.target.value)}
                    className={inputClass}
                  />
                </div>
                <div>
                  <label className={uiLabel}>Forma de pago</label>
                  <select
                    value={metodoPago}
                    onChange={(e) => setMetodoPago(e.target.value)}
                    className={uiSelect}
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
                <label className={uiLabel}>Notas</label>
                <input
                  value={notasVenta}
                  onChange={(e) => setNotasVenta(e.target.value)}
                  className={inputClass}
                />
              </div>
              <div className="rounded-sm border border-gold/25 bg-cream/50 px-3 py-2 text-sm">
                <p className="font-semibold tabular-nums text-ink-dark">
                  Total a cobrar: {fmtPesos(total)}
                </p>
                {descuentoNum > 0 ? (
                  <p className="text-xs text-ink-muted">
                    Incluye descuento de {fmtPesos(descuentoNum)}
                  </p>
                ) : null}
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  className={uiBtnSecondary}
                  disabled={pending}
                  onClick={() => setModalCobroAbierto(false)}
                >
                  Cancelar
                </button>
                <button type="submit" disabled={pending} className={uiBtnPrimary}>
                  Confirmar e imprimir ticket
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : null}
    </div>
  );
}
