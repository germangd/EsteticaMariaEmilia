"use client";

import type { VentaDetalle } from "@/lib/caja-repo";
import { fmtPesos } from "@/lib/fmt-pesos";

const METODO_LABEL: Record<string, string> = {
  efectivo: "Efectivo",
  transferencia: "Transferencia",
  debito: "D\u00e9bito",
  credito: "Cr\u00e9dito",
  otro: "Otro",
};

function fmtFecha(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString("es-AR", {
    dateStyle: "short",
    timeStyle: "short",
  });
}

export function AdminCajaTicket({ venta }: { venta: VentaDetalle }) {
  const anulada = venta.estado === "anulada";

  return (
    <div className="ticket-print mx-auto max-w-md bg-white p-6 text-ink-dark print:m-0 print:max-w-none print:p-4">
      <div className="mb-4 border-b border-dashed border-ink/30 pb-4 text-center">
        <p className="text-[0.65rem] font-bold uppercase tracking-[0.25em] text-gold-dark">
          {"Mar\u00eda Emilia Est\u00e9tica"}
        </p>
        <h1 className="font-serif text-xl font-light">Comprobante de venta</h1>
        {anulada ? (
          <p className="mt-2 text-sm font-bold uppercase text-red-700">
            Anulado
          </p>
        ) : null}
      </div>

      <dl className="mb-4 space-y-1 text-sm">
        <div className="flex justify-between gap-4">
          <dt className="text-ink-muted">{"Ticket N\u00b0"}</dt>
          <dd className="font-semibold">
            {venta.sessionId}-{String(venta.numero).padStart(4, "0")}
          </dd>
        </div>
        <div className="flex justify-between gap-4">
          <dt className="text-ink-muted">Fecha</dt>
          <dd>{fmtFecha(venta.createdAt)}</dd>
        </div>
        {venta.clienteNombre ? (
          <div className="flex justify-between gap-4">
            <dt className="text-ink-muted">Cliente</dt>
            <dd className="text-right">{venta.clienteNombre}</dd>
          </div>
        ) : null}
        {venta.clienteTelefono ? (
          <div className="flex justify-between gap-4">
            <dt className="text-ink-muted">{"Tel\u00e9fono"}</dt>
            <dd>{venta.clienteTelefono}</dd>
          </div>
        ) : null}
        <div className="flex justify-between gap-4">
          <dt className="text-ink-muted">Pago</dt>
          <dd className="capitalize">
            {METODO_LABEL[venta.metodoPago] ?? venta.metodoPago}
          </dd>
        </div>
      </dl>

      <table className="mb-4 w-full border-collapse text-sm">
        <thead>
          <tr className="border-b border-ink/20 text-left text-[0.65rem] uppercase tracking-wide text-ink-muted">
            <th className="py-1 pr-2">Concepto</th>
            <th className="py-1 px-1 text-center">Cant.</th>
            <th className="py-1 pl-2 text-right">Importe</th>
          </tr>
        </thead>
        <tbody>
          {venta.lineas.map((l) => (
            <tr key={l.id} className="border-b border-ink/10">
              <td className="py-2 pr-2">{l.descripcion}</td>
              <td className="py-2 px-1 text-center">{l.cantidad}</td>
              <td className="py-2 pl-2 text-right tabular-nums">
                {fmtPesos(l.totalLineaPesos)}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <div className="space-y-1 border-t border-dashed border-ink/30 pt-3 text-sm">
        <div className="flex justify-between">
          <span>Subtotal</span>
          <span className="tabular-nums">{fmtPesos(venta.subtotalPesos)}</span>
        </div>
        {venta.descuentoPesos > 0 ? (
          <div className="flex justify-between text-ink-muted">
            <span>Descuento</span>
            <span className="tabular-nums">
              -{fmtPesos(venta.descuentoPesos)}
            </span>
          </div>
        ) : null}
        <div className="flex justify-between text-base font-bold">
          <span>Total</span>
          <span className="tabular-nums">{fmtPesos(venta.totalPesos)}</span>
        </div>
      </div>

      {venta.notas ? (
        <p className="mt-4 text-xs text-ink-muted">
          <span className="font-semibold">Notas:</span> {venta.notas}
        </p>
      ) : null}

      <p className="mt-6 text-center text-[0.65rem] text-ink-muted">
        {"Comprobante interno \u2014 no v\u00e1lido como factura fiscal"}
      </p>

      <div className="mt-6 flex flex-wrap justify-center gap-3 print:hidden">
        <button
          type="button"
          onClick={() => window.print()}
          className="rounded-sm bg-gold px-5 py-2.5 text-[0.72rem] font-semibold uppercase tracking-wider text-white shadow-sm hover:bg-gold-dark"
        >
          Imprimir
        </button>
        <a
          href="/admin/caja"
          className="rounded-sm border border-gold/55 bg-white px-5 py-2.5 text-[0.72rem] font-semibold uppercase tracking-wider text-gold-dark hover:bg-cream"
        >
          Volver a caja
        </a>
      </div>
    </div>
  );
}
