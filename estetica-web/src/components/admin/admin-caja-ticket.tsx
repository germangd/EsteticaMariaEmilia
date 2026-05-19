"use client";

import { getBusinessName } from "@/config/site";
import type { VentaDetalle } from "@/lib/caja-repo";
import { fmtPesos } from "@/lib/fmt-pesos";

const METODO_LABEL: Record<string, string> = {
  efectivo: "Efectivo",
  transferencia: "Transferencia",
  debito: "Debito",
  credito: "Credito",
  otro: "Otro",
};

function fmtFecha(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString("es-AR", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

export function AdminCajaTicket({ venta }: { venta: VentaDetalle }) {
  const anulada = venta.estado === "anulada";
  const ticketNum = String(venta.numeroTicket).padStart(6, "0");

  return (
    <div className="ticket-thermal mx-auto w-[80mm] max-w-[80mm] bg-white px-3 py-4 font-mono text-[11px] leading-snug text-black print:m-0 print:w-[80mm] print:max-w-[80mm] print:p-0">
      <div className="border-b border-dashed border-black/40 pb-2 text-center">
        <p className="text-[10px] font-bold uppercase tracking-widest">
          {getBusinessName()}
        </p>
        <p className="mt-1 text-[12px] font-bold">COMPROBANTE DE VENTA</p>
        {anulada ? (
          <p className="mt-1 text-[11px] font-bold uppercase">*** ANULADO ***</p>
        ) : null}
      </div>

      <div className="my-2 space-y-0.5">
        <Row label="Ticket" value={ticketNum} bold />
        <Row label="Fecha" value={fmtFecha(venta.createdAt)} />
        {venta.clienteNombre ? (
          <Row label="Cliente" value={venta.clienteNombre} />
        ) : null}
        {venta.clienteTelefono ? (
          <Row label="Tel" value={venta.clienteTelefono} />
        ) : null}
        <Row
          label="Pago"
          value={METODO_LABEL[venta.metodoPago] ?? venta.metodoPago}
        />
      </div>

      <div className="border-y border-dashed border-black/40 py-2">
        {venta.lineas.map((l) => (
          <div key={l.id} className="mb-2 last:mb-0">
            <p className="font-semibold leading-tight">{l.descripcion}</p>
            <p className="flex justify-between tabular-nums">
              <span>
                {l.cantidad} x {fmtPesos(l.precioUnitarioPesos)}
              </span>
              <span>{fmtPesos(l.totalLineaPesos)}</span>
            </p>
          </div>
        ))}
      </div>

      <div className="mt-2 space-y-0.5 tabular-nums">
        <Row label="Subtotal" value={fmtPesos(venta.subtotalPesos)} />
        {venta.descuentoPesos > 0 ? (
          <Row label="Descuento" value={`-${fmtPesos(venta.descuentoPesos)}`} />
        ) : null}
        <p className="flex justify-between border-t border-black/30 pt-1 text-[13px] font-bold">
          <span>TOTAL (este comprobante)</span>
          <span>{fmtPesos(venta.totalPesos)}</span>
        </p>
      </div>

      {venta.resumenAnticipo ? (
        <div className="mt-2 space-y-0.5 border-t border-dashed border-black/40 pt-2 tabular-nums">
          <p className="mb-1 text-center text-[9px] font-bold uppercase">
            Reserva con anticipo
          </p>
          <Row
            label="Valor del tratamiento"
            value={fmtPesos(venta.resumenAnticipo.precioTratamientoPesos)}
          />
          {venta.resumenAnticipo.anticipoReferenciaPesos > 0 ? (
            <Row
              label={`Anticipo ref. (${venta.resumenAnticipo.anticipoPorcentaje}%)`}
              value={fmtPesos(venta.resumenAnticipo.anticipoReferenciaPesos)}
            />
          ) : null}
          <Row
            label="Total abonado (turno)"
            value={fmtPesos(venta.resumenAnticipo.totalAbonadoPesos)}
            bold
          />
          {venta.resumenAnticipo.importeEsteComprobantePesos > 0 &&
          venta.resumenAnticipo.importeEsteComprobantePesos !==
            venta.resumenAnticipo.totalAbonadoPesos ? (
            <Row
              label="En este comprobante"
              value={fmtPesos(venta.resumenAnticipo.importeEsteComprobantePesos)}
            />
          ) : null}
          <p className="flex justify-between border-t border-black/30 pt-1 text-[12px] font-bold">
            <span>Saldo a abonar</span>
            <span>
              {venta.resumenAnticipo.saldoPendientePesos > 0
                ? fmtPesos(venta.resumenAnticipo.saldoPendientePesos)
                : "$ 0"}
            </span>
          </p>
          {venta.resumenAnticipo.saldoPendientePesos <= 0 ? (
            <p className="text-center text-[9px] leading-tight">
              Tratamiento abonado en su totalidad.
            </p>
          ) : null}
        </div>
      ) : null}

      {venta.notas ? (
        <p className="mt-2 text-[10px]">Notas: {venta.notas}</p>
      ) : null}

      {venta.auditoria.length > 0 ? (
        <div className="mt-3 border-t border-dashed border-black/30 pt-2 text-[9px]">
          <p className="mb-1 font-bold uppercase">Registro de cambios</p>
          {[...venta.auditoria].reverse().map((a) => (
            <p key={a.id} className="leading-tight">
              {fmtFecha(a.createdAt)} —{" "}
              {a.accion === "creada"
                ? "Alta"
                : a.accion === "modificada"
                  ? "Modificación"
                  : a.accion === "anulada"
                    ? "Anulación"
                    : a.accion}
              {a.detalle ? `: ${a.detalle}` : ""}
            </p>
          ))}
        </div>
      ) : null}

      <p className="mt-3 border-t border-dashed border-black/30 pt-2 text-center text-[9px] leading-tight">
        Comprobante interno. No valido como factura fiscal.
      </p>

      <div className="mt-4 flex justify-center gap-2 print:hidden">
        <button
          type="button"
          onClick={() => window.print()}
          className="rounded border border-black/30 bg-black px-4 py-2 text-[10px] font-bold uppercase text-white"
        >
          Imprimir
        </button>
        <a
          href="/admin/caja"
          className="rounded border border-black/30 px-4 py-2 text-[10px] font-bold uppercase"
        >
          Volver
        </a>
      </div>
    </div>
  );
}

function Row({
  label,
  value,
  bold,
}: {
  label: string;
  value: string;
  bold?: boolean;
}) {
  return (
    <p className={`flex justify-between gap-2 ${bold ? "font-bold" : ""}`}>
      <span>{label}</span>
      <span className="text-right">{value}</span>
    </p>
  );
}
