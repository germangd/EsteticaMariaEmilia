"use client";

import { useEffect, useState } from "react";
import type { VentaAuditoria, VentaDetalle } from "@/lib/caja-repo";
import { urlTicketVenta } from "@/lib/caja-url";
import { fmtPesos } from "@/lib/fmt-pesos";
import { uiBtnSecondary } from "@/lib/ui-classes";

const ACCION_LABEL: Record<string, string> = {
  creada: "Alta del comprobante",
  modificada: "Modificación",
  anulada: "Anulación",
};

function fmtFecha(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString("es-AR", {
    dateStyle: "medium",
    timeStyle: "short",
  });
}

function esObjeto(v: unknown): v is Record<string, unknown> {
  return v != null && typeof v === "object" && !Array.isArray(v);
}

function resumenSnapshot(data: unknown): string[] {
  if (data == null) return ["—"];
  if (!esObjeto(data)) return [String(data)];

  const lineas: string[] = [];
  if (data.estado != null) lineas.push(`Estado: ${String(data.estado)}`);
  if (data.clienteNombre != null) {
    lineas.push(`Cliente: ${String(data.clienteNombre)}`);
  }
  if (data.clienteTelefono != null) {
    lineas.push(`Teléfono: ${String(data.clienteTelefono)}`);
  }
  if (data.metodoPago != null) {
    lineas.push(`Pago: ${String(data.metodoPago)}`);
  }
  if (data.totalPesos != null) {
    lineas.push(`Total: ${fmtPesos(Number(data.totalPesos))}`);
  }
  if (data.subtotalPesos != null && data.descuentoPesos != null) {
    lineas.push(
      `Subtotal ${fmtPesos(Number(data.subtotalPesos))} · Desc. ${fmtPesos(Number(data.descuentoPesos))}`
    );
  }
  if (data.notas != null && String(data.notas).trim()) {
    lineas.push(`Notas: ${String(data.notas)}`);
  }
  if (Array.isArray(data.lineas) && data.lineas.length > 0) {
    lineas.push("Ítems:");
    for (const item of data.lineas) {
      if (!esObjeto(item)) continue;
      const desc = String(item.descripcion ?? "—");
      const cant = Number(item.cantidad ?? 1);
      const pu = Number(item.precioUnitarioPesos ?? item.totalLineaPesos ?? 0);
      lineas.push(`  · ${cant} × ${desc} — ${fmtPesos(pu)}`);
    }
  }
  if (lineas.length === 0) {
    lineas.push(JSON.stringify(data, null, 2));
  }
  return lineas;
}

function BloqueSnapshot({
  titulo,
  data,
}: {
  titulo: string;
  data: unknown | null;
}) {
  const [expandido, setExpandido] = useState(false);
  if (data == null) return null;

  const resumen = resumenSnapshot(data);
  const json =
    typeof data === "object"
      ? JSON.stringify(data, null, 2)
      : String(data);

  return (
    <div className="rounded border border-gold/25 bg-white/60 p-3 text-sm">
      <button
        type="button"
        onClick={() => setExpandido((x) => !x)}
        className="flex w-full items-center justify-between text-left font-medium text-ink-dark"
      >
        <span>{titulo}</span>
        <span className="text-xs text-gold-dark">
          {expandido ? "Ocultar JSON" : "Ver JSON"}
        </span>
      </button>
      <ul className="mt-2 space-y-0.5 text-ink-muted">
        {resumen.map((l, i) => (
          <li key={i} className="whitespace-pre-wrap">
            {l}
          </li>
        ))}
      </ul>
      {expandido ? (
        <pre className="mt-2 max-h-48 overflow-auto rounded bg-cream/80 p-2 text-[10px] leading-snug text-ink">
          {json}
        </pre>
      ) : null}
    </div>
  );
}

function EntradaAuditoria({ entrada }: { entrada: VentaAuditoria }) {
  return (
    <li className="border-l-2 border-gold/50 pl-4 pb-6 last:pb-0">
      <p className="text-xs text-ink-muted">{fmtFecha(entrada.createdAt)}</p>
      <p className="font-medium text-ink-dark">
        {ACCION_LABEL[entrada.accion] ?? entrada.accion}
      </p>
      {entrada.detalle ? (
        <p className="mt-1 text-sm text-ink-muted">
          Motivo: {entrada.detalle}
        </p>
      ) : null}
      <div className="mt-3 grid gap-2 sm:grid-cols-2">
        <BloqueSnapshot titulo="Antes" data={entrada.datosAntes} />
        <BloqueSnapshot titulo="Después" data={entrada.datosDespues} />
      </div>
    </li>
  );
}

export function AdminCajaAuditoriaPanel({
  ventaId,
  numeroTicket,
  onClose,
}: {
  ventaId: number;
  numeroTicket: number;
  onClose: () => void;
}) {
  const [venta, setVenta] = useState<VentaDetalle | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [pending, setPending] = useState(true);

  useEffect(() => {
    let cancel = false;
    (async () => {
      setPending(true);
      setError(null);
      try {
        const r = await fetch(`/api/admin/caja/ventas/${ventaId}`, {
          credentials: "same-origin",
        });
        const data = (await r.json()) as {
          ok?: boolean;
          venta?: VentaDetalle;
          mensaje?: string;
        };
        if (cancel) return;
        if (!r.ok || !data.ok || !data.venta) {
          setError(data.mensaje ?? "No se pudo cargar la auditoría.");
          return;
        }
        setVenta(data.venta);
      } catch {
        if (!cancel) setError("Error de conexión.");
      } finally {
        if (!cancel) setPending(false);
      }
    })();
    return () => {
      cancel = true;
    };
  }, [ventaId]);

  const entradas = venta
    ? [...venta.auditoria].sort(
        (a, b) =>
          new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()
      )
    : [];

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="auditoria-titulo"
    >
      <div className="flex max-h-[90vh] w-full max-w-2xl flex-col rounded-sm border border-gold/40 bg-surface shadow-lg">
        <div className="border-b border-gold/25 px-6 py-4">
          <h3
            id="auditoria-titulo"
            className="font-serif text-lg text-ink-dark"
          >
            Auditoría — ticket #
            {String(numeroTicket).padStart(6, "0")}
          </h3>
          <p className="mt-1 text-sm text-ink-muted">
            Registro de altas, cambios y anulaciones con el detalle guardado en
            cada momento.
          </p>
        </div>

        <div className="flex-1 overflow-y-auto px-6 py-4">
          {pending ? (
            <p className="text-sm text-ink-muted">Cargando…</p>
          ) : error ? (
            <p className="text-sm text-red-700">{error}</p>
          ) : entradas.length === 0 ? (
            <p className="text-sm text-ink-muted">
              Sin movimientos registrados para este comprobante.
            </p>
          ) : (
            <ol className="space-y-0">{entradas.map((e) => (
                <EntradaAuditoria key={e.id} entrada={e} />
              ))}</ol>
          )}
        </div>

        <div className="flex flex-wrap justify-end gap-2 border-t border-gold/25 px-6 py-4">
          <a
            href={urlTicketVenta(ventaId)}
            target="_blank"
            rel="noopener noreferrer"
            className="text-sm font-medium text-gold-dark underline"
          >
            Ver comprobante
          </a>
          <button type="button" className={uiBtnSecondary} onClick={onClose}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
}
