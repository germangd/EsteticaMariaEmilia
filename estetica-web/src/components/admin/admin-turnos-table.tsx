import { DateTime } from "luxon";
import type { AppointmentRow } from "@/db/schema";
import { AdminCancelCell } from "@/components/admin/admin-cancel-cell";

function fmtFechaEtiqueta(fechaIso: string, tz: string): string {
  const dt = DateTime.fromISO(fechaIso, { zone: tz });
  if (!dt.isValid) return fechaIso;
  return dt.setLocale("es").toFormat("ccc d MMM yyyy");
}

export function AdminTurnosTable({
  rows,
  tz,
  showActions,
}: {
  rows: AppointmentRow[];
  tz: string;
  showActions: boolean;
}) {
  if (rows.length === 0) {
    return (
      <p className="py-8 text-center text-sm text-ink-muted">
        No hay turnos en esta lista.
      </p>
    );
  }
  return (
    <div className="overflow-x-auto rounded-sm border border-gold/25 bg-white shadow-sm">
      <table className="min-w-[780px] w-full text-left text-sm">
        <thead className="border-b border-gold/20 bg-cream text-[0.65rem] font-medium uppercase tracking-[0.12em] text-ink-muted">
          <tr>
            <th className="px-3 py-3 pl-4">Fecha</th>
            <th className="px-3 py-3">Hora</th>
            <th className="px-3 py-3">Servicio</th>
            <th className="px-3 py-3">Cliente</th>
            <th className="px-3 py-3">Teléfono</th>
            <th className="px-3 py-3">Email</th>
            <th className="px-3 py-3">Responsable</th>
            <th className="px-3 py-3">Código</th>
            {showActions ? (
              <th className="px-3 py-3 pr-4 text-right">Acciones</th>
            ) : null}
          </tr>
        </thead>
        <tbody className="divide-y divide-gold/10">
          {rows.map((r) => (
            <tr key={r.id} className="text-ink hover:bg-cream/60">
              <td className="whitespace-nowrap px-3 py-2.5 pl-4 text-ink-muted">
                {fmtFechaEtiqueta(r.fecha, tz)}
              </td>
              <td className="whitespace-nowrap px-3 py-2.5 font-medium text-ink-dark">
                {r.hora}
              </td>
              <td
                className="max-w-[140px] truncate px-3 py-2.5"
                title={r.servicioNombre}
              >
                {r.servicioNombre}
              </td>
              <td
                className="max-w-[120px] truncate px-3 py-2.5"
                title={r.nombreCliente}
              >
                {r.nombreCliente}
              </td>
              <td className="whitespace-nowrap px-3 py-2.5 text-ink-muted">
                {r.telefono}
              </td>
              <td className="max-w-[160px] truncate px-3 py-2.5 text-ink-muted">
                {r.email ?? "—"}
              </td>
              <td className="whitespace-nowrap px-3 py-2.5 text-ink-muted">
                {r.responsable}
              </td>
              <td className="whitespace-nowrap px-3 py-2.5 font-mono text-xs text-gold-dark">
                {r.codigoCancelacion}
              </td>
              {showActions ? (
                <td className="px-3 py-2 pr-4 text-right">
                  <AdminCancelCell codigo={r.codigoCancelacion} />
                </td>
              ) : null}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
