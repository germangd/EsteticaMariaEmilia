import { DateTime } from "luxon";
import type { TurnoListado } from "@/lib/turnos-repo";
import { AdminConfirmarAnticipoButton } from "@/components/admin/admin-confirmar-anticipo-button";
import { AdminCancelCell } from "@/components/admin/admin-cancel-cell";
import { urlCobrarTurno } from "@/lib/caja-url";
import { uiTableHead, uiTableWrap } from "@/lib/ui-classes";

function fmtFechaEtiqueta(fechaIso: string, tz: string): string {
  const dt = DateTime.fromISO(fechaIso, { zone: tz });
  if (!dt.isValid) return fechaIso;
  return dt.setLocale("es").toFormat("ccc d MMM yyyy");
}

export function AdminTurnosPendientesTable({
  rows,
  tz,
}: {
  rows: TurnoListado[];
  tz: string;
}) {
  if (rows.length === 0) return null;

  return (
    <div className="mb-10 rounded-sm border border-amber-300/60 bg-amber-50/40 p-4 md:p-6">
      <h3 className="mb-1 font-serif text-lg text-ink-dark">
        Pendientes de confirmación ({rows.length})
      </h3>
      <p className="mb-4 text-sm text-ink-muted">
        Reservas web con anticipo. Usá{" "}
        <strong className="font-semibold text-ink-dark">Cobrar en caja</strong>{" "}
        para registrar el pago; luego{" "}
        <strong className="font-semibold text-ink-dark">Confirmar anticipo</strong>{" "}
        envía el aviso al cliente y lo muestra como confirmado en la agenda.
      </p>
      <div className={uiTableWrap}>
        <table className="w-full table-fixed text-left text-[0.8125rem] leading-snug">
          <thead className={uiTableHead}>
            <tr>
              <th className="px-3 py-3 pl-4">Fecha</th>
              <th className="px-3 py-3">Hora</th>
              <th className="px-3 py-3">Sede</th>
              <th className="px-3 py-3">Servicio</th>
              <th className="px-3 py-3">Cliente</th>
              <th className="px-3 py-3">Teléfono</th>
              <th className="px-3 py-3">Código</th>
              <th className="px-3 py-3 pr-4 text-right">Acciones</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gold/10">
            {rows.map((r) => (
              <tr
                key={r.id}
                className="bg-amber-50/30 font-medium text-ink hover:bg-amber-50/60"
              >
                <td className="whitespace-nowrap px-3 py-2.5 pl-4">
                  {fmtFechaEtiqueta(r.fecha, tz)}
                </td>
                <td className="whitespace-nowrap px-3 py-2.5 font-semibold text-ink-dark">
                  {r.hora}
                </td>
                <td className="px-3 py-2.5 text-ink-muted">{r.sedeNombre}</td>
                <td className="max-w-[140px] truncate px-3 py-2.5" title={r.servicioNombre}>
                  {r.servicioNombre}
                </td>
                <td className="max-w-[120px] truncate px-3 py-2.5">{r.nombreCliente}</td>
                <td className="whitespace-nowrap px-3 py-2.5">{r.telefono}</td>
                <td className="whitespace-nowrap px-3 py-2.5 font-mono text-xs font-semibold text-gold-dark">
                  {r.codigoCancelacion}
                </td>
                <td className="px-3 py-2 pr-4 text-right whitespace-nowrap">
                  <div className="flex flex-col items-end gap-1.5">
                    <div className="flex flex-wrap items-center justify-end gap-x-3 gap-y-1">
                      <a
                        href={urlCobrarTurno(r.id)}
                        className="text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark underline"
                      >
                        Cobrar en caja
                      </a>
                      <AdminCancelCell codigo={r.codigoCancelacion} />
                    </div>
                    <AdminConfirmarAnticipoButton appointmentId={r.id} />
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
