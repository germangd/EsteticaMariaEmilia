import { DateTime } from "luxon";
import type { TurnoListado } from "@/lib/turnos-repo";
import { AdminCancelCell } from "@/components/admin/admin-cancel-cell";
import { urlCobrarTurno } from "@/lib/caja-url";
import { uiTableHead, uiTableWrap } from "@/lib/ui-classes";

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
  rows: TurnoListado[];
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
    <div className={`${uiTableWrap} -mx-1 px-1`}>
      <table className="w-full table-fixed text-left text-[0.8125rem] leading-snug">
        <thead className={uiTableHead}>
          <tr>
            <th className="w-[12%] px-2 py-2.5 pl-3">Fecha</th>
            <th className="w-[6%] px-2 py-2.5">Hora</th>
            <th className="w-[11%] px-2 py-2.5">Sede</th>
            <th className="w-[14%] px-2 py-2.5">Servicio</th>
            <th className="w-[9%] px-2 py-2.5">Cliente</th>
            <th className="w-[10%] px-2 py-2.5">Teléfono</th>
            <th className="w-[14%] px-2 py-2.5">Email</th>
            <th className="w-[11%] px-2 py-2.5">Responsable</th>
            <th className="w-[7%] px-2 py-2.5">Código</th>
            {showActions ? (
              <th className="w-[10%] px-2 py-2 pr-3 text-right">Acciones</th>
            ) : null}
          </tr>
        </thead>
        <tbody className="divide-y divide-gold/10">
          {rows.map((r) => (
            <tr key={r.id} className="font-medium text-ink hover:bg-cream/80">
              <td className="px-2 py-2 pl-3 align-top text-ink">
                {fmtFechaEtiqueta(r.fecha, tz)}
              </td>
              <td className="whitespace-nowrap px-2 py-2 align-top font-semibold text-ink-dark">
                {r.hora}
              </td>
              <td className="break-words px-2 py-2 align-top text-ink-muted">
                {r.sedeNombre}
              </td>
              <td
                className="break-words px-2 py-2 align-top"
                title={r.servicioNombre}
              >
                {r.servicioNombre}
              </td>
              <td
                className="break-words px-2 py-2 align-top"
                title={r.nombreCliente}
              >
                {r.nombreCliente}
              </td>
              <td className="whitespace-nowrap px-2 py-2 align-top tabular-nums text-ink">
                {r.telefono}
              </td>
              <td
                className="break-all px-2 py-2 align-top text-[0.75rem] text-ink"
                title={r.email ?? undefined}
              >
                {r.email ?? "—"}
              </td>
              <td className="break-words px-2 py-2 align-top text-ink">
                {r.responsable}
              </td>
              <td className="whitespace-nowrap px-2 py-2 align-top font-mono text-[0.7rem] font-semibold text-gold-dark">
                {r.codigoCancelacion}
              </td>
              {showActions ? (
                <td className="px-2 py-2 pr-3 text-right align-top whitespace-nowrap">
                  <a
                    href={urlCobrarTurno(r.id)}
                    className="mr-2 text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark underline"
                  >
                    Cobrar
                  </a>
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
