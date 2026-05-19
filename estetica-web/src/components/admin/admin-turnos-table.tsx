import { DateTime } from "luxon";
import type { TurnoListado } from "@/lib/turnos-repo";
import { AdminCancelCell } from "@/components/admin/admin-cancel-cell";
import { urlCobrarTurno } from "@/lib/caja-url";
import { uiTableHead } from "@/lib/ui-classes";

function fmtFechaCorta(fechaIso: string, tz: string): string {
  const dt = DateTime.fromISO(fechaIso, { zone: tz });
  if (!dt.isValid) return fechaIso;
  return dt.setLocale("es").toFormat("d/M/yy");
}

function fmtFechaLarga(fechaIso: string, tz: string): string {
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
    <div className="overflow-hidden rounded-sm border border-gold/55 bg-white/50 ring-1 ring-gold/20">
      <table className="w-full table-fixed border-collapse text-left text-[0.75rem] leading-tight">
        <colgroup>
          <col className={showActions ? "w-[7%]" : "w-[8%]"} />
          <col className="w-[5%]" />
          <col className={showActions ? "w-[9%]" : "w-[10%]"} />
          <col className={showActions ? "w-[14%]" : "w-[16%]"} />
          <col className={showActions ? "w-[8%]" : "w-[9%]"} />
          <col className={showActions ? "w-[9%]" : "w-[10%]"} />
          <col className={showActions ? "w-[13%]" : "w-[15%]"} />
          <col className={showActions ? "w-[9%]" : "w-[10%]"} />
          <col className={showActions ? "w-[6%]" : "w-[7%]"} />
          {showActions ? <col className="w-[8%]" /> : null}
        </colgroup>
        <thead className={uiTableHead}>
          <tr>
            <th className="px-1.5 py-2 pl-2 font-bold">Fecha</th>
            <th className="px-1 py-2 font-bold">Hora</th>
            <th className="px-1 py-2 font-bold">Sede</th>
            <th className="px-1 py-2 font-bold">Servicio</th>
            <th className="px-1 py-2 font-bold">Cliente</th>
            <th className="px-1 py-2 font-bold">Tel.</th>
            <th className="px-1 py-2 font-bold">Email</th>
            <th className="px-1 py-2 font-bold">Resp.</th>
            <th className="px-1 py-2 font-bold">Cód.</th>
            {showActions ? (
              <th className="px-1.5 py-2 pr-2 text-right font-bold">Acc.</th>
            ) : null}
          </tr>
        </thead>
        <tbody className="divide-y divide-gold/10">
          {rows.map((r) => (
            <tr key={r.id} className="text-ink hover:bg-cream/80">
              <td
                className="px-1.5 py-1.5 pl-2 align-top tabular-nums"
                title={fmtFechaLarga(r.fecha, tz)}
              >
                {fmtFechaCorta(r.fecha, tz)}
              </td>
              <td className="whitespace-nowrap px-1 py-1.5 align-top font-semibold text-ink-dark">
                {r.hora}
              </td>
              <td
                className="line-clamp-2 px-1 py-1.5 align-top text-ink-muted"
                title={r.sedeNombre}
              >
                {r.sedeNombre}
              </td>
              <td
                className="line-clamp-2 px-1 py-1.5 align-top"
                title={r.servicioNombre}
              >
                {r.servicioNombre}
              </td>
              <td
                className="truncate px-1 py-1.5 align-top"
                title={r.nombreCliente}
              >
                {r.nombreCliente}
              </td>
              <td
                className="truncate px-1 py-1.5 align-top tabular-nums"
                title={r.telefono}
              >
                {r.telefono}
              </td>
              <td className="px-1 py-1.5 align-top">
                {r.email ? (
                  <a
                    href={`mailto:${r.email}`}
                    className="line-clamp-2 break-all text-[0.7rem] text-ink underline-offset-2 hover:text-gold-dark hover:underline"
                    title={r.email}
                  >
                    {r.email}
                  </a>
                ) : (
                  "—"
                )}
              </td>
              <td
                className="line-clamp-2 px-1 py-1.5 align-top"
                title={r.responsable}
              >
                {r.responsable}
              </td>
              <td
                className="truncate px-1 py-1.5 align-top font-mono text-[0.65rem] font-semibold text-gold-dark"
                title={r.codigoCancelacion}
              >
                {r.codigoCancelacion}
              </td>
              {showActions ? (
                <td className="px-1.5 py-1.5 pr-2 align-top">
                  <div className="flex flex-col items-end gap-1">
                    <a
                      href={urlCobrarTurno(r.id)}
                      className="text-[0.6rem] font-semibold uppercase leading-none tracking-wide text-gold-dark underline"
                    >
                      Cobrar
                    </a>
                    <AdminCancelCell codigo={r.codigoCancelacion} compact />
                  </div>
                </td>
              ) : null}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
