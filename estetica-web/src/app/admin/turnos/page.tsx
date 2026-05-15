import type { Metadata } from "next";
import Link from "next/link";
import { getAppTimeZone, hoyIsoEnZona } from "@/lib/agenda";
import { AdminTurnosTable } from "@/components/admin/admin-turnos-table";
import {
  listarNombresServiciosCatalogo,
  listarTurnosActivosFiltrados,
} from "@/lib/turnos-repo";
import { DateTime } from "luxon";

export const metadata: Metadata = {
  title: "Admin — Turnos | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

function fmtFechaEtiqueta(fechaIso: string, tz: string): string {
  const dt = DateTime.fromISO(fechaIso, { zone: tz });
  if (!dt.isValid) return fechaIso;
  return dt.setLocale("es").toFormat("ccc d MMM yyyy");
}

function isIsoDate(s: string | undefined): s is string {
  return Boolean(s && /^\d{4}-\d{2}-\d{2}$/.test(s));
}

function buildExportQuery(q: {
  servicio?: string;
  desde?: string;
  hasta?: string;
}): string {
  const u = new URLSearchParams();
  if (q.servicio?.trim()) u.set("servicio", q.servicio.trim());
  if (q.desde?.trim()) u.set("desde", q.desde.trim());
  if (q.hasta?.trim()) u.set("hasta", q.hasta.trim());
  const s = u.toString();
  return s ? `?${s}` : "";
}

export default async function AdminTurnosPage({
  searchParams,
}: {
  searchParams: Promise<{
    servicio?: string;
    desde?: string;
    hasta?: string;
  }>;
}) {
  const sp = await searchParams;
  const tz = getAppTimeZone();
  const hoy = hoyIsoEnZona(tz);

  let desde = sp.desde?.trim() ?? hoy;
  if (!isIsoDate(desde)) desde = hoy;

  const hastaRaw = sp.hasta?.trim();
  const hasta = isIsoDate(hastaRaw) ? hastaRaw : null;
  if (hasta && hasta < desde) {
    return (
      <main className="mx-auto max-w-6xl px-5 py-12 md:px-8">
        <p className="text-sm text-red-700">
          La fecha &quot;hasta&quot; no puede ser anterior a &quot;desde&quot;.
        </p>
        <Link href="/admin/turnos" className="mt-4 inline-block text-gold-dark underline">
          Quitar filtros
        </Link>
      </main>
    );
  }

  const servicio = sp.servicio?.trim() || null;
  const catalogo = await listarNombresServiciosCatalogo();

  const todos = await listarTurnosActivosFiltrados({
    fechaDesde: desde,
    fechaHasta: hasta,
    servicioNombre: servicio,
  });

  const turnosHoy = todos.filter((r) => r.fecha === hoy);
  const turnosFuturos = todos.filter((r) => r.fecha > hoy);

  const exportHref = `/api/admin/turnos/export${buildExportQuery({
    servicio: servicio ?? undefined,
    desde,
    hasta: hasta ?? undefined,
  })}`;

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/20 pb-8">
          <p className="mb-1 text-[0.65rem] font-medium uppercase tracking-[0.25em] text-gold">
            Turnos
          </p>
          <h1 className="font-serif text-3xl font-light text-ink-dark md:text-4xl">
            Agenda
          </h1>
          <p className="mt-2 text-sm text-ink-muted">
            Zona horaria: {tz} · Hoy calendario:{" "}
            <span className="font-medium text-ink">{fmtFechaEtiqueta(hoy, tz)}</span>
          </p>
        </div>

        <section className="mb-10 rounded-sm border border-gold/20 bg-white p-5 shadow-sm md:p-6">
          <h2 className="mb-4 font-serif text-lg font-normal text-ink-dark">
            Filtros y exportación
          </h2>
          <form
            method="get"
            className="flex flex-col gap-4 md:flex-row md:flex-wrap md:items-end"
          >
            <div className="min-w-[200px] flex-1">
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Servicio
              </label>
              <select
                name="servicio"
                defaultValue={servicio ?? ""}
                className="w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink"
              >
                <option value="">Todos</option>
                {catalogo.map((n) => (
                  <option key={n} value={n}>
                    {n}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Desde
              </label>
              <input
                type="date"
                name="desde"
                defaultValue={desde}
                className="w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink md:w-auto"
              />
            </div>
            <div>
              <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Hasta (opcional)
              </label>
              <input
                type="date"
                name="hasta"
                defaultValue={hasta ?? ""}
                className="w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink md:w-auto"
              />
            </div>
            <div className="flex flex-wrap gap-2">
              <button
                type="submit"
                className="rounded-sm bg-gold px-5 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-white hover:bg-gold-dark"
              >
                Aplicar
              </button>
              <Link
                href="/admin/turnos"
                className="rounded-sm border border-gold/40 px-5 py-2 text-center text-[0.72rem] font-medium uppercase tracking-wider text-gold-dark hover:bg-cream"
              >
                Limpiar
              </Link>
            </div>
          </form>
          <p className="mt-4 text-sm text-ink-muted">
            <a
              href={exportHref}
              className="font-medium text-gold-dark underline hover:text-gold"
            >
              Descargar CSV
            </a>{" "}
            con los mismos filtros (requiere sesión iniciada en este navegador).
          </p>
        </section>

        <section className="mb-14">
          <h2 className="mb-4 font-serif text-xl font-normal text-ink-dark">
            Hoy en calendario ({turnosHoy.length})
          </h2>
          <AdminTurnosTable rows={turnosHoy} tz={tz} showActions />
        </section>

        <section>
          <h2 className="mb-4 font-serif text-xl font-normal text-ink-dark">
            Después de hoy ({turnosFuturos.length})
          </h2>
          <p className="mb-4 text-sm text-ink-muted">
            Turnos activos con fecha posterior a hoy dentro del filtro aplicado,
            ordenados por fecha y hora.
          </p>
          <AdminTurnosTable rows={turnosFuturos} tz={tz} showActions />
        </section>
      </div>
    </main>
  );
}
