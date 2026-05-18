import type { Metadata } from "next";
import Link from "next/link";
import { getAppTimeZone, hoyIsoEnZona } from "@/lib/agenda";
import { AdminCargarTurnoForm } from "@/components/admin/admin-cargar-turno-form";
import { AdminTurnosTable } from "@/components/admin/admin-turnos-table";
import { rowToServicioApi } from "@/lib/servicio-format";
import { listarServiciosAdmin } from "@/lib/servicios-repo";
import {
  listarNombresServiciosCatalogo,
  listarTurnosActivosFiltrados,
} from "@/lib/turnos-repo";
import { DateTime } from "luxon";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
  uiPanelDesc,
  uiPanelKicker,
  uiPanelTitle,
  uiSectionTitle,
  uiSelect,
  uiSubsectionTitle,
} from "@/lib/ui-classes";

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
  const rowsServ = await listarServiciosAdmin();
  const serviciosAdmin = Array.isArray(rowsServ)
    ? rowsServ.map((r) => ({ id: r.id, ...rowToServicioApi(r) }))
    : [];

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
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Turnos</p>
          <h1 className={uiPanelTitle}>Agenda</h1>
          <p className={uiPanelDesc}>
            Zona horaria: {tz} · Hoy calendario:{" "}
            <span className="font-semibold text-ink-dark">{fmtFechaEtiqueta(hoy, tz)}</span>
          </p>
        </div>

        <section className={`mb-10 ${uiCard}`}>
          <h2 className={uiSubsectionTitle}>Cargar turno manual</h2>
          <p className="mb-4 text-sm font-medium text-ink">
            Para reservas por teléfono o WhatsApp. Respeta el cupo configurado en
            cada servicio.
          </p>
          <AdminCargarTurnoForm
            servicios={serviciosAdmin}
            fechaDefault={hoy}
          />
        </section>

        <section className={`mb-10 ${uiCard}`}>
          <h2 className={uiSubsectionTitle}>Filtros y exportación</h2>
          <form
            method="get"
            className="flex flex-col gap-4 md:flex-row md:flex-wrap md:items-end"
          >
            <div className="min-w-[200px] flex-1">
              <label className={uiLabel}>Servicio</label>
              <select
                name="servicio"
                defaultValue={servicio ?? ""}
                className={uiSelect}
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
              <label className={uiLabel}>Desde</label>
              <input
                type="date"
                name="desde"
                defaultValue={desde}
                className={`md:w-auto ${uiInput}`}
              />
            </div>
            <div>
              <label className={uiLabel}>Hasta (opcional)</label>
              <input
                type="date"
                name="hasta"
                defaultValue={hasta ?? ""}
                className={`md:w-auto ${uiInput}`}
              />
            </div>
            <div className="flex flex-wrap gap-2">
              <button type="submit" className={uiBtnPrimary}>
                Aplicar
              </button>
              <Link href="/admin/turnos" className={uiBtnSecondary}>
                Limpiar
              </Link>
            </div>
          </form>
          <p className="mt-4 text-sm font-medium text-ink">
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
          <h2 className={uiSectionTitle}>Hoy en calendario ({turnosHoy.length})</h2>
          <AdminTurnosTable rows={turnosHoy} tz={tz} showActions />
        </section>

        <section>
          <h2 className={uiSectionTitle}>Después de hoy ({turnosFuturos.length})</h2>
          <p className="mb-4 text-sm font-medium text-ink">
            Turnos activos con fecha posterior a hoy dentro del filtro aplicado,
            ordenados por fecha y hora.
          </p>
          <AdminTurnosTable rows={turnosFuturos} tz={tz} showActions />
        </section>
      </div>
    </main>
  );
}
