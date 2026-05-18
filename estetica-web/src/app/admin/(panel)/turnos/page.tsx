import type { Metadata } from "next";
import Link from "next/link";
import { getAppTimeZone, hoyIsoEnZona } from "@/lib/agenda";
import {
  buildTurnosQuery,
  parseVistaAgenda,
  rangoAgenda,
  type VistaAgenda,
} from "@/lib/agenda-rango";
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
  title: "Admin ? Turnos | Mar?a Emilia Est?tica",
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

const VISTAS: { id: VistaAgenda; label: string }[] = [
  { id: "dia", label: "D?a" },
  { id: "semana", label: "Semana" },
  { id: "mes", label: "Mes" },
  { id: "fecha", label: "Fecha" },
];

function tabClass(active: boolean): string {
  return `rounded-sm px-4 py-2 text-[0.7rem] font-semibold uppercase tracking-wider transition ${
    active
      ? "bg-gold text-white shadow-sm"
      : "border border-gold/55 bg-white/90 text-gold-dark hover:bg-cream"
  }`;
}

export default async function AdminTurnosPage({
  searchParams,
}: {
  searchParams: Promise<{
    vista?: string;
    ref?: string;
    servicio?: string;
    desde?: string;
    hasta?: string;
  }>;
}) {
  const sp = await searchParams;
  const tz = getAppTimeZone();
  const hoy = hoyIsoEnZona(tz);

  const vista = parseVistaAgenda(sp.vista);
  let ref = sp.ref?.trim() ?? hoy;
  if (!isIsoDate(ref)) ref = hoy;

  const servicio = sp.servicio?.trim() || null;

  let { desde, hasta, etiqueta } = rangoAgenda({
    vista,
    refIso: ref,
    tz,
    hoyIso: hoy,
  });

  if (isIsoDate(sp.desde) && !sp.vista) {
    desde = sp.desde;
    hasta = isIsoDate(sp.hasta) ? sp.hasta : sp.desde;
    etiqueta =
      desde === hasta
        ? fmtFechaEtiqueta(desde, tz)
        : `${fmtFechaEtiqueta(desde, tz)} ? ${fmtFechaEtiqueta(hasta, tz)}`;
  }

  if (hasta < desde) {
    return (
      <main className="mx-auto max-w-6xl px-5 py-12 md:px-8">
        <p className="text-sm text-red-700">Rango de fechas inv?lido.</p>
        <Link
          href="/admin/turnos"
          className="mt-4 inline-block text-gold-dark underline"
        >
          Volver a la agenda
        </Link>
      </main>
    );
  }

  const catalogo = await listarNombresServiciosCatalogo();
  const rowsServ = await listarServiciosAdmin();
  const serviciosAdmin = Array.isArray(rowsServ)
    ? rowsServ.map((r) => ({ id: r.id, ...rowToServicioApi(r) }))
    : [];

  const turnos = await listarTurnosActivosFiltrados({
    fechaDesde: desde,
    fechaHasta: hasta,
    servicioNombre: servicio,
  });

  const exportHref = `/api/admin/turnos/export?servicio=${encodeURIComponent(servicio ?? "")}&desde=${desde}&hasta=${hasta}`;

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Turnos</p>
          <h1 className={uiPanelTitle}>Agenda</h1>
          <p className={uiPanelDesc}>
            Zona horaria: {tz} ? Hoy:{" "}
            <span className="font-semibold text-ink-dark">
              {fmtFechaEtiqueta(hoy, tz)}
            </span>
          </p>
        </div>

        <section className={`mb-10 ${uiCard}`}>
          <h2 className={uiSubsectionTitle}>Cargar turno manual</h2>
          <p className="mb-4 text-sm font-medium text-ink">
            Para reservas por tel?fono o WhatsApp. Respeta el cupo configurado en
            cada servicio.
          </p>
          <AdminCargarTurnoForm
            servicios={serviciosAdmin}
            fechaDefault={hoy}
          />
        </section>

        <section className={`mb-10 ${uiCard}`}>
          <h2 className={uiSubsectionTitle}>Turnos asignados</h2>
          <p className="mb-4 text-sm font-medium text-ink">
            Un solo listado seg?n el per?odo elegido. Pod?s filtrar por servicio.
          </p>

          <div className="mb-4 flex flex-wrap gap-2">
            {VISTAS.map((v) => (
              <Link
                key={v.id}
                href={`/admin/turnos${buildTurnosQuery({
                  vista: v.id,
                  ref,
                  servicio: servicio ?? undefined,
                })}`}
                className={tabClass(vista === v.id)}
              >
                {v.id === "dia" && ref === hoy ? "Hoy" : v.label}
              </Link>
            ))}
          </div>

          <form
            method="get"
            className="flex flex-col gap-4 border-t border-gold/20 pt-4 md:flex-row md:flex-wrap md:items-end"
          >
            <input type="hidden" name="vista" value={vista} />
            <div>
              <label className={uiLabel}>
                {vista === "semana"
                  ? "D?a de referencia (semana)"
                  : vista === "mes"
                    ? "D?a de referencia (mes)"
                    : "Fecha"}
              </label>
              <input
                type="date"
                name="ref"
                defaultValue={ref}
                className={`md:w-auto ${uiInput}`}
              />
            </div>
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
            <div className="flex flex-wrap gap-2">
              <button type="submit" className={uiBtnPrimary}>
                Aplicar
              </button>
              <Link
                href={`/admin/turnos${buildTurnosQuery({ vista: "dia", ref: hoy })}`}
                className={uiBtnSecondary}
              >
                Hoy
              </Link>
              <Link href="/admin/turnos" className={uiBtnSecondary}>
                Limpiar
              </Link>
            </div>
          </form>

          <p className="mt-4 text-sm font-medium text-ink">
            <a
              href={exportHref}
              className="font-semibold text-gold-dark underline hover:text-gold"
            >
              Descargar CSV
            </a>{" "}
            del período mostrado.
          </p>
        </section>

        <section>
          <h2 className={uiSectionTitle}>
            {turnos.length} turno{turnos.length === 1 ? "" : "s"}
          </h2>
          <p className="mb-4 text-sm font-medium text-ink-muted">{etiqueta}</p>
          <AdminTurnosTable rows={turnos} tz={tz} showActions />
        </section>
      </div>
    </main>
  );
}
