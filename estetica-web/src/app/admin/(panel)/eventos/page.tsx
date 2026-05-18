import type { Metadata } from "next";
import { AdminEventosManager } from "@/components/admin/admin-eventos-manager";
import { listarEventosAdmin } from "@/lib/eventos-repo";
import { listarSedesActivas } from "@/lib/sedes-repo";
import { listarServiciosAdmin } from "@/lib/servicios-repo";
import { rowToServicioApi } from "@/lib/servicio-format";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata: Metadata = {
  title: "Admin — Eventos | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminEventosPage() {
  const eventosRaw = await listarEventosAdmin();
  const serviciosRaw = await listarServiciosAdmin();

  const servicios = Array.isArray(serviciosRaw)
    ? serviciosRaw.map((r) => ({
        id: r.id,
        ...rowToServicioApi(r),
        parentId: r.parentId,
        esGrupo: r.esGrupo,
      }))
    : [];

  const eventos = Array.isArray(eventosRaw) ? eventosRaw : [];
  const sedesRaw = await listarSedesActivas();
  const sedes = Array.isArray(sedesRaw) ? sedesRaw : [];

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Configuración</p>
          <h1 className={uiPanelTitle}>Eventos</h1>
          <p className={uiPanelDesc}>
            Programá eventos por franja horaria (varios el mismo día). En cada
            franja solo se reservan los servicios del evento; podés asignar
            cliente y precio acordado.
          </p>
        </div>

        {!Array.isArray(eventosRaw) ? (
          <p className="text-sm text-red-700">
            Base de datos no disponible. Aplicá la migración{" "}
            <code className="text-xs">drizzle/0007_availability_events.sql</code> y{" "}
            <code className="text-xs">0008_eventos_franjas_cliente_precio.sql</code> y{" "}
            <code className="text-xs">0010_sedes.sql</code> en Neon.
          </p>
        ) : servicios.length === 0 || sedes.length === 0 ? (
          <p className="text-sm text-ink-muted">
            Primero cargá servicios en{" "}
            <a href="/admin/servicios" className="text-gold-dark underline">
              Servicios
            </a>
            .
          </p>
        ) : (
          <AdminEventosManager
            initialEventos={eventos}
            servicios={servicios}
            sedes={sedes}
          />
        )}
      </div>
    </main>
  );
}
