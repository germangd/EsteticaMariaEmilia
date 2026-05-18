import type { Metadata } from "next";
import { AdminPaquetesManager } from "@/components/admin/admin-paquetes-manager";
import { listarServiciosAdmin } from "@/lib/servicios-repo";
import { rowToServicioApi } from "@/lib/servicio-format";
import { listarAsignacionesPaquete, listarPaquetesAdmin } from "@/lib/paquetes-repo";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata: Metadata = {
  title: "Admin — Paquetes | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminPaquetesPage() {
  const paquetesRaw = await listarPaquetesAdmin();
  const asignRaw = await listarAsignacionesPaquete(true);
  const serviciosRaw = await listarServiciosAdmin();

  const paquetes = Array.isArray(paquetesRaw) ? paquetesRaw : [];
  const asignaciones = Array.isArray(asignRaw) ? asignRaw : [];
  const servicios = Array.isArray(serviciosRaw)
    ? serviciosRaw.map((r) => ({ id: r.id, ...rowToServicioApi(r) }))
    : [];

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Configuración</p>
          <h1 className={uiPanelTitle}>Paquetes</h1>
          <p className={uiPanelDesc}>
            Armá combos con varias sesiones, vendelos en el salón y controlá
            cuántas quedan por cliente. La web pública sigue reservando servicios
            sueltos.
          </p>
        </div>

        {!Array.isArray(paquetesRaw) ? (
          <p className="text-sm text-red-700">
            Base de datos no disponible. Si es la primera vez, aplicá la
            migración{" "}
            <code className="text-xs">drizzle/0002_packages.sql</code> en Neon o{" "}
            <code className="text-xs">npm run db:push</code>.
          </p>
        ) : servicios.length === 0 ? (
          <p className="text-sm text-ink-muted">
            Primero cargá servicios en{" "}
            <a href="/admin/servicios" className="text-gold-dark underline">
              Servicios
            </a>
            .
          </p>
        ) : (
          <AdminPaquetesManager
            initialPaquetes={paquetes}
            initialAsignaciones={asignaciones}
            servicios={servicios}
          />
        )}
      </div>
    </main>
  );
}
