import type { Metadata } from "next";
import { AdminServiciosManager } from "@/components/admin/admin-servicios-manager";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";
import { listarServiciosAdmin } from "@/lib/servicios-repo";
import { rowToServicioApi } from "@/lib/servicio-format";

export const metadata: Metadata = {
  title: "Admin — Servicios | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminServiciosPage() {
  const rows = await listarServiciosAdmin();
  const servicios =
    Array.isArray(rows) ?
      rows.map((r) => ({
        id: r.id,
        ...rowToServicioApi(r),
      }))
    : [];

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Configuración</p>
          <h1 className={uiPanelTitle}>Servicios</h1>
          <p className={uiPanelDesc}>
            Lo que cargues acá aparece en la web de reservas. El cupo define
            cuántos clientes pueden tomar el mismo horario.
          </p>
        </div>

        {!Array.isArray(rows) ? (
          <p className="text-sm text-red-700">Base de datos no disponible.</p>
        ) : (
          <AdminServiciosManager initialServicios={servicios} />
        )}
      </div>
    </main>
  );
}
