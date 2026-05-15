import type { Metadata } from "next";
import { AdminServiciosManager } from "@/components/admin/admin-servicios-manager";
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
        <div className="mb-8 border-b border-gold/20 pb-8">
          <p className="mb-1 text-[0.65rem] font-medium uppercase tracking-[0.25em] text-gold">
            Configuración
          </p>
          <h1 className="font-serif text-3xl font-light text-ink-dark md:text-4xl">
            Servicios
          </h1>
          <p className="mt-2 text-sm text-ink-muted">
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
