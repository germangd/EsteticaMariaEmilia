import type { Metadata } from "next";
import { AdminSedesManager } from "@/components/admin/admin-sedes-manager";
import { listarSedesAdmin } from "@/lib/sedes-repo";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata: Metadata = {
  title: "Admin — Sedes | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminSedesPage() {
  const sedesRaw = await listarSedesAdmin();
  const sedes = Array.isArray(sedesRaw) ? sedesRaw : [];

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-3xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Configuración</p>
          <h1 className={uiPanelTitle}>Sedes</h1>
          <p className={uiPanelDesc}>
            Lugares de atención (Ensenada, Bartolomé Bavio, Magdalena). Cada
            sede tiene su horario en Admin → Horarios; los turnos y eventos se
            asocian a una sede.
          </p>
        </div>

        {!Array.isArray(sedesRaw) ? (
          <p className="text-sm text-red-700">
            Base de datos no disponible. Aplicá{" "}
            <code className="text-xs">drizzle/0010_sedes.sql</code> en Neon.
          </p>
        ) : (
          <AdminSedesManager initialSedes={sedes} />
        )}
      </div>
    </main>
  );
}
