import { AdminHorariosPorSede } from "@/components/admin/admin-horarios-por-sede";
import { adminMetadata } from "@/config/admin-metadata";
import { listarSedesActivas } from "@/lib/sedes-repo";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata = adminMetadata("Horarios");

export const dynamic = "force-dynamic";

export default async function AdminHorariosPage() {
  const sedesRaw = await listarSedesActivas();
  const sedes = Array.isArray(sedesRaw) ? sedesRaw : [];

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-3xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Configuración</p>
          <h1 className={uiPanelTitle}>Horario de atención</h1>
          <p className={uiPanelDesc}>
            Franjas por sede (mañana, tarde, etc.). Las reservas usan la sede
            elegida y el horario de cada servicio.
          </p>
        </div>

        {!Array.isArray(sedesRaw) ? (
          <p className="text-sm text-red-700">
            Base de datos no disponible. Aplicá{" "}
            <code className="text-xs">drizzle/0010_sedes.sql</code> en Neon.
          </p>
        ) : (
          <AdminHorariosPorSede sedes={sedes} />
        )}
      </div>
    </main>
  );
}
