import { AdminReportesManager } from "@/components/admin/admin-reportes-manager";
import { adminMetadata } from "@/config/admin-metadata";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata = adminMetadata("Reportes");

export const dynamic = "force-dynamic";

export default function AdminReportesPage() {
  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>Facturación</p>
          <h1 className={uiPanelTitle}>Reportes</h1>
          <p className={uiPanelDesc}>
            Resumen de ventas por día, semana o mes, e historial de turnos de
            caja con arqueo por método de pago.
          </p>
        </div>
        <AdminReportesManager />
      </div>
    </main>
  );
}
