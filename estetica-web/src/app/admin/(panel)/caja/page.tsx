import type { Metadata } from "next";
import { AdminCajaManager } from "@/components/admin/admin-caja-manager";
import type { CatalogoCaja } from "@/lib/caja-repo";
import {
  listarVentasSesion,
  obtenerCatalogoCaja,
  obtenerSesionAbierta,
} from "@/lib/caja-repo";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata: Metadata = {
  title: "Admin \u2014 Caja | Mar\u00eda Emilia Est\u00e9tica",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminCajaPage() {
  const sesionRaw = await obtenerSesionAbierta();
  const catalogoRaw = await obtenerCatalogoCaja();

  if (
    (sesionRaw && typeof sesionRaw === "object" && "ok" in sesionRaw) ||
    ("ok" in catalogoRaw && catalogoRaw.ok === false)
  ) {
    return (
      <main className="pb-16 pt-8">
        <div className="mx-auto max-w-6xl px-5 md:px-8">
          <p className="text-sm text-red-700">
            {"Base de datos no disponible. Si es la primera vez, aplic\u00e1 la migraci\u00f3n "}
            <code className="text-xs">drizzle/0004_caja.sql</code> en Neon.
          </p>
        </div>
      </main>
    );
  }

  const sesion =
    sesionRaw && typeof sesionRaw === "object" && "id" in sesionRaw
      ? sesionRaw
      : null;

  const ventasRaw = sesion
    ? await listarVentasSesion(sesion.id)
    : [];
  const ventas = Array.isArray(ventasRaw) ? ventasRaw : [];
  const catalogo = catalogoRaw as CatalogoCaja;

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>{"Facturaci\u00f3n"}</p>
          <h1 className={uiPanelTitle}>Caja</h1>
          <p className={uiPanelDesc}>
            {"Abr\u00ed el turno de caja, registr\u00e1 cobros con ticket imprimible y cerr\u00e1 al final del d\u00eda. Comprobante interno (no factura fiscal AFIP)."}
          </p>
        </div>

        <AdminCajaManager
          initialSesion={sesion}
          initialVentas={ventas}
          catalogo={catalogo}
        />
      </div>
    </main>
  );
}
