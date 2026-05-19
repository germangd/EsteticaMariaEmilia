import type { Metadata } from "next";
import { Suspense } from "react";
import { AdminCajaManager } from "@/components/admin/admin-caja-manager";
import { AdminCajaScrollToForm } from "@/components/admin/admin-caja-scroll";
import type { CatalogoCaja } from "@/lib/caja-repo";
import {
  listarVentasSesion,
  obtenerCatalogoCaja,
  obtenerPrefillCobroPaquete,
  obtenerPrefillCobroTurno,
  obtenerSesionAbierta,
} from "@/lib/caja-repo";
import type { PrefillCobroPaquete, PrefillCobroTurno } from "@/lib/caja-repo";
import { uiPanelDesc, uiPanelKicker, uiPanelTitle } from "@/lib/ui-classes";

export const metadata: Metadata = {
  title: "Admin \u2014 Caja | Mar\u00eda Emilia Est\u00e9tica",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminCajaPage({
  searchParams,
}: {
  searchParams: Promise<{ turno?: string; paquete?: string }>;
}) {
  const sp = await searchParams;
  const turnoId = Number(sp.turno);
  const paqueteId = Number(sp.paquete);
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
            <code className="text-xs">drizzle/0004_caja.sql</code> y{" "}
            <code className="text-xs">0005_services_precio.sql</code> en Neon.
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

  let prefillTurno: PrefillCobroTurno | null = null;
  if (Number.isFinite(turnoId) && turnoId > 0) {
    const prefillRaw = await obtenerPrefillCobroTurno(turnoId);
    if (prefillRaw && typeof prefillRaw === "object" && "appointmentId" in prefillRaw) {
      prefillTurno = prefillRaw;
    }
  }

  let prefillPaquete: PrefillCobroPaquete | null = null;
  if (Number.isFinite(paqueteId) && paqueteId > 0) {
    const prefillRaw = await obtenerPrefillCobroPaquete(paqueteId);
    if (
      prefillRaw &&
      typeof prefillRaw === "object" &&
      "clientPackageId" in prefillRaw
    ) {
      prefillPaquete = prefillRaw;
    }
  }

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/35 pb-8">
          <p className={uiPanelKicker}>{"Facturaci\u00f3n"}</p>
          <h1 className={uiPanelTitle}>Caja</h1>
          <p className={uiPanelDesc}>
            {
              "Abr\u00ed el turno de caja, registr\u00e1 cobros con ticket imprimible y cerr\u00e1 al final del d\u00eda. Todas las ventas quedan guardadas en el historial. Comprobante interno (no factura fiscal AFIP)."
            }
          </p>
        </div>

        <Suspense fallback={null}>
          <AdminCajaScrollToForm />
        </Suspense>
        <AdminCajaManager
          initialSesion={sesion}
          initialVentas={ventas}
          catalogo={catalogo}
          initialPrefillTurno={prefillTurno}
          initialPrefillPaquete={prefillPaquete}
        />
      </div>
    </main>
  );
}
