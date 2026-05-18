import type { Metadata } from "next";
import { Suspense } from "react";
import { AdminClientesPanel } from "@/components/admin/admin-clientes-panel";
import {
  listarClientesResumen,
  normalizarTelefono,
  obtenerClienteDetalle,
} from "@/lib/clientes-repo";

export const metadata: Metadata = {
  title: "Admin — Clientes | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminClientesPage({
  searchParams,
}: {
  searchParams: Promise<{ q?: string; tel?: string }>;
}) {
  const sp = await searchParams;
  const q = sp.q?.trim() ?? "";
  const telRaw = sp.tel?.trim() ?? "";
  const telefono = telRaw ? normalizarTelefono(telRaw) : null;

  const clientesRaw = await listarClientesResumen(q || undefined);
  const clientes = Array.isArray(clientesRaw) ? clientesRaw : [];

  let detalle = null;
  if (telefono && Array.isArray(clientesRaw)) {
    const d = await obtenerClienteDetalle(telefono);
    if (!("ok" in d)) detalle = d;
  }

  return (
    <main className="pb-16 pt-8">
      <div className="mx-auto max-w-6xl px-5 md:px-8">
        <div className="mb-8 border-b border-gold/20 pb-8">
          <p className="mb-1 text-[0.65rem] font-medium uppercase tracking-[0.25em] text-gold">
            Clientes
          </p>
          <h1 className="font-serif text-3xl font-light text-ink-dark md:text-4xl">
            Historial por cliente
          </h1>
          <p className="mt-2 text-sm text-ink-muted">
            Se arma solo con los turnos de la web o carga manual y los paquetes
            vendidos. Podés agregar notas en la ficha.
          </p>
        </div>

        {!Array.isArray(clientesRaw) ? (
          <div className="text-sm text-red-700">
            <p className="mb-3">
              Falta la tabla de fichas. En Neon → SQL Editor, pegá solo esto
              (sin comillas raras ni la ruta del archivo):
            </p>
            <pre className="overflow-x-auto rounded-sm border border-gold/20 bg-cream/80 p-4 text-xs text-ink-dark">
{`CREATE TABLE IF NOT EXISTS client_profiles (
  telefono text PRIMARY KEY NOT NULL,
  nombre text,
  email text,
  notas text,
  updated_at timestamptz DEFAULT now() NOT NULL
);`}
            </pre>
          </div>
        ) : (
          <Suspense
            fallback={
              <p className="text-sm text-ink-muted">Cargando clientes…</p>
            }
          >
            <AdminClientesPanel
              initialClientes={clientes}
              initialDetalle={detalle}
              telefonoSeleccionado={telefono}
              busquedaInicial={q}
            />
          </Suspense>
        )}
      </div>
    </main>
  );
}
