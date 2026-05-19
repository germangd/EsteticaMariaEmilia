import type { Metadata } from "next";
import Link from "next/link";
import { AdminCajaTicket } from "@/components/admin/admin-caja-ticket";
import { obtenerVentaDetalle } from "@/lib/caja-repo";

export const metadata: Metadata = {
  title: "Ticket | Caja admin",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminCajaTicketPage({
  params,
  searchParams,
}: {
  params: Promise<{ id: string }>;
  searchParams: Promise<{ imprimir?: string }>;
}) {
  const { id: idStr } = await params;
  const sp = await searchParams;
  const autoPrint = sp.imprimir === "1";
  const id = Number(idStr);

  if (!Number.isFinite(id) || id < 1) {
    return (
      <main className="py-12 text-center">
        <p className="text-sm text-red-700">{"ID de venta inv\u00e1lido."}</p>
        <Link href="/admin/caja" className="mt-4 inline-block text-gold-dark underline">
          Volver a caja
        </Link>
      </main>
    );
  }

  const venta = await obtenerVentaDetalle(id);

  if (venta && typeof venta === "object" && "ok" in venta) {
    return (
      <main className="py-12 text-center">
        <p className="text-sm text-red-700">Base de datos no disponible.</p>
      </main>
    );
  }

  if (!venta) {
    return (
      <main className="py-12 text-center">
        <p className="text-sm text-red-700">Venta no encontrada.</p>
        <Link href="/admin/caja" className="mt-4 inline-block text-gold-dark underline">
          Volver a caja
        </Link>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-surface py-8 print:bg-white print:py-0">
      <AdminCajaTicket venta={venta} autoPrint={autoPrint} />
    </main>
  );
}
