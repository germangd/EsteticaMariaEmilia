import { cookies } from "next/headers";
import Link from "next/link";
import { redirect } from "next/navigation";
import { AdminNav } from "@/components/admin/admin-nav";
import {
  ADMIN_SESSION_COOKIE,
  verifySignedSessionValue,
} from "@/lib/admin-session";

export default async function AdminPanelLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const store = await cookies();
  if (!verifySignedSessionValue(store.get(ADMIN_SESSION_COOKIE)?.value)) {
    redirect("/admin/login?next=%2Fadmin%2Fturnos");
  }

  return (
    <div className="min-h-screen bg-cream-dark">
      <div className="border-b border-gold/60 bg-panel shadow-sm ring-1 ring-gold/20">
        <div className="mx-auto max-w-6xl px-5 py-3 md:px-8">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <p className="text-[0.65rem] font-bold uppercase tracking-[0.2em] text-gold-dark">
              Panel admin
            </p>
            <div className="flex items-center gap-3">
              <Link
                href="/"
                className="text-[0.7rem] font-medium text-ink underline-offset-2 hover:text-gold-dark"
              >
                Sitio público
              </Link>
              <form action="/api/admin/logout" method="post">
                <button
                  type="submit"
                  className="rounded-sm border border-gold/55 bg-white px-3 py-1.5 text-[0.65rem] font-semibold uppercase tracking-wide text-gold-dark shadow-sm hover:bg-cream"
                >
                  Cerrar sesión
                </button>
              </form>
            </div>
          </div>
          <AdminNav />
        </div>
      </div>
      {children}
    </div>
  );
}
