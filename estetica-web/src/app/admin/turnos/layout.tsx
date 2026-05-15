import { cookies } from "next/headers";
import Link from "next/link";
import { redirect } from "next/navigation";
import {
  ADMIN_SESSION_COOKIE,
  verifySignedSessionValue,
} from "@/lib/admin-session";

export default async function AdminTurnosLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const store = await cookies();
  if (!verifySignedSessionValue(store.get(ADMIN_SESSION_COOKIE)?.value)) {
    redirect("/admin/login?next=%2Fadmin%2Fturnos");
  }

  return (
    <div className="min-h-screen bg-cream">
      <div className="border-b border-gold/20 bg-white/90">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3 px-5 py-3 md:px-8">
          <p className="text-[0.65rem] font-medium uppercase tracking-[0.2em] text-gold">
            Panel admin
          </p>
          <div className="flex items-center gap-3">
            <Link
              href="/"
              className="text-[0.7rem] text-ink-muted underline hover:text-gold-dark"
            >
              Sitio público
            </Link>
            <form action="/api/admin/logout" method="post">
              <button
                type="submit"
                className="rounded-sm border border-gold/40 px-3 py-1.5 text-[0.65rem] font-medium uppercase tracking-wide text-gold-dark hover:bg-cream"
              >
                Cerrar sesión
              </button>
            </form>
          </div>
        </div>
      </div>
      {children}
    </div>
  );
}
