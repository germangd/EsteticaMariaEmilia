import { Suspense } from "react";
import { cookies } from "next/headers";
import type { Metadata } from "next";
import Link from "next/link";
import { redirect } from "next/navigation";
import { AdminLoginForm } from "@/components/admin/admin-login-form";
import {
  ADMIN_SESSION_COOKIE,
  adminAuthConfigured,
  verifySignedSessionValue,
} from "@/lib/admin-session";

export const metadata: Metadata = {
  title: "Admin — Ingreso | María Emilia Estética",
  robots: { index: false, follow: false },
};

export const dynamic = "force-dynamic";

export default async function AdminLoginPage() {
  const store = await cookies();
  if (verifySignedSessionValue(store.get(ADMIN_SESSION_COOKIE)?.value)) {
    redirect("/admin/turnos");
  }

  if (!adminAuthConfigured()) {
    return (
      <main className="mx-auto min-h-screen max-w-xl px-5 py-16">
        <h1 className="mb-3 font-serif text-2xl font-light text-ink-dark">
          Administración
        </h1>
        <p className="text-sm leading-relaxed text-ink-muted">
          Configurá en el servidor (por ejemplo{" "}
          <code className="rounded bg-white px-1 text-xs">.env.local</code> y en
          Vercel):
        </p>
        <ul className="mt-4 list-disc space-y-2 pl-5 text-sm text-ink-muted">
          <li>
            <code className="text-xs">ADMIN_SESSION_SECRET</code> o bien{" "}
            <code className="text-xs">ADMIN_SECRET</code> — clave larga y
            aleatoria solo para el servidor (firma de la cookie de sesión).
          </li>
          <li>
            <code className="text-xs">ADMIN_PASSWORD</code> — contraseña que
            vas a escribir acá para entrar.
          </li>
        </ul>
        <p className="mt-6 text-sm">
          <Link href="/" className="text-gold-dark underline hover:text-gold">
            Volver al inicio
          </Link>
        </p>
      </main>
    );
  }

  return (
    <main className="mx-auto min-h-screen max-w-lg px-5 py-16">
      <h1 className="mb-2 font-serif text-2xl font-light text-ink-dark">
        Ingreso — Administración
      </h1>
      <p className="mb-8 text-sm text-ink-muted">
        Sesión segura por cookie (no uses la contraseña en la URL).
      </p>
      <Suspense
        fallback={
          <div className="rounded-sm border border-gold/20 bg-white p-8 text-sm text-ink-muted">
            Cargando formulario…
          </div>
        }
      >
        <AdminLoginForm />
      </Suspense>
      <p className="mt-8 text-center text-sm">
        <Link href="/" className="text-gold-dark underline hover:text-gold">
          Volver al inicio
        </Link>
      </p>
    </main>
  );
}
