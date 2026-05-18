"use client";

import { useRouter, useSearchParams } from "next/navigation";
import { useState } from "react";
import {
  uiBtnPrimaryFull,
  uiCardSoft,
  uiInput,
  uiLabel,
} from "@/lib/ui-classes";

export function AdminLoginForm() {
  const router = useRouter();
  const sp = useSearchParams();
  const rawNext = sp.get("next")?.trim() || "/admin/turnos";
  const next =
    rawNext.startsWith("/") && !rawNext.startsWith("//")
      ? rawNext
      : "/admin/turnos";
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setPending(true);
    try {
      const r = await fetch("/api/admin/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "same-origin",
        body: JSON.stringify({ password }),
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        error?: string;
      };
      if (!r.ok) {
        setError(data.mensaje ?? "No se pudo iniciar sesión.");
        return;
      }
      router.push(next);
      router.refresh();
    } catch {
      setError("Error de red. Intentá de nuevo.");
    } finally {
      setPending(false);
    }
  }

  return (
    <form onSubmit={(e) => void onSubmit(e)} className={`mx-auto max-w-md ${uiCardSoft}`}>
      <label className={uiLabel}>Contraseña de administración</label>
      <input
        type="password"
        autoComplete="current-password"
        value={password}
        onChange={(e) => setPassword(e.target.value)}
        className={`mb-4 ${uiInput}`}
        placeholder="••••••••"
        required
      />
      {error ? (
        <p className="mb-4 text-sm font-medium text-red-700" role="alert">
          {error}
        </p>
      ) : null}
      <button type="submit" disabled={pending} className={uiBtnPrimaryFull}>
        {pending ? "Ingresando…" : "Ingresar"}
      </button>
    </form>
  );
}
