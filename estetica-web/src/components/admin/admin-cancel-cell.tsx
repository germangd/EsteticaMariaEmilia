"use client";

import { useState } from "react";

export function AdminCancelCell({
  codigo,
  compact,
}: {
  codigo: string;
  compact?: boolean;
}) {
  const [pending, setPending] = useState(false);

  async function onCancel() {
    if (!window.confirm("¿Cancelar este turno? El cliente verá el código como no válido si ya lo tenía guardado.")) {
      return;
    }
    setPending(true);
    try {
      const r = await fetch("/api/admin/turnos/cancelar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "same-origin",
        body: JSON.stringify({ codigo }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok) {
        window.alert(data.mensaje ?? "No se pudo cancelar.");
        return;
      }
      window.location.reload();
    } finally {
      setPending(false);
    }
  }

  return (
    <button
      type="button"
      disabled={pending}
      onClick={() => void onCancel()}
      className={
        compact
          ? "rounded-sm border border-red-300/80 bg-white px-1 py-0.5 text-[0.55rem] font-medium uppercase leading-none tracking-wide text-red-800 transition hover:bg-red-50 disabled:opacity-50"
          : "rounded-sm border border-red-300/80 bg-white px-2 py-1 text-[0.65rem] font-medium uppercase tracking-wide text-red-800 transition hover:bg-red-50 disabled:opacity-50"
      }
    >
      {pending ? "…" : "Cancelar"}
    </button>
  );
}
