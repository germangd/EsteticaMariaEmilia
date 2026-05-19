"use client";

import { useRouter } from "next/navigation";
import { useState } from "react";
import { uiBtnPrimary } from "@/lib/ui-classes";

export function AdminConfirmarAnticipoButton({
  appointmentId,
  disabled,
}: {
  appointmentId: number;
  disabled?: boolean;
}) {
  const router = useRouter();
  const [pending, setPending] = useState(false);
  const [msg, setMsg] = useState<string | null>(null);

  async function onConfirmar() {
    if (
      !window.confirm(
        "¿Confirmar que recibiste el anticipo? Se enviará el mail de turno confirmado al cliente."
      )
    ) {
      return;
    }
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch("/api/admin/turnos/confirmar-anticipo", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: appointmentId }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo confirmar.");
        return;
      }
      router.refresh();
    } catch {
      setMsg("Error de red.");
    } finally {
      setPending(false);
    }
  }

  return (
    <span className="inline-flex flex-col items-end gap-1">
      <button
        type="button"
        disabled={disabled || pending}
        onClick={() => void onConfirmar()}
        className={`${uiBtnPrimary} !px-2 !py-1 text-[0.65rem]`}
      >
        {pending ? "…" : "Confirmar anticipo"}
      </button>
      {msg ? <span className="text-xs text-red-700">{msg}</span> : null}
    </span>
  );
}
