"use client";

import { TURNOS_ASIGNADOS_ANCHOR } from "@/lib/agenda-rango";
import { usePathname, useSearchParams } from "next/navigation";
import { useEffect } from "react";

/** Tras filtrar la agenda, lleva la vista al bloque de turnos asignados. */
export function AdminTurnosScrollToList() {
  const pathname = usePathname();
  const searchParams = useSearchParams();

  useEffect(() => {
    if (pathname !== "/admin/turnos") return;
    if (typeof window === "undefined") return;
    if (window.location.hash !== `#${TURNOS_ASIGNADOS_ANCHOR}`) return;

    const el = document.getElementById(TURNOS_ASIGNADOS_ANCHOR);
    if (!el) return;

    requestAnimationFrame(() => {
      el.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }, [pathname, searchParams]);

  return null;
}
