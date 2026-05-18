"use client";

import { useEffect } from "react";

/** Al llegar desde agenda (?turno=), scroll al formulario de cobro. */
export function AdminCajaScrollToForm() {
  useEffect(() => {
    if (typeof window === "undefined") return;
    if (!window.location.hash.includes("nueva-venta")) return;
    const el = document.getElementById("nueva-venta");
    if (!el) return;
    requestAnimationFrame(() => {
      el.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }, []);

  return null;
}
