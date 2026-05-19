"use client";

import { useEffect, useRef } from "react";
import { alertarCajaNoAbierta } from "@/lib/caja-sesion-client";

/** Aviso al llegar a Caja para cobrar sin sesión abierta (?turno=, ?paquete=, #nueva-venta). */
function hayIntentoCobroEnUrl(): boolean {
  const p = new URLSearchParams(window.location.search);
  return (
    Number(p.get("turno")) > 0 ||
    Number(p.get("paquete")) > 0 ||
    window.location.hash.includes("nueva-venta")
  );
}

export function AdminCajaAlertaSinSesion({
  sesionAbierta,
  intentoCobro,
}: {
  sesionAbierta: boolean;
  intentoCobro: boolean;
}) {
  const mostradoRef = useRef(false);

  useEffect(() => {
    const quiereCobrar = intentoCobro || hayIntentoCobroEnUrl();
    if (sesionAbierta || !quiereCobrar || mostradoRef.current) return;
    mostradoRef.current = true;
    alertarCajaNoAbierta();
    const estado = document.getElementById("estado-caja");
    estado?.scrollIntoView({ behavior: "smooth", block: "start" });
  }, [sesionAbierta, intentoCobro]);

  return null;
}
