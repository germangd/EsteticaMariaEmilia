"use client";

import type { ReactNode } from "react";
import {
  alertarCajaNoAbierta,
  haySesionCajaAbierta,
} from "@/lib/caja-sesion-client";

/** Enlace a caja que avisa si no hay sesión abierta antes de navegar. */
export function AdminLinkCobrarCaja({
  href,
  className,
  children,
}: {
  href: string;
  className?: string;
  children: ReactNode;
}) {
  async function onClick(e: React.MouseEvent<HTMLAnchorElement>) {
    e.preventDefault();
    const abierta = await haySesionCajaAbierta();
    if (!abierta) {
      alertarCajaNoAbierta();
    }
    window.location.assign(href);
  }

  return (
    <a href={href} className={className} onClick={(e) => void onClick(e)}>
      {children}
    </a>
  );
}
