"use client";

import { usePathname, useRouter } from "next/navigation";
import { uiLabel, uiSelect } from "@/lib/ui-classes";

const links = [
  { href: "/admin/turnos", label: "Agenda" },
  { href: "/admin/servicios", label: "Servicios" },
  { href: "/admin/sedes", label: "Sedes" },
  { href: "/admin/horarios", label: "Horarios" },
  { href: "/admin/paquetes", label: "Paquetes" },
  { href: "/admin/eventos", label: "Eventos" },
  { href: "/admin/clientes", label: "Clientes" },
  { href: "/admin/caja", label: "Caja" },
  { href: "/admin/reportes", label: "Reportes" },
] as const;

function hrefActivo(pathname: string, href: string): boolean {
  return pathname === href || pathname.startsWith(`${href}/`);
}

export function AdminNav() {
  const pathname = usePathname();
  const router = useRouter();

  const current =
    links.find((l) => hrefActivo(pathname, l.href))?.href ?? "/admin/turnos";

  return (
    <nav className="border-t border-gold/15 pt-3">
      <label htmlFor="admin-nav-select" className={`${uiLabel} mb-1.5`}>
        Sección del panel
      </label>
      <select
        id="admin-nav-select"
        className={uiSelect}
        value={current}
        onChange={(e) => router.push(e.target.value)}
      >
        {links.map(({ href, label }) => (
          <option key={href} value={href}>
            {label}
          </option>
        ))}
      </select>
    </nav>
  );
}
