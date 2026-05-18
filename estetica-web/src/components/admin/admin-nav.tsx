"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";

const links = [
  { href: "/admin/turnos", label: "Agenda" },
  { href: "/admin/servicios", label: "Servicios" },
  { href: "/admin/paquetes", label: "Paquetes" },
  { href: "/admin/clientes", label: "Clientes" },
  { href: "/admin/caja", label: "Caja" },
] as const;

export function AdminNav() {
  const pathname = usePathname();

  return (
    <nav className="flex flex-wrap gap-2 border-t border-gold/15 pt-3">
      {links.map(({ href, label }) => {
        const active =
          pathname === href || pathname.startsWith(`${href}/`);
        return (
          <Link
            key={href}
            href={href}
            className={`rounded-sm px-4 py-2 text-[0.7rem] font-semibold uppercase tracking-wider transition ${
              active
                ? "bg-gold text-white shadow-sm"
                : "border border-gold/55 bg-white/90 text-gold-dark shadow-sm hover:border-gold hover:bg-white"
            }`}
          >
            {label}
          </Link>
        );
      })}
    </nav>
  );
}
