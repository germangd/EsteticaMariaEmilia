"use client";

import {
  ServicioSelectOptgroups,
  type PaqueteSelectOption,
} from "@/components/admin/servicio-select-optgroups";

type ServicioOpt = {
  id: number;
  nombre: string;
  parentId?: number | null;
  categoriaNombre?: string | null;
  capacidad?: number;
};

type Props = {
  servicios: ServicioOpt[];
  paquetes: PaqueteSelectOption[];
  value: string;
  onChange: (clave: string) => void;
  className?: string;
  required?: boolean;
  disabled?: boolean;
};

export function ReservaCatalogoSelect({
  servicios,
  paquetes,
  value,
  onChange,
  className,
  required,
  disabled,
}: Props) {
  const tieneServicios = servicios.length > 0;
  const tienePaquetes = paquetes.length > 0;

  if (!tieneServicios && !tienePaquetes) {
    return (
      <p className="text-sm text-ink-muted">No hay servicios ni combos disponibles.</p>
    );
  }

  return (
    <ServicioSelectOptgroups
      servicios={servicios.map((s) => ({
        id: s.id,
        nombre: s.nombre,
        parentId: s.parentId ?? null,
        esGrupo: false,
        capacidad: s.capacidad,
        categoriaNombre: s.categoriaNombre,
      }))}
      paquetes={tienePaquetes ? paquetes : undefined}
      value={value}
      onChange={onChange}
      className={className}
      required={required}
      placeholder={
        tienePaquetes && tieneServicios
          ? "Eleg\u00ed servicio o combo"
          : tienePaquetes
            ? "Eleg\u00ed un combo"
            : "Eleg\u00ed un servicio"
      }
      valueMode="clave"
    />
  );
}
