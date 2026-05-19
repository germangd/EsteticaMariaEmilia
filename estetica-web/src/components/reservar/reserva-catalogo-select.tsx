"use client";

import { useEffect, useMemo, useState } from "react";
import {
  ServicioPasosSelect,
  type ServicioPasosOpt,
} from "@/components/admin/servicio-pasos-select";
import type { PaqueteSelectOption } from "@/components/admin/servicio-select-optgroups";
import { fmtPesos } from "@/lib/fmt-pesos";
import {
  claveReservaPaquete,
  claveReservaServicio,
  parsearClaveReserva,
} from "@/lib/reserva-claves";
import { uiLabel, uiSelect } from "@/lib/ui-classes";

type TipoReserva = "" | "combo" | "servicio";

type Props = {
  servicios: ServicioPasosOpt[];
  paquetes: PaqueteSelectOption[];
  value: string;
  onChange: (clave: string) => void;
  className?: string;
  selectClassName?: string;
  disabled?: boolean;
};

export function ReservaCatalogoSelect({
  servicios,
  paquetes,
  value,
  onChange,
  className,
  selectClassName,
  disabled,
}: Props) {
  const puedeCombo = paquetes.length > 0;
  const puedeServicio = servicios.length > 0;

  const [tipo, setTipo] = useState<TipoReserva>("");
  const [comboId, setComboId] = useState("");
  const servicioValor = useMemo(() => {
    const p = parsearClaveReserva(value);
    return p?.tipo === "servicio" ? p.nombre : "";
  }, [value]);

  useEffect(() => {
    const parsed = parsearClaveReserva(value);
    if (!parsed) {
      if (!puedeCombo && puedeServicio) setTipo("servicio");
      else if (puedeCombo && !puedeServicio) setTipo("combo");
      else setTipo("");
      setComboId("");
      return;
    }
    if (parsed.tipo === "paquete") {
      setTipo("combo");
      setComboId(String(parsed.id));
      return;
    }
    setTipo("servicio");
    setComboId("");
  }, [value, puedeCombo, puedeServicio]);

  const selectCls = selectClassName ?? uiSelect;

  function onTipoChange(next: TipoReserva) {
    setTipo(next);
    setComboId("");
    onChange("");
  }

  function onComboChange(id: string) {
    setComboId(id);
    const n = Number(id);
    if (Number.isFinite(n) && n > 0) onChange(claveReservaPaquete(n));
    else onChange("");
  }

  if (!puedeCombo && !puedeServicio) {
    return (
      <p className="text-sm text-ink-muted">No hay servicios ni combos disponibles.</p>
    );
  }

  return (
    <div className={className ?? "mb-4 space-y-4"}>
      {puedeCombo && puedeServicio ? (
        <div>
          <label className={uiLabel}>¿Qué querés reservar?</label>
          <select
            required
            disabled={disabled}
            className={selectCls}
            value={tipo}
            onChange={(e) => onTipoChange(e.target.value as TipoReserva)}
          >
            <option value="">Elegí una opción</option>
            <option value="combo">Combo / paquete</option>
            <option value="servicio">Servicio</option>
          </select>
        </div>
      ) : null}

      {tipo === "combo" || (!puedeServicio && puedeCombo) ? (
        <div>
          <label className={uiLabel}>Combo</label>
          <select
            required
            disabled={disabled}
            className={selectCls}
            value={comboId}
            onChange={(e) => onComboChange(e.target.value)}
          >
            <option value="">Elegí un combo</option>
            {paquetes.map((p) => {
              const incluye =
                (p.serviciosIncluidos?.length ?? 0) > 0
                  ? ` — ${p.serviciosIncluidos!.join(", ")}`
                  : "";
              const precio =
                (p.precioPesos ?? 0) > 0 ? ` (${fmtPesos(p.precioPesos!)})` : "";
              const sesiones =
                (p.sesionesTotal ?? 0) > 1
                  ? ` · ${p.sesionesTotal} sesiones`
                  : "";
              return (
                <option key={p.id} value={String(p.id)}>
                  {p.nombre}
                  {sesiones}
                  {precio}
                  {incluye}
                </option>
              );
            })}
          </select>
        </div>
      ) : null}

      {tipo === "servicio" || (!puedeCombo && puedeServicio) ? (
        <ServicioPasosSelect
          servicios={servicios}
          value={servicioValor}
          onChange={(nombre) =>
            onChange(nombre ? claveReservaServicio(nombre) : "")
          }
          valueMode="nombre"
          selectClassName={selectCls}
          disabled={disabled}
        />
      ) : null}
    </div>
  );
}
