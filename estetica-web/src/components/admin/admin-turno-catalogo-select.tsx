"use client";

import { useEffect, useState } from "react";
import {
  ServicioPasosSelect,
  type ServicioPasosOpt,
} from "@/components/admin/servicio-pasos-select";
import type { PaqueteSelectOption } from "@/components/admin/servicio-select-optgroups";
import { fmtPesos } from "@/lib/fmt-pesos";
import { uiLabel, uiSelect } from "@/lib/ui-classes";

export type AdminTurnoSeleccion =
  | { tipo: "servicio"; servicioId: number }
  | { tipo: "paquete"; paqueteId: number };

type TipoItem = "" | "combo" | "servicio";

type Props = {
  servicios: ServicioPasosOpt[];
  paquetes: PaqueteSelectOption[];
  seleccion: AdminTurnoSeleccion | null;
  onSeleccionChange: (s: AdminTurnoSeleccion | null) => void;
  className?: string;
  selectClassName?: string;
  disabled?: boolean;
};

export function AdminTurnoCatalogoSelect({
  servicios,
  paquetes,
  seleccion,
  onSeleccionChange,
  className,
  selectClassName,
  disabled,
}: Props) {
  const puedeCombo = paquetes.length > 0;
  const puedeServicio = servicios.length > 0;

  const [tipo, setTipo] = useState<TipoItem>("");
  const [comboId, setComboId] = useState("");
  const [servicioId, setServicioId] = useState("");

  const selectCls = selectClassName ?? uiSelect;

  useEffect(() => {
    if (!seleccion) {
      if (!puedeCombo && puedeServicio) setTipo("servicio");
      else if (puedeCombo && !puedeServicio) setTipo("combo");
      else setTipo("");
      setComboId("");
      setServicioId("");
      return;
    }
    if (seleccion.tipo === "paquete") {
      setTipo("combo");
      setComboId(String(seleccion.paqueteId));
      setServicioId("");
      return;
    }
    setTipo("servicio");
    setServicioId(String(seleccion.servicioId));
    setComboId("");
  }, [seleccion, puedeCombo, puedeServicio]);

  function onTipoChange(next: TipoItem) {
    setTipo(next);
    setComboId("");
    setServicioId("");
    onSeleccionChange(null);
  }

  function onComboChange(id: string) {
    setComboId(id);
    const n = Number(id);
    if (Number.isFinite(n) && n > 0) {
      onSeleccionChange({ tipo: "paquete", paqueteId: n });
    } else {
      onSeleccionChange(null);
    }
  }

  function onServicioChange(id: string) {
    setServicioId(id);
    const n = Number(id);
    if (Number.isFinite(n) && n > 0) {
      onSeleccionChange({ tipo: "servicio", servicioId: n });
    } else {
      onSeleccionChange(null);
    }
  }

  if (!puedeCombo && !puedeServicio) {
    return (
      <p className="text-sm text-ink-muted">No hay servicios ni combos disponibles.</p>
    );
  }

  return (
    <div className={className ?? "space-y-4"}>
      {puedeCombo && puedeServicio ? (
        <div>
          <label className={uiLabel}>¿Qué cargás?</label>
          <select
            required
            disabled={disabled}
            className={selectCls}
            value={tipo}
            onChange={(e) => onTipoChange(e.target.value as TipoItem)}
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
          value={servicioId}
          onChange={onServicioChange}
          valueMode="id"
          showCupo
          selectClassName={selectCls}
          disabled={disabled}
        />
      ) : null}
    </div>
  );
}
