"use client";

import { useEffect, useMemo, useState } from "react";
import type { PaqueteSelectOption } from "@/components/admin/servicio-select-optgroups";
import { fmtPesos } from "@/lib/fmt-pesos";
import {
  claveReservaPaquete,
  claveReservaServicio,
  parsearClaveReserva,
} from "@/lib/reserva-claves";
import { uiLabel, uiSelect } from "@/lib/ui-classes";

const SUELTOS_KEY = "__sueltos__";

type ServicioOpt = {
  id: number;
  nombre: string;
  parentId?: number | null;
  categoriaNombre?: string | null;
};

type TipoReserva = "" | "combo" | "servicio";

type Props = {
  servicios: ServicioOpt[];
  paquetes: PaqueteSelectOption[];
  value: string;
  onChange: (clave: string) => void;
  className?: string;
  selectClassName?: string;
  disabled?: boolean;
};

function etiquetaCategoria(key: string): string {
  return key === SUELTOS_KEY ? "Otros servicios" : key;
}

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
  const [categoriaKey, setCategoriaKey] = useState("");
  const [servicioNombre, setServicioNombre] = useState("");

  const categorias = useMemo(() => {
    const map = new Map<string, ServicioOpt[]>();
    for (const s of servicios) {
      const key = s.categoriaNombre?.trim() || SUELTOS_KEY;
      const arr = map.get(key) ?? [];
      arr.push(s);
      map.set(key, arr);
    }
    return [...map.entries()]
      .map(([key, items]) => ({
        key,
        label: etiquetaCategoria(key),
        items: items.sort((a, b) =>
          a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" })
        ),
      }))
      .sort((a, b) =>
        a.label.localeCompare(b.label, "es", { sensitivity: "base" })
      );
  }, [servicios]);

  const serviciosEnCategoria = useMemo(() => {
    const cat = categorias.find((c) => c.key === categoriaKey);
    return cat?.items ?? [];
  }, [categorias, categoriaKey]);

  const omitirCategoria =
    tipo === "servicio" && categorias.length === 1;

  useEffect(() => {
    const parsed = parsearClaveReserva(value);
    if (!parsed) {
      if (!puedeCombo && puedeServicio) setTipo("servicio");
      else if (puedeCombo && !puedeServicio) setTipo("combo");
      else setTipo("");
      setComboId("");
      setCategoriaKey("");
      setServicioNombre("");
      return;
    }
    if (parsed.tipo === "paquete") {
      setTipo("combo");
      setComboId(String(parsed.id));
      setCategoriaKey("");
      setServicioNombre("");
      return;
    }
    setTipo("servicio");
    setComboId("");
    const s = servicios.find((x) => x.nombre === parsed.nombre);
    const key = s?.categoriaNombre?.trim() || SUELTOS_KEY;
    setCategoriaKey(key);
    setServicioNombre(parsed.nombre);
  }, [value, servicios, puedeCombo, puedeServicio]);

  useEffect(() => {
    if (tipo === "servicio" && omitirCategoria && categorias[0]) {
      setCategoriaKey(categorias[0].key);
    }
  }, [tipo, omitirCategoria, categorias]);

  const selectCls = selectClassName ?? uiSelect;

  function onTipoChange(next: TipoReserva) {
    setTipo(next);
    setComboId("");
    setCategoriaKey("");
    setServicioNombre("");
    onChange("");
  }

  function onComboChange(id: string) {
    setComboId(id);
    const n = Number(id);
    if (Number.isFinite(n) && n > 0) onChange(claveReservaPaquete(n));
    else onChange("");
  }

  function onCategoriaChange(key: string) {
    setCategoriaKey(key);
    setServicioNombre("");
    onChange("");
  }

  function onServicioChange(nombre: string) {
    setServicioNombre(nombre);
    if (nombre) onChange(claveReservaServicio(nombre));
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
        <>
          {!omitirCategoria ? (
            <div>
              <label className={uiLabel}>Categoría</label>
              <select
                required
                disabled={disabled}
                className={selectCls}
                value={categoriaKey}
                onChange={(e) => onCategoriaChange(e.target.value)}
              >
                <option value="">Elegí una categoría</option>
                {categorias.map((c) => (
                  <option key={c.key} value={c.key}>
                    {c.label}
                  </option>
                ))}
              </select>
            </div>
          ) : null}

          {(omitirCategoria || categoriaKey) && serviciosEnCategoria.length > 0 ? (
            <div>
              <label className={uiLabel}>
                {categoriaKey === SUELTOS_KEY || omitirCategoria
                  ? "Servicio"
                  : "Sub-servicio"}
              </label>
              <select
                required
                disabled={disabled}
                className={selectCls}
                value={servicioNombre}
                onChange={(e) => onServicioChange(e.target.value)}
              >
                <option value="">
                  {omitirCategoria
                    ? "Elegí un servicio"
                    : "Elegí el servicio"}
                </option>
                {serviciosEnCategoria.map((s) => (
                  <option key={s.id} value={s.nombre}>
                    {s.nombre}
                  </option>
                ))}
              </select>
            </div>
          ) : categoriaKey && serviciosEnCategoria.length === 0 ? (
            <p className="text-sm text-ink-muted">
              No hay servicios en esta categoría.
            </p>
          ) : null}
        </>
      ) : null}
    </div>
  );
}
