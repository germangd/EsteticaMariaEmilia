"use client";

import { useEffect, useMemo, useState } from "react";
import { fmtPesos } from "@/lib/fmt-pesos";
import { uiLabel, uiSelect } from "@/lib/ui-classes";

export const SUELTOS_CATEGORIA_KEY = "__sueltos__";

export type ServicioPasosOpt = {
  id: number;
  nombre: string;
  categoriaNombre?: string | null;
  capacidad?: number;
  precioPesos?: number;
};

type Props = {
  servicios: ServicioPasosOpt[];
  value: string;
  onChange: (value: string) => void;
  /** Valor del select: id del servicio o nombre. */
  valueMode?: "id" | "nombre";
  showCupo?: boolean;
  showPrecio?: boolean;
  className?: string;
  selectClassName?: string;
  disabled?: boolean;
};

function etiquetaCategoria(key: string): string {
  return key === SUELTOS_CATEGORIA_KEY ? "Otros servicios" : key;
}

export function agruparServiciosPorCategoria(servicios: ServicioPasosOpt[]) {
  const map = new Map<string, ServicioPasosOpt[]>();
  for (const s of servicios) {
    const key = s.categoriaNombre?.trim() || SUELTOS_CATEGORIA_KEY;
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
}

export function ServicioPasosSelect({
  servicios,
  value,
  onChange,
  valueMode = "id",
  showCupo = false,
  showPrecio = false,
  className,
  selectClassName,
  disabled,
}: Props) {
  const [categoriaKey, setCategoriaKey] = useState("");
  const [servicioKey, setServicioKey] = useState("");

  const categorias = useMemo(
    () => agruparServiciosPorCategoria(servicios),
    [servicios]
  );

  const omitirCategoria = categorias.length === 1;

  const serviciosEnCategoria = useMemo(() => {
    const cat = categorias.find((c) => c.key === categoriaKey);
    return cat?.items ?? [];
  }, [categorias, categoriaKey]);

  const optValue = (s: ServicioPasosOpt) =>
    valueMode === "id" ? String(s.id) : s.nombre;

  const optLabel = (s: ServicioPasosOpt) => {
    const cupo =
      showCupo && s.capacidad != null ? ` (cupo ${s.capacidad})` : "";
    const precio =
      showPrecio && s.precioPesos != null && s.precioPesos > 0
        ? ` (${fmtPesos(s.precioPesos)})`
        : "";
    return `${s.nombre}${cupo}${precio}`;
  };

  useEffect(() => {
    if (!value) {
      setCategoriaKey(omitirCategoria && categorias[0] ? categorias[0].key : "");
      setServicioKey("");
      return;
    }
    const s = servicios.find((x) =>
      valueMode === "id" ? String(x.id) === value : x.nombre === value
    );
    if (!s) {
      setServicioKey("");
      return;
    }
    setCategoriaKey(s.categoriaNombre?.trim() || SUELTOS_CATEGORIA_KEY);
    setServicioKey(value);
  }, [value, servicios, valueMode, omitirCategoria, categorias]);

  useEffect(() => {
    if (omitirCategoria && categorias[0]) {
      setCategoriaKey(categorias[0].key);
    }
  }, [omitirCategoria, categorias]);

  const selectCls = selectClassName ?? uiSelect;

  function onCategoriaChange(key: string) {
    setCategoriaKey(key);
    setServicioKey("");
    onChange("");
  }

  function onServicioChange(next: string) {
    setServicioKey(next);
    onChange(next);
  }

  if (servicios.length === 0) {
    return (
      <p className="text-sm text-ink-muted">No hay servicios disponibles.</p>
    );
  }

  return (
    <div className={className ?? "space-y-4"}>
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
            {categoriaKey === SUELTOS_CATEGORIA_KEY || omitirCategoria
              ? "Servicio"
              : "Sub-servicio"}
          </label>
          <select
            required
            disabled={disabled}
            className={selectCls}
            value={servicioKey}
            onChange={(e) => onServicioChange(e.target.value)}
          >
            <option value="">
              {omitirCategoria ? "Elegí un servicio" : "Elegí el servicio"}
            </option>
            {serviciosEnCategoria.map((s) => (
              <option key={s.id} value={optValue(s)}>
                {optLabel(s)}
              </option>
            ))}
          </select>
        </div>
      ) : categoriaKey && serviciosEnCategoria.length === 0 ? (
        <p className="text-sm text-ink-muted">
          No hay servicios en esta categoría.
        </p>
      ) : null}
    </div>
  );
}
