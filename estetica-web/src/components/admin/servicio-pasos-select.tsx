"use client";

import { useEffect, useMemo, useState } from "react";
import { fmtPesos } from "@/lib/fmt-pesos";
import { uiLabel } from "@/lib/ui-classes";

export const SUELTOS_CATEGORIA_KEY = "__sueltos__";

export type ServicioPasosOpt = {
  id: number;
  nombre: string;
  categoriaNombre?: string | null;
  capacidad?: number;
  precioPesos?: number;
  anticipoRequerido?: boolean;
  anticipoPorcentaje?: number;
};

type Props = {
  servicios: ServicioPasosOpt[];
  value: string;
  onChange: (value: string) => void;
  /** Valor del select: id del servicio o nombre. */
  valueMode?: "id" | "nombre";
  showCupo?: boolean;
  showPrecio?: boolean;
  showAnticipo?: boolean;
  className?: string;
  selectClassName?: string;
  disabled?: boolean;
  required?: boolean;
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
  showAnticipo = false,
  className,
  selectClassName: _selectClassName,
  disabled,
  required = false,
}: Props) {
  const [expandedKey, setExpandedKey] = useState<string | null>(null);
  const [servicioKey, setServicioKey] = useState("");

  const categorias = useMemo(
    () => agruparServiciosPorCategoria(servicios),
    [servicios]
  );

  const omitirCategoria = categorias.length === 1;
  const categoriaUnica = omitirCategoria ? categorias[0] : null;

  const optValue = (s: ServicioPasosOpt) =>
    valueMode === "id" ? String(s.id) : s.nombre;

  const optLabel = (s: ServicioPasosOpt) => {
    const cupo =
      showCupo && s.capacidad != null ? ` (cupo ${s.capacidad})` : "";
    const precio =
      showPrecio && s.precioPesos != null && s.precioPesos > 0
        ? ` (${fmtPesos(s.precioPesos)})`
        : "";
    const anticipo =
      showAnticipo && s.anticipoRequerido && (s.anticipoPorcentaje ?? 0) > 0
        ? ` · ant. ${s.anticipoPorcentaje}%`
        : "";
    return `${s.nombre}${cupo}${precio}${anticipo}`;
  };

  useEffect(() => {
    if (!value) {
      setServicioKey("");
      if (!omitirCategoria) setExpandedKey(null);
      return;
    }
    const s = servicios.find((x) =>
      valueMode === "id" ? String(x.id) === value : x.nombre === value
    );
    if (!s) {
      setServicioKey("");
      return;
    }
    const catKey = s.categoriaNombre?.trim() || SUELTOS_CATEGORIA_KEY;
    setServicioKey(value);
    setExpandedKey(catKey);
  }, [value, servicios, valueMode, omitirCategoria]);

  useEffect(() => {
    if (omitirCategoria && categorias[0]) {
      setExpandedKey(categorias[0].key);
    }
  }, [omitirCategoria, categorias]);

  function onServicioChange(next: string) {
    setServicioKey(next);
    onChange(next);
  }

  function toggleCategoria(key: string) {
    if (disabled) return;
    if (expandedKey === key) {
      setExpandedKey(null);
      return;
    }
    setExpandedKey(key);
    const selected = servicios.find((x) =>
      valueMode === "id" ? String(x.id) === servicioKey : x.nombre === servicioKey
    );
    const selectedCat =
      selected?.categoriaNombre?.trim() || SUELTOS_CATEGORIA_KEY;
    if (selectedCat !== key) {
      setServicioKey("");
      onChange("");
    }
  }

  function renderItemButton(s: ServicioPasosOpt) {
    const val = optValue(s);
    const selected = servicioKey === val;
    return (
      <button
        key={s.id}
        type="button"
        disabled={disabled}
        onClick={() => onServicioChange(val)}
        className={`block w-full rounded-sm px-3 py-2 text-left text-sm transition-colors ${
          selected
            ? "bg-gold/25 font-medium text-ink-dark ring-1 ring-gold/40"
            : "text-ink hover:bg-cream/70"
        }`}
      >
        {optLabel(s)}
      </button>
    );
  }

  if (servicios.length === 0) {
    return (
      <p className="text-sm text-ink-muted">No hay servicios disponibles.</p>
    );
  }

  if (omitirCategoria && categoriaUnica) {
    return (
      <div className={className ?? "space-y-2"}>
        {required ? (
          <input
            tabIndex={-1}
            aria-hidden
            className="pointer-events-none absolute h-0 w-0 opacity-0"
            required
            value={servicioKey}
            readOnly
            onChange={() => {}}
          />
        ) : null}
        <label className={uiLabel}>Servicio</label>
        <div className="space-y-1 rounded-sm border border-gold/25 bg-white/60 p-2">
          {categoriaUnica.items.map(renderItemButton)}
        </div>
      </div>
    );
  }

  return (
    <div className={className ?? "space-y-2"}>
      {required ? (
        <input
          tabIndex={-1}
          aria-hidden
          className="pointer-events-none absolute h-0 w-0 opacity-0"
          required
          value={servicioKey}
          readOnly
          onChange={() => {}}
        />
      ) : null}
      <label className={uiLabel}>Categoría y servicio</label>
      <p className="text-xs text-ink-muted">
        Elegí una categoría para ver los servicios disponibles.
      </p>
      <div className="overflow-hidden rounded-sm border border-gold/25 bg-white/50 divide-y divide-gold/15">
        {categorias.map((c) => {
          const isOpen = expandedKey === c.key;
          const selectedInCat = c.items.some(
            (s) => optValue(s) === servicioKey
          );
          return (
            <div key={c.key}>
              <button
                type="button"
                disabled={disabled}
                onClick={() => toggleCategoria(c.key)}
                aria-expanded={isOpen}
                className={`flex w-full items-center justify-between gap-2 px-3 py-2.5 text-left text-sm transition-colors ${
                  isOpen
                    ? "bg-cream/90 font-medium text-ink-dark"
                    : "hover:bg-cream/50 text-ink"
                }`}
              >
                <span>{c.label}</span>
                <span className="shrink-0 text-xs text-ink-muted">
                  {selectedInCat && !isOpen ? "· elegido " : ""}
                  {c.items.length}
                  <span className="ml-1.5 inline-block w-4 text-center">
                    {isOpen ? "▴" : "▾"}
                  </span>
                </span>
              </button>
              {isOpen ? (
                <div className="space-y-0.5 border-t border-gold/15 bg-white/70 px-2 py-2">
                  {c.items.map(renderItemButton)}
                </div>
              ) : null}
            </div>
          );
        })}
      </div>
    </div>
  );
}
