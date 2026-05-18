"use client";

import { useCallback, useEffect, useState } from "react";
import { uiBtnSecondary, uiInput, uiLabel } from "@/lib/ui-classes";

export function AdminServicioFechas({
  serviceId,
  esGrupo,
}: {
  serviceId: number;
  esGrupo: boolean;
}) {
  const [fechas, setFechas] = useState<string[]>([]);
  const [nuevaFecha, setNuevaFecha] = useState("");
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState<string | null>(null);

  const cargar = useCallback(async () => {
    setLoading(true);
    try {
      const r = await fetch(`/api/admin/servicios/${serviceId}/fechas`, {
        credentials: "same-origin",
      });
      const data = (await r.json()) as { ok?: boolean; fechas?: string[] };
      if (data.ok && Array.isArray(data.fechas)) setFechas(data.fechas);
    } finally {
      setLoading(false);
    }
  }, [serviceId]);

  useEffect(() => {
    if (!esGrupo && serviceId > 0) void cargar();
  }, [cargar, esGrupo, serviceId]);

  if (esGrupo) {
    return (
      <p className="text-sm text-ink-muted">
        {"Las categor\u00edas no reciben turnos. Configur\u00e1 fechas en cada sub-servicio."}
      </p>
    );
  }

  async function agregar() {
    if (!nuevaFecha) return;
    setLoading(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/servicios/${serviceId}/fechas`, {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fecha: nuevaFecha }),
      });
      const data = (await r.json()) as { ok?: boolean; fechas?: string[] };
      if (!r.ok || !data.ok) {
        setMsg("No se pudo agregar la fecha.");
        return;
      }
      setFechas(data.fechas ?? []);
      setNuevaFecha("");
    } finally {
      setLoading(false);
    }
  }

  async function quitar(fecha: string) {
    setLoading(true);
    setMsg(null);
    try {
      const r = await fetch(
        `/api/admin/servicios/${serviceId}/fechas?fecha=${encodeURIComponent(fecha)}`,
        { method: "DELETE", credentials: "same-origin" }
      );
      const data = (await r.json()) as { ok?: boolean; fechas?: string[] };
      if (data.ok && data.fechas) setFechas(data.fechas);
    } finally {
      setLoading(false);
    }
  }

  async function limpiarCalendario() {
    if (!window.confirm("¿Quitar todas las fechas? El servicio volverá a aceptar cualquier día (salvo domingos y eventos).")) return;
    setLoading(true);
    try {
      const r = await fetch(`/api/admin/servicios/${serviceId}/fechas`, {
        method: "PUT",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fechas: [] }),
      });
      const data = (await r.json()) as { ok?: boolean; fechas?: string[] };
      if (data.ok) setFechas(data.fechas ?? []);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="mt-6 rounded-sm border border-gold/25 bg-cream/40 p-4">
      <h3 className="mb-2 font-serif text-lg text-ink-dark">
        {"Fechas habilitadas para turnos"}
      </h3>
      <p className="mb-3 text-sm text-ink-muted">
        {fechas.length === 0
          ? "Sin fechas cargadas: el servicio usa el calendario normal (lun\u2013s\u00e1b, horario del servicio)."
          : "Solo las fechas listadas admiten reservas para este servicio."}
      </p>
      <div className="mb-3 flex flex-wrap items-end gap-2">
        <div>
          <label className={uiLabel}>Agregar fecha</label>
          <input
            type="date"
            className={uiInput}
            value={nuevaFecha}
            onChange={(e) => setNuevaFecha(e.target.value)}
            disabled={loading}
          />
        </div>
        <button
          type="button"
          disabled={loading || !nuevaFecha}
          onClick={() => void agregar()}
          className={uiBtnSecondary}
        >
          Agregar
        </button>
        {fechas.length > 0 ? (
          <button
            type="button"
            disabled={loading}
            onClick={() => void limpiarCalendario()}
            className="text-xs text-red-800 underline"
          >
            Quitar restricción (todas las fechas)
          </button>
        ) : null}
      </div>
      {fechas.length > 0 ? (
        <ul className="flex flex-wrap gap-2">
          {fechas.map((f) => (
            <li
              key={f}
              className="flex items-center gap-2 rounded-sm border border-gold/30 bg-white px-2 py-1 text-sm"
            >
              {f}
              <button
                type="button"
                className="text-red-800"
                disabled={loading}
                onClick={() => void quitar(f)}
                aria-label={`Quitar ${f}`}
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      ) : null}
      {msg ? <p className="mt-2 text-sm text-red-800">{msg}</p> : null}
    </div>
  );
}
