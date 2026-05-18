"use client";

import { useCallback, useEffect, useState } from "react";
import type { HorarioAtencionPorDia } from "@/lib/horario-atencion-repo";
import {
  uiBtnPrimary,
  uiBtnSecondary,
  uiCard,
  uiInput,
  uiLabel,
} from "@/lib/ui-classes";

type DiaState = HorarioAtencionPorDia & {
  franjas: { horarioInicio: string; horarioFin: string }[];
};

function cloneDias(dias: HorarioAtencionPorDia[]): DiaState[] {
  return dias.map((d) => ({
    ...d,
    franjas: d.franjas.map((f) => ({ ...f })),
  }));
}

export function AdminHorariosAtencionManager({
  sedeId,
  initialDias,
}: {
  sedeId: number;
  initialDias: HorarioAtencionPorDia[];
}) {
  const [dias, setDias] = useState<DiaState[]>(() => cloneDias(initialDias));
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  useEffect(() => {
    setDias(cloneDias(initialDias));
  }, [initialDias]);

  const agregarFranja = useCallback((diaSemana: number) => {
    setDias((prev) =>
      prev.map((d) =>
        d.diaSemana === diaSemana
          ? {
              ...d,
              franjas: [
                ...d.franjas,
                { horarioInicio: "09:00", horarioFin: "13:00" },
              ],
            }
          : d
      )
    );
  }, []);

  const quitarFranja = useCallback((diaSemana: number, index: number) => {
    setDias((prev) =>
      prev.map((d) =>
        d.diaSemana === diaSemana
          ? {
              ...d,
              franjas: d.franjas.filter((_, i) => i !== index),
            }
          : d
      )
    );
  }, []);

  const actualizarFranja = useCallback(
    (
      diaSemana: number,
      index: number,
      campo: "horarioInicio" | "horarioFin",
      valor: string
    ) => {
      setDias((prev) =>
        prev.map((d) => {
          if (d.diaSemana !== diaSemana) return d;
          const franjas = [...d.franjas];
          franjas[index] = { ...franjas[index]!, [campo]: valor };
          return { ...d, franjas };
        })
      );
    },
    []
  );

  const copiarLunesATodos = useCallback(() => {
    setDias((prev) => {
      const lunes = prev.find((d) => d.diaSemana === 1);
      if (!lunes) return prev;
      const copia = lunes.franjas.map((f) => ({ ...f }));
      return prev.map((d) =>
        d.diaSemana === 1 ? d : { ...d, franjas: copia.map((f) => ({ ...f })) }
      );
    });
    setMsg("Franjas del lunes copiadas a martes\u2013s\u00e1bado.");
  }, []);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const franjas = dias.flatMap((d) =>
        d.franjas.map((f) => ({
          sedeId,
          diaSemana: d.diaSemana,
          horarioInicio: f.horarioInicio,
          horarioFin: f.horarioFin,
        }))
      );
      const r = await fetch("/api/admin/horarios-atencion", {
        method: "PUT",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sedeId, franjas }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo guardar.");
        return;
      }
      setMsg("Horario de atenci\u00f3n guardado.");
      const refresh = await fetch(
        `/api/admin/horarios-atencion?sedeId=${sedeId}`,
        { credentials: "same-origin" }
      );
      const refreshed = (await refresh.json()) as {
        ok?: boolean;
        dias?: HorarioAtencionPorDia[];
      };
      if (refreshed.ok && refreshed.dias) setDias(cloneDias(refreshed.dias));
    } finally {
      setPending(false);
    }
  }

  return (
    <div className="space-y-8">
      <section className={uiCard}>
        <p className="mb-4 text-sm text-ink-muted">
          {
            "Defin\u00ed las franjas en las que el local atiende cada d\u00eda (ej. 09:00\u201313:00 y 17:00\u201321:00). Los turnos solo se ofrecen dentro de esas franjas y del horario de cada servicio. Domingo sin atenci\u00f3n."
          }
        </p>
        <div className="mb-4">
          <button
            type="button"
            className={uiBtnSecondary}
            onClick={copiarLunesATodos}
            disabled={pending}
          >
            Copiar lunes a todos los d\u00edas
          </button>
        </div>
        <form onSubmit={(e) => void onSubmit(e)} className="space-y-8">
          {dias.map((dia) => (
            <div
              key={dia.diaSemana}
              className="border-t border-gold/20 pt-6 first:border-t-0 first:pt-0"
            >
              <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
                <h3 className="font-serif text-lg text-ink-dark">{dia.label}</h3>
                <button
                  type="button"
                  className="text-xs text-gold-dark underline"
                  onClick={() => agregarFranja(dia.diaSemana)}
                  disabled={pending}
                >
                  + Agregar franja
                </button>
              </div>
              {dia.franjas.length === 0 ? (
                <p className="text-sm text-ink-muted">
                  Sin franjas: no hay turnos este d\u00eda.
                </p>
              ) : (
                <ul className="space-y-3">
                  {dia.franjas.map((f, idx) => (
                    <li
                      key={`${dia.diaSemana}-${idx}`}
                      className="flex flex-wrap items-end gap-3"
                    >
                      <div>
                        <label className={uiLabel}>Desde</label>
                        <input
                          type="time"
                          required
                          className={uiInput}
                          value={f.horarioInicio}
                          onChange={(e) =>
                            actualizarFranja(
                              dia.diaSemana,
                              idx,
                              "horarioInicio",
                              e.target.value
                            )
                          }
                        />
                      </div>
                      <div>
                        <label className={uiLabel}>Hasta</label>
                        <input
                          type="time"
                          required
                          className={uiInput}
                          value={f.horarioFin}
                          onChange={(e) =>
                            actualizarFranja(
                              dia.diaSemana,
                              idx,
                              "horarioFin",
                              e.target.value
                            )
                          }
                        />
                      </div>
                      <button
                        type="button"
                        className="mb-0.5 text-xs text-red-800 underline"
                        onClick={() => quitarFranja(dia.diaSemana, idx)}
                        disabled={pending}
                      >
                        Quitar
                      </button>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          ))}
          <div className="flex flex-wrap gap-2 border-t border-gold/20 pt-6">
            <button type="submit" disabled={pending} className={uiBtnPrimary}>
              {pending ? "Guardando\u2026" : "Guardar horarios"}
            </button>
          </div>
        </form>
        {msg ? <p className="mt-3 text-sm font-medium text-ink">{msg}</p> : null}
      </section>
    </div>
  );
}
