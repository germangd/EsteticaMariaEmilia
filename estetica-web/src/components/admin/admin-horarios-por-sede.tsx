"use client";

import { useCallback, useEffect, useState } from "react";
import { AdminHorariosAtencionManager } from "@/components/admin/admin-horarios-atencion-manager";
import type { HorarioAtencionPorDia } from "@/lib/horario-atencion-repo";
import { uiLabel, uiSelect } from "@/lib/ui-classes";

type SedeOpt = { id: number; nombre: string };

export function AdminHorariosPorSede({ sedes }: { sedes: SedeOpt[] }) {
  const [sedeId, setSedeId] = useState(sedes[0]?.id ?? 0);
  const [dias, setDias] = useState<HorarioAtencionPorDia[]>([]);
  const [loading, setLoading] = useState(true);

  const cargar = useCallback(async (id: number) => {
    if (id < 1) {
      setDias([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    try {
      const r = await fetch(`/api/admin/horarios-atencion?sedeId=${id}`, {
        credentials: "same-origin",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        dias?: HorarioAtencionPorDia[];
      };
      if (data.ok && data.dias) setDias(data.dias);
      else setDias([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void cargar(sedeId);
  }, [cargar, sedeId]);

  if (sedes.length === 0) {
    return (
      <p className="text-sm text-ink-muted">
        {"No hay sedes. Cre\u00e1 sedes en Admin \u2192 Sedes."}
      </p>
    );
  }

  return (
    <div>
      <div className="mb-6">
        <label className={uiLabel}>Sede</label>
        <select
          className={uiSelect}
          value={sedeId}
          onChange={(e) => setSedeId(Number(e.target.value))}
        >
          {sedes.map((s) => (
            <option key={s.id} value={s.id}>
              {s.nombre}
            </option>
          ))}
        </select>
      </div>
      {loading ? (
        <p className="text-sm text-ink-muted">{"Cargando horarios\u2026"}</p>
      ) : (
        <AdminHorariosAtencionManager
          key={sedeId}
          sedeId={sedeId}
          initialDias={dias}
        />
      )}
    </div>
  );
}
