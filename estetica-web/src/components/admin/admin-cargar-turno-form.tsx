"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  AdminTurnoCatalogoSelect,
  type AdminTurnoSeleccion,
} from "@/components/admin/admin-turno-catalogo-select";
import type { PaqueteSelectOption } from "@/components/admin/servicio-select-optgroups";
import type { ServicioAdmin } from "@/components/admin/admin-servicios-manager";
import { claveReservaPaquete, claveReservaServicio } from "@/lib/reserva-claves";
import { filtrarServiciosReservables } from "@/lib/servicio-tree";
import {
  uiBtnPrimary,
  uiHint,
  uiInput,
  uiLabel,
  uiSelect,
  uiTimeSlot,
  uiTimeSlotActive,
} from "@/lib/ui-classes";

export function AdminCargarTurnoForm({
  servicios,
  paquetes,
  sedes,
  fechaDefault,
}: {
  servicios: ServicioAdmin[];
  paquetes: PaqueteSelectOption[];
  sedes: { id: number; nombre: string }[];
  fechaDefault: string;
}) {
  const [sedeId, setSedeId] = useState(sedes[0]?.id ?? 0);
  const [seleccion, setSeleccion] = useState<AdminTurnoSeleccion | null>(null);
  const [fecha, setFecha] = useState(fechaDefault);
  const [hora, setHora] = useState("");
  const [horarios, setHorarios] = useState<string[]>([]);
  const [loadingHorarios, setLoadingHorarios] = useState(false);
  const [horariosError, setHorariosError] = useState<string | null>(null);
  const [nombre, setNombre] = useState("");
  const [telefono, setTelefono] = useState("");
  const [email, setEmail] = useState("");
  const [enviarMail, setEnviarMail] = useState(true);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const inputClass = uiInput;
  const selectClass = uiSelect;

  const serviciosReservables = useMemo(
    () => filtrarServiciosReservables(servicios),
    [servicios]
  );

  const servicioSel = useMemo(() => {
    if (seleccion?.tipo !== "servicio") return undefined;
    return serviciosReservables.find((s) => s.id === seleccion.servicioId);
  }, [seleccion, serviciosReservables]);

  const paqueteSel = useMemo(() => {
    if (seleccion?.tipo !== "paquete") return undefined;
    return paquetes.find((p) => p.id === seleccion.paqueteId);
  }, [seleccion, paquetes]);

  const claveHorarios = useMemo(() => {
    if (seleccion?.tipo === "servicio" && servicioSel) {
      return claveReservaServicio(servicioSel.nombre);
    }
    if (seleccion?.tipo === "paquete") {
      return claveReservaPaquete(seleccion.paqueteId);
    }
    return "";
  }, [seleccion, servicioSel]);

  const duracionMin = useMemo(() => {
    if (seleccion?.tipo === "servicio") return servicioSel?.duracion;
    if (seleccion?.tipo === "paquete") return paqueteSel?.duracion;
    return undefined;
  }, [seleccion, servicioSel, paqueteSel]);

  const duracionEtiqueta = useMemo(() => {
    const min = duracionMin;
    if (!min) return null;
    let base: string;
    if (min < 60) base = `${min} min`;
    else {
      const h = Math.floor(min / 60);
      const m = min % 60;
      base = m > 0 ? `${h} h ${m} min` : `${h} h`;
    }
    if (
      seleccion?.tipo === "paquete" &&
      paqueteSel &&
      (paqueteSel.sesionesTotal ?? 0) > 1
    ) {
      return `${base} por sesión · paquete de ${paqueteSel.sesionesTotal} sesiones`;
    }
    return base;
  }, [duracionMin, seleccion, paqueteSel]);

  const cargarHorarios = useCallback(async () => {
    if (!claveHorarios || !fecha) {
      setHorarios([]);
      setHora("");
      setHorariosError(null);
      return;
    }
    setLoadingHorarios(true);
    setHorariosError(null);
    setHora("");
    try {
      const q = new URLSearchParams({
        clave: claveHorarios,
        fecha,
        sedeId: String(sedeId),
      });
      const r = await fetch(`/api/horarios?${q.toString()}`, {
        cache: "no-store",
      });
      const data = (await r.json()) as {
        ok?: boolean;
        horarios?: string[];
        mensaje?: string;
      };
      if (!r.ok || data.ok === false) {
        setHorarios([]);
        setHorariosError(data.mensaje ?? "No se pudieron cargar horarios.");
        return;
      }
      const lista = Array.isArray(data.horarios) ? data.horarios : [];
      setHorarios(lista);
      if (lista.length === 1) setHora(lista[0]!);
    } catch {
      setHorarios([]);
      setHorariosError("Error de red al cargar horarios.");
    } finally {
      setLoadingHorarios(false);
    }
  }, [claveHorarios, fecha, sedeId]);

  useEffect(() => {
    void cargarHorarios();
  }, [cargarHorarios]);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!hora || !seleccion) {
      setMsg("Elegí servicio o combo y un horario disponible.");
      return;
    }
    setPending(true);
    setMsg(null);
    try {
      const body =
        seleccion.tipo === "paquete"
          ? {
              paqueteId: seleccion.paqueteId,
              sedeId,
              fecha,
              hora,
              nombre,
              telefono,
              email: email.trim() || undefined,
              enviarMail,
            }
          : {
              servicioId: seleccion.servicioId,
              sedeId,
              fecha,
              hora,
              nombre,
              telefono,
              email: email.trim() || undefined,
              enviarMail,
            };

      const r = await fetch("/api/admin/turnos/crear", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = (await r.json()) as {
        ok?: boolean;
        mensaje?: string;
        codigo?: string;
      };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo cargar el turno.");
        return;
      }
      setMsg(
        data.codigo
          ? `Turno guardado. Código: ${data.codigo}`
          : (data.mensaje ?? "Turno guardado.")
      );
      setNombre("");
      setTelefono("");
      setEmail("");
      window.setTimeout(() => window.location.reload(), 1200);
    } finally {
      setPending(false);
    }
  }

  const tieneCatalogo =
    serviciosReservables.length > 0 || paquetes.length > 0;

  if (!tieneCatalogo) {
    return (
      <p className="text-sm font-medium text-ink">
        Primero cargá servicios en{" "}
        <a href="/admin/servicios" className="text-gold-dark underline">
          Servicios
        </a>{" "}
        o combos en{" "}
        <a href="/admin/paquetes" className="text-gold-dark underline">
          Paquetes
        </a>
        .
      </p>
    );
  }

  return (
    <form onSubmit={(e) => void onSubmit(e)} className="grid gap-4 md:grid-cols-2">
      <div className="md:col-span-2">
        <label className={uiLabel}>Sede</label>
        <select
          required
          className={selectClass}
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
      <div className="md:col-span-2">
        <AdminTurnoCatalogoSelect
          servicios={serviciosReservables.map((s) => ({
            id: s.id,
            nombre: s.nombre,
            categoriaNombre: s.categoriaNombre,
            capacidad: s.capacidad,
          }))}
          paquetes={paquetes}
          seleccion={seleccion}
          onSeleccionChange={setSeleccion}
          selectClassName={selectClass}
        />
      </div>
      <div>
        <label className={uiLabel}>Fecha</label>
        <input
          type="date"
          required
          className={inputClass}
          value={fecha}
          onChange={(e) => setFecha(e.target.value)}
        />
      </div>
      <div className="md:col-span-2">
        <label className={uiLabel}>Horario disponible</label>
        {duracionEtiqueta ? (
          <p className={`mb-2 ${uiHint}`}>
            Duración estimada: <strong>{duracionEtiqueta}</strong>. Solo se
            muestran turnos libres según cupo y reservas ya cargadas.
          </p>
        ) : null}
        {loadingHorarios ? (
          <p className="text-sm font-medium text-ink">Buscando horarios…</p>
        ) : horariosError ? (
          <p className="rounded-sm bg-amber-50 px-3 py-2 text-sm text-amber-900">
            {horariosError}
          </p>
        ) : !seleccion || !fecha ? (
          <p className="text-sm font-medium text-ink-muted">
            Elegí servicio o combo, categoría si aplica, y fecha.
          </p>
        ) : horarios.length === 0 ? (
          <p className="text-sm font-medium text-ink-muted">
            No hay horarios libres para esta combinación en esa fecha.
          </p>
        ) : (
          <div className="flex flex-wrap gap-2">
            {horarios.map((h) => (
              <button
                key={h}
                type="button"
                onClick={() => setHora(h)}
                className={hora === h ? uiTimeSlotActive : uiTimeSlot}
              >
                {h}
              </button>
            ))}
          </div>
        )}
      </div>
      <div>
        <label className={uiLabel}>Nombre cliente</label>
        <input
          required
          className={inputClass}
          value={nombre}
          onChange={(e) => setNombre(e.target.value)}
        />
      </div>
      <div>
        <label className={uiLabel}>Teléfono</label>
        <input
          required
          className={inputClass}
          value={telefono}
          onChange={(e) => setTelefono(e.target.value)}
        />
      </div>
      <div className="md:col-span-2">
        <label className={uiLabel}>Email (opcional)</label>
        <input
          type="email"
          className={inputClass}
          value={email}
          onChange={(e) => setEmail(e.target.value)}
        />
      </div>
      <label className="flex items-center gap-2 text-sm font-medium text-ink md:col-span-2">
        <input
          type="checkbox"
          checked={enviarMail}
          onChange={(e) => setEnviarMail(e.target.checked)}
          className="h-4 w-4 rounded border-gold/55 text-gold focus:ring-gold/40"
        />
        Enviar mail de confirmación (si hay configuración de correo)
      </label>
      <div className="md:col-span-2">
        <button
          type="submit"
          disabled={pending || !hora || !seleccion}
          className={uiBtnPrimary}
        >
          {pending ? "Guardando…" : "Cargar turno"}
        </button>
        {msg ? <p className="mt-3 text-sm font-medium text-ink">{msg}</p> : null}
      </div>
    </form>
  );
}
