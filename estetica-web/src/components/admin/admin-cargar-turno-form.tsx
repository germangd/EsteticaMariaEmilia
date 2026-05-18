"use client";

import { useState } from "react";
import type { ServicioAdmin } from "@/components/admin/admin-servicios-manager";
import {
  uiBtnPrimary,
  uiInput,
  uiLabel,
  uiSelect,
} from "@/lib/ui-classes";

export function AdminCargarTurnoForm({
  servicios,
  fechaDefault,
}: {
  servicios: ServicioAdmin[];
  fechaDefault: string;
}) {
  const [servicioId, setServicioId] = useState(
    servicios[0]?.id ? String(servicios[0].id) : ""
  );
  const [fecha, setFecha] = useState(fechaDefault);
  const [hora, setHora] = useState("10:00");
  const [nombre, setNombre] = useState("");
  const [telefono, setTelefono] = useState("");
  const [email, setEmail] = useState("");
  const [enviarMail, setEnviarMail] = useState(true);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const inputClass = uiInput;
  const selectClass = uiSelect;

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch("/api/admin/turnos/crear", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          servicioId: Number(servicioId),
          fecha,
          hora,
          nombre,
          telefono,
          email: email.trim() || undefined,
          enviarMail,
        }),
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

  if (servicios.length === 0) {
    return (
      <p className="text-sm text-ink-muted">
        Primero cargá servicios en{" "}
        <a href="/admin/servicios" className="text-gold-dark underline">
          Servicios
        </a>
        .
      </p>
    );
  }

  return (
    <form onSubmit={(e) => void onSubmit(e)} className="grid gap-4 md:grid-cols-2">
      <div className="md:col-span-2">
        <label className={uiLabel}>Servicio</label>
        <select
          required
          className={selectClass}
          value={servicioId}
          onChange={(e) => setServicioId(e.target.value)}
        >
          {servicios.map((s) => (
            <option key={s.id} value={s.id}>
              {s.nombre} (cupo {s.capacidad})
            </option>
          ))}
        </select>
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
      <div>
        <label className={uiLabel}>Hora</label>
        <input
          type="time"
          required
          className={inputClass}
          value={hora}
          onChange={(e) => setHora(e.target.value)}
        />
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
        <button type="submit" disabled={pending} className={uiBtnPrimary}>
          {pending ? "Guardando…" : "Cargar turno"}
        </button>
        {msg ? <p className="mt-3 text-sm font-medium text-ink">{msg}</p> : null}
      </div>
    </form>
  );
}
