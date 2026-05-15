"use client";

import { useState } from "react";
import type { ServicioAdmin } from "@/components/admin/admin-servicios-manager";

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

  const inputClass =
    "w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink";

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
        <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
          Servicio
        </label>
        <select
          required
          className={inputClass}
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
        <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
          Fecha
        </label>
        <input
          type="date"
          required
          className={inputClass}
          value={fecha}
          onChange={(e) => setFecha(e.target.value)}
        />
      </div>
      <div>
        <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
          Hora
        </label>
        <input
          type="time"
          required
          className={inputClass}
          value={hora}
          onChange={(e) => setHora(e.target.value)}
        />
      </div>
      <div>
        <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
          Nombre cliente
        </label>
        <input
          required
          className={inputClass}
          value={nombre}
          onChange={(e) => setNombre(e.target.value)}
        />
      </div>
      <div>
        <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
          Teléfono
        </label>
        <input
          required
          className={inputClass}
          value={telefono}
          onChange={(e) => setTelefono(e.target.value)}
        />
      </div>
      <div className="md:col-span-2">
        <label className="mb-1 block text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
          Email (opcional)
        </label>
        <input
          type="email"
          className={inputClass}
          value={email}
          onChange={(e) => setEmail(e.target.value)}
        />
      </div>
      <label className="flex items-center gap-2 text-sm text-ink-muted md:col-span-2">
        <input
          type="checkbox"
          checked={enviarMail}
          onChange={(e) => setEnviarMail(e.target.checked)}
          className="rounded border-gold/40"
        />
        Enviar mail de confirmación (si hay configuración de correo)
      </label>
      <div className="md:col-span-2">
        <button
          type="submit"
          disabled={pending}
          className="rounded-sm bg-gold px-5 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-white hover:bg-gold-dark disabled:opacity-50"
        >
          {pending ? "Guardando…" : "Cargar turno"}
        </button>
        {msg ? <p className="mt-3 text-sm text-ink-muted">{msg}</p> : null}
      </div>
    </form>
  );
}
