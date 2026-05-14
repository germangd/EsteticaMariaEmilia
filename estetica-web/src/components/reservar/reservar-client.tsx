"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useState } from "react";

type ServicioApi = {
  nombre: string;
  duracion: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
};

type Tab = "reservar" | "cancelar";

export function ReservarClient() {
  const [tab, setTab] = useState<Tab>("reservar");

  const [servicios, setServicios] = useState<ServicioApi[]>([]);
  const [serviciosError, setServiciosError] = useState<string | null>(null);
  const [loadingServicios, setLoadingServicios] = useState(true);

  const [servicio, setServicio] = useState("");
  const [fecha, setFecha] = useState("");
  const [horarios, setHorarios] = useState<string[]>([]);
  const [loadingHorarios, setLoadingHorarios] = useState(false);
  const [horariosError, setHorariosError] = useState<string | null>(null);
  const [hora, setHora] = useState("");

  const [nombre, setNombre] = useState("");
  const [telefono, setTelefono] = useState("");
  const [email, setEmail] = useState("");

  const [reservaMsg, setReservaMsg] = useState<{
    type: "ok" | "err";
    text: string;
  } | null>(null);
  const [reservando, setReservando] = useState(false);

  const [codigoCancel, setCodigoCancel] = useState("");
  const [cancelMsg, setCancelMsg] = useState<{
    type: "ok" | "err";
    text: string;
  } | null>(null);
  const [cancelando, setCancelando] = useState(false);

  const minFecha = useMemo(() => {
    const d = new Date();
    return d.toISOString().slice(0, 10);
  }, []);

  const maxFecha = useMemo(() => {
    const d = new Date();
    d.setDate(d.getDate() + 120);
    return d.toISOString().slice(0, 10);
  }, []);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingServicios(true);
      setServiciosError(null);
      try {
        const r = await fetch("/api/servicios", { cache: "no-store" });
        const data = (await r.json()) as {
          ok?: boolean;
          servicios?: ServicioApi[];
          mensaje?: string;
        };
        if (cancelled) return;
        if (!r.ok || !data.ok || !Array.isArray(data.servicios)) {
          setServiciosError(
            data.mensaje ?? "No se pudieron cargar los servicios."
          );
          setServicios([]);
          return;
        }
        setServicios(data.servicios);
      } catch {
        if (!cancelled) setServiciosError("Error de red al cargar servicios.");
      } finally {
        if (!cancelled) setLoadingServicios(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const cargarHorarios = useCallback(async () => {
    if (!servicio || !fecha) {
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
        servicio,
        fecha,
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
        setHorariosError(data.mensaje ?? "No se pudieron obtener horarios.");
        return;
      }
      setHorarios(Array.isArray(data.horarios) ? data.horarios : []);
    } catch {
      setHorarios([]);
      setHorariosError("Error de red al cargar horarios.");
    } finally {
      setLoadingHorarios(false);
    }
  }, [servicio, fecha]);

  useEffect(() => {
    void cargarHorarios();
  }, [cargarHorarios]);

  async function enviarReserva(e: React.FormEvent) {
    e.preventDefault();
    setReservaMsg(null);
    if (!servicio || !fecha || !hora || !nombre.trim() || !telefono.trim()) {
      setReservaMsg({
        type: "err",
        text: "Completá servicio, fecha, horario, nombre y teléfono.",
      });
      return;
    }
    setReservando(true);
    try {
      const r = await fetch("/api/turnos", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          servicio,
          fecha,
          hora,
          nombre: nombre.trim(),
          telefono: telefono.trim(),
          email: email.trim() || undefined,
        }),
      });
      const data = (await r.json()) as {
        exito?: boolean;
        mensaje?: string;
        codigo?: string;
      };
      if (data.exito) {
        setReservaMsg({
          type: "ok",
          text:
            data.mensaje ??
            (data.codigo
              ? `Turno confirmado. Guardá tu código: ${data.codigo}`
              : "Turno confirmado."),
        });
        setNombre("");
        setTelefono("");
        setEmail("");
        setHora("");
        void cargarHorarios();
      } else {
        setReservaMsg({
          type: "err",
          text: data.mensaje ?? "No se pudo guardar el turno.",
        });
      }
    } catch {
      setReservaMsg({ type: "err", text: "Error de red al reservar." });
    } finally {
      setReservando(false);
    }
  }

  async function enviarCancelacion(e: React.FormEvent) {
    e.preventDefault();
    setCancelMsg(null);
    const c = codigoCancel.trim();
    if (!c) {
      setCancelMsg({ type: "err", text: "Ingresá el código de cancelación." });
      return;
    }
    setCancelando(true);
    try {
      const r = await fetch("/api/turnos/cancelar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ codigo: c }),
      });
      const data = (await r.json()) as { exito?: boolean; mensaje?: string };
      if (data.exito) {
        setCancelMsg({
          type: "ok",
          text: data.mensaje ?? "Turno cancelado.",
        });
        setCodigoCancel("");
      } else {
        setCancelMsg({
          type: "err",
          text: data.mensaje ?? "No se pudo cancelar.",
        });
      }
    } catch {
      setCancelMsg({ type: "err", text: "Error de red al cancelar." });
    } finally {
      setCancelando(false);
    }
  }

  return (
    <div className="min-h-screen bg-cream pb-16 pt-24 text-ink">
      <div className="mx-auto max-w-lg px-4">
        <p className="mb-1 text-center text-xs font-medium uppercase tracking-[0.25em] text-gold">
          María Emilia Estética
        </p>
        <h1 className="mb-2 text-center font-serif text-2xl font-light text-ink-dark md:text-3xl">
          Turnos online
        </h1>
        <p className="mb-8 text-center text-sm font-light text-ink-muted">
          Elegí servicio, fecha y horario. El código de cancelación te llega por
          mail si configuraste email.
        </p>

        <div className="mb-8 flex rounded border border-gold/25 bg-white/80 p-1 shadow-sm">
          <button
            type="button"
            onClick={() => setTab("reservar")}
            className={`flex-1 rounded py-2.5 text-xs font-medium uppercase tracking-wide transition-colors ${
              tab === "reservar"
                ? "bg-gold text-white"
                : "text-ink-muted hover:text-ink"
            }`}
          >
            Reservar
          </button>
          <button
            type="button"
            onClick={() => setTab("cancelar")}
            className={`flex-1 rounded py-2.5 text-xs font-medium uppercase tracking-wide transition-colors ${
              tab === "cancelar"
                ? "bg-gold text-white"
                : "text-ink-muted hover:text-ink"
            }`}
          >
            Cancelar
          </button>
        </div>

        {tab === "reservar" && (
          <form
            onSubmit={enviarReserva}
            className="rounded border border-gold/20 bg-white/90 p-6 shadow-sm"
          >
            {loadingServicios ? (
              <p className="text-center text-sm text-ink-muted">
                Cargando servicios…
              </p>
            ) : serviciosError ? (
              <p className="rounded bg-red-50 px-3 py-2 text-center text-sm text-red-800">
                {serviciosError}
              </p>
            ) : servicios.length === 0 ? (
              <p className="text-center text-sm text-ink-muted">
                No hay servicios cargados en la base. Usá{" "}
                <code className="text-xs">seed_example.sql</code> o el panel de
                datos.
              </p>
            ) : (
              <>
                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  Servicio
                </label>
                <select
                  required
                  value={servicio}
                  onChange={(e) => setServicio(e.target.value)}
                  className="mb-4 w-full rounded border border-gold/25 bg-cream px-3 py-2.5 text-sm outline-none ring-gold/30 focus:ring-2"
                >
                  <option value="">Elegí un servicio</option>
                  {servicios.map((s) => (
                    <option key={s.nombre} value={s.nombre}>
                      {s.nombre}
                    </option>
                  ))}
                </select>

                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  Fecha
                </label>
                <input
                  type="date"
                  required
                  min={minFecha}
                  max={maxFecha}
                  value={fecha}
                  onChange={(e) => setFecha(e.target.value)}
                  className="mb-4 w-full rounded border border-gold/25 bg-cream px-3 py-2.5 text-sm outline-none ring-gold/30 focus:ring-2"
                />

                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  Horario
                </label>
                {loadingHorarios ? (
                  <p className="mb-4 text-sm text-ink-muted">
                    Buscando horarios…
                  </p>
                ) : horariosError ? (
                  <p className="mb-4 rounded bg-amber-50 px-3 py-2 text-sm text-amber-900">
                    {horariosError}
                  </p>
                ) : !servicio || !fecha ? (
                  <p className="mb-4 text-sm text-ink-muted">
                    Elegí servicio y fecha para ver horarios.
                  </p>
                ) : horarios.length === 0 ? (
                  <p className="mb-4 text-sm text-ink-muted">
                    No hay horarios disponibles para esa combinación.
                  </p>
                ) : (
                  <div className="mb-4 flex flex-wrap gap-2">
                    {horarios.map((h) => (
                      <button
                        key={h}
                        type="button"
                        onClick={() => setHora(h)}
                        className={`rounded-full border px-3 py-1.5 text-sm transition-colors ${
                          hora === h
                            ? "border-gold-dark bg-gold text-white"
                            : "border-gold/30 bg-cream text-ink hover:border-gold"
                        }`}
                      >
                        {h}
                      </button>
                    ))}
                  </div>
                )}

                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  Nombre completo
                </label>
                <input
                  required
                  value={nombre}
                  onChange={(e) => setNombre(e.target.value)}
                  className="mb-4 w-full rounded border border-gold/25 bg-cream px-3 py-2.5 text-sm outline-none ring-gold/30 focus:ring-2"
                />

                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  Teléfono (código de área, sin 0 ni 15)
                </label>
                <input
                  required
                  type="tel"
                  inputMode="numeric"
                  placeholder="Ej: 2215918286"
                  value={telefono}
                  onChange={(e) => setTelefono(e.target.value)}
                  className="mb-4 w-full rounded border border-gold/25 bg-cream px-3 py-2.5 text-sm outline-none ring-gold/30 focus:ring-2"
                />

                <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
                  Email (opcional, para confirmación)
                </label>
                <input
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="mb-4 w-full rounded border border-gold/25 bg-cream px-3 py-2.5 text-sm outline-none ring-gold/30 focus:ring-2"
                />

                <p className="mb-4 rounded border-l-4 border-gold bg-cream/80 px-3 py-2 text-xs leading-relaxed text-ink-muted">
                  Cancelaciones con código hasta{" "}
                  <strong>24 h antes</strong> del turno. Con menos tiempo,
                  contactanos por WhatsApp desde la home.
                </p>

                {reservaMsg && (
                  <p
                    className={`mb-4 rounded px-3 py-2 text-sm ${
                      reservaMsg.type === "ok"
                        ? "bg-emerald-50 text-emerald-900"
                        : "bg-red-50 text-red-800"
                    }`}
                  >
                    {reservaMsg.text}
                  </p>
                )}

                <button
                  type="submit"
                  disabled={reservando || !hora}
                  className="w-full rounded bg-gold py-3 text-xs font-semibold uppercase tracking-widest text-white transition-colors hover:bg-gold-dark disabled:cursor-not-allowed disabled:opacity-50"
                >
                  {reservando ? "Enviando…" : "Solicitar turno"}
                </button>
              </>
            )}
          </form>
        )}

        {tab === "cancelar" && (
          <form
            onSubmit={enviarCancelacion}
            className="rounded border border-gold/20 bg-white/90 p-6 shadow-sm"
          >
            <p className="mb-4 text-sm leading-relaxed text-ink-muted">
              El código te lo enviamos por mail al reservar. Podés cancelar sin
              cargo hasta <strong>24 h antes</strong> del horario.
            </p>
            <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-ink-muted">
              Código de cancelación
            </label>
            <input
              value={codigoCancel}
              onChange={(e) => setCodigoCancel(e.target.value)}
              placeholder="Ej: A1B2C3"
              className="mb-4 w-full rounded border border-gold/25 bg-cream px-3 py-2.5 text-sm uppercase outline-none ring-gold/30 focus:ring-2"
            />
            {cancelMsg && (
              <p
                className={`mb-4 rounded px-3 py-2 text-sm ${
                  cancelMsg.type === "ok"
                    ? "bg-emerald-50 text-emerald-900"
                    : "bg-red-50 text-red-800"
                }`}
              >
                {cancelMsg.text}
              </p>
            )}
            <button
              type="submit"
              disabled={cancelando}
              className="w-full rounded bg-ink-dark py-3 text-xs font-semibold uppercase tracking-widest text-white transition-opacity hover:opacity-90 disabled:opacity-50"
            >
              {cancelando ? "Procesando…" : "Cancelar mi turno"}
            </button>
          </form>
        )}

        <p className="mt-10 text-center">
          <Link
            href="/"
            className="text-sm font-medium text-gold-dark underline-offset-4 hover:underline"
          >
            ← Volver al inicio
          </Link>
        </p>
      </div>
    </div>
  );
}
