"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useState } from "react";
import { MeLogo } from "@/components/landing/me-logo";
import { ReservaCatalogoSelect } from "@/components/reservar/reserva-catalogo-select";
import { dedupeServiciosPorNombre } from "@/lib/servicio-format";
import { parsearClaveReserva } from "@/lib/reserva-claves";
import {
  uiBtnDark,
  uiBtnPrimaryFull,
  uiReservarForm,
  uiHint,
  uiInput,
  uiLabel,
  uiSelect,
  uiTabActive,
  uiTabBar,
  uiTabInactive,
  uiTimeSlot,
  uiTimeSlotActive,
} from "@/lib/ui-classes";
import { buildWhatsAppTurnoUrl } from "@/lib/whatsapp";

type ServicioApi = {
  id: number;
  nombre: string;
  duracion: number;
  responsable: string;
  capacidad: number;
  horarioInicio: string;
  horarioFin: string;
  parentId?: number | null;
  esGrupo?: boolean;
  categoriaNombre?: string | null;
};

type PaqueteApi = {
  id: number;
  nombre: string;
  descripcion: string | null;
  precioPesos: number;
  sesionesTotal: number;
  serviciosIncluidos: string[];
  duracion: number;
};

type Tab = "reservar" | "cancelar";

export function ReservarClient() {
  const [tab, setTab] = useState<Tab>("reservar");

  const [sedes, setSedes] = useState<{ id: number; nombre: string }[]>([]);
  const [sedeId, setSedeId] = useState<number>(0);
  const [sedesError, setSedesError] = useState<string | null>(null);

  const [servicios, setServicios] = useState<ServicioApi[]>([]);
  const [paquetes, setPaquetes] = useState<PaqueteApi[]>([]);
  const [catalogoError, setCatalogoError] = useState<string | null>(null);
  const [loadingCatalogo, setLoadingCatalogo] = useState(true);

  const [claveReserva, setClaveReserva] = useState("");
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
  const [ultimaReserva, setUltimaReserva] = useState<{
    codigo: string;
    servicio: string;
    fecha: string;
    hora: string;
    nombre: string;
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

  const itemSel = useMemo(() => {
    const parsed = parsearClaveReserva(claveReserva);
    if (!parsed) return null;
    if (parsed.tipo === "servicio") {
      const s = servicios.find((x) => x.nombre === parsed.nombre);
      if (!s) return null;
      return {
        tipo: "servicio" as const,
        nombre: s.nombre,
        duracion: s.duracion,
        sesionesTotal: undefined as number | undefined,
      };
    }
    const p = paquetes.find((x) => x.id === parsed.id);
    if (!p) return null;
    return {
      tipo: "paquete" as const,
      nombre: p.nombre,
      duracion: p.duracion,
      sesionesTotal: p.sesionesTotal,
    };
  }, [claveReserva, servicios, paquetes]);

  const duracionEtiqueta = useMemo(() => {
    const min = itemSel?.duracion;
    if (!min) return null;
    let base: string;
    if (min < 60) base = `${min} min`;
    else {
      const h = Math.floor(min / 60);
      const m = min % 60;
      base = m > 0 ? `${h} h ${m} min` : `${h} h`;
    }
    if (itemSel?.tipo === "paquete" && (itemSel.sesionesTotal ?? 0) > 1) {
      return `${base} por sesión · paquete de ${itemSel.sesionesTotal} sesiones`;
    }
    return base;
  }, [itemSel]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const rSedes = await fetch("/api/sedes", { cache: "no-store" });
        const dataSedes = (await rSedes.json()) as {
          ok?: boolean;
          sedes?: { id: number; nombre: string }[];
        };
        if (cancelled) return;
        if (dataSedes.ok && dataSedes.sedes?.length) {
          setSedes(dataSedes.sedes);
          setSedeId(dataSedes.sedes[0]!.id);
        } else {
          setSedesError("No hay sedes disponibles.");
        }
      } catch {
        if (!cancelled) setSedesError("Error al cargar sedes.");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingCatalogo(true);
      setCatalogoError(null);
      try {
        const [rSvc, rPkg] = await Promise.all([
          fetch("/api/servicios", { cache: "no-store" }),
          fetch("/api/paquetes", { cache: "no-store" }),
        ]);
        const dataSvc = (await rSvc.json()) as {
          ok?: boolean;
          servicios?: ServicioApi[];
          mensaje?: string;
        };
        const dataPkg = (await rPkg.json()) as {
          ok?: boolean;
          paquetes?: PaqueteApi[];
        };
        if (cancelled) return;
        if (!rSvc.ok || !dataSvc.ok || !Array.isArray(dataSvc.servicios)) {
          setCatalogoError(
            dataSvc.mensaje ?? "No se pudieron cargar los servicios."
          );
          setServicios([]);
          setPaquetes([]);
          return;
        }
        setServicios(dedupeServiciosPorNombre(dataSvc.servicios));
        setPaquetes(
          rPkg.ok && dataPkg.ok && Array.isArray(dataPkg.paquetes)
            ? dataPkg.paquetes
            : []
        );
      } catch {
        if (!cancelled) setCatalogoError("Error de red al cargar el catálogo.");
      } finally {
        if (!cancelled) setLoadingCatalogo(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const cargarHorarios = useCallback(async () => {
    if (!claveReserva || !fecha || sedeId < 1) {
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
        clave: claveReserva,
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
        setHorariosError(data.mensaje ?? "No se pudieron obtener horarios.");
        return;
      }
      const lista = Array.isArray(data.horarios) ? data.horarios : [];
      setHorarios(lista);
      if (lista.length === 0 && data.mensaje) {
        setHorariosError(data.mensaje);
      }
    } catch {
      setHorarios([]);
      setHorariosError("Error de red al cargar horarios.");
    } finally {
      setLoadingHorarios(false);
    }
  }, [claveReserva, fecha, sedeId]);

  useEffect(() => {
    void cargarHorarios();
  }, [cargarHorarios]);

  async function enviarReserva(e: React.FormEvent) {
    e.preventDefault();
    setReservaMsg(null);
    setUltimaReserva(null);
    if (
      !claveReserva ||
      !itemSel ||
      !fecha ||
      !hora ||
      !nombre.trim() ||
      !telefono.trim()
    ) {
      setReservaMsg({
        type: "err",
        text: "Completá servicio o combo, fecha, horario, nombre y teléfono.",
      });
      return;
    }
    setReservando(true);
    try {
      const r = await fetch("/api/turnos", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          clave: claveReserva,
          sedeId,
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
        const nombreGuardado = nombre.trim();
        if (data.codigo) {
          setUltimaReserva({
            codigo: data.codigo,
            servicio: itemSel.nombre,
            fecha,
            hora,
            nombre: nombreGuardado,
          });
          setReservaMsg({
            type: "ok",
            text: "Turno confirmado. Guardá el código de abajo.",
          });
        } else {
          setReservaMsg({
            type: "ok",
            text: data.mensaje ?? "Turno confirmado.",
          });
        }
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
    <div className="min-h-screen bg-surface pb-16 pt-24 text-ink">
      <div className="mx-auto max-w-lg px-4">
        <header className="mb-8 text-center">
          <Link
            href="/"
            className="group inline-flex flex-col items-center gap-3 rounded-sm px-2 py-1 transition-opacity hover:opacity-85"
            aria-label="Volver a la página principal"
          >
            <span className="rounded-full border border-gold/25 bg-white/80 p-3 shadow-sm transition group-hover:border-gold/45">
              <MeLogo gradientId="meReservarNav" className="h-11 w-[4.75rem]" />
            </span>
            <span className="text-xs font-medium uppercase tracking-[0.25em] text-gold-dark">
              María Emilia Estética
            </span>
          </Link>
        </header>
        <h1 className="mb-2 text-center font-serif text-2xl font-medium text-ink-dark md:text-3xl">
          Turnos online
        </h1>
        <p className="mb-8 text-center text-sm font-medium text-ink">
          Elegí servicio o combo, fecha y horario. Al confirmar verás un{" "}
          <strong className="font-semibold text-ink-dark">código en pantalla</strong>:
          guardalo para cancelar o para consultarnos.
        </p>

        <div className={uiTabBar}>
          <button
            type="button"
            onClick={() => setTab("reservar")}
            className={tab === "reservar" ? uiTabActive : uiTabInactive}
          >
            Reservar
          </button>
          <button
            type="button"
            onClick={() => setTab("cancelar")}
            className={tab === "cancelar" ? uiTabActive : uiTabInactive}
          >
            Cancelar
          </button>
        </div>

        {tab === "reservar" && (
          <form onSubmit={enviarReserva} className={uiReservarForm}>
            {loadingCatalogo ? (
              <p className="text-center text-sm text-ink-muted">
                Cargando servicios y combos…
              </p>
            ) : catalogoError ? (
              <p className="rounded bg-red-50 px-3 py-2 text-center text-sm text-red-800">
                {catalogoError}
              </p>
            ) : servicios.length === 0 && paquetes.length === 0 ? (
              <p className="text-center text-sm text-ink-muted">
                No hay servicios ni combos disponibles. Configuralos en el panel
                de administración.
              </p>
            ) : sedesError ? (
              <p className="rounded bg-red-50 px-3 py-2 text-center text-sm text-red-800">
                {sedesError}
              </p>
            ) : (
              <>
                <label className={uiLabel}>Sede</label>
                <select
                  required
                  className={`mb-4 ${uiSelect}`}
                  value={sedeId}
                  onChange={(e) => setSedeId(Number(e.target.value))}
                >
                  {sedes.map((s) => (
                    <option key={s.id} value={s.id}>
                      {s.nombre}
                    </option>
                  ))}
                </select>

                <ReservaCatalogoSelect
                  servicios={servicios.map((s) => ({
                    id: s.id,
                    nombre: s.nombre,
                    parentId: s.parentId ?? null,
                    categoriaNombre: s.categoriaNombre,
                  }))}
                  paquetes={paquetes.map((p) => ({
                    id: p.id,
                    nombre: p.nombre,
                    precioPesos: p.precioPesos,
                    sesionesTotal: p.sesionesTotal,
                    serviciosIncluidos: p.serviciosIncluidos,
                  }))}
                  value={claveReserva}
                  onChange={setClaveReserva}
                  selectClassName={uiSelect}
                />

                <label className={uiLabel}>Fecha</label>
                <input
                  type="date"
                  required
                  min={minFecha}
                  max={maxFecha}
                  value={fecha}
                  onChange={(e) => setFecha(e.target.value)}
                  className={`mb-4 ${uiInput}`}
                />

                <label className={uiLabel}>Horario</label>
                {duracionEtiqueta ? (
                  <p className={`mb-2 ${uiHint}`}>
                    Duración estimada: <strong>{duracionEtiqueta}</strong>.
                    Los turnos se ofrecen cada ese intervalo.
                  </p>
                ) : null}
                {loadingHorarios ? (
                  <p className="mb-4 text-sm text-ink-muted">
                    Buscando horarios…
                  </p>
                ) : horariosError ? (
                  <p className="mb-4 rounded bg-amber-50 px-3 py-2 text-sm text-amber-900">
                    {horariosError}
                  </p>
                ) : !claveReserva || !fecha ? (
                  <p className="mb-4 text-sm text-ink-muted">
                    Completá la selección y la fecha para ver horarios.
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
                        className={hora === h ? uiTimeSlotActive : uiTimeSlot}
                      >
                        {h}
                      </button>
                    ))}
                  </div>
                )}

                <label className={uiLabel}>Nombre completo</label>
                <input
                  required
                  value={nombre}
                  onChange={(e) => setNombre(e.target.value)}
                  className={`mb-4 ${uiInput}`}
                />

                <label className={uiLabel}>Teléfono (código de área, sin 0 ni 15)</label>
                <input
                  required
                  type="tel"
                  inputMode="numeric"
                  placeholder="Ej: 2215918286"
                  value={telefono}
                  onChange={(e) => setTelefono(e.target.value)}
                  className={`mb-4 ${uiInput}`}
                />

                <label className={uiLabel}>Email (opcional)</label>
                <p className={`mb-2 ${uiHint}`}>
                  Si lo dejás, podemos enviarte confirmación cuando el correo del
                  salón esté activo.{" "}
                  <strong className="font-semibold text-ink-dark">
                    El código oficial aparece siempre al confirmar abajo.
                  </strong>
                </p>
                <input
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className={`mb-4 ${uiInput}`}
                />

                <p className="mb-4 rounded-sm border-l-4 border-gold bg-white/70 px-3 py-2.5 text-xs font-medium leading-relaxed text-ink">
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

                {ultimaReserva && reservaMsg?.type === "ok" ? (
                  <div className="mb-4 rounded-sm border border-gold/45 bg-white/80 p-4 text-center shadow-sm">
                    <p className="text-[0.65rem] font-bold uppercase tracking-wider text-ink-dark">
                      Tu código de cancelación
                    </p>
                    <p className="my-2 font-mono text-2xl font-semibold tracking-[0.2em] text-gold-dark">
                      {ultimaReserva.codigo}
                    </p>
                    <p className="mb-4 text-xs text-ink-muted">
                      {ultimaReserva.servicio} · {ultimaReserva.fecha}{" "}
                      {ultimaReserva.hora}
                    </p>
                    <a
                      href={buildWhatsAppTurnoUrl(ultimaReserva)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex w-full items-center justify-center gap-2 rounded bg-[#25D366] py-3 text-xs font-semibold uppercase tracking-widest text-white transition-opacity hover:opacity-90"
                    >
                      Consultar por WhatsApp
                    </a>
                    <p className="mt-2 text-[0.65rem] text-ink-muted">
                      Abrís el chat con el mensaje listo; tocá Enviar en
                      WhatsApp.
                    </p>
                  </div>
                ) : null}

                <button
                  type="submit"
                  disabled={reservando || !hora}
                  className={uiBtnPrimaryFull}
                >
                  {reservando ? "Enviando…" : "Solicitar turno"}
                </button>
              </>
            )}
          </form>
        )}

        {tab === "cancelar" && (
          <form onSubmit={enviarCancelacion} className={uiReservarForm}>
            <p className={`mb-4 ${uiHint}`}>
              El código aparece en pantalla al reservar (anotalo). Podés
              cancelar sin cargo hasta <strong className="font-semibold">24 h antes</strong> del horario.
            </p>
            <label className={uiLabel}>Código de cancelación</label>
            <input
              value={codigoCancel}
              onChange={(e) => setCodigoCancel(e.target.value)}
              placeholder="Ej: A1B2C3"
              className={`mb-4 uppercase ${uiInput}`}
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
              className={uiBtnDark}
            >
              {cancelando ? "Procesando…" : "Cancelar mi turno"}
            </button>
          </form>
        )}

        <footer className="mt-10 border-t border-gold/15 pt-6 text-center">
          <Link
            href="/"
            className="text-sm font-medium text-gold-dark underline-offset-4 hover:underline"
          >
            ← Volver al inicio
          </Link>
          <p className="mt-4">
            <Link
              href="/admin/login"
              className="text-[0.65rem] font-normal uppercase tracking-[0.12em] text-ink-muted/55 transition-colors hover:text-ink-muted"
            >
              Acceso staff
            </Link>
          </p>
        </footer>
      </div>
    </div>
  );
}
