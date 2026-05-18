"use client";

import { useRouter } from "next/navigation";
import { useCallback, useEffect, useState } from "react";
import type { ClienteDetalle, ClienteResumen } from "@/lib/clientes-repo";

function fmtFecha(iso: string): string {
  const p = iso.split("-");
  if (p.length !== 3) return iso;
  return `${p[2]}/${p[1]}/${p[0]}`;
}

function fmtPesos(n: number | null): string {
  if (n == null) return "—";
  return n.toLocaleString("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  });
}

export function AdminClientesPanel({
  initialClientes,
  initialDetalle,
  telefonoSeleccionado,
  busquedaInicial,
}: {
  initialClientes: ClienteResumen[];
  initialDetalle: ClienteDetalle | null;
  telefonoSeleccionado: string | null;
  busquedaInicial: string;
}) {
  const router = useRouter();
  const [clientes, setClientes] = useState(initialClientes);
  const [detalle, setDetalle] = useState<ClienteDetalle | null>(initialDetalle);
  const [q, setQ] = useState(busquedaInicial);
  const [msg, setMsg] = useState<string | null>(null);
  const [pending, setPending] = useState(false);

  const [nombre, setNombre] = useState(initialDetalle?.perfil.nombre ?? "");
  const [email, setEmail] = useState(initialDetalle?.perfil.email ?? "");
  const [notas, setNotas] = useState(initialDetalle?.perfil.notas ?? "");

  const tel = telefonoSeleccionado;

  useEffect(() => {
    setClientes(initialClientes);
    setDetalle(initialDetalle);
    setNombre(initialDetalle?.perfil.nombre ?? "");
    setEmail(initialDetalle?.perfil.email ?? "");
    setNotas(initialDetalle?.perfil.notas ?? "");
  }, [initialClientes, initialDetalle]);

  const irCliente = useCallback(
    (telefono: string | null, busqueda?: string) => {
      const p = new URLSearchParams();
      const b = busqueda ?? q;
      if (b.trim()) p.set("q", b.trim());
      if (telefono) p.set("tel", telefono);
      const s = p.toString();
      router.push(s ? `/admin/clientes?${s}` : "/admin/clientes");
    },
    [router, q]
  );

  async function buscar(e: React.FormEvent) {
    e.preventDefault();
    irCliente(tel, q);
  }

  async function guardarPerfil(e: React.FormEvent) {
    e.preventDefault();
    if (!tel) return;
    setPending(true);
    setMsg(null);
    try {
      const r = await fetch(`/api/admin/clientes/${encodeURIComponent(tel)}`, {
        method: "PATCH",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ nombre, email, notas }),
      });
      const data = (await r.json()) as { ok?: boolean; mensaje?: string };
      if (!r.ok || !data.ok) {
        setMsg(data.mensaje ?? "No se pudo guardar.");
        return;
      }
      setMsg("Datos del cliente guardados.");
      irCliente(tel, q);
    } finally {
      setPending(false);
    }
  }

  const inputClass =
    "w-full rounded-sm border border-gold/30 bg-cream px-3 py-2 text-sm text-ink";

  return (
    <div className="grid gap-8 lg:grid-cols-[minmax(0,1fr)_minmax(0,1.2fr)]">
      <section>
        <form onSubmit={(e) => void buscar(e)} className="mb-4 flex gap-2">
          <input
            type="search"
            placeholder="Nombre, teléfono o email…"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            className={`${inputClass} flex-1`}
          />
          <button
            type="submit"
            className="rounded-sm bg-gold px-4 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-white hover:bg-gold-dark"
          >
            Buscar
          </button>
        </form>

        <p className="mb-3 text-sm text-ink-muted">
          {clientes.length} cliente{clientes.length === 1 ? "" : "s"}
        </p>

        <div className="max-h-[70vh] overflow-y-auto rounded-sm border border-gold/25 bg-white shadow-sm">
          {clientes.length === 0 ? (
            <p className="p-6 text-center text-sm text-ink-muted">
              No hay clientes con turnos o paquetes aún.
            </p>
          ) : (
            <ul className="divide-y divide-gold/10">
              {clientes.map((c) => (
                <li key={c.telefono}>
                  <button
                    type="button"
                    onClick={() => irCliente(c.telefono)}
                    className={`w-full px-4 py-3 text-left transition hover:bg-cream/80 ${
                      tel === c.telefono ? "bg-cream" : ""
                    }`}
                  >
                    <div className="font-medium text-ink-dark">{c.nombre}</div>
                    <div className="text-xs text-ink-muted">{c.telefono}</div>
                    <div className="mt-1 flex flex-wrap gap-2 text-[0.65rem] text-ink-muted">
                      <span>{c.turnosTotal} turno{c.turnosTotal === 1 ? "" : "s"}</span>
                      {c.paquetesActivos > 0 ? (
                        <span>· {c.paquetesActivos} paquete activo</span>
                      ) : null}
                      {c.ultimaFecha ? (
                        <span>· último {fmtFecha(c.ultimaFecha)}</span>
                      ) : null}
                      {c.tieneNotas ? <span>· nota</span> : null}
                    </div>
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
      </section>

      <section className="rounded-sm border border-gold/20 bg-white p-5 shadow-sm md:p-6">
        {!tel || !detalle ? (
          <p className="text-sm text-ink-muted">
            Elegí un cliente de la lista para ver su historial de servicios y
            paquetes.
          </p>
        ) : (
          <>
            <div className="mb-6 border-b border-gold/15 pb-4">
              <h2 className="font-serif text-xl font-normal text-ink-dark">
                {detalle.perfil.nombre ?? "Cliente"}
              </h2>
              <p className="text-sm text-ink-muted">{detalle.telefono}</p>
            </div>

            <form
              onSubmit={(e) => void guardarPerfil(e)}
              className="mb-8 grid gap-3 border-b border-gold/15 pb-6"
            >
              <h3 className="text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
                Ficha del cliente
              </h3>
              <div>
                <label className="mb-1 block text-xs text-ink-muted">Nombre</label>
                <input
                  className={inputClass}
                  value={nombre}
                  onChange={(e) => setNombre(e.target.value)}
                />
              </div>
              <div>
                <label className="mb-1 block text-xs text-ink-muted">Email</label>
                <input
                  type="email"
                  className={inputClass}
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                />
              </div>
              <div>
                <label className="mb-1 block text-xs text-ink-muted">Notas</label>
                <textarea
                  rows={3}
                  className={inputClass}
                  value={notas}
                  onChange={(e) => setNotas(e.target.value)}
                />
              </div>
              <button
                type="submit"
                disabled={pending}
                className="rounded-sm border border-gold/40 px-4 py-2 text-[0.72rem] font-medium uppercase tracking-wider text-gold-dark hover:bg-cream disabled:opacity-50"
              >
                Guardar ficha
              </button>
            </form>

            <h3 className="mb-3 text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Historial de turnos ({detalle.turnos.length})
            </h3>
            {detalle.turnos.length === 0 ? (
              <p className="mb-6 text-sm text-ink-muted">Sin turnos registrados.</p>
            ) : (
              <div className="mb-8 overflow-x-auto">
                <table className="w-full min-w-[480px] text-left text-sm">
                  <thead className="border-b border-gold/20 text-[0.65rem] uppercase tracking-wide text-ink-muted">
                    <tr>
                      <th className="py-2 pr-2">Fecha</th>
                      <th className="py-2 pr-2">Hora</th>
                      <th className="py-2 pr-2">Servicio</th>
                      <th className="py-2">Estado</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gold/10">
                    {detalle.turnos.map((t) => (
                      <tr key={t.id}>
                        <td className="py-2 pr-2 whitespace-nowrap">
                          {fmtFecha(t.fecha)}
                        </td>
                        <td className="py-2 pr-2">{t.hora}</td>
                        <td className="py-2 pr-2">{t.servicioNombre}</td>
                        <td className="py-2 capitalize">{t.estado}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <h3 className="mb-3 text-[0.65rem] font-medium uppercase tracking-wider text-ink-muted">
              Paquetes ({detalle.paquetes.length})
            </h3>
            {detalle.paquetes.length === 0 ? (
              <p className="text-sm text-ink-muted">Sin paquetes.</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full min-w-[480px] text-left text-sm">
                  <thead className="border-b border-gold/20 text-[0.65rem] uppercase tracking-wide text-ink-muted">
                    <tr>
                      <th className="py-2 pr-2">Paquete</th>
                      <th className="py-2 pr-2">Sesiones</th>
                      <th className="py-2 pr-2">Cobrado</th>
                      <th className="py-2">Estado</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gold/10">
                    {detalle.paquetes.map((p) => (
                      <tr key={p.id}>
                        <td className="py-2 pr-2">{p.paqueteNombre}</td>
                        <td className="py-2 pr-2">
                          {p.sesionesRestantes}/{p.sesionesIniciales}
                        </td>
                        <td className="py-2 pr-2">
                          {fmtPesos(p.precioCobradoPesos)}
                        </td>
                        <td className="py-2 capitalize">{p.estado}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {msg ? <p className="mt-4 text-sm text-ink-muted">{msg}</p> : null}
          </>
        )}
      </section>
    </div>
  );
}
