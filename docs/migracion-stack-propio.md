# Migración a stack propio (Vercel + API + DB)

Objetivo: **una sola URL** (dominio propio), landing + reserva + admin + mails, **sin** depender del HTML servido por Apps Script ni de `google.script.run`.

Referencia actual del negocio: `Código.gs`, `index.html`, `agenda.html`, hojas **`Config`** y **`Turnos`**.

---

## Stack recomendado (equipo pequeño)

| Capa | Elección típica |
|------|------------------|
| Frontend + rutas API | **Next.js** (App Router) en **Vercel** |
| Base de datos | **Neon** (Postgres; cliente en `estetica-web/src/lib/db.ts`) + Prisma/Drizzle cuando haya esquema |
| Email | **Resend** (simple) o SendGrid |
| Admin | Sesión con **cookie httpOnly** + secreto en env, o **Supabase Auth** solo para rol admin |
| Jobs (recordatorios) | **Vercel Cron** → ruta `/api/cron/recordatorios` |

Seguir usando **Google Sheets** como base vía API es posible, pero para concurrencia, locks y consultas complejas **Postgres conviene más** que replicar toda la lógica contra Sheets.

---

## Mapeo funcional (de `Código.gs` a tu API)

| Hoy (Apps Script) | Después (ejemplo) |
|-------------------|---------------------|
| `obtenerServicios()` | `GET /api/servicios` → tabla `services` (Drizzle + Neon en `estetica-web`) |
| `obtenerHorariosDisponibles` | `GET /api/horarios?servicio=&fecha=` (query) |
| `guardarTurno` | `POST /api/turnos` (JSON; Resend: `RESEND_API_KEY`, `EMAIL_FROM`, `OWNER_EMAIL`) |
| `cancelarTurno` | `POST /api/turnos/cancelar` con `{ "codigo": "..." }` |
| `obtenerTurnosHoy` / futuros | `GET /api/admin/turnos` con cookie de sesión admin |
| `verificarPasswordAdmin` + token | `POST /api/admin/login` → JWT o sesión firmada |
| `solicitarRecuperacionAdmin` | `POST /api/admin/recuperar` + token en tabla `password_reset_tokens` + Resend |
| `restablecerPasswordAdmin` | `POST /api/admin/restablecer` |

---

## Modelo de datos mínimo (Postgres)

- **`services`**: nombre, duración_min, responsable, capacidad, hora_inicio, hora_fin  
- **`appointments`**: id, fecha (`YYYY-MM-DD`), hora, nombre_cliente, teléfono, email (opcional), servicio_nombre, responsable, codigo_cancelacion (único), estado (`activo` \| `cancelado`), created_at  
- **`password_reset_tokens`**: token_hash, email, expires_at, used_at (nullable)

Índices únicos compuestos para evitar doble reserva: `(servicio_id, fecha, hora)` donde estado = activo, o lógica por `responsable` + cupo como hoy.

---

## Fases sugeridas

1. **Bootstrap** — Carpeta `estetica-web/` (Next.js), deploy en Vercel, `GET /api/health`.  
2. **Datos** — Migrar `Config` / `Turnos` a SQL (script one-off o CSV) o leer Sheets con service account (transitorio).  
3. **API pública** — Servicios + horarios + crear turno + cancelar (paridad con cliente actual).  
4. **Emails** — Resend con mismos textos que los HTML del mail en `Código.gs`.  
5. **Admin** — Login, listados, sin exponer datos sin sesión.  
6. **Recuperación** — Flujo con token y página `/admin/restablecer`.  
7. **Landing** — Portar diseño desde `index.html` a componentes React.  
8. **DNS** — Dominio apuntando a Vercel; retirar o redirigir el `/exec` viejo.

---

## Variables de entorno (Vercel)

Ver `estetica-web/.env.example`. Nunca commitees `.env.local` ni claves.

---

## Esfuerzo aproximado

Para un dev que ya conoce el negocio: **varias jornadas** solo en API + datos + mails + admin con calidad similar a lo actual. Sumar pruebas, límites de tasa y observabilidad si el sitio es público.

La carpeta **`estetica-web`** del repo es el arranque; el código legacy en la raíz puede convivir hasta completar la migración.
