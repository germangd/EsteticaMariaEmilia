# estetica-web (Next.js → Vercel)

Esqueleto para la **opción 3**: sitio propio con API en el mismo proyecto. La guía de migración desde Apps Script está en [`../docs/migracion-stack-propio.md`](../docs/migracion-stack-propio.md).

## Requisitos

- Node.js 20+

## Instalación

```bash
cd estetica-web
npm install
npm run dev
```

Desde la **raíz del repo** también podés: `npm run dev`, `npm run db:push`, etc. Variables en **`estetica-web/.env.local`** (o `.env.local` en la raíz; ver `drizzle.config.ts`). El script `dev` libera **3000/3001** antes de arrancar para evitar EPERM en `.next/trace` en Windows; si necesitás Next sin eso: `npm run dev:raw`.

Abrí [http://localhost:3000](http://localhost:3000) y probá [http://localhost:3000/api/health](http://localhost:3000/api/health). La **home** replica la landing del `index.html` del repo (hero, servicios, zonas, CTA). **`/reservar`** carga servicios desde **`GET /api/servicios`**, horarios con **`GET /api/horarios`**, reserva con **`POST /api/turnos`** y cancelación con **`POST /api/turnos/cancelar`**.

## Base de datos (Neon)

### Rama única (`production`) y `estetica-web/.env.local`

No puedo entrar a tu cuenta de Neon ni editar tu `.env.local` (son secretos en tu máquina). Sí podés alinearlo en un minuto:

1. [Neon Console](https://console.neon.tech) → tu proyecto.
2. **Branches** → abrí **`production`** (o la que diga **Default**). Para el día a día evitá ramas `import-...` salvo que quieras trabajar solo en esa rama.
3. Arriba en la consola: rama = **production**, base de datos = **`neondb`**.
4. **Connect** → copiá **Pooled** → pegalo en `estetica-web/.env.local` como **`DATABASE_URL="..."`**.
5. (Opcional) En el mismo modal, conexión **directa** (sin pooling / sin `-pooler` en el host) → **`DATABASE_URL_DIRECT="..."`** ([Neon + Drizzle](https://neon.com/docs/guides/drizzle-migrations)).
6. Guardá el archivo. Reiniciá `npm run dev` si ya estaba corriendo.

Luego:

1. Volvé a cargar `/api/health`: debería responder `"db": "ok"`.
2. **Esquema**: `npm run db:push` **o** pegá `drizzle/0000_init.sql` en el SQL Editor de Neon **o** `npm run db:migrate`. Si `db:push` se corta, usá **`DATABASE_URL_DIRECT`**; el proyecto ya incluye **`ws`** en `devDependencies` y `drizzle.config.ts` registra WebSocket para Node.
3. (Opcional) Datos de prueba: `drizzle/seed_example.sql` en el SQL Editor.
4. Probá **`GET /api/servicios`**: debe devolver `{ ok: true, servicios: [...] }` en el mismo formato que `obtenerServicios()` del Apps Script (`nombre`, `duracion`, `responsable`, `capacidad`, `horarioInicio`, `horarioFin`).
5. Tras **`npm run db:push`** (o migración `0001_appointments.sql`), APIs de agenda:
   - **`GET /api/horarios?servicio=Manicura&fecha=2026-05-20`**
   - **`POST /api/turnos`** con JSON `{ "servicio", "fecha", "hora", "nombre", "telefono", "email"? }` → `{ exito, mensaje, codigo }`
   - **`POST /api/turnos/cancelar`** con `{ "codigo": "..." }` → `{ exito, mensaje }`

Si todavía no tenés proyecto en Neon: creá uno en [neon.tech](https://neon.tech); en **Connect** verás **Pooled** y **Direct**.

### Scripts de base de datos

| Comando | Uso |
|--------|-----|
| `npm run db:generate` | Genera SQL en `drizzle/` a partir de `src/db/schema.ts`. |
| `npm run db:push` | Aplica el esquema al Neon configurado en `.env.local`. |
| `npm run db:migrate` | Aplica migraciones pendientes desde `drizzle/` (tabla `__drizzle_migrations`). |
| `npm run db:studio` | UI local (usa la misma URL que Drizzle Kit: `DATABASE_URL_DIRECT` si existe). |

El cliente está en `src/lib/db.ts` (`@neondatabase/serverless`). El ORM y el esquema: `src/db/`.

## Deploy en Vercel

1. Importá el repo en Vercel.
2. **Root Directory**: `estetica-web` (obligatorio para que Vercel lea este `package.json` con `next`). Si ves *“No Next.js version detected”*, el Root Directory no apunta a esta carpeta.
3. **Framework Preset**: **Next.js** (no “Other”). Si el proyecto se creó como estático, en la raíz del repo existe **`vercel.json`** que fuerza Next cuando el root del proyecto es el repo entero; igual lo más estable es root **`estetica-web`**.
4. **Output Directory**: sin override (vacío). Si quedó `public` del preset *Other*, borralo.
5. Build: `npm run build` (por defecto si el root del proyecto es `estetica-web`; desde la raíz del repo ya está cableado en el `package.json` de la raíz).
6. Variables de entorno: `DATABASE_URL`, `EMAIL_FROM`, `RESEND_API_KEY`, `OWNER_EMAIL` (y las opcionales de `.env.example`).

### Resend (mails al reservar)

Tras crear un turno con **`POST /api/turnos`**, si hay **`RESEND_API_KEY`** y **`EMAIL_FROM`**, se envía mail al cliente (si mandó `email`) y al dueño si definís **`OWNER_EMAIL`**. Dominio del remitente: [Resend → Domains](https://resend.com/domains).

## Seguridad

- Mantené `next` y `eslint-config-next` alineados en versiones parcheadas (`npm outdated` / avisos de Vercel).
