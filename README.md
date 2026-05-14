# María Emilia Estética

Sitio y agenda de turnos para **María Emilia Estética**: landing en `index.html`, reservas y administración integradas con **Google Apps Script** y una hoja de cálculo (Google Sheets).

## Contenido del repositorio

| Archivo | Uso |
|--------|-----|
| `index.html` | Landing + vista de agenda embebida (reservar, cancelar, admin con sesión). |
| `agenda.html` | Solo agenda (despliegue `?page=agenda`). |
| `adminReset.html` | Página de nueva contraseña tras el enlace de recuperación (`?page=adminReset&token=…`). |
| `vercel-site/` | Landing estática para Vercel que enlaza a la app de Apps Script. |
| `estetica-web/` | Next.js 15 (stack propio futuro): API + front en Vercel. Ver `docs/migracion-stack-propio.md`. |

**Next.js / Drizzle:** desde la raíz podés `npm run dev`, `npm run db:push`, etc. (los scripts hacen `cd estetica-web`). La primera vez: `cd estetica-web && npm install` o `npm install --prefix estetica-web`. **`DATABASE_URL`** (pooled) en `estetica-web/.env.local` para la API; para **`db:push`** en Windows, Neon recomienda también **`DATABASE_URL_DIRECT`** (conexión directa, sin pooler). Detalle en `estetica-web/.env.example`.

| `Código.gs` | Backend actual (Apps Script): servicios, horarios, turnos, mails, admin. |

## Requisitos en Google

1. **Hoja de cálculo** vinculada al proyecto de Apps Script, con hojas:
   - `Config` — servicios (columnas según tu implementación actual).
   - `Turnos` — reservas.
2. Proyecto de **Apps Script** con archivos HTML con **el mismo nombre** que aquí (`index`, `agenda`, `adminReset`) y un archivo `.gs` con el código de `Código.gs`.
3. **Implementación** como aplicación web:
   - Ejecutar como: quien corresponda.
   - Acceso: **Cualquiera** (para que clientes anónimos reserven) o el nivel que elijas según política.

## Configuración recomendada

En **Apps Script → Archivo → Propiedades del proyecto → Propiedades de secuencias de comandos** podés definir:

| Propiedad | Descripción |
|-----------|-------------|
| `ADMIN_PASSWORD` | Contraseña del panel de administración. |
| `ADMIN_RECOVERY_EMAIL` | Email (minúsculas) que recibe el enlace de “olvidé mi contraseña”. Si no existe, se usa el valor de `EMAIL_DUENIO` en `Código.gs`. |
| `WEB_APP_URL` | URL completa del despliegue (`…/exec`) si no se obtiene bien con `ScriptApp.getService().getUrl()`. |

Revisá también en `Código.gs` la constante `EMAIL_DUENIO` (avisos de nuevos turnos) y **no subas contraseñas en claro** al repositorio: usá propiedades del script o valores por defecto solo en entornos de prueba.

## URLs del despliegue

- Inicio / landing: URL de implementación sin parámetros (o `?page=index`).
- Agenda sola: `?page=agenda`.
- Reset de contraseña: `?page=adminReset&token=TOKEN` (el mail de recuperación arma el enlace).

## Personalización en HTML

- **Mapa “Cómo llegar”**: buscá `google.com/maps/search` en los HTML y reemplazá por el enlace **Compartir** de Google Maps de tu local.
- **Open Graph / Twitter**: metas `og:image` y `twitter:image` en `index.html` y `agenda.html`; conviene usar una imagen propia (URL absoluta pública).

## Sincronizar con GitHub

1. Cloná o descargá el repo.
2. Copiá el contenido de los archivos al editor de Apps Script **o** usá [clasp](https://github.com/google/clasp) para empujar/pull entre local y el proyecto en la nube.

Este repo sirve como **fuente de verdad** del código frente a lo que pegás en el editor web.

## Landing en Vercel (dominio público)

Los turnos **no pueden ejecutarse en Vercel** con el código actual (dependen de `google.script.run` y Sheets). Lo habitual es:

1. Publicar la agenda en **Apps Script** (URL `/exec`).
2. Publicar una **landing estática** en Vercel que derive a esa URL.

En este repo tenés la carpeta **`vercel-site/`**: landing mínima lista para Vercel. Configurá ahí `APP_SCRIPT_EXEC_URL` y en Vercel poné **Root Directory = `vercel-site`**. Los pasos detallados están en [`vercel-site/README.md`](vercel-site/README.md).

## Stack propio (futuro): Next.js en Vercel

Para **reemplazar** Apps Script por API + base de datos en el mismo dominio, existe la carpeta **`estetica-web/`** (Next.js 15). Guía de arquitectura y fases: [`docs/migracion-stack-propio.md`](docs/migracion-stack-propio.md). En Vercel usá **Root Directory = `estetica-web`**.

## Licencia

Propiedad del titular del negocio / desarrollador. Ajustá esta sección si publicás bajo una licencia abierta concreta.
