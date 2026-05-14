# María Emilia Estética

Sitio y agenda de turnos para **María Emilia Estética**: landing en `index.html`, reservas y administración integradas con **Google Apps Script** y una hoja de cálculo (Google Sheets).

## Contenido del repositorio

| Archivo | Uso |
|--------|-----|
| `index.html` | Landing + vista de agenda embebida (reservar, cancelar, admin con sesión). |
| `agenda.html` | Solo agenda (despliegue `?page=agenda`). |
| `adminReset.html` | Página de nueva contraseña tras el enlace de recuperación (`?page=adminReset&token=…`). |
| `Código.gs` | Backend: servicios, horarios, turnos, emails, sesión admin, recuperación de contraseña. |

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

## Licencia

Propiedad del titular del negocio / desarrollador. Ajustá esta sección si publicás bajo una licencia abierta concreta.
