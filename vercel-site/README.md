# Landing en Vercel + turnos en Google

Vercel solo sirve **archivos estáticos** (esta carpeta). El **motor de turnos** (Sheets, mails, horarios) sigue en **Google Apps Script**. La landing de acá envía al usuario a tu URL `/exec`.

---

## Parte 1 — Google (desde cero)

### 1. Hoja de cálculo

1. Creá una hoja en [Google Sheets](https://sheets.google.com).
2. Dos hojas con nombres exactos: **`Config`** y **`Turnos`** (como espera `Código.gs` en la raíz del repo).
3. Completá `Config` con tus servicios según la lógica que ya tenés en el proyecto.

### 2. Apps Script

1. En la hoja: **Extensiones → Apps Script**.
2. Renombrá el proyecto (ej. “María Emilia Turnos”).
3. Subí o pegá el contenido de:
   - `Código.gs` (código del servidor),
   - **Archivo → Nuevo → Archivo HTML** con nombres **`index`**, **`agenda`**, **`adminReset`** y el mismo contenido que los `.html` de la raíz del repo (sin la extensión en el editor; el archivo en disco sigue siendo `.html` al clonar).

### 3. Publicar la aplicación web

1. **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. Ejecutar como: tu cuenta.
4. Quién tiene acceso: **Cualquiera** (para que clientes sin login reserven).
5. **Implementar** y copiá la **URL** (debe terminar en `/exec`).

### 4. Propiedades (recomendado)

**Archivo → Propiedades del proyecto → Propiedades de secuencias de comandos**: `ADMIN_PASSWORD`, `ADMIN_RECOVERY_EMAIL`, `WEB_APP_URL` (esta última = la misma URL `/exec` si el mail de recuperación no arma bien el enlace).

---

## Parte 2 — Vercel

### Opción A — Desde GitHub (recomendado)

1. Subí este repo a GitHub (si no está).
2. Entrá a [vercel.com](https://vercel.com) → **Add New → Project** → importá el repo.
3. **Root Directory**: elegí **`vercel-site`** (importante: no la raíz del repo, así no se publica el `index.html` de Apps Script por error).
4. Framework: **Other** o detección automática; no hace falta comando de build.
5. **Deploy**.

### Opción B — Desde la consola

```bash
cd vercel-site
npx vercel
```

Seguí el asistente; enlazá con tu cuenta de Vercel.

### Después del deploy

1. Abrí `vercel-site/index.html` en el editor.
2. Buscá la línea `var APP_SCRIPT_EXEC_URL = "PEGAR_AQUI_TU_URL_DE_APPS_SCRIPT_EXEC";`
3. Reemplazá el valor por tu URL real (`https://script.google.com/macros/s/.../exec`).
4. Volvé a hacer **commit** y **push**; Vercel redeploya solo, o ejecutá `vercel --prod` de nuevo.

---

## Resumen

| Dónde | Qué hace |
|--------|----------|
| **Vercel** (`vercel-site`) | Landing pública, marca, botón a WhatsApp y botón “Reservar” → Apps Script. |
| **Apps Script** (`/exec`) | Formulario de turnos, cancelación, admin, mails, hoja `Turnos`. |

Si más adelante querés **todo** en Vercel (API + base de datos), es otro desarrollo: habría que reemplazar `Código.gs` por backend propio.
