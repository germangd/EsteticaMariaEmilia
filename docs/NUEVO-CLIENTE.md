# Nueva implementación (plantilla Estética / turnos + caja)

Este repositorio funciona como **motor reutilizable**: reservas, panel admin, caja y reportes. La marca y la landing se configuran en un solo archivo.

## Enfoque recomendado

| Opción | Cuándo usarla |
|--------|----------------|
| **Fork de este repo** (recomendado hoy) | Cada cliente = copia en GitHub + Vercel + Neon propios |
| Repo plantilla `estetica-platform-template` | Mismo código; GitHub → “Use this template” al crear el repo del cliente |
| Multi-tenant (un solo deploy) | Solo si más adelante querés SaaS; requiere mucho más desarrollo |

No hace falta un repo aparte para arrancar: podés **duplicar este proyecto** y editar la config. Si querés un repo “en blanco” sin historial de María Emilia, en GitHub: **Use this template** o cloná y borrá `.git` antes del primer commit del cliente.

---

## Archivo central de marca

Editá:

`estetica-web/src/config/site.config.ts`

Ahí van:

- Nombre del negocio, tagline, descripción SEO
- WhatsApp, mapa, Instagram
- Textos de la landing (hero, nosotros, zonas, tarjetas de servicios)
- Pie de mails

Helpers (no editar salvo casos especiales): `estetica-web/src/config/site.ts`

Variables de entorno opcionales (Vercel):

| Variable | Uso |
|----------|-----|
| `NEXT_PUBLIC_SITE_NAME` | Sobrescribe `businessName` |
| `NEXT_PUBLIC_WHATSAPP_NUMBER` | Número WA (solo dígitos, AR) |
| `DATABASE_URL` | Postgres (Neon) **del cliente** |
| `ADMIN_PASSWORD` | Clave del panel |
| `ADMIN_SESSION_SECRET` | Secreto de sesión admin |
| `RESEND_API_KEY` / SMTP | Mails de turnos |
| `EMAIL_FROM` | Remitente (ej. `Mi Salón <onboarding@resend.dev>`) |

---

## Checklist — nuevo cliente

### 1. Repositorio y deploy

- [ ] Fork o copia del repo (nombre del cliente)
- [ ] Proyecto en **Vercel** apuntando a `estetica-web`
- [ ] Dominio propio (ej. `misalon.com.ar`)

### 2. Base de datos (Neon)

- [ ] Proyecto Neon **nuevo** (no compartir con otro cliente)
- [ ] `DATABASE_URL` en Vercel
- [ ] Ejecutar migraciones en orden: `estetica-web/drizzle/*.sql` (0000 → última)
- [ ] Opcional: `npm run db:push` desde `estetica-web` si usás Drizzle push en dev

### 3. Configuración de marca

- [ ] Completar `src/config/site.config.ts`
- [ ] Reemplazar imágenes en `estetica-web/public/landing/hero/` y `public/landing/servicios/<carpeta>/`
- [ ] Alinear carpetas de servicios con `SERVICIO_MEDIA_FOLDERS` en `src/lib/servicio-media-folders.ts` y el script `scripts/landing-sync-manifest.cjs`
- [ ] `npm run build` en `estetica-web` (regenera manifest de landing)

### 4. Panel admin (datos operativos)

- [ ] Login `/admin/login`
- [ ] **Sedes** y **Horarios**
- [ ] **Servicios** (catálogo real; precios, anticipos, categorías)
- [ ] **Paquetes** / **Eventos** si aplican
- [ ] Probar **Reservar** (`/reservar`) y un turno de prueba
- [ ] Abrir **Caja**, venta de prueba, **Reportes**

### 5. Puesta en producción limpia

- [ ] Ejecutar `estetica-web/drizzle/reset-datos-prueba.sql` en Neon si hubo pruebas (borra ventas y agenda; **no** borra servicios)
- [ ] Verificar mails y WhatsApp con un turno real

### 6. Lo que no hay que tocar para otro cliente

- Código de APIs, repos, componentes admin (salvo personalización mayor)
- Lógica de cupo, anticipos, caja, reportes

---

## Landing a medida

Si la landing del cliente es **muy distinta** (otro rubro, otra estructura):

1. Igual editá `site.config.ts` para marca, contacto y SEO.
2. Podés reescribir `src/app/page.tsx` manteniendo imports de `@/config/site` para WA, mapa y nombre.
3. O crear `src/app/page.cliente-x.tsx` y exportar desde `page.tsx` según env (solo si tenés pocos clientes en un mono-repo).

Para la mayoría de centros de estética alcanza con cambiar **config + fotos**.

---

## Crear repo plantilla en GitHub (opcional)

1. En este repo: Settings → marcar **Template repository**
2. Nuevo repo del cliente: **Use this template**
3. Seguir checklist de arriba

---

## Soporte

Desarrollo base: [GestiónYa](https://gestion-ya.vercel.app)
