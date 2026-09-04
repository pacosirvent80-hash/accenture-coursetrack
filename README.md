# CourseTrack Pro Max — Gestor de Cursos eLearning

Herramienta web de gestión del ciclo de vida de cursos eLearning para oficinas técnicas de formación. Permite registrar y controlar el estado de cada curso desde la primera entrega del proveedor hasta su publicación en el LMS, con control de versiones, registro de incidencias con comentarios, trazabilidad por usuario, vista global de incidencias y sincronización en tiempo real con Google Sheets.

---

## Acceso

| Recurso | URL |
|---|---|
| **Aplicación web** | https://pacosirvent80-hash.github.io/accenture-coursetrack/ |
| **Repositorio GitHub** | https://github.com/pacosirvent80-hash/accenture-coursetrack |

> ⚠️ La URL del Google Apps Script otorga acceso de lectura y escritura sin autenticación.  
> No publicarla en canales abiertos. Cada usuario la introduce manualmente en ⚙ Configuración — nunca está embebida en el código fuente.

---

## Estructura de ficheros

```
accenture-coursetrack/
├── index.html                                  # Aplicación completa (HTML + CSS + JS, fichero único autocontenido)
├── README.md                                   # Este fichero
├── .gitignore
└── Documentacion/
    ├── CourseTrack_Manual_Usuario.docx         # Manual completo de usuario (v2.0)
    ├── gen_manual_coursetrack.py               # Script Python para regenerar el manual
    ├── CourseTrack_Guia_Usuario.docx           # Guía de usuario (versión anterior)
    ├── gen_guia_coursetrack.py                 # Script Python para regenerar la guía
    ├── CourseTrack_Novedades.docx              # Novedades por release
    ├── gen_novedades_coursetrack.py            # Script Python para regenerar novedades
    ├── CourseTrack_EspecificacionesTecnicas.docx
    ├── gen_coursetrack_tech.py
    └── cursos_prueba.json                      # 5 cursos de ejemplo para pruebas de importación
```

La aplicación es un fichero HTML autocontenido sin dependencias externas: sin frameworks, sin librerías externas, sin servidor de aplicaciones.

---

## Tecnologías

| Capa | Tecnología |
|---|---|
| Frontend | HTML5, CSS3, JavaScript ES2020+ (Vanilla) |
| Persistencia local | `localStorage` del navegador (claves: `coursetrack_v1`, `coursetrack_user`, `coursetrack_gas_url`, `coursetrack_theme`) |
| Persistencia compartida | Google Sheets + Google Apps Script REST |
| Hosting | GitHub Pages — rama `main`, raíz `/` |
| Responsive | CSS Grid, Flexbox, Media Queries (900px / 768px / 600px / 420px / 400px) |

---

## Funcionalidades principales

### Gestión de cursos

| Acción | Cómo se hace |
|---|---|
| Crear curso | Botón **+ Nuevo curso** en la cabecera |
| Editar curso | Botón **✏ Editar curso** en el pie del panel de detalle |
| Cambiar estado | Botón **↔ Cambiar estado** en el pie del panel |
| Nueva versión | Botón **+ Nueva versión** en el pie del panel |
| Duplicar curso | Botón **⊕ Duplicar** en el pie del panel — crea copia con versión 1.0 limpia |
| Eliminar curso | Botón **🗑** en el pie del panel (pide confirmación) |

### Incidencias

| Acción | Cómo se hace |
|---|---|
| Registrar incidencia | Botón **⚡ Nueva incidencia** en el pie del panel |
| Editar incidencia | Botón **✏ Editar incidencia** en cada tarjeta |
| Resolver incidencia | Botón **Resolver ✓** en cada tarjeta abierta |
| Añadir comentario | Campo de texto + botón **Enviar** (o Enter) en cada tarjeta de incidencia |
| Vista global | Toggle **⚡ Incidencias** en el toolbar — tabla plana de todas las incidencias abiertas del portfolio |

### Exportación e importación

| Formato | Descripción |
|---|---|
| **JSON** | Copia de seguridad completa (cursos, versiones, incidencias, comentarios) |
| **CSV** | Tabla plana para Excel u otras herramientas |
| **Excel (.xlsx)** | Fichero Excel listo para abrir |
| **Informe ejecutivo** | HTML autocontenido con KPIs, tabla de portfolio e incidencias abiertas por prioridad |
| **Importar JSON** | Modo fusión: añade cursos nuevos sin sobreescribir los existentes |

### Otras funcionalidades

- **Stats bar** con 8 contadores clicables: Total, Entregado, En revisión, Incidencias (total abiertas), En corrección, Aprobados, Publicados, En pausa.
- **Búsqueda en tiempo real** por nombre, código y proveedor.
- **Filtros** por estado, proveedor y usuario (acumulables).
- **Ordenación** por cualquier columna de la tabla (desktop).
- **Tema claro / oscuro** con persistencia en localStorage.
- **Historial de versiones** con línea de tiempo por curso.
- **Celebración gamificada** al publicar un curso (modal con confetti, cuenta atrás de 5 s y nombre del usuario).
- **Responsive completo**: desktop, tablet y móvil. Botón ← Volver y soporte del botón físico Atrás en Android.
- **Sincronización** con Google Sheets vía Apps Script (last-write-wins).

---

## Estados del curso

| Estado | Significado |
|---|---|
| 📥 Entregado | Recibido del proveedor, pendiente de revisión |
| 🔍 En revisión | El equipo L&D está revisando el contenido |
| ⚡ Incidencia | Bloqueado por incidencias críticas abiertas |
| 🔧 En corrección | El proveedor está aplicando correcciones |
| ✅ Aprobado | Validado, pendiente de publicar en el LMS |
| 🚀 Publicado | Activo y disponible para los alumnos |
| ⏸ En pausa | Desarrollo detenido temporalmente |

---

## Trazabilidad por usuario

Cada acción registrada incluye el nombre del usuario activo:

| Acción | Campo guardado |
|---|---|
| Crear curso / cambiar estado / nueva versión | `modificadoPor` en el objeto de versión |
| Registrar incidencia | `creadaPor` en el objeto de incidencia |
| Editar incidencia | `editadoPor` + `fechaEdicion` |
| Resolver incidencia | `resueltaPor` + `fechaResolucion` |
| Añadir comentario | `autor` + `fecha` en el objeto de comentario |

---

## Listas predefinidas

Tres arrays en `index.html` controlan los valores de los selectores. Para modificarlos: editar y hacer `git push` — GitHub Pages actualiza en 1–2 minutos.

```javascript
const USERS = ['Paco', 'Lourdes', 'Natalia'];       // miembros del equipo
const PROVIDERS = ['Accenture', 'INAP', 'Mercadona', 'Test Provider'];
const LANGUAGES = ['Español', 'Inglés', 'Catalán', 'Euskera', 'Gallego'];
```

---

## Configuración de Google Sheets

### 1. Preparar la hoja de cálculo
Crear una Google Sheet con dos pestañas: `Data` y `Log`.

### 2. Desplegar el Apps Script
- En la hoja: **Extensiones → Apps Script**.
- Pegar el código que aparece en el modal ⚙ Configuración de la aplicación.
- **Implementar → Nueva implementación → Aplicación web**:
  - Ejecutar como: **Yo**
  - Acceso: **Cualquiera (sin iniciar sesión)**
- Copiar la URL del Web App generada.

### 3. Conectar la app
Abrir la aplicación → ⚙ → pegar la URL del Web App → **Guardar y sincronizar**.

### Cómo se almacenan los datos
Los datos se serializan como JSON y se guardan en la celda **A1** de la pestaña `Data`. La pestaña `Log` registra cada escritura con fecha/hora y tamaño en bytes.

### Resetear los datos

Para empezar de cero necesitas limpiar los dos sitios donde viven los datos:

**1. Datos locales (este navegador)** → ⚙ Ajustes → **🗑 Limpiar datos locales** (zona de peligro al final del modal). Borra el `coursetrack_v1` del navegador y reinicia con los datos de ejemplo. No afecta a Google Sheets ni a otros usuarios. Cada compañero debe hacerlo en su propio navegador.

**2. Google Sheets (datos compartidos)** → Abrir la Google Sheet → pestaña **Data** → borrar el contenido de la celda **A1** → guardar. La próxima sincronización sobreescribirá A1 con el estado local de quien guarde primero.

> Exporta siempre una copia de seguridad JSON (**⬇ Exportar → JSON**) antes de borrar.

---

## Publicar cambios en el código

```bash
git clone https://github.com/pacosirvent80-hash/accenture-coursetrack.git
# editar index.html
git add index.html
git commit -m "descripción del cambio"
git push
```

GitHub Pages actualiza la URL pública en 1–2 minutos tras el push.

---

## Seguridad y privacidad

- La URL del Apps Script **no está en el código fuente** — cada usuario la introduce manualmente y se guarda en su `localStorage`.
- Los datos viajan cifrados por HTTPS entre el navegador y los servidores de Google.
- El modelo de concurrencia es **last-write-wins**. Para equipos pequeños (< 5 personas) esto es aceptable.
- La selección de usuario **no es autenticación**: está pensada para equipos de confianza donde la trazabilidad importa más que el control de acceso.

---

## Historial de versiones

| Versión | Fecha | Cambios destacados |
|---|---|---|
| v1.0 | may 2026 | Alta de cursos, estados, versiones, incidencias, Google Sheets, exportación JSON/CSV/Excel |
| v1.1 | jun 2026 | Responsive completo (tablet y móvil), soporte botón Atrás Android, importación en modo fusión |
| v1.2 | sep 2026 | Stats bar con 8 cajas (Entregado, En pausa), corrección conteo incidencias |
| v2.0 | sep 2026 | Vista global de incidencias, informe ejecutivo HTML, duplicar cursos, comentarios en incidencias, modal de celebración gamificado al publicar |
