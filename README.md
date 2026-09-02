# CourseTrack — Gestor de Cursos eLearning

Herramienta web de gestión del ciclo de vida de cursos eLearning para oficinas técnicas de formación. Permite registrar y controlar el estado de cada curso desde la primera entrega del proveedor hasta su publicación en el LMS, con control de versiones, registro de incidencias, trazabilidad por usuario y sincronización en tiempo real con Google Sheets.

---

## Acceso

| Recurso | URL |
|---|---|
| **Aplicación web** | https://pacosirvent80-hash.github.io/accenture-coursetrack/ |
| **Repositorio GitHub** | https://github.com/pacosirvent80-hash/accenture-coursetrack |
| **Google Apps Script (backend)** | `https://script.google.com/macros/s/AKfycbziQw_JwEW5MZDlOk3AaBLzTfhtfB3U4riXq_NnU1Fzy2VTJTITllKWNlAi7Jwvy2d0/exec` |

> ⚠️ La URL del Apps Script otorga acceso de lectura y escritura sin autenticación.  
> No publicarla en canales abiertos. Cada usuario la introduce manualmente en la configuración de su navegador — nunca está embebida en el código fuente..

---

## Estructura de ficheros

```
accenture-coursetrack/
├── index.html                              # Aplicación completa (HTML + CSS + JS, fichero único autocontenido)
├── README.md                               # Este fichero
├── .gitignore
└── Documentacion/
    ├── CourseTrack_Guia_Usuario.docx       # Guía de usuario completa
    ├── gen_guia_coursetrack.py             # Script Python para regenerar la guía de usuario
    ├── CourseTrack_Novedades.docx          # Novedades de la aplicación (actualizado con cada release)
    ├── gen_novedades_coursetrack.py        # Script Python para regenerar el documento de novedades
    ├── CourseTrack_EspecificacionesTecnicas.docx  # Especificaciones técnicas y buenas prácticas
    └── gen_coursetrack_tech.py             # Script Python para regenerar las especificaciones técnicas
```

La aplicación es un fichero HTML autocontenido sin dependencias externas: sin frameworks, sin librerías externas, sin servidor de aplicaciones.

---

## Tecnologías

| Capa | Tecnología |
|---|---|
| Frontend | HTML5, CSS3, JavaScript ES2020+ (Vanilla) |
| Persistencia local | `localStorage` del navegador (claves: `coursetrack_v1`, `coursetrack_user`, `coursetrack_gas_url`) |
| Persistencia compartida | Google Sheets + Google Apps Script REST |
| Hosting | GitHub Pages — rama `main`, raíz `/` |
| Responsive | CSS Grid, Flexbox, Media Queries, JS layout dinámico |

---

## Listas predefinidas

Tres arrays hardcodeados en `index.html` (~línea 792) controlan los valores disponibles en los selectores. Para modificarlos: editar el array y hacer `git push` — GitHub Pages actualiza en 1–2 minutos.

```javascript
const USERS = [
  'Paco',
  'Lourdes',
  'Natalia',
  // añadir aquí más miembros del equipo
];

const PROVIDERS = [
  'Accenture',
  'INAP',
  'Mercadona',
  'Test Provider',
  // añadir aquí más proveedores
];

const LANGUAGES = [
  'Español',
  'Inglés',
  'Catalán',
  'Euskera',
  'Gallego',
];
```

**Usuarios:** al abrir la app por primera vez en un navegador, aparece un selector que no puede cerrarse sin elegir un nombre. El nombre queda en `localStorage`. El chip `👤` de la cabecera permite cambiarlo en cualquier momento.

**Proveedores:** el filtro del toolbar siempre muestra los valores de `PROVIDERS`. Si hay cursos con proveedores no incluidos en la lista, también aparecen (compatibilidad con datos existentes).

**Idiomas:** selector en el formulario de alta de curso. El primero de la lista (`Español`) es el valor por defecto.

---

## Funcionalidades principales

| Acción | Cómo se hace |
|---|---|
| Crear curso | Botón **+ Nuevo curso** en la cabecera |
| **Editar curso** | Botón **✏ Editar curso** en el pie del panel de detalle |
| Cambiar estado | Botón **↔ Cambiar estado** en el pie del panel |
| Nueva versión mayor | Botón **+ Nueva versión** en el pie del panel |
| Registrar incidencia | Botón **⚡ Nueva incidencia** en el pie del panel |
| **Editar incidencia** | Botón **✏ Editar incidencia** en cada tarjeta de incidencia |
| Resolver incidencia | Botón **Resolver ✓** en cada tarjeta de incidencia abierta |
| Eliminar curso | Botón **🗑** en el pie del panel (pide confirmación) |

---

## Trazabilidad por usuario

Cada acción registrada en la app incluye el nombre del usuario activo:

| Acción | Campo guardado |
|---|---|
| Crear curso / cambiar estado / nueva versión | `modificadoPor` en el objeto de versión |
| Registrar incidencia | `creadaPor` en el objeto de incidencia |
| Resolver incidencia | `resueltaPor` en el objeto de incidencia |

Estos campos forman parte del JSON que se sincroniza con Google Sheets. El toolbar incluye un filtro **"Todos los usuarios"** que se autocompleta con los nombres que hayan realizado cambios. La tabla desktop incluye una columna **"Por"** ordenable.

---

## Publicar cambios

```bash
# Clonar el repositorio (una sola vez)
git clone https://github.com/pacosirvent80-hash/accenture-coursetrack.git

# Editar index.html (con Claude Code u otro editor)

# Publicar
git add index.html
git commit -m "descripción del cambio"
git push
```

GitHub Pages actualiza la URL pública en 1–2 minutos tras el push.

---

## Configuración de Google Sheets

### 1. Preparar la hoja de cálculo
- Crear una Google Sheet con dos pestañas: `Data` y `Log`.

### 2. Desplegar el Apps Script
- En la hoja: **Extensiones → Apps Script**.
- Pegar el código que aparece en el modal ⚙ Configuración de la aplicación.
- **Implementar → Nueva implementación → Aplicación web**.
  - Ejecutar como: **Yo**
  - Acceso: **Cualquiera (sin iniciar sesión)**
- Copiar la URL del Web App generada.

### 3. Conectar la app
- Abrir la aplicación web en el navegador.
- Pulsar ⚙ → pegar la URL del Web App → **Guardar y sincronizar**.

### Cómo se almacenan los datos
Los datos (cursos, versiones, incidencias y trazabilidad de usuarios) se serializan como JSON y se guardan en la celda **A1** de la pestaña `Data`. La pestaña `Log` registra cada escritura con fecha/hora y tamaño en bytes. No hay base de datos; el JSON de A1 es la fuente de verdad compartida.

---

## Seguridad y privacidad

- La URL del Apps Script **no está en el código fuente** del repositorio público.
- Cada usuario la introduce manualmente; se guarda solo en su `localStorage`.
- Los datos viajan cifrados por HTTPS entre el navegador y los servidores de Google.
- El modelo de concurrencia es **last-write-wins**: si dos usuarios guardan simultáneamente, gana el último. Para equipos pequeños (< 5 personas) esto es aceptable.
- La selección de usuario **no es autenticación**: cualquier persona con acceso a la URL puede elegir cualquier nombre. Está pensado para equipos de confianza donde la trazabilidad importa más que la seguridad.
