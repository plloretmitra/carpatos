# Metodología: Construcción de una Guía de Viaje Personalizada

Proceso seguido para crear una guía de viaje estilo Lonely Planet en formato Word, replicable para cualquier destino.

---

## 1. Definición del viaje

Antes de escribir nada, fijar los parámetros del viaje:

- Destino y duración (número de días)
- Fechas (mes y año)
- Aeropuerto de salida
- Perfil del viajero (solo, pareja, familia...)
- Estilo de viaje (coche de alquiler, transporte público, mochilero, mid-range...)
- Presupuesto orientativo
- Restricciones o preferencias especiales

---

## 2. Construcción del itinerario

- Definir las paradas principales del destino y su orden lógico geográfico
- Asignar días a cada parada (días de base vs. días de tránsito)
- Calcular distancias y tiempos de conducción entre paradas
- Identificar qué ver y hacer en cada lugar: monumentos, actividades, experiencias
- Señalar desvíos opcionales con su coste real en tiempo y kilómetros
- Gestionar expectativas sobre paradas sobrevaloradas o que no merecen el esfuerzo
- Identificar "El momento del día" al final de cada jornada (el instante más especial)

---

## 3. Contenido complementario

Secciones que enriquecen la guía más allá del día a día:

- **Gastronomía**: platos típicos, vinos, bebidas, dulces regionales
- **Frases útiles** en el idioma local (12-15 frases prácticas)
- **Logística y supervivencia**: documentación, carreteras, peajes, apps útiles, seguridad
- **Qué reservar con antelación**: meses antes / semanas antes / al llegar
- **Errores a evitar**: en carretera, en ciudades, en la planificación

---

## 4. Estructura y diseño del documento

- Formato A4, estilo editorial tipo Lonely Planet
- Portada con título, subtítulo e ilustración
- Tipografías: Calibri para textos funcionales, Georgia para contenido narrativo

---

## 5. Generación del documento

- Implementación en Node.js usando la librería `docx`. Hay que instalarla con `npm install docx` en la carpeta del proyecto antes de la primera ejecución: no viene preinstalada
- `node_modules/` va en `.gitignore` y no se sube al repositorio. Al clonar el proyecto en una máquina nueva, reinstalar las dependencias con `npm install` antes de generar nada
- Un único fichero `.js` que genera el `.docx` completo
- Funciones auxiliares reutilizables: cabecera de día, caja informativa, momento del día
- Verificación visual en dos pasos:
  1. Convertir el `.docx` a PDF con LibreOffice: `soffice --headless --convert-to pdf <fichero>.docx`
  2. Rasterizar el PDF a una imagen por página con Ghostscript: `gs -dNOPAUSE -dBATCH -sDEVICE=jpeg -r100 -sOutputFile=pagina-%02d.jpg <fichero>.pdf`
- Revisar las imágenes resultantes para detectar páginas en blanco, cajas partidas entre páginas y cabeceras huérfanas al pie

---

## 6. Ilustración de portada

- Imagen generada por IA (Gemini, Midjourney u otro) con prompt descriptivo del destino
- Insertada directamente como imagen en la portada, sin recreación ni modificación

---

## 7. Iteración y mejoras

- Revisión del contenido con otras IAs para detectar errores, omisiones o mejoras
- Filtrado crítico de sugerencias: no todas las propuestas externas son válidas
- Verificación de datos geográficos (distancias, tiempos, ubicaciones)
- Corrección de expectativas sobre atracciones sobrevendidas
- Regla de trabajo: **no actualizar el documento hasta recibir instrucción explícita**

---

## 8. Posible evolución: PWA

Una vez cerrada la guía Word, se puede convertir en una Progressive Web App (PWA):

- HTML + CSS + JS autocontenido, instalable en móvil como app sin pasar por tiendas
- Añade interactividad: enlaces GPS a Google Maps, datos de hotel y vuelo, check-lists
- No requiere compilación ni publicación en Google Play
- Los datos de hotel, vuelo y reservas se incorporan conforme se van confirmando

---

## 9. La PWA en producción

### Qué es y qué contiene

Tres ficheros:

- **`<nombre>.html`** — la aplicación entera: estructura, estilos, contenido y lógica en un solo fichero. Sin dependencias externas ni build step.
- **`manifest.json`** — nombre, icono, colores y modo de visualización. Es lo que permite instalarla como app standalone.
- **`sw.js`** — Service Worker independiente que gestiona la caché offline. Debe servirse desde el mismo origen que el HTML.

El contenido sale del mismo material que el Word: los días con sus lugares y descripciones, más las secciones complementarias (gastronomía, logística, etc.).

### URL de producción

```
https://<usuario>.github.io/<repo>/<nombre>.html
```

Servida por GitHub Pages desde la rama `master`, raíz del repositorio. Si el HTML se llama `index.html` la URL se acorta a `https://<usuario>.github.io/<repo>/`.

### Estructura de navegación típica

Una barra de navegación inferior fija con 3-5 pestañas:

- **Inicio** — portada con los datos clave del viaje y accesos a las otras secciones
- **Días** — el itinerario día a día, desplegable, con lugares, descripciones y hotel
- **Gastro** — gastronomía local organizada por categorías
- **Info** — logística, frases útiles, qué reservar con antelación, errores a evitar

Cada viaje puede añadir o quitar pestañas según el destino.

### Cómo actualizar el contenido

1. Editar el HTML (o `manifest.json`) en Claude Code
2. `git add`, `git commit`, `git push`
3. GitHub Pages despliega en ~2 minutos

No hay que compilar nada. Si el móvil sigue mostrando la versión antigua es la caché del Service Worker: cerrar y volver a abrir la app, o incrementar el número de versión de la caché en el código del SW.

### Cómo instalarla en el móvil

1. Abrir la URL de producción en Chrome (Android) o Safari (iOS)
2. Menú → **Añadir a pantalla de inicio**
3. Se instala como app: icono propio, sin barra de navegador, con acceso offline

### Lecciones aprendidas

- **El manifest tiene que ser un fichero real servido por HTTP.** Ni los `data:` URI ni los blob URL funcionan en Chrome Android para el modo standalone. Solo un `manifest.json` real, con su propio Content-Type, instala la app correctamente.
- **El Service Worker tiene que ser un fichero `sw.js` separado, servido desde el mismo origen.** Registrarlo desde un `blob:` URL (la aproximación inline, generando el código del SW como template literal dentro del HTML y pasándolo por `URL.createObjectURL`) es rechazado por los navegadores: el SW nunca llega a instalarse y la app no funciona offline. El registro correcto es:
  ```js
  navigator.serviceWorker.register('./sw.js')
  ```
- **El array de caché debe incluir explícitamente todos los recursos que la app necesita offline**: el HTML, el `manifest.json` y cualquier imagen referenciada externamente (como `portada.jpg`). No incluir `'/'` si no hay `index.html` en la raíz: `addAll` es atómico y un solo 404 cancela toda la caché.
- **Cuidado con el orden del JavaScript.** Si cualquier línea de JS lanza un error antes de llegar al `navigator.serviceWorker.register(...)`, el SW nunca se registra y la app no funciona offline. Un TypeError no rompe solo su línea: mata todo lo que viene detrás.
- **Con `file://` no funciona nada.** Ni manifest, ni Service Worker, ni instalación. La PWA necesita servirse por HTTP aunque sea para probarla.
- **El repositorio solo lleva fuentes.** `node_modules/`, los `.docx` generados, los PDF y las imágenes de verificación van en `.gitignore`. Las dependencias se reinstalan con `npm install`; el Word y sus verificaciones se regeneran con el script.

### Repositorio

Patrón recomendado:

```
git@github.com:<usuario>/<destino>.git    rama master
```

Ficheros típicos: el HTML de la PWA, `manifest.json`, `sw.js`, el generador del Word, la imagen de portada, el fichero de metodología, `package.json`, `package-lock.json` y `.gitignore`.
