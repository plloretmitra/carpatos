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

- Implementación en Node.js usando la librería `docx` (preinstalada, no requiere npm install)
- Un único fichero `.js` que genera el `.docx` completo
- Funciones auxiliares reutilizables: cabecera de día, caja informativa, momento del día
- Verificación visual: conversión a PDF con LibreOffice y revisión de páginas con pdftoppm

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

Dos ficheros, nada más:

- **`carpatos_pwa.html`** — la aplicación entera en un solo fichero: estructura, estilos, contenido y lógica. Sin dependencias externas ni build.
- **`manifest.json`** — fichero separado con el nombre, el icono, los colores y el modo de visualización. Es lo que permite instalarla como app.

El contenido sale del mismo material que el Word: los 12 días con sus lugares y descripciones, más gastronomía e información práctica.

### URL de producción

```
https://plloretmitra.github.io/carpatos/carpatos_pwa.html
```

Servida por GitHub Pages desde la rama `master`, raíz del repositorio. La raíz del sitio devuelve 404 porque no hay `index.html`: hay que usar la URL completa. Renombrar el fichero a `index.html` la acortaría a `https://plloretmitra.github.io/carpatos/`.

### Estructura de navegación

Cuatro pestañas en una barra inferior fija:

| Pestaña | Contenido |
|---|---|
| 🏠 **Inicio** | Portada, datos del viaje, tarjeta de vuelo y accesos a las otras secciones |
| 🗺️ **Días** | Los 12 días. Cada uno se despliega y muestra la ruta, todos los lugares con descripción y enlace a Google Maps, la tarjeta de hotel y "el momento del día" |
| 🍽️ **Gastro** | Platos principales, quesos, vinos, destilados, bebidas sin alcohol y dulces |
| 🧭 **Info** | Logística, frases útiles, qué reservar con antelación y errores a evitar |

### Cómo actualizar el contenido

1. Editar `carpatos_pwa.html` (o `manifest.json`) en Claude Code
2. `git add`, `git commit`, `git push`
3. GitHub Pages despliega solo en unos 2 minutos

No hay que compilar nada. Si el móvil sigue mostrando la versión antigua, es la caché del Service Worker: cerrar la app y volver a abrirla, o subir el número de versión de la caché en el código.

### Cómo instalarla en el móvil

1. Abrir la URL de producción en Chrome
2. Menú de tres puntos → **Añadir a pantalla de inicio**
3. Se instala como app: icono propio, sin barra de navegador y con acceso offline

### Lecciones aprendidas

- **El manifest tiene que ser un fichero real servido por HTTP.** Ni los `data:` URI ni los blob URL funcionan en Chrome Android para el modo standalone. Se probaron las tres vías y solo la tercera instala la app correctamente.
- **Cuidado con el orden del JavaScript.** El Service Worker va embebido en el propio HTML como blob URL, y durante la sesión no llegaba a registrarse nunca: el bloque de manifest inline que lo precedía hacía `document.getElementById('manifest-link').href = ...` sobre un elemento que no existía en el DOM. `getElementById` devolvía `null`, el TypeError cortaba el script en seco y el registro del Service Worker, que venía justo después, no se ejecutaba. Al pasar al `manifest.json` real ese bloque desapareció y con él el problema. Moraleja: un error de JavaScript no rompe solo su línea, mata todo lo que viene detrás.
- **Con `file://` no funciona nada.** Ni manifest ni Service Worker ni instalación. La PWA hay que servirla por HTTP sí o sí, aunque sea para probarla.
- **El repositorio solo lleva fuentes.** `node_modules/`, `*.docx`, `*.pdf` y `pagina-*.jpg` están en `.gitignore`: las dependencias se reinstalan con `npm install` y el Word y sus imágenes de verificación se regeneran con el script.

### Repositorio

```
git@github.com:plloretmitra/carpatos.git    rama master
```

Ocho ficheros: la PWA y su manifest, el generador del Word, la imagen de portada, esta metodología, los dos `package*.json` y el `.gitignore`.
