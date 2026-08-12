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
