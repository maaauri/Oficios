# Handoff: App Oficios CGE — Rediseño V5 (Minimal)

## Overview
Rediseño de la aplicación interna de Gestión de Oficios de CGE (Comercial · Servicio al Cliente). La aplicación lee PDFs de oficios recibidos desde la SEC, los clasifica automáticamente por área responsable mediante un agente, los registra en Excel, genera informes de multa cuando corresponde, y muestra estadísticas. Este handoff cubre el rediseño "V5 Minimal", que reemplaza la UI actual (Tkinter con ventanas separadas y botones saturados) por una sola ventana moderna con navegación por tabs.

## About the Design Files
Los archivos dentro de `prototipo/` son **referencias de diseño en HTML+React** — prototipos que muestran el look y comportamiento esperados. **No son código de producción para copiar directamente**. La tarea es **recrear estos diseños dentro del entorno real de la app**: actualmente la app está escrita en Python (Tkinter), por lo que la implementación objetivo es probablemente **CustomTkinter** (que fue la preferencia indicada) o, si se decide migrar, PyWebView / Flet / Tauri. Usa los patrones y librerías establecidos del proyecto.

## Fidelity
**Alta fidelidad (hifi)**: colores, tipografía, espaciado y estados están definidos con valores finales. Debe recrearse lo más fielmente posible, adaptando a los idioms del framework destino.

## Target stack (recomendado)
- **CustomTkinter** (`pip install customtkinter`) sobre Python 3.10+
- `tkinter.ttk.Treeview` para tablas densas
- `Pillow` para iconos / preview PDF
- `matplotlib` embebido o canvas nativo para gráficos
- Mantener la lógica existente del agente, lectura de PDFs y generación de Excel

## Pantallas

### 1. Header global (persistente en todas las pantallas)
- Altura ~62px, fondo `panel` (`#ffffff` claro / `#17212f` oscuro), borde inferior 1px `border`.
- Izquierda: cuadrado 32x32 `accent` con texto "CGE" blanco, peso 700, radio 7.
- A la derecha del logo: "Gestión de Oficios" (15px, peso 600) sobre "Comercial · Servicio al Cliente" (11px, `subtext`).
- Centro-derecha: segmented control (tabs `Bandeja` / `Estadísticas`) sobre fondo `soft`, padding 3, radio 8. Tab activo fondo `panel` con sombra suave.
- Derecha: botón primario "▶ Ejecutar análisis" (fondo `accent`, texto blanco).

### 2. Footer global
- Altura ~34px, fondo `panel`, borde superior. Punto verde `success` (7x7), texto 11.5px: `Completado: 6 PDFs procesados · hace 12 min`. Versión a la derecha (10.5px `dim`).

### 3. Pantalla: Bandeja (Inicio)
- Padding 20px/28px (normal) o 16px/24px (compact).
- **KPI row**: grid 4 columnas, gap 12. Cada KPI: card (`panel`, radio 10, borde `border`, padding 14x16) con label 11.5px `subtext`, valor 28px peso 600 con `font-variant-numeric: tabular-nums`, subtítulo 11px. Ícono en esquina superior derecha 18px `dim` opacity 0.5.
  - `128 Oficios totales · +6 hoy`
  - `100% Accuracy agente` (color `success`)
  - `10 Multas detectadas` (color `warn`)
  - `3 Plazos críticos · menos de 5 días` (color `danger`)
- **Alert bar**: fondo `warnSoft`, borde `warn` con alpha 0.2, radio 10, padding 12x16. Círculo `warn` 28x28 con "!" blanco. Título "3 oficios con plazo en menos de 5 días" + detalle en `subtext`. Botón "Revisar" secundario a la derecha.
- **Tabs + búsqueda**: línea de tabs (Hoy / Por vencer (warn) / Multas (danger) / Histórico) con underline 2px `accent` en el activo. Badge circular de conteo por tab. Buscador 240px con ícono ⌕ a la izquierda.
- **Cards grid**: 2 columnas, gap 12. Cada card:
  - Número oficio (JetBrains Mono, peso 600) · tipo en `subtext`
  - Chip "MULTA" (fondo `warnSoft`, texto `warn`, uppercase 10px) si aplica
  - Chip de días restantes con color según umbral: ≤3d `danger`, ≤5d `warn`, resto `success`
  - Asunto 13px, line-height 1.4, min-height 36px
  - Footer: punto de color por área, nombre del área, fecha plazo, botones ghost "Revaluar" y "Informe" (si multa)

### 4. Pantalla: Estadísticas
- Header de pantalla: título 18px + subtítulo + selector rango + botón "Exportar".
- **KPI row** (4 columnas) igual a Inicio, con métricas distintas (totales, accuracy, categorías, monto UTM).
- **Grid 1.3fr / 1fr**:
  - Panel "Volumen mensual": barras 140px altura, 1 barra por mes. Barra completa `blue`; porción inferior representa multas en `warn`. Etiqueta numérica arriba, mes abajo. Leyenda con 2 swatches.
  - Panel "Por tipo de oficio": filas con nombre, conteo y %, barra de progreso 4px.
- **Grid 1fr / 1.3fr**:
  - Panel "Distribución por área": SVG donut 150x150 (radio 55, stroke 22, segmentos con stroke-dasharray) + leyenda con swatches cuadrados 9x9 radio 2, nombre, conteo, %.
  - Panel "Desempeño del agente": 3 mini-stats (tiempo promedio, clasificación correcta, correcciones) + lista de actividad reciente con punto de color + mensaje + tiempo relativo.

### 5. Pantalla: Revaluar
- Breadcrumb arriba: `← Bandeja` en `blue` / `Revaluar oficio` en `text`.
- **Grid 1fr / 1.4fr**:
  - **Panel izquierdo "Oficios registrados"**: input filtro + lista scroll con altura máx 460px. Cada item: nro mono (60px) + dot de área + nombre área en `subtext` + días restantes. Item activo: fondo `blueSoft`, borde izq 3px `accent`.
  - **Panel derecho**: 
    - Card de detalle: asunto 14px, grid 3 columnas con "área asignada / plazo / multa" (cada celda en fondo `softer`, label uppercase 10.5px, valor 13.5px peso 600). Banda inferior `softer` con confianza del agente y keywords.
    - Card "Corregir asignación": grid 2 cols con `<select>` de área y `<input>` de plazo. Botones segmentados para "¿Es multa?" (sin cambio / Sí / No). Footer con "Cancelar", indicador "● Cambios sin guardar" en `warn`, "✓ Guardar corrección" primario (disabled si no hay cambios).

### 6. Pantalla: Informe de Multa
- Breadcrumb.
- **Banner degradé**: fondo `linear-gradient(135deg, warn15, danger15)`, borde `warn40`, radio 10, padding 16x20. Icono ◆ 42x42 fondo `warn`. Título "Oficio {nro} · {asunto}". A la derecha: "Monto estimado · UTM 250" (22px, `warn`).
- **Grid 1.5fr / 1fr**:
  - **Panel "Informe generado"** con secciones numeradas (1. Identificación con grid key/value · 2. Hechos imputados · 3. Normativa infringida (lista) · 4. Propuesta de descargos). Títulos de sección uppercase 11px, separación 18px.
  - **Columna derecha** (3 panels apilados):
    - "Detalles": filas key/value con severidad (punto + texto), monto base, agravantes, atenuantes, total.
    - "Deadline": número gigante 44px (color según días) + "días para responder a SEC" + fecha mono.
    - Panel acciones: "Exportar informe (Word)" primario, "Exportar PDF" ghost, "Enviar al responsable" ghost, "No es multa · revaluar" ghost `danger`.

## Interactions & Behavior

- **Navegación**: `go(screen, payload)` cambia `screen` state entre `inicio | stats | revaluar | multa`. `Revaluar` e `Informe de multa` reciben el oficio como payload. El breadcrumb vuelve a `inicio`.
- **Tabs en Bandeja**: filtran `SAMPLE_OFICIOS` por `diasRest <= 5` (por vencer) o `multa === true` (multas).
- **Búsqueda**: filtra por match case-insensitive en `nro + asunto + area`.
- **Revaluar**: los tres inputs (`newArea`, `newPlazo`, `multaCorrection`) habilitan el botón "Guardar corrección" solo cuando hay al menos un cambio.
- **Ejecutar análisis**: dispara el agente sobre los PDFs nuevos (conectar con la lógica existente).
- **Hover en cards**: transición `border-color 0.15s`.

## State Management
- `screen: 'inicio' | 'stats' | 'revaluar' | 'multa'`
- `selectedOficio`, `multaFor`: objeto oficio seleccionado
- `tab` en Bandeja: `'hoy' | 'vencer' | 'multas' | 'todos'`
- `q`: búsqueda
- En Revaluar: `sel`, `newArea`, `newPlazo`, `multaCorrection`, `filter`
- Tweaks (persistentes): `dark: bool`, `density: 'compact' | 'normal'`, `accent: 'cge-blue' | 'deep-navy' | 'electric' | 'graphite'`

## Design Tokens

### Paleta claro
```
bg           #f7f8fa
panel        #ffffff
soft         #eff2f6
softer       #f4f6f9
border       #e2e6ec
borderStrong #d4dae3
text         #1c2633
subtext      #6b7684
dim          #9aa3b0
accent       #0B3D6B   (primario CGE blue)
blue         #1E6FB8
blueSoft     #e8f0fa
success      #1a7f5a   successSoft #e4f3ec
warn         #c47a00   warnSoft    #fdf2e0
danger       #b42d2d   dangerSoft  #fbe7e7
lilac        #8a4fb5   lilacSoft   #f2ebf8
teal         #0e8a82   tealSoft    #dff2f0
neutral      #556270   neutralSoft #eaedf1
```

### Paleta oscura
```
bg           #0f1722
panel        #17212f
soft         #1c2838
softer       #1a2433
border       #253244
text         #e4eaf2
subtext      #9aa8bc
dim          #6b7a8f
accent       #3d8bd9
blue         #5aa9e6   blueSoft #1c3550
success      #4db87d   warn     #e0a54a   danger  #e06464
lilac        #a88be8   teal     #4ec9b0   neutral #8798b0
```

### Acentos alternativos (tweakable)
```
cge-blue    accent #0B3D6B  blue #1E6FB8
deep-navy   accent #1a2847  blue #2d5a9b
electric    accent #0057b7  blue #2a89e6
graphite    accent #2a3441  blue #4a5a70
```

### Mapeo área → color
```
Conexiones          → blue
PMGD                → warn
Servicio al Cliente → success
Pérdidas            → lilac
Sin área            → neutral
Cobranza            → danger
Lectura             → teal
```

### Tipografía
- UI: **Inter** (400/500/600/700). Fallback: system-ui.
- Números / códigos oficio: **JetBrains Mono** (400/500/600). Usar `font-variant-numeric: tabular-nums` en todos los valores numéricos.
- Tamaños usados: 10, 10.5, 11, 11.5, 12, 12.5, 13, 13.5, 14, 15, 17, 18, 22, 28, 44.
- Pesos: 400 normal, 500 medio, 600 semibold, 700 bold (solo logos).
- Letter-spacing: títulos grandes `-0.3` a `-0.5`; uppercase labels `+0.4` a `+0.8`.

### Espaciado
- Padding cards: 14x16 (normal) / 12x14 (compact)
- Padding pantalla: 20x28 (normal) / 16x24 (compact)
- Gaps grid: 12 default, 14 entre paneles, 10 dentro de paneles
- Border radius: chips 3-4, inputs/buttons 6, cards/panels 10, logo 7

### Sombras
- Cards (hover activo): `0 1px 2px rgba(0,0,0,0.03)` a `0 1px 3px rgba(0,0,0,0.08)`
- Ventana (chrome): `0 1px 3px rgba(0,0,0,0.08), 0 20px 60px rgba(0,0,0,0.12)`
- Segmented control tab activa: `0 1px 2px rgba(0,0,0,0.06)`

## Assets
- No hay imágenes/íconos externos; todos los glifos son Unicode (`▶ ⟳ ▣ ◆ ⇩ ✓ ⚑ ⏱ ▦ ◈ ✦ ⌕ ←`) o SVG inline (donut chart).
- Logo: cuadrado de color con texto "CGE" (sustituir por logo oficial si corresponde).
- PDFs: se leen desde `C:\CGE\Oficios\Entrada\` (ruta de la app actual).

## Datos de ejemplo (mock)
Ver `SAMPLE_OFICIOS` en `prototipo/v5-deep/v5-app.jsx` — 8 oficios con todos los campos del modelo:
```
nro, tipo, area, plazo, diasRest, asunto, multa, conf, fecha, multaMonto?
```

## Implementación en CustomTkinter — pistas
- Usar `ctk.CTkTabview` para los tabs del header.
- `CTkFrame` con `fg_color=panel`, `border_width=1`, `border_color=border`, `corner_radius=10` = las "cards".
- Para los KPIs: `CTkLabel` grande (tamaño 28 peso bold) + dos labels auxiliares.
- Tablas de oficios: `ttk.Treeview` con estilo custom (quita bordes, usa `rowheight=32`, alterna fondo con `tag_configure`).
- Donut y barras: `tkinter.Canvas` o `matplotlib` embebido con `FigureCanvasTkAgg`. Usar exactamente los colores del token.
- Preview PDF (si se agrega luego): `tkPDFViewer` o convertir 1a página con `pdf2image` + mostrar en `CTkImage`.
- Tweaks → guardar en `~/.oficios_cge/config.json`; releer al inicio para aplicar tema/densidad/acento.

## Files
- `prototipo/App Oficios CGE - V5 Deep.html` — shell con tweaks y chrome
- `prototipo/v5-deep/v5-app.jsx` — componente React con toda la UI (Bandeja, Stats, Revaluar, Informe Multa) + paletas + tokens
