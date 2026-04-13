# Gestión de Oficios CGE

## Uso normal (interfaz gráfica)

Al ejecutar el programa sin argumentos, se abre una ventana con cuatro botones:

| Botón | Acción |
|-------|--------|
| **▶ Ejecutar una vez** | Procesa todos los PDFs nuevos de la carpeta y llena el Excel |
| **↺ Resetear valores** | Borra la memoria de PDFs procesados para que se vuelvan a analizar |
| **✎ Revaluar oficio** | Abre el formulario de corrección de oficios ya procesados |
| **📊 Estadísticas** | Muestra estadísticas agregadas de todos los oficios en el Excel |

```
python oficios_service.py
```

O bien, ejecutar directamente el `.exe` generado por el compilador.

---

## Qué hace
- Revisa una carpeta local con PDFs que inicien con los prefijos **OC**, **Ord.** o **RE**.
- Elimina automáticamente copias de archivos (ej: `archivo (1).pdf`, `archivo - copia.pdf`).
- Llama a la API de OpenAI para extraer metadatos del documento:
  - número de oficio, categoría, fecha, concepto, gerencia responsable, plazo, oficio relacionado
- Si el PDF menciona un oficio relacionado, busca ese PDF en el mismo directorio y **analiza ambos documentos juntos** para determinar el área.
- Agrega una fila al Excel con estructura predefinida.
- Crea tareas en **Microsoft Planner** para oficios con plazo vigente.
- Muestra un **popup de resumen** al finalizar, destacando multas detectadas.
- Muestra un **popup de alerta** con oficios que vencen en los próximos 5 días.
- Si detecta una **multa o formulación de cargos**, pregunta si se desea generar el **Informe de Zona por Multa SEC** en formato Word (generado con la API de Anthropic/Claude).

---

## Categorías y prefijos de archivo

| Prefijo | Categoría         |
|---------|--------------------|
| `OC`    | Oficio circular    |
| `Ord.`  | Oficio ordinario   |
| `RE`    | Resolución exenta  |

Los archivos sin estos prefijos se omiten y se registra el motivo en el log.

---

## Detección de multas

El sistema detecta automáticamente multas y formulación de cargos por palabras clave en el concepto (`multa`, `formulación de cargos`, `sanción`, `cargo sancionatorio`).

Cuando se detecta una multa, el popup de resumen lo indica con alerta visual y, al finalizar el procesamiento, el sistema **pregunta si se desea generar el Informe de Zona por Multa SEC**.

Si el usuario acepta, se llama a la **API de Anthropic (Claude)** con el modelo configurado (por defecto `claude-sonnet-4-20250514`) para extraer los campos del informe y se genera automáticamente un archivo Word en la carpeta configurada en `informe_multa.output_dir`.

### Template Word del informe

El archivo `informe_multa_template.docx` debe estar en el mismo directorio que el ejecutable. Si no existe, se crea automáticamente con la estructura del Informe de Zona por Multa SEC (secciones 1–7).

El template usa placeholders `{{CAMPO}}` que se reemplazan con la información extraída del PDF.

---

## Oficios vinculados

Cuando un PDF menciona explícitamente otro oficio previo (frases como *"en respuesta a"*, *"en relación al Oficio N°"*, *"complementa el Ord."*), el sistema:

1. Extrae el número del documento relacionado.
2. Busca ese PDF en el directorio configurado.
3. Si lo encuentra, **re-analiza ambos PDFs juntos** con OpenAI para determinar con mayor precisión el área responsable y demás metadatos.
4. Si no lo encuentra, continúa con la información del documento principal y registra el aviso en el log.

---

## Sistema de aprendizaje

El agente acumula experiencia de cuatro fuentes combinadas que se inyectan en el prompt en cada ejecución (fusión de `oficios_service` + `clasificador_oficios_v2`):

### 1. Historial del Excel
Antes de cada extracción, lee el Excel y selecciona hasta 5 ejemplos recientes por área. Se excluyen los oficios con concepto **"solicita más información"**.

### 2. Historial del agente (`historial_oficios.json`)
Cada oficio procesado se guarda con su `area_propuesta`, `area_final`, `fue_corregido`, `keywords`, `remitente` y `confianza`. Cuando el usuario corrige un área en Revaluar, la entrada se marca como `fue_corregido=True`.

El prompt incluye hasta 10 ejemplos **few-shot dinámicos** priorizando correcciones sobre confirmaciones (ratio 60/40), ya que las correcciones contienen la señal más valiosa.

### 3. Correcciones manuales (`corrections.json`)
Registra cada ajuste en el formulario de Revaluar con mayor prioridad que el historial general.

### 4. Reglas auto-generadas (`reglas_clasificacion.json`)
Cada 20 oficios procesados, el sistema llama a **Claude** para analizar el historial completo y destilar reglas de clasificación accionables por área (incluyendo reglas negativas *"NO clasificar como X si..."*). Estas reglas se inyectan con la prioridad más alta. Se pueden regenerar manualmente con:

```bash
python oficios_service.py --regenerar-reglas
```

### Campos extraídos por el agente
Además de los campos base (número, categoría, fecha, concepto, gerencia, plazo), cada oficio incluye:
- **`remitente`** — quien firma/envía el oficio
- **`keywords`** — 3-6 términos clave que justifican la clasificación
- **`confianza`** — score 0.0-1.0 de qué tan seguro está el modelo del área

---

## Revaloración de oficios

```bash
python oficios_service.py --revaluar
```

Abre una interfaz gráfica donde puedes seleccionar cualquier oficio ya procesado y corregir:
- **Área responsable** (dropdown con las 6 áreas válidas)
- **Plazo de respuesta** (formato DD-MM-YYYY)
- **¿Es multa?** (Sí / No / sin cambio)

Los cambios se aplican tanto al Excel como a `corrections.json`.

---

## Estadísticas

El botón **📊 Estadísticas** en la interfaz gráfica muestra un resumen del Excel con:
- Total de oficios registrados
- **Accuracy del agente** (porcentaje de decisiones confirmadas sin corrección)
- Desglose por categoría (con porcentaje)
- Desglose por área responsable (con porcentaje) + gráfico de torta
- Cantidad de multas / formulaciones de cargos
- Rango de fechas de los oficios
- **Accuracy por área** con barras visuales
- **Errores más frecuentes** (qué área se confunde con cuál)
- Estado de las reglas aprendidas (fecha de generación y base)

También disponible por línea de comandos:

```bash
python oficios_service.py --stats
```

---

## Cálculo de plazo relativo

| Tipo | Regla |
|------|-------|
| `días hábiles` | Excluye sábado y domingo |
| `días corridos` | Suma días calendario |
| `días` (sin aclaración) | Se trata como `días corridos` |

> Los feriados de Chile no se descuentan actualmente.

---

## Integración con Microsoft Planner

Se crean tareas automáticamente en Planner para oficios con plazo vigente.

```json
"planner": {
    "enabled": true,
    "tenant_id": "tu-tenant-id",
    "client_id": "tu-client-id",
    "client_secret": "tu-client-secret",
    "plan_id": "id-del-plan",
    "bucket_id": "id-del-bucket"
}
```

Requiere una app registrada en **Azure AD** con permiso `Tasks.ReadWrite.All`.

---

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `oficios_service.py` | Servicio principal |
| `config.json` | Configuración |
| `corrections.json` | Correcciones del usuario (generado automáticamente) |
| `historial_oficios.json` | Historial de decisiones del agente con flag `fue_corregido` (generado) |
| `reglas_clasificacion.json` | Reglas aprendidas por Claude del historial (generado cada 20 oficios) |
| `informe_multa_template.docx` | Template Word del informe de multa |
| `requirements.txt` | Dependencias Python |
| `build.bat` | Script para compilar el ejecutable en Windows |

---

## Instalación

```bash
pip install -r requirements.txt
```

---

## Compilar el ejecutable (.exe)

```bat
build.bat
```

Genera la carpeta `dist\GestionOficios\` con el ejecutable y todas las dependencias. El inicio es inmediato (no requiere descomprimir en cada ejecución). Copia también `config.json` e `informe_multa_template.docx` dentro de esa carpeta.

> Para distribución, comparte toda la carpeta `GestionOficios\`. Puedes crear un acceso directo a `GestionOficios.exe` para mayor comodidad.

---

## Comandos de línea de comandos

| Comando | Descripción |
|---------|-------------|
| `python oficios_service.py` | Abre la interfaz gráfica (por defecto) |
| `python oficios_service.py --run-once` | Procesa una vez sin GUI |
| `python oficios_service.py --service` | Modo servicio continuo sin GUI |
| `python oficios_service.py --reset` | Resetea la memoria de PDFs procesados |
| `python oficios_service.py --revaluar` | Abre el formulario de revaloración |
| `python oficios_service.py --stats` | Imprime métricas de accuracy del agente |
| `python oficios_service.py --regenerar-reglas` | Regenera reglas de clasificación con Claude |
| `python oficios_service.py --create-template` | Crea la plantilla Excel |

---

## Configuración (`config.json`)

```json
{
  "watch_dir": "ruta/carpeta/pdfs",
  "excel_path": "ruta/oficios.xlsx",
  "openai_api_key": "sk-...",
  "model": "gpt-4o-mini",
  "informe_multa": {
    "api_key": "sk-ant-...",
    "model": "claude-sonnet-4-20250514",
    "output_dir": "ruta/informes"
  },
  "gerentes": {
    "PMGD": {"nombre": "Nombre", "email": ""},
    "Conexiones": {"nombre": "Nombre", "email": ""},
    "Lectura": {"nombre": "Nombre", "email": ""},
    "Servicio al Cliente": {"nombre": "Nombre", "email": ""},
    "Cobranza": {"nombre": "Nombre", "email": ""},
    "Pérdidas": {"nombre": "Nombre", "email": ""}
  }
}
```

> `openai_api_key` se usa para la extracción de metadatos de todos los oficios.
> `informe_multa.api_key` debe ser una API key de **Anthropic** (`sk-ant-...`) para la generación del informe de multa.

---

## Control de duplicados
- **Archivos copiados**: detecta y elimina copias por nombre antes de procesar.
- **Hash SHA256**: evita reprocesar archivos ya analizados.
- **Duplicados en Excel**: compara Nro + Categoría + Fecha de Oficio.

## Archivos no accesibles
Archivos de OneDrive "solo en la nube" se omiten con warning en el log sin interrumpir el procesamiento.

## Limitaciones
- Los días hábiles excluyen solo sábado y domingo; los feriados no se descuentan.
- Si el modelo no identifica la fecha del oficio, no puede calcular plazos relativos.
- La búsqueda de oficio relacionado requiere que el PDF correspondiente esté en el mismo directorio de vigilancia.
