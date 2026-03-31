# Gestión de Oficios CGE

## Uso normal (interfaz gráfica)

Al ejecutar el programa sin argumentos, se abre una ventana con tres botones:

| Botón | Acción |
|-------|--------|
| **▶ Ejecutar una vez** | Procesa todos los PDFs nuevos de la carpeta y llena el Excel |
| **↺ Resetear valores** | Borra la memoria de PDFs procesados para que se vuelvan a analizar |
| **✎ Revaluar oficio** | Abre el formulario de corrección de oficios ya procesados |

```
python oficios_service.py
```

O bien, ejecutar directamente el `.exe` generado por el compilador.

---

## Qué hace
- Revisa una carpeta local con PDFs que inicien con los prefijos **OC**, **Ord.** o **RE**.
- Elimina automáticamente copias de archivos (ej: `archivo (1).pdf`, `archivo - copia.pdf`).
- Llama a la API de OpenAI para extraer metadatos del documento:
  - número de oficio, categoría, fecha, concepto, gerencia responsable, plazo
- Agrega una fila al Excel con estructura predefinida.
- Crea tareas en **Microsoft Planner** para oficios con plazo vigente.
- Muestra un **popup de resumen** al finalizar, destacando multas detectadas.
- Muestra un **popup de alerta** con oficios que vencen en los próximos 5 días.
- Si detecta una **multa o formulación de cargos**, pregunta si se desea generar el **Informe de Zona por Multa SEC** en formato Word.

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

Si el usuario acepta, se llama a OpenAI con un modelo configurable (por defecto `gpt-4o-mini`) para extraer los campos del informe y se genera automáticamente un archivo Word en la carpeta configurada en `informe_multa.output_dir`.

### Template Word del informe

El archivo `informe_multa_template.docx` debe estar en el mismo directorio que el ejecutable. Si no existe, se crea automáticamente con la estructura del Informe de Zona por Multa SEC (secciones 1–7).

El template usa placeholders `{{CAMPO}}` que se reemplazan con la información extraída del PDF.

---

## Revaloración de oficios (correcciones y aprendizaje)

```bash
python oficios_service.py --revaluar
```

Abre una interfaz gráfica donde puedes corregir:
- **Área responsable** (dropdown con las 6 áreas válidas)
- **Plazo de respuesta** (formato DD-MM-YYYY)
- **¿Es multa?** (Sí / No / sin cambio)

Las correcciones se guardan en `corrections.json` y se usan como **ejemplos de aprendizaje** en el prompt de OpenAI para mejorar futuras clasificaciones.

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

Genera `dist\GestionOficios.exe`. Copia también `config.json` e `informe_multa_template.docx` en la misma carpeta que el `.exe`.

---

## Comandos de línea de comandos

| Comando | Descripción |
|---------|-------------|
| `python oficios_service.py` | Abre la interfaz gráfica (por defecto) |
| `python oficios_service.py --run-once` | Procesa una vez sin GUI |
| `python oficios_service.py --service` | Modo servicio continuo sin GUI |
| `python oficios_service.py --reset` | Resetea la memoria de PDFs procesados |
| `python oficios_service.py --revaluar` | Abre el formulario de revaloración |
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
    "api_key": "",
    "model": "gpt-4o-mini",
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

> Si `informe_multa.api_key` está vacío, se usa la misma `openai_api_key` principal.

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
