# Servicio de procesamiento de oficios PDF

## Qué hace
- Revisa una carpeta local con PDFs que inicien con los prefijos **OC**, **Ord.** o **RE**.
- Elimina automáticamente copias de archivos (ej: `archivo (1).pdf`, `archivo - copia.pdf`).
- Llama a la API de OpenAI para extraer metadatos del documento:
  - número de oficio
  - categoría (Oficio circular, Oficio ordinario o Resolución exenta)
  - fecha de oficio
  - concepto
  - gerencia responsable
  - plazo respuesta
  - plazo relativo (si el documento dice, por ejemplo, "10 días hábiles")
- Agrega una fila al Excel con esta estructura:
  - Nro
  - Categoría
  - Fecha de Oficio
  - Concepto
  - Dirección Responsable
  - Gerencia Responsable
  - Gerente Responsable
  - Equipo
  - Plazo Respuesta
- Crea tareas en **Microsoft Planner** para oficios con plazo vigente (no vencido).
- Muestra un **popup de resumen** al finalizar con la cantidad de PDFs procesados, desglosados por categoría y área. Si algún oficio trata sobre **multas o formulación de cargos**, se destacan en el resumen con una alerta visual.
- Muestra un **popup de alerta** con los oficios que vencen en los próximos 5 días, indicando número, categoría, fecha, área y gerente responsable.

## Categorías y prefijos de archivo

Solo se procesan PDFs cuyo nombre comience con uno de estos prefijos:

| Prefijo | Categoría         |
|---------|--------------------|
| `OC`    | Oficio circular    |
| `Ord.`  | Oficio ordinario   |
| `RE`    | Resolución exenta  |

Los archivos que no cumplan con estos prefijos se omiten y se registra el motivo en el log.

## Detección de multas y formulación de cargos

El script analiza el concepto de cada oficio procesado buscando palabras clave como "multa", "formulación de cargos", "sanción" o "cargo sancionatorio". Si detecta alguno:
- Se destaca en el popup de resumen con el icono de advertencia.
- El título del popup indica cuántas multas se detectaron.
- Se listan los oficios afectados con su número, categoría, área y concepto.

## Cálculo de plazo relativo
Si el PDF trae una fecha exacta de vencimiento, el script usa esa fecha.

Si no trae una fecha exacta pero sí un plazo relativo, por ejemplo:
- `10 días hábiles`
- `5 días corridos`
- `30 días`

el script calcula `Plazo Respuesta` usando la **fecha del oficio** (`Fecha de Oficio`) como base.

### Regla actual
- `días hábiles`: excluye sábado y domingo.
- `días corridos`: suma días calendario.
- Si el texto dice solo `días` sin aclaración, se trata como `días corridos`.
- Aún **no** descuenta feriados de Chile.

## Revaloración (correcciones y aprendizaje)

Si la IA clasificó un oficio con el área o plazo incorrecto, puedes corregirlo con:

```bash
python oficios_service.py --config config.json --revaluar
```

Esto abre una **interfaz gráfica** donde puedes:
1. Seleccionar un oficio de la lista.
2. Cambiar el **área responsable** (dropdown con las 5 áreas válidas).
3. Cambiar el **plazo de respuesta** (formato DD-MM-YYYY).
4. Guardar la corrección.

Las correcciones se almacenan en `corrections.json` y se usan como **ejemplos de aprendizaje** en futuras ejecuciones. El prompt de OpenAI incluye las últimas 20 correcciones como referencia para que el modelo mejore su clasificación en oficios similares.

Las correcciones también actualizan el Excel inmediatamente.

## Integración con Microsoft Planner

El script puede crear tareas automáticamente en Microsoft Planner para cada oficio con plazo de respuesta vigente.

### Requisitos
1. Registrar una aplicación en **Azure AD** (portal.azure.com > App registrations).
2. Otorgar el permiso `Tasks.ReadWrite.All` de tipo Application.
3. Crear un client secret.
4. Configurar la sección `planner` en `config.json`:

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

Si `enabled` es `false`, la integración se desactiva y el script funciona sin Planner.

## Archivos
- `oficios_service.py`: servicio principal.
- `config.json`: archivo de configuración.
- `corrections.json`: correcciones del usuario (generado automáticamente por `--revaluar`).
- `requirements.txt`: dependencias (`requests`, `openpyxl`, `msal`).

## Configuración
1. Edita `config.json` con:
   - rutas de carpeta de entrada y salida
   - nombres de gerentes
   - `openai_api_key`
   - sección `planner` (opcional)
2. Crea la carpeta de entrada para los PDFs.
3. Crea la carpeta de salida del Excel y logs.

## Instalación
```bash
pip install -r requirements.txt
```

## Comandos disponibles

### Crear plantilla Excel
```bash
python oficios_service.py --config config.json --create-template
```

### Procesar una sola vez
```bash
python oficios_service.py --config config.json --run-once
```

### Modo servicio continuo
```bash
python oficios_service.py --config config.json
```
El proceso queda corriendo y ejecuta una vez por día a la hora programada.

### Resetear memoria de PDFs procesados
```bash
python oficios_service.py --config config.json --reset
```
Borra los hashes almacenados para que todos los PDFs se reprocesen en la siguiente ejecución.

### Revaloración de oficios
```bash
python oficios_service.py --config config.json --revaluar
```
Abre una interfaz gráfica para corregir el área responsable o el plazo de respuesta de oficios ya procesados. Las correcciones alimentan el aprendizaje del modelo.

## Recomendación para Windows
Si no quieres dejar una terminal abierta todo el día, crea una tarea con el Programador de tareas de Windows para que el script se ejecute al iniciar sesión. El script se mantiene vivo y corre la revisión diaria a las 16:00.

## Control de duplicados
- **Archivos copiados**: detecta y elimina copias por nombre (ej: `archivo (1).pdf`, `archivo - copia.pdf`) antes de procesar.
- **Hash de archivo**: evita reprocesar PDFs ya analizados usando SHA256.
- **Duplicados en Excel**: evita insertar filas duplicadas comparando Nro + Categoría + Fecha de Oficio.

## Archivos no accesibles
Los archivos que existen en el directorio pero no se pueden leer (por ejemplo, archivos de OneDrive solo en la nube) se omiten con un warning en el log y no interrumpen el procesamiento.

## Limitaciones
- Si el modelo no logra identificar correctamente la fecha del oficio, no podrá calcular un plazo relativo.
- Si el modelo no logra identificar una gerencia responsable válida, la fila se agregará sin gerente responsable.
- Los días hábiles actualmente excluyen solo sábado y domingo; los feriados no se descuentan.
