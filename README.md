# Servicio de procesamiento de oficios PDF

## Qué hace
- Revisa una carpeta local con PDFs.
- Todos los días a la hora definida en `config.json` (por defecto 16:00 de Chile) procesa los PDFs nuevos.
- Llama a la API de OpenAI para extraer:
  - número de oficio
  - categoría
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

## Archivos
- `oficios_service.py`: servicio principal.
- `config.example.json`: plantilla de configuración.
- `requirements.txt`: dependencias.
- `plantilla_oficios.xlsx`: plantilla inicial del Excel.

## Configuración
1. Copia `config.example.json` como `config.json`.
2. Reemplaza:
   - rutas de carpeta
   - nombres de gerentes
   - `openai_api_key`
3. Crea la carpeta de entrada para los PDFs.
4. Crea la carpeta de salida del Excel y logs.

## Instalación
```bash
pip install -r requirements.txt
```

## Crear plantilla Excel
```bash
python oficios_service.py --config config.json --create-template
```

## Prueba manual una sola vez
```bash
python oficios_service.py --config config.json --run-once
```

## Modo servicio continuo
```bash
python oficios_service.py --config config.json
```

El proceso queda corriendo y ejecuta una vez por día a la hora programada.

## Recomendación para Windows
Si no quieres dejar una terminal abierta todo el día, crea una tarea con el Programador de tareas de Windows para que el script se ejecute al iniciar sesión. El script se mantiene vivo y corre la revisión diaria a las 16:00.

## Control de duplicados
El script evita reprocesar PDFs por hash de archivo y además evita insertar duplicados en Excel usando:
- Nro
- Categoría
- Fecha de Oficio

## Limitaciones
- Si el modelo no logra identificar correctamente la fecha del oficio, no podrá calcular un plazo relativo.
- Si el modelo no logra identificar una gerencia responsable válida, la fila se agregará sin gerente responsable.
- Los días hábiles actualmente excluyen solo sábado y domingo; los feriados no se descuentan.
