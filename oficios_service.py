from __future__ import annotations

import argparse
import base64
import hashlib
import json
import logging
import os
import re
import sys
import threading
import time
import tkinter as tk
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

try:
    from docx import Document as DocxDocument
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

import msal
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

EXPECTED_COLUMNS = [
    "Nro",
    "Categoría",
    "Fecha de Oficio",
    "Concepto",
    "Dirección Responsable",
    "Gerencia Responsable",
    "Gerente Responsable",
    "Equipo",
    "Plazo Respuesta",
    "Multa",
]

AREAS_VALIDAS = [
    "PMGD", "Conexiones", "Lectura",
    "Servicio al Cliente", "Cobranza", "Pérdidas",
]

SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "numero_oficio": {"type": ["string", "null"]},
        "categoria": {
            "type": ["string", "null"],
            "enum": ["Resolución exenta", "Oficio ordinario", "Oficio circular", None],
        },
        "fecha_oficio": {"type": ["string", "null"]},
        "concepto": {"type": ["string", "null"]},
        "gerencia_responsable": {
            "type": ["string", "null"],
            "enum": ["PMGD", "Conexiones", "Lectura", "Servicio al Cliente", "Cobranza", "Pérdidas", None],
        },
        "plazo_respuesta": {"type": ["string", "null"]},
        "plazo_relativo_cantidad": {"type": ["integer", "null"], "minimum": 1},
        "plazo_relativo_tipo": {
            "type": ["string", "null"],
            "enum": ["dias_corridos", "dias_habiles", None],
        },
        "oficio_relacionado": {"type": ["string", "null"]},
    },
    "required": [
        "numero_oficio",
        "categoria",
        "fecha_oficio",
        "concepto",
        "gerencia_responsable",
        "plazo_respuesta",
        "plazo_relativo_cantidad",
        "plazo_relativo_tipo",
        "oficio_relacionado",
    ],
}

PROMPT = """Eres un extractor de metadatos regulatorios para oficios y resoluciones recibidos por CGE.

Tu tarea es leer el PDF adjunto y devolver SOLO un JSON válido que cumpla exactamente el schema entregado.

Reglas de extracción y normalización:

1. Usa exclusivamente información presente en el PDF.
2. numero_oficio:
   - devuelve solo el número principal del oficio o resolución
   - sin \"N°\", sin \"Nº\", sin \"No.\", sin prefijos
   - ejemplo: \"Resolución Exenta Electrónica N° 38222\" => \"38222\"
3. categoria:
   normaliza a uno de estos tres valores exactos según el prefijo del nombre del archivo:
   - \"RE\" => \"Resolución exenta\"
   - \"Ord.\" => \"Oficio ordinario\"
   - \"OC\" => \"Oficio circular\"
   Usa siempre el prefijo del nombre del archivo para determinar la categoría.
4. fecha_oficio:
   - es la fecha de emisión/envío del oficio o resolución
   - formato exacto YYYY-MM-DD
   - si no está clara, devuelve null
5. concepto:
   - resume el asunto principal del documento
   - una sola línea
   - máximo 250 caracteres
   - sin comillas, sin saltos de línea
   - debe ser ejecutivo y fiel al contenido
6. gerencia_responsable:
   normaliza a uno de estos valores exactos:
   - \"PMGD\"
   - \"Conexiones\"
   - \"Lectura\"
   - \"Servicio al Cliente\"
   - \"Cobranza\"
   - \"Pérdidas\"
   Criterios:
   - \"PMGD\": generación distribuida, conexión de PMGD, plataformas o procesos PMGD
   - \"Conexiones\": conexión, empalme, factibilidad, puesta en servicio, plazos de conexión, obras o procesos de conexión
   - \"Lectura\": lectura de medidores, medición, consumos leídos, toma de lectura
   - \"Servicio al Cliente\": atención, reclamos, canales, calidad de servicio, respuesta al cliente
   - \"Cobranza\": deuda, mora, pago, repactación, suspensión/corte por deuda, cobranza
   - \"Pérdidas\": pérdidas de energía, hurto de energía, consumo no registrado, fraude eléctrico, intervención de medidor, irregularidades en consumo
   Si hay más de un tema, elige el principal.
7. plazo_respuesta:
   - devuelve una fecha explícita de vencimiento en formato YYYY-MM-DD solo si en el PDF existe una fecha calendario inequívoca
   - si no existe una fecha exacta, devuelve null
8. plazo_relativo_cantidad y plazo_relativo_tipo:
   - si el documento establece un plazo relativo, extráelo
   - ejemplos:
     - \"10 días hábiles\" => plazo_relativo_cantidad=10, plazo_relativo_tipo=\"dias_habiles\"
     - \"5 días corridos\" => plazo_relativo_cantidad=5, plazo_relativo_tipo=\"dias_corridos\"
     - \"30 días\" sin aclaración expresa => trátalo como \"dias_corridos\"
   - si existe una fecha exacta y además un plazo relativo, devuelve ambos si ambos aparecen explícitamente
   - si no existe plazo relativo, devuelve null en ambos campos
9. oficio_relacionado:
   - si el documento hace referencia explícita a otro oficio, resolución u orden de compra anterior
     (expresiones como "en respuesta a", "en relación al Oficio N°", "complementa el Ord.", "adjunto al OC"),
     devuelve el número de ese documento (solo cifras o código, sin prefijos)
   - si no hay referencia a otro documento, devuelve null
10. Si un dato no está claro o no existe, devuelve null.
11. No agregues ningún texto fuera del JSON.
12. No inventes datos.
"""

INFORME_MULTA_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "sociedad":                 {"type": ["string", "null"]},
        "regional":                 {"type": ["string", "null"]},
        "zonal":                    {"type": ["string", "null"]},
        "evento":                   {"type": ["string", "null"]},
        "sector_activo":            {"type": ["string", "null"]},
        "clientes_afectados":       {"type": ["string", "null"]},
        "fecha_evento":             {"type": ["string", "null"]},
        "fecha_notificacion_multa": {"type": ["string", "null"]},
        "monto_utm":                {"type": ["string", "null"]},
        "numero_resolucion":        {"type": ["string", "null"]},
        "etapa_multa":              {"type": ["string", "null"]},
        "descripcion_motivo_sec":   {"type": ["string", "null"]},
        "descripcion_tecnica_zona": {"type": ["string", "null"]},
        "causa_raiz":               {"type": ["string", "null"]},
        "cronologia":               {"type": ["string", "null"]},
        "que_paso":                 {"type": ["string", "null"]},
        "que_se_hizo":              {"type": ["string", "null"]},
        "plan_mejora":              {"type": ["string", "null"]},
        "area_responsable_multa":   {"type": ["string", "null"]},
        "area_responsable_mejora":  {"type": ["string", "null"]},
    },
    "required": [
        "sociedad", "regional", "zonal", "evento", "sector_activo",
        "clientes_afectados", "fecha_evento", "fecha_notificacion_multa",
        "monto_utm", "numero_resolucion", "etapa_multa",
        "descripcion_motivo_sec", "descripcion_tecnica_zona", "causa_raiz",
        "cronologia", "que_paso", "que_se_hizo", "plan_mejora",
        "area_responsable_multa", "area_responsable_mejora",
    ],
}

INFORME_MULTA_PROMPT = """Eres un analista regulatorio de CGE. A partir del PDF adjunto (una multa o resolución sancionatoria de la SEC), extrae SOLO un JSON válido con los campos del informe de zona.

Instrucciones:
1. sociedad: nombre de la empresa sancionada (ej: "CGE Distribución S.A.").
2. regional: región geográfica involucrada (ej: "Metropolitana", "Tarapacá").
3. zonal: zona operativa específica si se menciona.
4. evento: descripción breve del tipo de evento sancionado.
5. sector_activo: sector eléctrico o activo físico involucrado.
6. clientes_afectados: número o descripción de clientes afectados, o null.
7. fecha_evento: fecha del evento sancionado (formato DD-MM-YYYY o texto si no es exacta).
8. fecha_notificacion_multa: fecha en que CGE fue notificada de la multa.
9. monto_utm: monto de la multa en UTM como texto (ej: "100 UTM").
10. numero_resolucion: número de la resolución exenta SEC.
11. etapa_multa: etapa actual (ej: "Notificación", "Descargos", "Resolución firme").
12. descripcion_motivo_sec: descripción objetiva según la SEC del motivo de la multa. Sé detallado.
13. descripcion_tecnica_zona: descripción técnica del evento desde la perspectiva operativa.
14. causa_raiz: posible causa raíz interna identificable en el documento.
15. cronologia: secuencia cronológica de los hechos relevantes.
16. que_paso: resumen de qué ocurrió y cuál es la problemática sancionada.
17. que_se_hizo: qué acciones se tomaron o dejaron de tomar.
18. plan_mejora: si existe algún plan de mejora mencionado, descríbelo. Si no, null.
19. area_responsable_multa: área interna de CGE responsable de la infracción.
20. area_responsable_mejora: área responsable de implementar correcciones.

Si un dato no está en el PDF, devuelve null. No inventes datos. Solo JSON, sin texto adicional.
"""


def get_base_dir() -> Path:
    """Retorna el directorio base del ejecutable o del script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


@dataclass
class Gerente:
    nombre: str
    email: str = ""


@dataclass
class PlannerConfig:
    enabled: bool = False
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    plan_id: str = ""
    bucket_id: str = ""


@dataclass
class OutlookConfig:
    enabled: bool = False
    user_email: str = ""
    folder_name: str = "Oficios"


@dataclass
class Config:
    watch_dir: Path
    excel_path: Path
    processed_state_path: Path
    corrections_path: Path
    log_path: Path
    timezone: str
    run_time: str
    openai_api_key: str
    model: str
    gerentes: Dict[str, Gerente]
    planner: PlannerConfig = field(default_factory=PlannerConfig)
    informe_multa_api_key: str = ""
    informe_multa_model: str = "claude-sonnet-4-20250514"
    informe_output_dir: Path = field(default_factory=lambda: Path("."))
    request_timeout_seconds: int = 180
    scan_extensions: tuple[str, ...] = (".pdf",)


def load_config(path: Path) -> Config:
    with path.open("r", encoding="utf-8") as f:
        try:
            raw = json.load(f)
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"El archivo de configuración tiene un error de formato JSON:\n"
                f"  Archivo: {path}\n"
                f"  Detalle: {exc}\n\n"
                f"Revise que no haya comas sobrantes, comillas sin cerrar o caracteres inválidos."
            ) from exc

    gerentes: Dict[str, Gerente] = {}
    for area, item in raw.get("gerentes", {}).items():
        gerentes[area] = Gerente(nombre=item.get("nombre", ""), email=item.get("email", ""))

    api_key = raw.get("openai_api_key") or os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        raise ValueError("Falta openai_api_key en config.json o variable de entorno OPENAI_API_KEY.")

    planner_raw = raw.get("planner", {})
    planner = PlannerConfig(
        enabled=planner_raw.get("enabled", False),
        tenant_id=planner_raw.get("tenant_id", ""),
        client_id=planner_raw.get("client_id", ""),
        client_secret=planner_raw.get("client_secret", ""),
        plan_id=planner_raw.get("plan_id", ""),
        bucket_id=planner_raw.get("bucket_id", ""),
    )

    informe_raw = raw.get("informe_multa", {})
    excel_parent = str(Path(raw["excel_path"]).parent)
    informe_output_dir = Path(informe_raw.get("output_dir", excel_parent))

    return Config(
        watch_dir=Path(raw["watch_dir"]),
        excel_path=Path(raw["excel_path"]),
        processed_state_path=Path(raw.get("processed_state_path", "processed_state.json")),
        corrections_path=Path(raw.get("corrections_path", "corrections.json")),
        log_path=Path(raw.get("log_path", "oficios_service.log")),
        timezone=raw.get("timezone", "America/Santiago"),
        run_time=raw.get("run_time", "16:00"),
        openai_api_key=api_key,
        model=raw.get("model", "gpt-5.4-mini"),
        gerentes=gerentes,
        planner=planner,
        informe_multa_api_key=informe_raw.get("api_key", ""),
        informe_multa_model=informe_raw.get("model", "claude-sonnet-4-20250514"),
        informe_output_dir=informe_output_dir,
        request_timeout_seconds=int(raw.get("request_timeout_seconds", 180)),
    )


def setup_logging(log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def create_excel_template(path: Path, sheet_name: str = "Oficios") -> None:
    ensure_parent(path)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.freeze_panes = "A2"

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_gray = Side(style="thin", color="D9E2F3")

    for col_idx, name in enumerate(EXPECTED_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin_gray)

    widths = {
        "A": 12,
        "B": 22,
        "C": 16,
        "D": 60,
        "E": 28,
        "F": 24,
        "G": 26,
        "H": 14,
        "I": 18,
        "J": 10,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for col in ["C", "I"]:
        for row in range(2, 5000):
            ws[f"{col}{row}"].number_format = "DD-MM-YYYY"

    wb.save(path)


def ensure_excel_exists(path: Path) -> None:
    if not path.exists():
        create_excel_template(path)
        logging.info("Plantilla Excel creada en %s", path)
        return

    wb = load_workbook(path)
    ws = wb.active
    existing = [ws.cell(row=1, column=i).value for i in range(1, len(EXPECTED_COLUMNS) + 1)]
    if existing == EXPECTED_COLUMNS:
        return

    # Migrar: si solo falta la columna "Multa" al final, agregarla automáticamente
    old_columns = EXPECTED_COLUMNS[:-1]
    existing_old = [ws.cell(row=1, column=i).value for i in range(1, len(old_columns) + 1)]
    if existing_old == old_columns:
        col_idx = len(EXPECTED_COLUMNS)
        header_cell = ws.cell(row=1, column=col_idx, value="Multa")
        header_cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        header_cell.font = Font(color="FFFFFF", bold=True)
        header_cell.alignment = Alignment(horizontal="center", vertical="center")
        header_cell.border = Border(bottom=Side(style="thin", color="D9E2F3"))
        ws.column_dimensions["J"].width = 10
        # Rellenar multa para filas existentes según concepto
        for row_idx in range(2, ws.max_row + 1):
            concepto = str(ws.cell(row=row_idx, column=4).value or "")
            if concepto and es_multa_o_cargos(concepto):
                cell = ws.cell(row=row_idx, column=col_idx, value="Sí")
                cell.alignment = Alignment(horizontal="center", vertical="center")
        wb.save(path)
        logging.info("Columna 'Multa' agregada al Excel existente.")
        return

    raise ValueError(
        f"El Excel existente no tiene los encabezados esperados. Esperado: {EXPECTED_COLUMNS} | Actual: {existing}"
    )


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"processed_hashes": [], "last_run_date": None}
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_state(path: Path, state: dict[str, Any]) -> None:
    ensure_parent(path)
    with path.open("w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_date_yyyy_mm_dd(value: Optional[str]) -> Optional[date]:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def add_business_days(start_date: date, business_days: int) -> date:
    current = start_date
    added = 0
    while added < business_days:
        current += timedelta(days=1)
        if current.weekday() < 5:  # 0-4 = lunes-viernes
            added += 1
    return current


def compute_due_date(extracted: dict[str, Any]) -> Optional[date]:
    explicit_due = parse_date_yyyy_mm_dd(extracted.get("plazo_respuesta"))
    if explicit_due:
        return explicit_due

    fecha_oficio = parse_date_yyyy_mm_dd(extracted.get("fecha_oficio"))
    if not fecha_oficio:
        return None

    cantidad = extracted.get("plazo_relativo_cantidad")
    tipo = extracted.get("plazo_relativo_tipo")

    if not isinstance(cantidad, int) or cantidad <= 0 or tipo not in {"dias_corridos", "dias_habiles"}:
        return None

    if tipo == "dias_habiles":
        return add_business_days(fecha_oficio, cantidad)

    return fecha_oficio + timedelta(days=cantidad)


def extract_output_text(api_response: dict[str, Any]) -> str:
    if isinstance(api_response.get("output_text"), str):
        return api_response["output_text"]

    pieces: List[str] = []
    for item in api_response.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"} and isinstance(content.get("text"), str):
                pieces.append(content["text"])
    if pieces:
        return "\n".join(pieces).strip()

    raise ValueError("No se pudo extraer output_text de la respuesta de OpenAI.")


def call_openai_extract(
    config: Config,
    pdf_path: Path,
    corrections_prompt: str = "",
    related_pdf_path: Optional[Path] = None,
) -> dict[str, Any]:
    pdf_b64 = base64.b64encode(pdf_path.read_bytes()).decode("ascii")
    full_prompt = PROMPT + corrections_prompt

    if related_pdf_path is not None:
        intro_text = (
            f"Se adjuntan DOS documentos vinculados entre sí. "
            f"Documento principal: {pdf_path.name}. "
            f"Documento relacionado: {related_pdf_path.name}. "
            f"Estudia ambos para determinar el área responsable y los demás metadatos. "
            f"El JSON debe corresponder al documento principal."
        )
        related_b64 = base64.b64encode(related_pdf_path.read_bytes()).decode("ascii")
        user_content = [
            {"type": "input_text", "text": intro_text},
            {"type": "input_file", "filename": pdf_path.name,
             "file_data": f"data:application/pdf;base64,{pdf_b64}"},
            {"type": "input_file", "filename": related_pdf_path.name,
             "file_data": f"data:application/pdf;base64,{related_b64}"},
        ]
    else:
        user_content = [
            {"type": "input_text",
             "text": f"Analiza el siguiente PDF. Nombre del archivo: {pdf_path.name}"},
            {"type": "input_file", "filename": pdf_path.name,
             "file_data": f"data:application/pdf;base64,{pdf_b64}"},
        ]

    payload = {
        "model": config.model,
        "max_output_tokens": 450,
        "text": {
            "format": {
                "type": "json_schema",
                "name": "oficio_sec_extract",
                "strict": True,
                "schema": SCHEMA,
            }
        },
        "input": [
            {
                "role": "developer",
                "content": [{"type": "input_text", "text": full_prompt}],
            },
            {
                "role": "user",
                "content": user_content,
            },
        ],
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {config.openai_api_key}",
    }

    response = requests.post(
    "https://api.openai.com/v1/responses",
    headers=headers,
    json=payload,
    timeout=config.request_timeout_seconds,
    )

    if not response.ok:
        logging.error("OpenAI devolvió %s: %s", response.status_code, response.text)
        response.raise_for_status()

    data = response.json()

    output_text = extract_output_text(data)
    try:
        parsed = json.loads(output_text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"La respuesta de OpenAI no es JSON válido: {output_text}") from exc

    usage = data.get("usage", {})
    if usage:
        logging.info(
            "Uso OpenAI para %s | input_tokens=%s | output_tokens=%s | total_tokens=%s",
            pdf_path.name,
            usage.get("input_tokens"),
            usage.get("output_tokens"),
            usage.get("total_tokens"),
        )

    return parsed


def map_row(extracted: dict[str, Any], gerentes: Dict[str, Gerente]) -> List[Any]:
    gerencia = extracted.get("gerencia_responsable")
    gerente = gerentes.get(gerencia, Gerente(nombre=""))
    fecha_oficio = parse_date_yyyy_mm_dd(extracted.get("fecha_oficio"))
    plazo = compute_due_date(extracted)

    concepto = extracted.get("concepto") or ""
    multa = "Sí" if es_multa_o_cargos(concepto) else ""

    return [
        extracted.get("numero_oficio") or "",
        extracted.get("categoria") or "",
        fecha_oficio,
        concepto,
        "Comercial y Servicio al Cliente",
        gerencia or "",
        gerente.nombre,
        "",
        plazo,
        multa,
    ]


def row_exists(ws, nro: str, categoria: str, fecha_oficio: Optional[date]) -> bool:
    target_date_str = fecha_oficio.isoformat() if fecha_oficio else ""
    for row in ws.iter_rows(min_row=2, values_only=True):
        existing_nro = str(row[0] or "").strip()
        existing_cat = str(row[1] or "").strip()
        existing_fecha = row[2]
        if hasattr(existing_fecha, "date"):
            existing_fecha = existing_fecha.date()
        existing_fecha_str = existing_fecha.isoformat() if existing_fecha else ""
        if existing_nro == nro and existing_cat == categoria and existing_fecha_str == target_date_str:
            return True
    return False


def first_empty_row(ws) -> int:
    for row_idx in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, len(EXPECTED_COLUMNS) + 1)]
        if all(v is None or str(v).strip() == "" for v in row_values):
            return row_idx
    return ws.max_row + 1


def append_to_excel(excel_path: Path, row: List[Any]) -> bool:
    wb = load_workbook(excel_path)
    ws = wb.active

    nro = str(row[0] or "").strip()
    categoria = str(row[1] or "").strip()
    fecha_oficio = row[2] if isinstance(row[2], date) else None

    if row_exists(ws, nro, categoria, fecha_oficio):
        logging.info("Se omite inserción duplicada en Excel: nro=%s | categoría=%s | fecha=%s", nro, categoria, fecha_oficio)
        return False

    next_row = first_empty_row(ws)

    for col_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=next_row, column=col_idx, value=value)
        if col_idx in (3, 9):
            cell.number_format = "DD-MM-YYYY"
            cell.alignment = Alignment(horizontal="center")
        elif col_idx == 10:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in (1, 2, 5, 6, 7, 8):
            cell.alignment = Alignment(vertical="center")
        else:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(excel_path)
    return True


_COPY_PATTERN = re.compile(
    r"^(?P<base>.+?)"
    r"(?:"
    r"\s*-\s*(?:cop(?:y|ia))(?:\s*\(\d+\))?"  # " - Copy", " - copia", " - Copy (2)"
    r"|\s+\(\d+\)"                              # " (1)", " (2)"
    r")"
    r"(?P<ext>\.[^.]+)$",
    re.IGNORECASE,
)


def remove_duplicate_files(watch_dir: Path, extensions: tuple[str, ...]) -> int:
    """Detecta archivos que son copias (por nombre) y los elimina si el original existe."""
    removed = 0
    for path in sorted(watch_dir.iterdir()):
        if not path.is_file() or path.suffix.lower() not in extensions:
            continue
        m = _COPY_PATTERN.match(path.name)
        if not m:
            continue
        original = watch_dir / (m.group("base") + m.group("ext"))
        if original.exists() and original != path:
            logging.info("Archivo duplicado detectado: %s (original: %s). Eliminando copia.", path.name, original.name)
            path.unlink()
            removed += 1
    if removed:
        logging.info("Se eliminaron %d copia(s) de archivos.", removed)
    return removed


_VALID_PREFIXES = re.compile(r"^(OC|Ord\.|RE)\s", re.IGNORECASE)


def find_related_pdf(watch_dir: Path, related_nro: str) -> Optional[Path]:
    """Busca en watch_dir un PDF cuyo nombre contenga el número de oficio relacionado."""
    if not related_nro or not watch_dir.exists():
        return None
    # Normalizar: quitar espacios y guiones para comparación flexible
    nro_clean = re.sub(r"[\s\-/]", "", related_nro).lower()
    for path in watch_dir.iterdir():
        if not path.name.lower().endswith(".pdf"):
            continue
        name_clean = re.sub(r"[\s\-/]", "", path.stem).lower()
        if nro_clean in name_clean:
            return path
    return None


def find_pending_pdfs(config: Config, processed_hashes: set[str]) -> List[Path]:
    if not config.watch_dir.exists():
        raise FileNotFoundError(f"No existe el directorio a revisar: {config.watch_dir}")

    pdfs: List[Path] = []
    for path in sorted(config.watch_dir.iterdir()):
        if path.is_dir():
            logging.info("Omitido (es un directorio): %s", path.name)
            continue
        if not path.name.lower().endswith(".pdf"):
            logging.info("Omitido (no es un archivo PDF): %s", path.name)
            continue
        if not _VALID_PREFIXES.match(path.name):
            logging.info("Omitido (no inicia con OC, Ord. o RE): %s", path.name)
            continue
        try:
            file_hash = sha256_file(path)
        except OSError as exc:
            logging.warning("Omitido (no se pudo leer, puede estar solo en la nube): %s — %s", path.name, exc)
            continue
        if file_hash in processed_hashes:
            logging.info("Omitido (ya fue procesado anteriormente): %s", path.name)
            continue
        pdfs.append(path)
    return pdfs


_MULTA_KEYWORDS = re.compile(
    r"multa|formulaci[oó]n de cargos|cargo[s]? sancionatorio|sanci[oó]n",
    re.IGNORECASE,
)


def es_multa_o_cargos(concepto: str) -> bool:
    return bool(_MULTA_KEYWORDS.search(concepto))


@dataclass
class ProcessedOficio:
    nro: str
    categoria: str
    concepto: str
    gerencia: str
    es_multa: bool


@dataclass
class ProcessingStats:
    total: int = 0
    errores: int = 0
    categorias: Counter = field(default_factory=Counter)
    areas: Counter = field(default_factory=Counter)
    oficios_multa: List[ProcessedOficio] = field(default_factory=list)

    def registrar(self, extracted: dict[str, Any]) -> None:
        self.total += 1
        cat = extracted.get("categoria") or "Sin categoría"
        area = extracted.get("gerencia_responsable") or "Sin área"
        concepto = extracted.get("concepto") or ""
        self.categorias[cat] += 1
        self.areas[area] += 1
        if es_multa_o_cargos(concepto):
            self.oficios_multa.append(ProcessedOficio(
                nro=extracted.get("numero_oficio") or "S/N",
                categoria=cat,
                concepto=concepto,
                gerencia=area,
                es_multa=True,
            ))

    def resumen(self) -> str:
        lines = [f"PDFs nuevos procesados: {self.total}"]
        if self.errores:
            lines.append(f"Errores: {self.errores}")
        lines.append("")
        lines.append("Por categoría:")
        for cat, n in sorted(self.categorias.items()):
            lines.append(f"  {cat}: {n}")
        lines.append("")
        lines.append("Por área:")
        for area, n in sorted(self.areas.items()):
            lines.append(f"  {area}: {n}")
        if self.oficios_multa:
            lines.append("")
            lines.append(f"MULTAS / FORMULACIÓN DE CARGOS: {len(self.oficios_multa)}")
            for o in self.oficios_multa:
                lines.append(f"  !! Nro {o.nro} ({o.categoria}) — {o.gerencia}")
                lines.append(f"     {o.concepto[:120]}")
        return "\n".join(lines)


def show_summary_popup(stats: ProcessingStats) -> None:
    root = tk.Tk()
    root.withdraw()
    icon = "warning" if stats.oficios_multa else "info"
    title = "Resumen de procesamiento"
    if stats.oficios_multa:
        title += f"  —  {len(stats.oficios_multa)} MULTA(S) DETECTADA(S)"
    if icon == "warning":
        messagebox.showwarning(title, stats.resumen())
    else:
        messagebox.showinfo(title, stats.resumen())
    root.destroy()


# ---------------------------------------------------------------------------
# Informe de Multa — OpenAI + Word
# ---------------------------------------------------------------------------

def is_multa(nro: str, concepto: str, corrections: List[Dict[str, Any]]) -> bool:
    """Retorna True si el oficio es una multa según keywords o corrección manual."""
    for c in reversed(corrections):
        if str(c.get("nro", "")) == nro and c.get("campo") == "es_multa":
            return bool(c.get("valor_nuevo", False))
    return bool(_MULTA_KEYWORDS.search(concepto))


def call_anthropic_informe_multa(config: Config, pdf_path: Path) -> Dict[str, Any]:
    """Llama a la API de Anthropic (Claude) para extraer datos del informe de multa."""
    pdf_b64 = base64.b64encode(pdf_path.read_bytes()).decode("ascii")

    schema_instruction = (
        "Responde SOLO con un JSON válido que cumpla exactamente este schema:\n"
        + json.dumps(INFORME_MULTA_SCHEMA, ensure_ascii=False, indent=2)
    )

    payload = {
        "model": config.informe_multa_model,
        "max_tokens": 4096,
        "system": INFORME_MULTA_PROMPT + "\n\n" + schema_instruction,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": f"Analiza este PDF de multa SEC. Archivo: {pdf_path.name}",
                    },
                ],
            }
        ],
    }
    headers = {
        "Content-Type": "application/json",
        "x-api-key": config.informe_multa_api_key,
        "anthropic-version": "2023-06-01",
    }
    resp = requests.post("https://api.anthropic.com/v1/messages", headers=headers,
                         json=payload, timeout=config.request_timeout_seconds)
    if not resp.ok:
        logging.error("Anthropic informe multa error %s: %s", resp.status_code, resp.text)
        resp.raise_for_status()

    data = resp.json()
    text_content = ""
    for block in data.get("content", []):
        if block.get("type") == "text":
            text_content += block.get("text", "")

    text_content = text_content.strip()
    if text_content.startswith("```"):
        text_content = re.sub(r"^```(?:json)?\s*", "", text_content)
        text_content = re.sub(r"\s*```$", "", text_content)

    return json.loads(text_content)


def _replace_in_paragraph(para: Any, replacements: Dict[str, str]) -> None:
    """Reemplaza placeholders {{KEY}} en un párrafo de python-docx."""
    full = "".join(r.text for r in para.runs)
    if "{{" not in full:
        return
    for key, val in replacements.items():
        full = full.replace(f"{{{{{key}}}}}", val or "")
    for i, run in enumerate(para.runs):
        run.text = full if i == 0 else ""


def create_informe_template(output_path: Path) -> None:
    """Crea el archivo Word template del informe de multa con placeholders."""
    if not DOCX_AVAILABLE:
        logging.warning("python-docx no instalado; no se puede crear el template.")
        return
    doc = DocxDocument()
    # Encabezado
    h = doc.add_heading("INFORME DE ZONA POR MULTA SEC", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Sección 1
    doc.add_heading("1. Identificación General del Evento", level=2)
    tbl = doc.add_table(rows=11, cols=2)
    tbl.style = "Table Grid"
    fields1 = [
        ("Sociedad", "{{SOCIEDAD}}"), ("Regional", "{{REGIONAL}}"),
        ("Zonal", "{{ZONAL}}"), ("Evento", "{{EVENTO}}"),
        ("Sector / Activo involucrado", "{{SECTOR_ACTIVO}}"),
        ("Clientes afectados (si aplica)", "{{CLIENTES_AFECTADOS}}"),
        ("Fecha del Evento", "{{FECHA_EVENTO}}"),
        ("Fecha de notificación de la multa SEC", "{{FECHA_NOTIFICACION}}"),
        ("Monto de la multa (UTM)", "{{MONTO_UTM}}"),
        ("N° Resolución Exenta SEC", "{{NUMERO_RESOLUCION}}"),
        ("Etapa actual de la multa", "{{ETAPA_MULTA}}"),
    ]
    for i, (label, ph) in enumerate(fields1):
        tbl.cell(i, 0).text = label
        tbl.cell(i, 1).text = ph

    # Sección 2
    doc.add_heading("2. Causa de Infracción Normativa", level=2)
    doc.add_paragraph("Descripción objetiva del motivo de la multa (según SEC):")
    doc.add_paragraph("{{DESCRIPCION_MOTIVO_SEC}}")
    doc.add_paragraph("Descripción Técnica del Evento (según Zona):")
    doc.add_paragraph("{{DESCRIPCION_TECNICA_ZONA}}")
    doc.add_paragraph("Declaración de la causa raíz interna (si se identifica):")
    doc.add_paragraph("{{CAUSA_RAIZ}}")

    # Sección 3
    doc.add_heading("3. Cronología de la sanción", level=2)
    doc.add_paragraph("{{CRONOLOGIA}}")

    # Sección 4
    doc.add_heading("4. Antecedentes Técnicos del Caso", level=2)
    doc.add_paragraph("¿Qué pasó? ¿Cuál es la problemática que se sanciona?")
    doc.add_paragraph("{{QUE_PASO}}")
    doc.add_paragraph("¿Qué se hizo y qué se dejó de hacer para que la multa fuera cursada?")
    doc.add_paragraph("{{QUE_SE_HIZO}}")
    doc.add_paragraph("Plan de mejora (si existió) y avance al momento de la multa:")
    doc.add_paragraph("{{PLAN_MEJORA}}")

    # Sección 5
    doc.add_heading("5. Área Responsable de la aplicación de la multa", level=2)
    doc.add_paragraph("{{AREA_RESPONSABLE_MULTA}}")

    # Sección 6
    doc.add_heading("6. Área Responsable de implementar el plan de mejora", level=2)
    doc.add_paragraph("{{AREA_RESPONSABLE_MEJORA}}")

    # Sección 7
    doc.add_heading("7. Firmas", level=2)
    tbl2 = doc.add_table(rows=2, cols=2)
    tbl2.cell(0, 0).text = "______________________________"
    tbl2.cell(0, 1).text = "______________________________"
    tbl2.cell(1, 0).text = "(nombre)\nDirector Regional"
    tbl2.cell(1, 1).text = "(nombre)\nDirector (Of. Central)"

    doc.add_paragraph(f"\nGenerado: {{{{FECHA_GENERACION}}}}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    logging.info("Template de informe de multa creado en %s", output_path)


def fill_informe_multa(config: Config, informe_data: Dict[str, Any],
                       nro: str, output_path: Path) -> bool:
    """Rellena el template Word con los datos extraídos y lo guarda."""
    if not DOCX_AVAILABLE:
        logging.error("python-docx no instalado; no se puede generar el informe.")
        return False

    template_path = get_base_dir() / "informe_multa_template.docx"
    if not template_path.exists():
        create_informe_template(template_path)

    def v(key: str) -> str:
        return str(informe_data.get(key) or "")

    replacements = {
        "SOCIEDAD": v("sociedad"),
        "REGIONAL": v("regional"),
        "ZONAL": v("zonal"),
        "EVENTO": v("evento"),
        "SECTOR_ACTIVO": v("sector_activo"),
        "CLIENTES_AFECTADOS": v("clientes_afectados"),
        "FECHA_EVENTO": v("fecha_evento"),
        "FECHA_NOTIFICACION": v("fecha_notificacion_multa"),
        "MONTO_UTM": v("monto_utm"),
        "NUMERO_RESOLUCION": v("numero_resolucion"),
        "ETAPA_MULTA": v("etapa_multa"),
        "DESCRIPCION_MOTIVO_SEC": v("descripcion_motivo_sec"),
        "DESCRIPCION_TECNICA_ZONA": v("descripcion_tecnica_zona"),
        "CAUSA_RAIZ": v("causa_raiz"),
        "CRONOLOGIA": v("cronologia"),
        "QUE_PASO": v("que_paso"),
        "QUE_SE_HIZO": v("que_se_hizo"),
        "PLAN_MEJORA": v("plan_mejora"),
        "AREA_RESPONSABLE_MULTA": v("area_responsable_multa"),
        "AREA_RESPONSABLE_MEJORA": v("area_responsable_mejora"),
        "FECHA_GENERACION": date.today().strftime("%d-%m-%Y"),
    }

    doc = DocxDocument(str(template_path))
    for para in doc.paragraphs:
        _replace_in_paragraph(para, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_in_paragraph(para, replacements)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    logging.info("Informe de multa generado: %s", output_path)
    return True


def ask_and_generate_informe(config: Config, pdf_path: Path,
                              extracted: Dict[str, Any]) -> None:
    """Pregunta al usuario si desea generar el informe de multa y lo genera."""
    nro = extracted.get("numero_oficio") or "S/N"
    answer = messagebox.askyesno(
        "Informe de Multa",
        f"El oficio Nro {nro} parece ser una multa o formulación de cargos.\n\n"
        "¿Desea generar automáticamente el Informe de Zona por Multa SEC?"
    )
    if not answer:
        return
    if not DOCX_AVAILABLE:
        messagebox.showerror(
            "Dependencia faltante",
            "Instale python-docx para generar informes Word:\n  pip install python-docx"
        )
        return
    output_path = config.informe_output_dir / f"Informe_Multa_Nro{nro}.docx"
    try:
        informe_data = call_anthropic_informe_multa(config, pdf_path)
        ok = fill_informe_multa(config, informe_data, nro, output_path)
        if ok:
            messagebox.showinfo(
                "Informe generado",
                f"Informe de multa guardado en:\n{output_path}"
            )
    except Exception as exc:
        logging.exception("Error generando informe de multa: %s", exc)
        messagebox.showerror("Error", f"No se pudo generar el informe:\n{exc}")


# ---------------------------------------------------------------------------
# Microsoft Planner integration
# ---------------------------------------------------------------------------

def get_planner_token(planner: PlannerConfig) -> Optional[str]:
    authority = f"https://login.microsoftonline.com/{planner.tenant_id}"
    app = msal.ConfidentialClientApplication(
        planner.client_id,
        authority=authority,
        client_credential=planner.client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    logging.error("No se pudo obtener token de Planner: %s", result.get("error_description", result))
    return None


def create_planner_task(token: str, planner: PlannerConfig, title: str, due_date: date) -> bool:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    payload = {
        "planId": planner.plan_id,
        "bucketId": planner.bucket_id,
        "title": title,
        "dueDateTime": f"{due_date.isoformat()}T23:59:59Z",
    }
    resp = requests.post(
        "https://graph.microsoft.com/v1.0/planner/tasks",
        headers=headers,
        json=payload,
        timeout=30,
    )
    if resp.ok:
        logging.info("Tarea creada en Planner: %s", title)
        return True
    logging.error("Error creando tarea en Planner: %s — %s", resp.status_code, resp.text)
    return False


def sync_to_planner(config: Config, extracted: dict[str, Any], due_date: date) -> None:
    if not config.planner.enabled:
        return
    if due_date <= date.today():
        logging.info("Plazo ya vencido, no se crea tarea en Planner: %s", extracted.get("numero_oficio"))
        return
    nro = extracted.get("numero_oficio") or "S/N"
    cat = extracted.get("categoria") or ""
    concepto = extracted.get("concepto") or ""
    gerencia = extracted.get("gerencia_responsable") or ""
    title = f"[{cat}] Nro {nro} — {gerencia} — {concepto[:80]}"
    token = get_planner_token(config.planner)
    if token:
        create_planner_task(token, config.planner, title, due_date)


# ---------------------------------------------------------------------------
# Outlook attachment downloader
# ---------------------------------------------------------------------------


def find_outlook_folder_id(token: str, user_email: str, folder_name: str) -> Optional[str]:
    """Busca el ID de una carpeta de correo por nombre."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders"
    params = {"$filter": f"displayName eq '{folder_name}'", "$top": "50"}
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    if not resp.ok:
        logging.error("Error listando carpetas de correo: %s — %s", resp.status_code, resp.text)
        return None
    folders = resp.json().get("value", [])
    for f in folders:
        if f.get("displayName") == folder_name:
            return f["id"]
    logging.warning("No se encontró la carpeta '%s' en el buzón de %s", folder_name, user_email)
    return None


def download_outlook_attachments(config: Config, state: dict[str, Any]) -> int:
    """Descarga adjuntos PDF de la carpeta Outlook configurada. Retorna cantidad de archivos nuevos."""
    if not config.outlook.enabled:
        return 0

    if not config.outlook.user_email:
        logging.warning("Outlook habilitado pero falta user_email en config.")
        return 0

    token = get_planner_token(config.planner)
    if not token:
        logging.error("No se pudo obtener token para Outlook.")
        return 0

    folder_id = find_outlook_folder_id(token, config.outlook.user_email, config.outlook.folder_name)
    if not folder_id:
        return 0

    processed_ids: set[str] = set(state.get("outlook_processed_ids", []))
    headers = {"Authorization": f"Bearer {token}"}
    user = config.outlook.user_email
    saved = 0

    page_size = 50
    skip = 0
    while True:
        url = f"https://graph.microsoft.com/v1.0/users/{user}/mailFolders/{folder_id}/messages"
        params = {
            "$top": str(page_size),
            "$skip": str(skip),
            "$select": "id,subject,hasAttachments,receivedDateTime",
            "$orderby": "receivedDateTime desc",
        }
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        if not resp.ok:
            logging.error("Error listando mensajes: %s — %s", resp.status_code, resp.text)
            break

        messages = resp.json().get("value", [])
        if not messages:
            break

        all_already_processed = True
        for msg in messages:
            msg_id = msg["id"]
            if msg_id in processed_ids:
                continue
            all_already_processed = False

            if not msg.get("hasAttachments"):
                processed_ids.add(msg_id)
                continue

            att_url = f"https://graph.microsoft.com/v1.0/users/{user}/messages/{msg_id}/attachments"
            att_resp = requests.get(att_url, headers=headers, timeout=30)
            if not att_resp.ok:
                logging.error("Error obteniendo adjuntos del mensaje %s: %s", msg_id, att_resp.status_code)
                continue

            for att in att_resp.json().get("value", []):
                if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
                    continue
                name = att.get("name", "")
                if not name.lower().endswith(".pdf"):
                    continue
                content_bytes = att.get("contentBytes")
                if not content_bytes:
                    continue

                dest = config.watch_dir / name
                if dest.exists():
                    logging.info("Adjunto ya existe en disco, omitido: %s", name)
                else:
                    dest.write_bytes(base64.b64decode(content_bytes))
                    logging.info("Adjunto descargado: %s (mensaje: %s)", name, msg.get("subject", ""))
                    saved += 1

            processed_ids.add(msg_id)

        if all_already_processed:
            break
        skip += page_size

    state["outlook_processed_ids"] = sorted(processed_ids)
    save_state(config.processed_state_path, state)

    if saved:
        logging.info("Se descargaron %d adjunto(s) nuevos desde Outlook.", saved)
    else:
        logging.info("No hay adjuntos nuevos en Outlook.")
    return saved


# ---------------------------------------------------------------------------
# Correcciones / aprendizaje
# ---------------------------------------------------------------------------

def load_corrections(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def build_corrections_prompt(corrections: List[Dict[str, Any]]) -> str:
    if not corrections:
        return ""
    lines = [
        "\n\nCorrecciones previas del usuario (usa estos ejemplos para aprender y mejorar):"
    ]
    for c in corrections[-20:]:
        lines.append(
            f"  - Oficio Nro {c['nro']}: "
            f"campo '{c['campo']}' corregido de '{c['valor_anterior']}' a '{c['valor_nuevo']}'"
        )
        if c.get("concepto"):
            lines.append(f"    Concepto del oficio: {c['concepto'][:150]}")
    lines.append(
        "\nConsidera estas correcciones como referencia para clasificar oficios similares."
    )
    return "\n".join(lines)


def build_history_prompt(excel_path: Path, max_per_area: int = 5) -> str:
    """Lee el Excel y construye un bloque de ejemplos históricos por área.

    Selecciona hasta *max_per_area* oficios recientes de cada área para que
    el modelo aprenda los patrones de clasificación reales.
    """
    if not excel_path.exists():
        return ""

    try:
        wb = load_workbook(excel_path, read_only=True)
        ws = wb.active
    except Exception:
        return ""

    _skip = re.compile(r"solicita\s+m[aá]s\s+informaci[oó]n", re.IGNORECASE)

    # area -> lista de (nro, concepto_corto)
    by_area: Dict[str, List[Tuple[str, str]]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        nro = str(row[0] or "").strip()
        concepto = str(row[3] or "").strip()
        area = str(row[5] or "").strip()
        if not nro or not area or not concepto:
            continue
        if _skip.search(concepto):
            continue  # excluir oficios de solicitud de más información
        by_area.setdefault(area, []).append((nro, concepto[:120]))
    wb.close()

    if not by_area:
        return ""

    lines = [
        "\n\nEjemplos históricos de oficios clasificados correctamente por área "
        "(usa estos ejemplos para aprender los patrones de clasificación):"
    ]
    for area in sorted(by_area):
        examples = by_area[area][-max_per_area:]  # los más recientes
        lines.append(f"\n  Área «{area}»:")
        for nro, concepto in examples:
            lines.append(f"    - Nro {nro}: {concepto}")

    lines.append(
        "\nClasifica el nuevo oficio en el área que mejor se ajuste "
        "a los patrones observados arriba. Las correcciones del usuario "
        "(si las hay) tienen prioridad sobre los ejemplos históricos."
    )
    return "\n".join(lines)


def save_corrections(path: Path, corrections: List[Dict[str, Any]]) -> None:
    ensure_parent(path)
    with path.open("w", encoding="utf-8") as f:
        json.dump(corrections, f, ensure_ascii=False, indent=2)


def update_excel_row(excel_path: Path, nro: str, categoria: str,
                     updates: Dict[str, Any], gerentes: Dict[str, Gerente]) -> None:
    """Actualiza una fila existente en el Excel según Nro + Categoría."""
    wb = load_workbook(excel_path)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        cell_nro = str(row[0].value or "").strip()
        cell_cat = str(row[1].value or "").strip()
        if cell_nro == nro and cell_cat == categoria:
            if "gerencia_responsable" in updates:
                new_area = updates["gerencia_responsable"]
                row[5].value = new_area  # col F = Gerencia Responsable
                gerente = gerentes.get(new_area, Gerente(nombre=""))
                row[6].value = gerente.nombre  # col G = Gerente Responsable
            if "plazo_respuesta" in updates:
                p = parse_date_yyyy_mm_dd(updates["plazo_respuesta"])
                if p:
                    row[8].value = p  # col I = Plazo Respuesta
                    row[8].number_format = "DD-MM-YYYY"
            if "es_multa" in updates:
                row[9].value = "Sí" if updates["es_multa"] else ""  # col J = Multa
            wb.save(excel_path)
            return
    wb.close()


# ---------------------------------------------------------------------------
# Design system (paleta y helpers de UI)
# ---------------------------------------------------------------------------

UI = {
    # Colores principales
    "primary":       "#1F4E78",
    "primary_dark":  "#163A5A",
    "primary_light": "#2E6BA1",
    "accent":        "#C55A11",
    "success":       "#2E7D4F",
    "warning":       "#B7791F",
    "danger":        "#922B21",
    "info":          "#5B2C6F",
    # Fondos
    "bg":            "#F4F6FA",
    "surface":       "#FFFFFF",
    "surface_alt":   "#EEF2F7",
    # Texto y bordes
    "text":          "#2C3E50",
    "text_muted":    "#6B7A8A",
    "border":        "#D6DEE7",
    # Tipografía
    "font_base":     ("Segoe UI", 10),
    "font_bold":     ("Segoe UI", 10, "bold"),
    "font_title":    ("Segoe UI", 16, "bold"),
    "font_heading":  ("Segoe UI", 12, "bold"),
    "font_small":    ("Segoe UI", 9),
    "font_mono":     ("Consolas", 9),
}


def _hover(widget: tk.Widget, base_bg: str, hover_bg: str) -> None:
    """Agrega efecto hover a un widget (botón)."""
    widget.bind("<Enter>", lambda _e: widget.config(bg=hover_bg))
    widget.bind("<Leave>", lambda _e: widget.config(bg=base_bg))


def ui_header(parent: tk.Widget, title: str, subtitle: Optional[str] = None,
              bg: Optional[str] = None, height: int = 70) -> tk.Frame:
    """Cabecera coloreada con título y subtítulo opcional."""
    bg_color = bg or UI["primary"]
    frame = tk.Frame(parent, bg=bg_color, height=height)
    frame.pack(fill=tk.X)
    frame.pack_propagate(False)
    inner = tk.Frame(frame, bg=bg_color)
    inner.pack(expand=True)
    tk.Label(inner, text=title, bg=bg_color, fg="white",
             font=UI["font_title"]).pack()
    if subtitle:
        tk.Label(inner, text=subtitle, bg=bg_color, fg="#CFE0EF",
                 font=UI["font_small"]).pack()
    return frame


def ui_button(parent: tk.Widget, text: str, command: Any,
              variant: str = "primary", width: int = 26, height: int = 2,
              icon: str = "") -> tk.Button:
    """Crea un botón con el estilo y variantes del design system."""
    palette = {
        "primary": (UI["primary"], UI["primary_dark"]),
        "accent":  (UI["accent"],  "#9E4710"),
        "success": (UI["success"], "#215C39"),
        "warning": (UI["warning"], "#8B5A17"),
        "danger":  (UI["danger"],  "#6E1F18"),
        "info":    (UI["info"],    "#431F50"),
        "ghost":   (UI["surface_alt"], UI["border"]),
    }
    base_bg, hover_bg = palette.get(variant, palette["primary"])
    fg = UI["text"] if variant == "ghost" else "white"
    label = f"{icon}   {text}" if icon else text
    btn = tk.Button(
        parent, text=label, command=command,
        bg=base_bg, fg=fg,
        activebackground=hover_bg, activeforeground=fg,
        font=UI["font_bold"],
        relief=tk.FLAT, bd=0, cursor="hand2",
        width=width, height=height,
        padx=10, pady=4,
    )
    _hover(btn, base_bg, hover_bg)
    return btn


def ui_card(parent: tk.Widget, **pack_kw) -> tk.Frame:
    """Contenedor tipo tarjeta con fondo blanco y borde sutil."""
    outer = tk.Frame(parent, bg=UI["border"])
    outer.pack(**pack_kw)
    inner = tk.Frame(outer, bg=UI["surface"])
    inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
    return inner


def show_revaluar_gui(config: Config) -> None:
    excel_path = config.excel_path
    if not excel_path.exists():
        logging.error("No existe el Excel: %s", excel_path)
        return

    wb = load_workbook(excel_path, read_only=True)
    ws = wb.active
    rows_data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        nro = str(row[0] or "").strip()
        if not nro:
            continue
        cat = str(row[1] or "")
        concepto = str(row[3] or "")
        gerencia = str(row[5] or "")
        plazo = row[8]
        if hasattr(plazo, "date"):
            plazo = plazo.date()
        plazo_str = plazo.strftime("%d-%m-%Y") if isinstance(plazo, date) else ""
        rows_data.append({
            "nro": nro, "categoria": cat, "concepto": concepto,
            "gerencia": gerencia, "plazo_str": plazo_str,
        })
    wb.close()

    if not rows_data:
        messagebox.showinfo("Revaloración", "No hay oficios en el Excel.")
        return

    root = tk.Tk()
    root.title("Revaloración de oficios")
    root.geometry("880x600")
    root.configure(bg=UI["bg"])
    root.resizable(True, True)

    ui_header(root, "Revaloración de Oficios",
              subtitle="Seleccione un oficio y corrija sus datos",
              bg=UI["success"])

    # --- Lista de oficios (tarjeta) ---
    list_card = ui_card(root, fill=tk.BOTH, expand=True, padx=16, pady=(14, 8))
    tk.Label(list_card, text="Oficios registrados",
             font=UI["font_heading"], bg=UI["surface"], fg=UI["text"]) \
        .pack(anchor=tk.W, padx=14, pady=(12, 6))

    list_wrap = tk.Frame(list_card, bg=UI["surface"])
    list_wrap.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))
    listbox = tk.Listbox(
        list_wrap, font=UI["font_mono"], selectmode=tk.SINGLE,
        bg=UI["surface"], fg=UI["text"],
        selectbackground=UI["primary_light"], selectforeground="white",
        relief=tk.FLAT, highlightthickness=1,
        highlightbackground=UI["border"], highlightcolor=UI["primary_light"],
        activestyle="none",
    )
    scrollbar = tk.Scrollbar(list_wrap, orient=tk.VERTICAL, command=listbox.yview)
    listbox.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.pack(fill=tk.BOTH, expand=True)

    for rd in rows_data:
        listbox.insert(
            tk.END,
            f"  Nro {rd['nro']:<8}  │  {rd['categoria']:<18}  │  "
            f"{rd['gerencia']:<18}  │  {rd['plazo_str']:<10}  │  {rd['concepto'][:60]}"
        )

    # --- Formulario de corrección (tarjeta) ---
    form_card = ui_card(root, fill=tk.X, padx=16, pady=(0, 8))
    tk.Label(form_card, text="Corregir oficio seleccionado",
             font=UI["font_heading"], bg=UI["surface"], fg=UI["text"]) \
        .pack(anchor=tk.W, padx=14, pady=(12, 8))

    frame_form = tk.Frame(form_card, bg=UI["surface"])
    frame_form.pack(fill=tk.X, padx=14, pady=(0, 14))

    lbl_cfg = {"bg": UI["surface"], "fg": UI["text_muted"], "font": UI["font_base"]}

    tk.Label(frame_form, text="Nueva área responsable", **lbl_cfg) \
        .grid(row=0, column=0, sticky=tk.W, pady=6, padx=(0, 12))
    area_var = tk.StringVar(root, value="(sin cambio)")
    area_options = ["(sin cambio)"] + AREAS_VALIDAS
    area_menu = tk.OptionMenu(frame_form, area_var, *area_options)
    area_menu.config(
        width=24, bg=UI["surface_alt"], fg=UI["text"],
        font=UI["font_base"], relief=tk.FLAT, bd=0,
        activebackground=UI["border"], highlightthickness=1,
        highlightbackground=UI["border"], cursor="hand2",
    )
    area_menu["menu"].config(bg=UI["surface"], fg=UI["text"], font=UI["font_base"])
    area_menu.grid(row=0, column=1, sticky=tk.W, pady=6)

    tk.Label(frame_form, text="Nuevo plazo (DD-MM-YYYY)", **lbl_cfg) \
        .grid(row=1, column=0, sticky=tk.W, pady=6, padx=(0, 12))
    plazo_entry = tk.Entry(
        frame_form, width=22, font=UI["font_base"],
        bg=UI["surface"], fg=UI["text"],
        relief=tk.FLAT, highlightthickness=1,
        highlightbackground=UI["border"], highlightcolor=UI["primary_light"],
    )
    plazo_entry.grid(row=1, column=1, sticky=tk.W, pady=6, ipady=4)

    tk.Label(frame_form, text="¿Es multa / formulación de cargos?", **lbl_cfg) \
        .grid(row=2, column=0, sticky=tk.W, pady=6, padx=(0, 12))
    multa_var = tk.StringVar(value="(sin cambio)")
    multa_frame = tk.Frame(frame_form, bg=UI["surface"])
    multa_frame.grid(row=2, column=1, sticky=tk.W, pady=6)
    for label_text, val in [("(sin cambio)", "(sin cambio)"),
                            ("Sí, es multa", "si"),
                            ("No es multa", "no")]:
        tk.Radiobutton(
            multa_frame, text=label_text, variable=multa_var, value=val,
            bg=UI["surface"], fg=UI["text"], font=UI["font_base"],
            activebackground=UI["surface"], selectcolor=UI["surface"],
            highlightthickness=0, cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 12))

    status_label = tk.Label(root, text="", fg=UI["success"],
                             font=UI["font_small"], bg=UI["bg"])
    status_label.pack(pady=(4, 0))

    def on_save():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Revaloración", "Seleccione un oficio primero.")
            return
        idx = sel[0]
        rd = rows_data[idx]
        corrections = load_corrections(config.corrections_path)
        updates: Dict[str, Any] = {}
        new_area = area_var.get()
        new_plazo_raw = plazo_entry.get().strip()
        new_multa = multa_var.get()

        if new_area and new_area != "(sin cambio)":
            corrections.append({
                "nro": rd["nro"],
                "campo": "gerencia_responsable",
                "valor_anterior": rd["gerencia"],
                "valor_nuevo": new_area,
                "concepto": rd["concepto"],
            })
            updates["gerencia_responsable"] = new_area

        if new_plazo_raw:
            try:
                parsed_plazo = datetime.strptime(new_plazo_raw, "%d-%m-%Y").date()
                corrections.append({
                    "nro": rd["nro"],
                    "campo": "plazo_respuesta",
                    "valor_anterior": rd["plazo_str"],
                    "valor_nuevo": parsed_plazo.isoformat(),
                    "concepto": rd["concepto"],
                })
                updates["plazo_respuesta"] = parsed_plazo.isoformat()
            except ValueError:
                messagebox.showerror("Error", "Formato de fecha inválido. Use DD-MM-YYYY.")
                return

        if new_multa in ("si", "no"):
            corrections.append({
                "nro": rd["nro"],
                "campo": "es_multa",
                "valor_nuevo": (new_multa == "si"),
                "concepto": rd["concepto"],
            })
            updates["es_multa"] = (new_multa == "si")

        if not updates and new_multa == "(sin cambio)":
            messagebox.showinfo("Revaloración", "No se indicó ninguna corrección.")
            return

        update_excel_row(excel_path, rd["nro"], rd["categoria"], updates, config.gerentes)
        save_corrections(config.corrections_path, corrections)

        # Actualizar listbox
        new_gerencia = updates.get("gerencia_responsable", rd["gerencia"])
        new_plazo_str = rd["plazo_str"]
        if "plazo_respuesta" in updates:
            p = parse_date_yyyy_mm_dd(updates["plazo_respuesta"])
            new_plazo_str = p.strftime("%d-%m-%Y") if p else rd["plazo_str"]
        rd["gerencia"] = new_gerencia
        rd["plazo_str"] = new_plazo_str
        listbox.delete(idx)
        listbox.insert(
            idx,
            f"  Nro {rd['nro']:<8}  │  {rd['categoria']:<18}  │  "
            f"{rd['gerencia']:<18}  │  {rd['plazo_str']:<10}  │  {rd['concepto'][:60]}"
        )

        status_label.config(text=f"Oficio Nro {rd['nro']} corregido correctamente.")
        logging.info("Corrección aplicada: Nro %s — %s", rd["nro"], updates)

        area_var.set("(sin cambio)")
        plazo_entry.delete(0, tk.END)
        multa_var.set("(sin cambio)")

    btn_row = tk.Frame(root, bg=UI["bg"])
    btn_row.pack(pady=(6, 14))
    ui_button(btn_row, "Guardar corrección", command=on_save,
              variant="success", width=22, icon="✓").pack(side=tk.LEFT, padx=5)
    ui_button(btn_row, "Cerrar", command=root.destroy,
              variant="ghost", width=12).pack(side=tk.LEFT, padx=5)

    root.mainloop()


# ---------------------------------------------------------------------------
# Alerta de oficios próximos a vencer
# ---------------------------------------------------------------------------

def get_upcoming_deadlines(excel_path: Path, days: int = 5) -> List[Dict[str, Any]]:
    if not excel_path.exists():
        return []
    wb = load_workbook(excel_path, read_only=True)
    ws = wb.active
    today = date.today()
    limit = today + timedelta(days=days)
    upcoming: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        plazo = row[8]  # columna I = Plazo Respuesta
        if plazo is None:
            continue
        if hasattr(plazo, "date"):
            plazo = plazo.date()
        if not isinstance(plazo, date):
            continue
        if today <= plazo <= limit:
            upcoming.append({
                "nro": str(row[0] or ""),
                "categoria": str(row[1] or ""),
                "concepto": str(row[3] or ""),
                "gerencia": str(row[5] or ""),
                "gerente": str(row[6] or ""),
                "plazo": plazo,
            })
    wb.close()
    return upcoming


def show_upcoming_deadlines_popup(excel_path: Path) -> None:
    upcoming = get_upcoming_deadlines(excel_path)
    if not upcoming:
        return
    lines = [f"Oficios que vencen en los próximos 5 días: {len(upcoming)}", ""]
    for item in sorted(upcoming, key=lambda x: x["plazo"]):
        dias_restantes = (item["plazo"] - date.today()).days
        lines.append(
            f"  • Nro {item['nro']} ({item['categoria']})"
            f" — Vence: {item['plazo'].strftime('%d-%m-%Y')}"
            f" ({dias_restantes}d)"
        )
        lines.append(f"    Área: {item['gerencia']} | Gerente: {item['gerente']}")
        if item["concepto"]:
            lines.append(f"    {item['concepto'][:100]}")
        lines.append("")
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning("Oficios próximos a vencer", "\n".join(lines))
    root.destroy()


def process_directory(
    config: Config, state: dict[str, Any]
) -> Tuple[ProcessingStats, List[Tuple[Path, Dict[str, Any]]]]:
    """Procesa los PDFs pendientes. Retorna (stats, lista_de_multas)."""
    empty_stats = ProcessingStats()
    ensure_excel_exists(config.excel_path)
    remove_duplicate_files(config.watch_dir, config.scan_extensions)
    processed_hashes: set[str] = set(state.get("processed_hashes", []))
    pending = find_pending_pdfs(config, processed_hashes)

    if not pending:
        logging.info("No hay PDFs nuevos para procesar en %s", config.watch_dir)
        return empty_stats, []

    logging.info("Se encontraron %s PDF(s) nuevos para procesar.", len(pending))

    corrections = load_corrections(config.corrections_path)
    corrections_prompt = build_corrections_prompt(corrections)
    history_prompt = build_history_prompt(config.excel_path)
    learning_prompt = history_prompt + corrections_prompt

    stats = ProcessingStats()
    multa_pdfs: List[Tuple[Path, Dict[str, Any]]] = []
    changed = False
    for pdf_path in pending:
        logging.info("Procesando %s", pdf_path.name)
        file_hash = sha256_file(pdf_path)
        try:
            extracted = call_openai_extract(config, pdf_path, learning_prompt)

            # Si detectó un oficio relacionado, buscar su PDF y re-extraer con ambos
            related_nro = extracted.get("oficio_relacionado")
            if related_nro:
                related_path = find_related_pdf(config.watch_dir, related_nro)
                if related_path and related_path != pdf_path:
                    logging.info(
                        "Oficio relacionado encontrado: %s → re-extrayendo con ambos PDFs.",
                        related_path.name,
                    )
                    extracted = call_openai_extract(
                        config, pdf_path, learning_prompt, related_pdf_path=related_path
                    )
                elif related_nro:
                    logging.info(
                        "Oficio relacionado Nro %s mencionado pero no encontrado en directorio.",
                        related_nro,
                    )

            due_date = compute_due_date(extracted)
            if due_date and not extracted.get("plazo_respuesta") and extracted.get("plazo_relativo_cantidad"):
                logging.info(
                    "Plazo relativo calculado para %s | fecha_oficio=%s | cantidad=%s | tipo=%s | vencimiento=%s",
                    pdf_path.name,
                    extracted.get("fecha_oficio"),
                    extracted.get("plazo_relativo_cantidad"),
                    extracted.get("plazo_relativo_tipo"),
                    due_date.isoformat(),
                )
            row = map_row(extracted, config.gerentes)
            append_to_excel(config.excel_path, row)
            if due_date:
                sync_to_planner(config, extracted, due_date)
            stats.registrar(extracted)
            nro = extracted.get("numero_oficio") or ""
            concepto = extracted.get("concepto") or ""
            if is_multa(nro, concepto, corrections):
                multa_pdfs.append((pdf_path, extracted))
            processed_hashes.add(file_hash)
            changed = True
            logging.info("PDF procesado correctamente: %s", pdf_path.name)
        except Exception as exc:
            stats.errores += 1
            logging.exception("Error procesando %s: %s", pdf_path.name, exc)

    if changed:
        state["processed_hashes"] = sorted(processed_hashes)
        save_state(config.processed_state_path, state)

    return stats, multa_pdfs


def parse_run_time(run_time: str) -> tuple[int, int]:
    hour_str, minute_str = run_time.split(":")
    return int(hour_str), int(minute_str)


def service_loop(config: Config) -> None:
    tz = ZoneInfo(config.timezone)
    logging.info("Servicio iniciado. Zona horaria=%s | hora programada=%s", config.timezone, config.run_time)

    while True:
        state = load_state(config.processed_state_path)
        now = datetime.now(tz)
        hour, minute = parse_run_time(config.run_time)
        today = now.date().isoformat()

        if (now.hour > hour or (now.hour == hour and now.minute >= minute)) and state.get("last_run_date") != today:
            logging.info("Ejecución diaria iniciada.")
            stats, multa_pdfs = process_directory(config, state)
            if stats.total or stats.errores:
                show_summary_popup(stats)
            show_upcoming_deadlines_popup(config.excel_path)
            for pdf_path, extracted in multa_pdfs:
                ask_and_generate_informe(config, pdf_path, extracted)
            state["last_run_date"] = today
            save_state(config.processed_state_path, state)
            logging.info("Ejecución diaria finalizada.")

        time.sleep(30)


def run_once(config: Config) -> None:
    state = load_state(config.processed_state_path)
    stats, multa_pdfs = process_directory(config, state)
    if stats.total or stats.errores:
        show_summary_popup(stats)
    show_upcoming_deadlines_popup(config.excel_path)
    for pdf_path, extracted in multa_pdfs:
        ask_and_generate_informe(config, pdf_path, extracted)


def reset_state(config: Config) -> None:
    state = {"processed_hashes": [], "last_run_date": None}
    save_state(config.processed_state_path, state)
    logging.info("Estado reseteado. Se reprocesarán todos los PDFs en la próxima ejecución.")


# ---------------------------------------------------------------------------
# Generar informe de multa manual
# ---------------------------------------------------------------------------

def show_generar_informe_gui(config: Config) -> None:
    """Permite seleccionar una multa del Excel y generar su informe Word."""
    excel_path = config.excel_path
    if not excel_path.exists():
        messagebox.showinfo("Informe de multa", "No existe el archivo Excel.")
        return

    wb = load_workbook(excel_path, read_only=True)
    ws = wb.active
    multas_data: List[Dict[str, str]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        nro = str(row[0] or "").strip()
        if not nro:
            continue
        multa_flag = str(row[9] or "").strip().lower()
        if multa_flag not in ("sí", "si"):
            continue
        multas_data.append({
            "nro": nro,
            "categoria": str(row[1] or ""),
            "concepto": str(row[3] or ""),
            "gerencia": str(row[5] or ""),
        })
    wb.close()

    if not multas_data:
        messagebox.showinfo("Informe de multa", "No hay oficios marcados como multa en el Excel.")
        return

    win = tk.Toplevel()
    win.title("Generar Informe de Multa")
    win.geometry("820x500")
    win.configure(bg=UI["bg"])
    win.resizable(True, True)

    ui_header(win, "Informe de Multa",
              subtitle="Seleccione una multa para generar el informe Word",
              bg=UI["danger"])

    list_card = ui_card(win, fill=tk.BOTH, expand=True, padx=16, pady=(14, 8))
    tk.Label(list_card, text="Multas registradas",
             font=UI["font_heading"], bg=UI["surface"], fg=UI["text"]) \
        .pack(anchor=tk.W, padx=14, pady=(12, 6))

    frame_list = tk.Frame(list_card, bg=UI["surface"])
    frame_list.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))

    listbox = tk.Listbox(
        frame_list, font=UI["font_mono"], selectmode=tk.SINGLE,
        bg=UI["surface"], fg=UI["text"],
        selectbackground=UI["danger"], selectforeground="white",
        relief=tk.FLAT, highlightthickness=1,
        highlightbackground=UI["border"], highlightcolor=UI["danger"],
        activestyle="none",
    )
    scrollbar = tk.Scrollbar(frame_list, orient=tk.VERTICAL, command=listbox.yview)
    listbox.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.pack(fill=tk.BOTH, expand=True)

    for md in multas_data:
        listbox.insert(
            tk.END,
            f"  Nro {md['nro']:<8}  │  {md['categoria']:<18}  │  "
            f"{md['gerencia']:<18}  │  {md['concepto'][:70]}"
        )

    status_label = tk.Label(win, text="", fg=UI["danger"],
                             font=UI["font_small"], bg=UI["bg"])
    status_label.pack(pady=(4, 0))

    def on_generate() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Informe de multa", "Seleccione una multa primero.")
            return
        md = multas_data[sel[0]]
        nro = md["nro"]

        # Buscar el PDF en watch_dir
        pdf_path = find_related_pdf(config.watch_dir, nro)
        if not pdf_path:
            messagebox.showerror(
                "PDF no encontrado",
                f"No se encontró el PDF del oficio Nro {nro} en:\n{config.watch_dir}\n\n"
                "Asegúrese de que el archivo PDF esté en la carpeta de oficios."
            )
            return

        if not DOCX_AVAILABLE:
            messagebox.showerror(
                "Dependencia faltante",
                "Instale python-docx para generar informes Word:\n  pip install python-docx"
            )
            return

        output_path = config.informe_output_dir / f"Informe_Multa_Nro{nro}.docx"
        status_label.config(text=f"Generando informe para Nro {nro}, espere...")
        win.update_idletasks()

        def worker() -> None:
            try:
                informe_data = call_anthropic_informe_multa(config, pdf_path)
                ok = fill_informe_multa(config, informe_data, nro, output_path)
                if ok:
                    win.after(0, lambda: status_label.config(
                        text=f"Informe generado: {output_path.name}"))
                    win.after(0, lambda: messagebox.showinfo(
                        "Informe generado",
                        f"Informe de multa guardado en:\n{output_path}"
                    ))
                else:
                    win.after(0, lambda: status_label.config(text="Error al generar el informe."))
            except Exception as exc:
                logging.exception("Error generando informe de multa: %s", exc)
                win.after(0, lambda: status_label.config(text="Error al generar el informe."))
                win.after(0, lambda: messagebox.showerror(
                    "Error", f"No se pudo generar el informe:\n{exc}"))

        threading.Thread(target=worker, daemon=True).start()

    btn_frame = tk.Frame(win, bg=UI["bg"])
    btn_frame.pack(pady=(6, 14))
    ui_button(btn_frame, "Generar informe", command=on_generate,
              variant="danger", width=22, icon="📝").pack(side=tk.LEFT, padx=5)
    ui_button(btn_frame, "Cerrar", command=win.destroy,
              variant="ghost", width=12).pack(side=tk.LEFT, padx=5)


# ---------------------------------------------------------------------------
# Estadísticas
# ---------------------------------------------------------------------------

_PIE_COLORS = [
    "#1F4E78", "#C55A11", "#2E7D4F", "#5B2C6F",
    "#922B21", "#1A7D8E", "#D4AC0D", "#117A65",
]


def _draw_pie_chart(canvas: tk.Canvas, data: Dict[str, int],
                    cx: int, cy: int, r: int) -> None:
    """Dibuja un gráfico de torta tipo donut con leyenda a la derecha."""
    total = sum(data.values())
    if total == 0:
        return

    items = sorted(data.items(), key=lambda x: -x[1])

    # Sombra sutil
    canvas.create_oval(cx - r + 3, cy - r + 4, cx + r + 3, cy + r + 4,
                       fill="#D6DEE7", outline="")

    start = 90.0  # comenzar arriba
    for i, (label, count) in enumerate(items):
        extent = -count / total * 360  # sentido horario
        color = _PIE_COLORS[i % len(_PIE_COLORS)]
        canvas.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=start, extent=extent,
            fill=color, outline="white", width=3,
        )
        start += extent

    # Agujero central (efecto donut)
    hole = int(r * 0.55)
    canvas.create_oval(cx - hole, cy - hole, cx + hole, cy + hole,
                       fill=UI["surface"], outline="")
    canvas.create_text(cx, cy - 8, text=f"{total}",
                       fill=UI["text"], font=("Segoe UI", 18, "bold"))
    canvas.create_text(cx, cy + 14, text="oficios",
                       fill=UI["text_muted"], font=UI["font_small"])

    # Leyenda a la derecha
    lx = cx + r + 30
    ly0 = cy - r + 6
    for i, (label, count) in enumerate(items):
        color = _PIE_COLORS[i % len(_PIE_COLORS)]
        ly = ly0 + i * 24
        canvas.create_oval(lx, ly + 1, lx + 12, ly + 13,
                           fill=color, outline=color)
        pct = count * 100 / total
        canvas.create_text(lx + 22, ly + 7, anchor=tk.W,
                           text=f"{label}",
                           fill=UI["text"], font=UI["font_bold"])
        canvas.create_text(lx + 22, ly + 20, anchor=tk.W,
                           text=f"{count} oficios · {pct:.1f}%",
                           fill=UI["text_muted"], font=UI["font_small"])
        ly0 += 10  # espacio extra por línea inferior


def show_estadisticas_gui(config: Config) -> None:
    """Muestra una ventana con estadísticas agregadas del Excel."""
    excel_path = config.excel_path
    if not excel_path.exists():
        messagebox.showinfo("Estadísticas", "No existe el archivo Excel.")
        return

    wb = load_workbook(excel_path, read_only=True)
    ws = wb.active

    total = 0
    categorias: Counter = Counter()
    areas: Counter = Counter()
    multas = 0
    fechas: List[date] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        nro = str(row[0] or "").strip()
        if not nro:
            continue
        total += 1
        categorias[str(row[1] or "Sin categoría")] += 1
        areas[str(row[5] or "Sin área")] += 1
        if str(row[9] or "").strip().lower() in ("sí", "si"):
            multas += 1
        fecha = row[2]
        if hasattr(fecha, "date"):
            fecha = fecha.date()
        if isinstance(fecha, date):
            fechas.append(fecha)
    wb.close()

    if total == 0:
        messagebox.showinfo("Estadísticas", "No hay oficios registrados en el Excel.")
        return

    win = tk.Toplevel()
    win.title("Estadísticas de Oficios")
    win.geometry("880x720")
    win.configure(bg=UI["bg"])
    win.resizable(True, True)

    ui_header(win, "Estadísticas de Oficios",
              subtitle=f"Total de {total} oficios registrados",
              bg=UI["info"])

    # --- Fila de métricas (KPI cards) ---
    kpi_row = tk.Frame(win, bg=UI["bg"])
    kpi_row.pack(fill=tk.X, padx=16, pady=(14, 8))

    def _kpi(parent: tk.Widget, value: str, label: str, color: str) -> None:
        card = ui_card(parent, side=tk.LEFT, expand=True, fill=tk.X, padx=4)
        tk.Label(card, text=value, font=("Segoe UI", 22, "bold"),
                 bg=UI["surface"], fg=color).pack(pady=(12, 0))
        tk.Label(card, text=label, font=UI["font_small"],
                 bg=UI["surface"], fg=UI["text_muted"]).pack(pady=(2, 14))

    _kpi(kpi_row, str(total), "Oficios totales", UI["primary"])
    _kpi(kpi_row, str(len(areas)), "Áreas distintas", UI["success"])
    _kpi(kpi_row, str(len(categorias)), "Categorías", UI["accent"])
    _kpi(kpi_row, str(multas), "Multas detectadas", UI["danger"])

    # --- Gráfico de torta (tarjeta) ---
    chart_card = ui_card(win, fill=tk.X, padx=16, pady=8)
    tk.Label(chart_card, text="Distribución por área responsable",
             font=UI["font_heading"], bg=UI["surface"], fg=UI["text"]) \
        .pack(anchor=tk.W, padx=14, pady=(12, 4))

    n_areas = len(areas)
    chart_w = max(620, 340 + n_areas * 28)
    chart_h = max(280, 60 + n_areas * 34)
    canvas = tk.Canvas(chart_card, width=chart_w, height=chart_h,
                       bg=UI["surface"], highlightthickness=0)
    canvas.pack(padx=14, pady=(0, 14))
    _draw_pie_chart(canvas, dict(areas), cx=140, cy=chart_h // 2, r=110)

    # --- Desglose detallado (tarjeta con texto) ---
    detail_card = ui_card(win, fill=tk.BOTH, expand=True, padx=16, pady=8)
    tk.Label(detail_card, text="Detalle",
             font=UI["font_heading"], bg=UI["surface"], fg=UI["text"]) \
        .pack(anchor=tk.W, padx=14, pady=(12, 4))

    lines = ["Por categoría:"]
    for cat, n in sorted(categorias.items(), key=lambda x: -x[1]):
        pct = n * 100 / total
        lines.append(f"   • {cat:<24} {n:>4}   ({pct:>5.1f}%)")
    lines.append("")
    lines.append("Por área responsable:")
    for area, n in sorted(areas.items(), key=lambda x: -x[1]):
        pct = n * 100 / total
        lines.append(f"   • {area:<24} {n:>4}   ({pct:>5.1f}%)")
    if fechas:
        lines.append("")
        lines.append(
            f"Rango de fechas: {min(fechas).strftime('%d-%m-%Y')} "
            f"→ {max(fechas).strftime('%d-%m-%Y')}"
        )

    text = tk.Text(detail_card, font=UI["font_mono"], wrap=tk.WORD,
                   bg=UI["surface"], fg=UI["text"],
                   relief=tk.FLAT, bd=0, padx=14, pady=8,
                   height=10)
    text.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 12))
    text.insert(tk.END, "\n".join(lines))
    text.config(state=tk.DISABLED)

    btn_row = tk.Frame(win, bg=UI["bg"])
    btn_row.pack(pady=(4, 14))
    ui_button(btn_row, "Cerrar", command=win.destroy,
              variant="primary", width=15).pack()


# ---------------------------------------------------------------------------
# GUI principal
# ---------------------------------------------------------------------------

def launch_main_gui(config: Config) -> None:
    root = tk.Tk()
    root.title("Gestión de Oficios CGE")
    root.geometry("500x620")
    root.configure(bg=UI["bg"])
    root.resizable(False, False)

    ui_header(
        root, "Gestión de Oficios CGE",
        subtitle="Centro de control · Comercial y Servicio al Cliente",
        height=78,
    )

    status_var = tk.StringVar(value="● Sistema listo")

    def set_status(msg: str) -> None:
        root.after(0, lambda: status_var.set(msg))

    def on_run_complete(result: Any) -> None:
        stats, multa_pdfs = result
        for b in (btn_run, btn_reset, btn_revaluar, btn_informe, btn_stats):
            b.config(state=tk.NORMAL)
        if stats.total or stats.errores:
            show_summary_popup(stats)
        show_upcoming_deadlines_popup(config.excel_path)
        for pdf_path, extracted in multa_pdfs:
            ask_and_generate_informe(config, pdf_path, extracted)
        set_status(
            f"✓ Completado: {stats.total} PDF(s) procesado(s)"
            if stats.total else "● Sin PDFs nuevos"
        )

    def on_run_error(exc: Exception) -> None:
        for b in (btn_run, btn_reset, btn_revaluar, btn_informe, btn_stats):
            b.config(state=tk.NORMAL)
        messagebox.showerror("Error", str(exc))
        set_status("✗ Error durante el procesamiento")

    def run_once_action() -> None:
        for b in (btn_run, btn_reset, btn_revaluar, btn_informe, btn_stats):
            b.config(state=tk.DISABLED)
        set_status("⏳ Procesando PDFs, por favor espere...")

        def worker() -> None:
            try:
                state = load_state(config.processed_state_path)
                result = process_directory(config, state)
                root.after(0, lambda: on_run_complete(result))
            except Exception as exc:
                root.after(0, lambda: on_run_error(exc))

        threading.Thread(target=worker, daemon=True).start()

    def reset_action() -> None:
        if messagebox.askyesno(
            "Resetear valores",
            "¿Resetear la memoria de PDFs procesados?\n"
            "Todos los PDFs serán analizados de nuevo en la próxima ejecución."
        ):
            reset_state(config)
            messagebox.showinfo("Listo", "Memoria reseteada correctamente.")
            set_status("✓ Memoria reseteada")

    def revaluar_action() -> None:
        show_revaluar_gui(config)

    def informe_multa_action() -> None:
        show_generar_informe_gui(config)

    def estadisticas_action() -> None:
        show_estadisticas_gui(config)

    # --- Tarjeta principal con los botones ---
    card = ui_card(root, fill=tk.BOTH, expand=True, padx=24, pady=20)

    tk.Label(card, text="Acciones disponibles",
             font=UI["font_heading"], bg=UI["surface"], fg=UI["text"]) \
        .pack(anchor=tk.W, padx=22, pady=(18, 4))
    tk.Label(card, text="Seleccione una operación",
             font=UI["font_small"], bg=UI["surface"], fg=UI["text_muted"]) \
        .pack(anchor=tk.W, padx=22, pady=(0, 14))

    btn_frame = tk.Frame(card, bg=UI["surface"])
    btn_frame.pack(expand=True, pady=(0, 18))

    btn_run = ui_button(btn_frame, "Ejecutar una vez",
                        command=run_once_action,
                        variant="primary", icon="▶")
    btn_run.pack(pady=5)

    btn_reset = ui_button(btn_frame, "Resetear valores",
                          command=reset_action,
                          variant="accent", icon="↺")
    btn_reset.pack(pady=5)

    btn_revaluar = ui_button(btn_frame, "Revaluar oficio",
                             command=revaluar_action,
                             variant="success", icon="✎")
    btn_revaluar.pack(pady=5)

    btn_informe = ui_button(btn_frame, "Informe de multa",
                            command=informe_multa_action,
                            variant="danger", icon="📝")
    btn_informe.pack(pady=5)

    btn_stats = ui_button(btn_frame, "Estadísticas",
                          command=estadisticas_action,
                          variant="info", icon="📊")
    btn_stats.pack(pady=5)

    # --- Barra de estado minimalista ---
    status_bar = tk.Frame(root, bg=UI["primary_dark"], height=28)
    status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    status_bar.pack_propagate(False)
    tk.Label(status_bar, textvariable=status_var,
             bg=UI["primary_dark"], fg="white",
             font=UI["font_small"], anchor=tk.W,
             padx=14).pack(side=tk.LEFT, fill=tk.Y)

    root.mainloop()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gestión de Oficios CGE — abre la interfaz gráfica si no se pasan argumentos."
    )
    default_config = get_base_dir() / "config.json"
    parser.add_argument("--config", default=str(default_config), help="Ruta al archivo de configuración JSON.")
    parser.add_argument("--run-once", action="store_true", help="Ejecuta una sola vez y termina (sin GUI).")
    parser.add_argument("--service", action="store_true", help="Modo servicio continuo (sin GUI).")
    parser.add_argument(
        "--create-template",
        action="store_true",
        help="Crea solo la plantilla Excel definida en config.json y termina.",
    )
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Resetea la memoria de PDFs procesados para que se vuelvan a analizar.",
    )
    parser.add_argument(
        "--revaluar",
        action="store_true",
        help="Abre interfaz para corregir área o plazo de oficios ya procesados.",
    )
    args = parser.parse_args()

    config = load_config(Path(args.config))
    setup_logging(config.log_path)

    if args.create_template:
        create_excel_template(config.excel_path)
        logging.info("Plantilla creada: %s", config.excel_path)
        return

    if args.reset:
        reset_state(config)
        return

    if args.revaluar:
        show_revaluar_gui(config)
        return

    if args.run_once:
        run_once(config)
    elif args.service:
        service_loop(config)
    else:
        # Comportamiento por defecto: abrir la interfaz gráfica
        launch_main_gui(config)


if __name__ == "__main__":
    main()
