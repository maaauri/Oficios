from __future__ import annotations

import argparse
import base64
import hashlib
import json
import logging
import os
import re
import sys
import time
import tkinter as tk
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from pathlib import Path
from tkinter import messagebox
from typing import Any, Dict, List, Optional
from zoneinfo import ZoneInfo

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
            "enum": ["PMGD", "Conexiones", "Lectura", "Servicio al Cliente", "Cobranza", None],
        },
        "plazo_respuesta": {"type": ["string", "null"]},
        "plazo_relativo_cantidad": {"type": ["integer", "null"], "minimum": 1},
        "plazo_relativo_tipo": {
            "type": ["string", "null"],
            "enum": ["dias_corridos", "dias_habiles", None],
        },
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
   Criterios:
   - \"PMGD\": generación distribuida, conexión de PMGD, plataformas o procesos PMGD
   - \"Conexiones\": conexión, empalme, factibilidad, puesta en servicio, plazos de conexión, obras o procesos de conexión
   - \"Lectura\": lectura de medidores, medición, consumos leídos, toma de lectura
   - \"Servicio al Cliente\": atención, reclamos, canales, calidad de servicio, respuesta al cliente
   - \"Cobranza\": deuda, mora, pago, repactación, suspensión/corte por deuda, cobranza
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
9. Si un dato no está claro o no existe, devuelve null.
10. No agregues ningún texto fuera del JSON.
11. No inventes datos.
"""


@dataclass
class Gerente:
    nombre: str
    email: str = ""


@dataclass
class Config:
    watch_dir: Path
    excel_path: Path
    processed_state_path: Path
    log_path: Path
    timezone: str
    run_time: str
    openai_api_key: str
    model: str
    gerentes: Dict[str, Gerente]
    request_timeout_seconds: int = 180
    scan_extensions: tuple[str, ...] = (".pdf",)


def load_config(path: Path) -> Config:
    with path.open("r", encoding="utf-8") as f:
        raw = json.load(f)

    gerentes: Dict[str, Gerente] = {}
    for area, item in raw.get("gerentes", {}).items():
        gerentes[area] = Gerente(nombre=item.get("nombre", ""), email=item.get("email", ""))

    api_key = raw.get("openai_api_key") or os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        raise ValueError("Falta openai_api_key en config.json o variable de entorno OPENAI_API_KEY.")

    return Config(
        watch_dir=Path(raw["watch_dir"]),
        excel_path=Path(raw["excel_path"]),
        processed_state_path=Path(raw.get("processed_state_path", "processed_state.json")),
        log_path=Path(raw.get("log_path", "oficios_service.log")),
        timezone=raw.get("timezone", "America/Santiago"),
        run_time=raw.get("run_time", "16:00"),
        openai_api_key=api_key,
        model=raw.get("model", "gpt-5.4-mini"),
        gerentes=gerentes,
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
    if existing != EXPECTED_COLUMNS:
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


def call_openai_extract(config: Config, pdf_path: Path) -> dict[str, Any]:
    pdf_b64 = base64.b64encode(pdf_path.read_bytes()).decode("ascii")

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
                "content": [
                    {"type": "input_text", "text": PROMPT}
                ],
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": f"Analiza el siguiente PDF. Nombre del archivo: {pdf_path.name}",
                    },
                    {
                        "type": "input_file",
                        "filename": pdf_path.name,
                        "file_data": f"data:application/pdf;base64,{pdf_b64}",
                    },
                ],
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

    return [
        extracted.get("numero_oficio") or "",
        extracted.get("categoria") or "",
        fecha_oficio,
        extracted.get("concepto") or "",
        "Comercial y Servicio al Cliente",
        gerencia or "",
        gerente.nombre,
        "",
        plazo,
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
        file_hash = sha256_file(path)
        if file_hash in processed_hashes:
            logging.info("Omitido (ya fue procesado anteriormente): %s", path.name)
            continue
        pdfs.append(path)
    return pdfs


@dataclass
class ProcessingStats:
    total: int = 0
    errores: int = 0
    categorias: Counter = field(default_factory=Counter)
    areas: Counter = field(default_factory=Counter)

    def registrar(self, extracted: dict[str, Any]) -> None:
        self.total += 1
        cat = extracted.get("categoria") or "Sin categoría"
        area = extracted.get("gerencia_responsable") or "Sin área"
        self.categorias[cat] += 1
        self.areas[area] += 1

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
        return "\n".join(lines)


def show_summary_popup(stats: ProcessingStats) -> None:
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Resumen de procesamiento", stats.resumen())
    root.destroy()


def process_directory(config: Config, state: dict[str, Any]) -> None:
    ensure_excel_exists(config.excel_path)
    remove_duplicate_files(config.watch_dir, config.scan_extensions)
    processed_hashes: set[str] = set(state.get("processed_hashes", []))
    pending = find_pending_pdfs(config, processed_hashes)

    if not pending:
        logging.info("No hay PDFs nuevos para procesar en %s", config.watch_dir)
        return

    logging.info("Se encontraron %s PDF(s) nuevos para procesar.", len(pending))

    stats = ProcessingStats()
    changed = False
    for pdf_path in pending:
        logging.info("Procesando %s", pdf_path.name)
        file_hash = sha256_file(pdf_path)
        try:
            extracted = call_openai_extract(config, pdf_path)
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
            stats.registrar(extracted)
            processed_hashes.add(file_hash)
            changed = True
            logging.info("PDF procesado correctamente: %s", pdf_path.name)
        except Exception as exc:
            stats.errores += 1
            logging.exception("Error procesando %s: %s", pdf_path.name, exc)

    if changed:
        state["processed_hashes"] = sorted(processed_hashes)
        save_state(config.processed_state_path, state)

    if stats.total or stats.errores:
        show_summary_popup(stats)


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
            process_directory(config, state)
            state["last_run_date"] = today
            save_state(config.processed_state_path, state)
            logging.info("Ejecución diaria finalizada.")

        time.sleep(30)


def run_once(config: Config) -> None:
    state = load_state(config.processed_state_path)
    process_directory(config, state)


def reset_state(config: Config) -> None:
    state = {"processed_hashes": [], "last_run_date": None}
    save_state(config.processed_state_path, state)
    logging.info("Estado reseteado. Se reprocesarán todos los PDFs en la próxima ejecución.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Procesa PDFs de oficios y llena un Excel usando OpenAI API.")
    parser.add_argument("--config", default="config.json", help="Ruta al archivo de configuración JSON.")
    parser.add_argument("--run-once", action="store_true", help="Ejecuta una sola vez y termina.")
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

    if args.run_once:
        run_once(config)
    else:
        service_loop(config)


if __name__ == "__main__":
    main()
