#!/usr/bin/env python
"""Servidor HTTP local para conectar a interface web à análise geoespacial."""

from __future__ import annotations

import csv
import gzip
import json
import logging
import mimetypes
import os
import re
import shutil
import sqlite3
import tempfile
import threading
import uuid
from contextlib import closing
from email.message import Message
from functools import lru_cache
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path, PurePosixPath
from types import SimpleNamespace
from typing import Any, BinaryIO
from urllib.parse import parse_qs, urlparse

import geopandas as gpd
import polars as pl
import shapely
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from analise_sobreposicao import (
    DEFAULT_REF_COLUMNS,
    EXCEL_MAX_ROWS,
    _excel_cell_value,
    infer_delimiter,
    inspect_input_file,
    plan_table_consolidation,
    resolve_csv_encoding,
    resolve_join_key_column,
    run,
    write_geopackage,
    write_geoparquet,
)

LOGGER = logging.getLogger("web_app")
BASE_DIR = Path(__file__).resolve().parent
JOB_DIR = BASE_DIR / ".web_jobs"
MAX_UPLOAD_BYTES = int(os.environ.get("GEO_MAX_UPLOAD_BYTES", str(10 * 1024**3)))
MAX_INSPECT_BYTES = 200 * 1024 * 1024
MAX_REQUEST_OVERHEAD = 16 * 1024**2
UPLOAD_CHUNK_BYTES = 64 * 1024
MAX_RETAINED_LOGS = 500
MAP_FEATURE_LIMIT = 500
GZIP_MIN_BYTES = 1024
LEGACY_ROLE_PARAMS = {
    "municipio": "municipio",
    "uf": "uf",
    "ano": "year",
    "cnpj_if": "institution",
}
ALLOWED_EXTENSIONS = {
    "plots": {".csv", ".gpkg", ".geojson", ".json", ".parquet"},
    "tables": {".csv", ".parquet"},
    "indigenous": {".gpkg", ".geojson", ".json", ".parquet"},
}
STATIC_FILES = {
    "/": (BASE_DIR / "interface.html", "text/html; charset=utf-8"),
    "/interface.html": (BASE_DIR / "interface.html", "text/html; charset=utf-8"),
    "/README.md": (BASE_DIR / "README.md", "text/markdown; charset=utf-8"),
}

_jobs: dict[str, dict[str, Any]] = {}
_jobs_lock = threading.Lock()
_active_job_lock = threading.Lock()
_active_job_id: str | None = None


class MultipartReader:
    """Leitor multipart incremental para aceitar uploads grandes sem carregá-los na RAM."""

    def __init__(self, stream: BinaryIO, content_length: int, boundary: bytes) -> None:
        self.stream = stream
        self.remaining = content_length
        self.boundary = b"\r\n--" + boundary
        self.buffer = bytearray()

    def _fill(self, minimum: int = 1) -> None:
        while len(self.buffer) < minimum and self.remaining > 0:
            data = self.stream.read(min(UPLOAD_CHUNK_BYTES, self.remaining))
            if not data:
                raise ValueError("A conexão foi interrompida durante o upload.")
            self.remaining -= len(data)
            self.buffer.extend(data)

    def read_line(self, limit: int = 64 * 1024) -> bytes:
        while True:
            marker = self.buffer.find(b"\r\n")
            if marker >= 0:
                line = bytes(self.buffer[:marker])
                del self.buffer[: marker + 2]
                return line
            if len(self.buffer) > limit:
                raise ValueError("Cabeçalho multipart excede o tamanho permitido.")
            if self.remaining == 0:
                raise ValueError("Corpo multipart incompleto.")
            self._fill(len(self.buffer) + 1)

    def copy_part(self, destination: BinaryIO, max_bytes: int) -> int:
        written = 0
        while True:
            marker = self.buffer.find(self.boundary)
            if marker >= 0:
                data = self.buffer[:marker]
                if data:
                    written += len(data)
                    if written > max_bytes:
                        raise ValueError("O tamanho total dos arquivos excede o limite permitido.")
                    destination.write(data)
                del self.buffer[: marker + len(self.boundary)]
                return written

            if self.remaining == 0:
                raise ValueError("Separador multipart final não encontrado.")
            flush_size = len(self.buffer) - len(self.boundary) + 1
            if flush_size > 0:
                data = self.buffer[:flush_size]
                written += len(data)
                if written > max_bytes:
                    raise ValueError("O tamanho total dos arquivos excede o limite permitido.")
                destination.write(data)
                del self.buffer[:flush_size]
            self._fill(len(self.buffer) + 1)

    def read_exact(self, size: int) -> bytes:
        self._fill(size)
        if len(self.buffer) < size:
            raise ValueError("Corpo multipart incompleto.")
        data = bytes(self.buffer[:size])
        del self.buffer[:size]
        return data


class JobLogHandler(logging.Handler):
    def __init__(self, job_id: str) -> None:
        super().__init__()
        self.job_id = job_id
        self.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))

    def emit(self, record: logging.LogRecord) -> None:
        with _jobs_lock:
            job = _jobs.get(self.job_id)
            if job is not None:
                job["logs"].append(self.format(record))
                del job["logs"][:-MAX_RETAINED_LOGS]


def _safe_filename(filename: str, extensions: set[str]) -> str:
    normalized = filename.replace("\\", "/")
    safe_name = PurePosixPath(normalized).name
    if not safe_name or safe_name in {".", ".."} or "\x00" in safe_name:
        raise ValueError("Nome de arquivo inválido.")
    if Path(safe_name).suffix.lower() not in extensions:
        raise ValueError(f"{safe_name}: extensão não permitida.")
    return safe_name


def _parse_content_disposition(value: str) -> tuple[str | None, str | None]:
    message = Message()
    message["content-disposition"] = value
    return message.get_param("name", header="content-disposition"), message.get_filename()


def _parse_multipart(
    handler: BaseHTTPRequestHandler,
    job_root: Path,
    content_length: int,
    max_bytes: int = MAX_UPLOAD_BYTES,
) -> tuple[dict[str, list[Path]], dict[str, str]]:
    content_type = Message()
    content_type["content-type"] = handler.headers.get("Content-Type", "")
    boundary_text = content_type.get_param("boundary", header="content-type")
    if content_type.get_content_type() != "multipart/form-data" or not boundary_text:
        raise ValueError("O envio deve usar multipart/form-data.")
    boundary = boundary_text.encode("ascii", errors="strict")
    reader = MultipartReader(handler.rfile, content_length, boundary)
    if reader.read_line() != b"--" + boundary:
        raise ValueError("Corpo multipart inválido.")

    uploads: dict[str, list[Path]] = {field: [] for field in ALLOWED_EXTENSIONS}
    fields: dict[str, str] = {}
    seen_names: dict[str, set[str]] = {field: set() for field in ALLOWED_EXTENSIONS}
    total_uploaded = 0
    upload_root = job_root / "uploads"

    while True:
        header_lines: list[bytes] = []
        while True:
            line = reader.read_line()
            if not line:
                break
            header_lines.append(line)
        part_headers = Message()
        part_headers.set_type("text/plain")
        for line in header_lines:
            name, separator, value = line.decode("latin-1").partition(":")
            if not separator:
                raise ValueError("Cabeçalho de parte multipart inválido.")
            part_headers[name.strip()] = value.strip()

        field_name, filename = _parse_content_disposition(
            part_headers.get("content-disposition", "")
        )
        if not field_name:
            raise ValueError("Campo multipart sem nome.")

        if filename is not None:
            extensions = ALLOWED_EXTENSIONS.get(field_name)
            if extensions is None:
                raise ValueError(f"Campo de arquivo desconhecido: {field_name}.")
            filename = _safe_filename(filename, extensions)
            if filename.casefold() in seen_names[field_name]:
                raise ValueError(
                    f"Arquivos com o mesmo nome foram selecionados: {filename}."
                )
            seen_names[field_name].add(filename.casefold())
            output_dir = upload_root / field_name
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / filename
            with output_path.open("wb") as target:
                size = reader.copy_part(target, max_bytes - total_uploaded)
            if size == 0:
                raise ValueError(f"O arquivo {filename} está vazio.")
            total_uploaded += size
            uploads[field_name].append(output_path)
        else:
            from io import BytesIO

            value_buffer = BytesIO()
            reader.copy_part(value_buffer, 64 * 1024)
            try:
                fields[field_name] = value_buffer.getvalue().decode("utf-8").strip()
            except UnicodeDecodeError as error:
                raise ValueError(f"O campo {field_name} não está codificado em UTF-8.") from error

        ending = reader.read_exact(2)
        if ending == b"--":
            if reader.buffer.startswith(b"\r\n"):
                del reader.buffer[:2]
            return uploads, fields
        if ending != b"\r\n":
            raise ValueError("Final de parte multipart inválido.")


def _send_bytes(
    handler: BaseHTTPRequestHandler,
    status: int,
    payload: bytes,
    content_type: str,
    cache_control: str = "no-store",
) -> None:
    accept_encoding = handler.headers.get("Accept-Encoding", "").lower()
    use_gzip = "gzip" in accept_encoding and len(payload) >= GZIP_MIN_BYTES
    body = gzip.compress(payload, compresslevel=5) if use_gzip else payload
    try:
        handler.send_response(status)
        handler.send_header("Content-Type", content_type)
        handler.send_header("Content-Length", str(len(body)))
        handler.send_header("Cache-Control", cache_control)
        handler.send_header("X-Content-Type-Options", "nosniff")
        if use_gzip:
            handler.send_header("Content-Encoding", "gzip")
            handler.send_header("Vary", "Accept-Encoding")
        handler.end_headers()
        handler.wfile.write(body)
    except (BrokenPipeError, ConnectionResetError):
        LOGGER.info("Cliente encerrou a resposta HTTP antecipadamente.")


def _send_json(handler: BaseHTTPRequestHandler, status: int, body: dict[str, Any]) -> None:
    payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
    _send_bytes(handler, status, payload, "application/json; charset=utf-8")


@lru_cache(maxsize=2)
def _read_financial_institutions_cached(
    path_str: str, mtime_ns: int, size: int
) -> dict[str, str]:
    del mtime_ns, size
    path = Path(path_str)
    encoding = resolve_csv_encoding(path, "auto")
    delimiter = infer_delimiter(path, encoding)
    with path.open("r", encoding=encoding, newline="") as source:
        reader = csv.DictReader(source, delimiter=delimiter)
        headers = {
            (name or "").strip().lstrip("#").strip().upper(): name
            for name in reader.fieldnames or []
        }
        cnpj_column = headers.get("CNPJ_IF")
        name_column = headers.get("NOME_IF")
        if cnpj_column is None or name_column is None:
            return {}
        return {
            str(row[cnpj_column]).strip(): str(row[name_column]).strip()
            for row in reader
            if row.get(cnpj_column) and row.get(name_column)
        }


def _read_financial_institutions() -> dict[str, str]:
    """Carrega opcionalmente o catálogo de instituições financeiras com cache em memória."""
    path = BASE_DIR / "dados" / "SICOR_LISTA_IFS.csv"
    if not path.is_file():
        return {}
    try:
        stat = path.stat()
        return _read_financial_institutions_cached(
            str(path), stat.st_mtime_ns, stat.st_size
        )
    except Exception:
        LOGGER.warning("Não foi possível ler %s; exibindo códigos originais.", path, exc_info=True)
        return {}


def _load_dashboard_schema(
    connection: sqlite3.Connection,
) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]], dict[str, str]]:
    columns = [
        dict(row)
        for row in connection.execute(
            "SELECT col_order, col_name, label, data_type, semantic_role, "
            "default_visible, filterable, non_null_count, distinct_count "
            "FROM dashboard_columns ORDER BY col_order"
        ).fetchall()
    ]
    by_name = {col["col_name"]: col for col in columns}
    meta = {
        row["key"]: row["value"]
        for row in connection.execute("SELECT key, value FROM dashboard_meta").fetchall()
    }
    return columns, by_name, meta


def _quote_ident(column_name: str, columns_by_name: dict[str, dict[str, Any]]) -> str:
    if column_name not in columns_by_name:
        raise ValueError(f"Coluna não reconhecida: {column_name}")
    return '"' + column_name.replace('"', '""') + '"'


def _find_col_by_role(columns_list: list[dict[str, Any]], role: str) -> str | None:
    for col in columns_list:
        if col["semantic_role"] == role:
            return str(col["col_name"])
    return None


def _escape_like(term: str) -> str:
    return term.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _sqlite_table_exists(connection: sqlite3.Connection, table_name: str) -> bool:
    row = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        [table_name],
    ).fetchone()
    return row is not None


def _build_dashboard_where(
    query: dict[str, list[str]],
    columns_list: list[dict[str, Any]],
    columns_by_name: dict[str, dict[str, Any]],
    meta: dict[str, str],
    institutions: dict[str, str],
) -> tuple[str, list[Any], dict[str, str], list[tuple[str, str]], str]:
    clauses: list[str] = []
    parameters: list[Any] = []
    active_filters: dict[str, str] = {}
    filter_display: list[tuple[str, str]] = []

    search_text = query.get("q", [""])[0].strip()
    if search_text:
        pattern = f"%{_escape_like(search_text)}%"
        search_candidates = [
            col
            for col in columns_list
            if col["data_type"] in {"text", "date"}
            and col["semantic_role"] != "plot_uid"
            and int(col["non_null_count"]) > 0
        ]
        search_candidates.sort(
            key=lambda col: (
                {
                    "effective_key": 0,
                    "key": 1,
                    "base_layer": 2,
                    "borrower": 3,
                    "institution": 4,
                    "uf": 5,
                    "municipio": 6,
                }.get(col["semantic_role"], 10 if col["default_visible"] else 20),
                int(col["col_order"]),
            )
        )
        or_clauses: list[str] = []
        for col in search_candidates[:25]:
            quoted = _quote_ident(col["col_name"], columns_by_name)
            or_clauses.append(f"CAST({quoted} AS TEXT) LIKE ? ESCAPE '\\'")
            parameters.append(pattern)

        inst_col = meta.get("institution_column", "")
        if inst_col and inst_col in columns_by_name and institutions:
            needle = search_text.casefold()
            matched_codes = [
                code
                for code, name in institutions.items()
                if needle in name.casefold()
            ][:100]
            if matched_codes:
                placeholders = ", ".join("?" for _ in matched_codes)
                or_clauses.append(
                    f"{_quote_ident(inst_col, columns_by_name)} IN ({placeholders})"
                )
                parameters.extend(matched_codes)

        if or_clauses:
            clauses.append(f"({' OR '.join(or_clauses)})")
        active_filters["q"] = search_text
        filter_display.append(("Busca textual", search_text))

    camada_base = query.get("camada_base", [""])[0].strip()
    if camada_base:
        clauses.append(
            "PLOT_UID IN ("
            "SELECT mpl.PLOT_UID FROM map_plot_lands mpl "
            "JOIN map_lands ml ON ml.LAND_UID = mpl.LAND_UID "
            "WHERE ml.TI_NAME = ?)"
        )
        parameters.append(camada_base)
        active_filters["camada_base"] = camada_base
        filter_display.append(("Camada Base", camada_base))

    min_area_raw = query.get("min_area_ha", [""])[0].strip().replace(",", ".")
    if min_area_raw:
        try:
            min_area_val = float(min_area_raw)
        except ValueError as error:
            raise ValueError("Área mínima deve ser um número válido.") from error
        if min_area_val < 0:
            raise ValueError("Área mínima não pode ser negativa.")
        area_col = _find_col_by_role(columns_list, "area_overlap") or "AREA_SOBREPOSTA_HA"
        if area_col in columns_by_name:
            clauses.append(f"CAST({_quote_ident(area_col, columns_by_name)} AS REAL) >= ?")
            parameters.append(min_area_val)
            active_filters["min_area_ha"] = min_area_raw
            filter_display.append(("Área sobreposta mínima (ha)", min_area_raw))

    min_pct_raw = query.get("min_percent", [""])[0].strip().replace(",", ".")
    if min_pct_raw:
        try:
            min_pct_val = float(min_pct_raw)
        except ValueError as error:
            raise ValueError("Percentual mínimo deve ser um número válido.") from error
        if min_pct_val < 0 or min_pct_val > 100:
            raise ValueError("Percentual mínimo deve estar entre 0 e 100.")
        pct_col = _find_col_by_role(columns_list, "overlap_pct") or "PERCENTUAL_SOBREPOSICAO"
        if pct_col in columns_by_name:
            clauses.append(f"CAST({_quote_ident(pct_col, columns_by_name)} AS REAL) >= ?")
            parameters.append(min_pct_val)
            active_filters["min_percent"] = min_pct_raw
            filter_display.append(("Sobreposição mínima (%)", min_pct_raw))

    match_status = query.get("match_status", [""])[0].strip().lower()
    if match_status in {"matched", "unmatched"}:
        ops_src_col = meta.get("operations_source_column", "")
        if meta.get("has_operations") == "1" and ops_src_col and ops_src_col in columns_by_name:
            quoted_src = _quote_ident(ops_src_col, columns_by_name)
            if match_status == "matched":
                clauses.append(f"({quoted_src} IS NOT NULL AND {quoted_src} != '')")
                filter_display.append(("Correspondência em tabelas", "Com correspondência"))
            else:
                clauses.append(f"({quoted_src} IS NULL OR {quoted_src} = '')")
                filter_display.append(("Correspondência em tabelas", "Sem correspondência"))
            active_filters["match_status"] = match_status

    seen_cols: set[str] = set()
    for param, values in query.items():
        value = (values[0] if values else "").strip()
        if not value:
            continue
        target_col: str | None = None
        if param.startswith("f_"):
            candidate = param[2:]
            if candidate in columns_by_name:
                target_col = candidate
        elif param in LEGACY_ROLE_PARAMS:
            target_col = _find_col_by_role(columns_list, LEGACY_ROLE_PARAMS[param])

        if target_col is None or target_col in seen_cols:
            continue
        seen_cols.add(target_col)
        spec = columns_by_name[target_col]
        quoted = _quote_ident(target_col, columns_by_name)
        if spec["data_type"] in {"text", "date"}:
            clauses.append(f"{quoted} = ?")
        else:
            clauses.append(f"CAST({quoted} AS TEXT) = ?")
        parameters.append(value)
        active_filters[param] = value
        display_val = (
            institutions.get(value, value)
            if spec["semantic_role"] == "institution"
            else value
        )
        filter_display.append((str(spec["label"]), display_val))

    where = f" WHERE {' AND '.join(clauses)}" if clauses else ""
    return where, parameters, active_filters, filter_display, camada_base


def _filtered_dashboard_summary(
    connection: sqlite3.Connection,
    where: str,
    parameters: list[Any],
    record_count: int,
    columns_list: list[dict[str, Any]],
    columns_by_name: dict[str, dict[str, Any]],
    meta: dict[str, str],
) -> list[dict[str, Any]]:
    area_overlap_col = _find_col_by_role(columns_list, "area_overlap") or "AREA_SOBREPOSTA_HA"
    quoted_overlap = _quote_ident(area_overlap_col, columns_by_name)
    plot_count, overlap_area = connection.execute(
        "SELECT COUNT(*), COALESCE(SUM(AREA_SOBREPOSICAO), 0) FROM ("
        f"SELECT PLOT_UID, MAX(CAST({quoted_overlap} AS REAL)) AS AREA_SOBREPOSICAO "
        f"FROM dashboard_rows{where} GROUP BY PLOT_UID)",
        parameters,
    ).fetchone()

    cards: list[dict[str, Any]] = [
        {"label": "Registros no filtro", "value": int(record_count), "format": "integer"},
        {
            "label": "Glebas com sobreposição no filtro",
            "value": int(plot_count),
            "format": "integer",
        },
        {
            "label": "Soma da área sobreposta (ha)",
            "value": float(overlap_area),
            "format": "area_full",
        },
    ]

    finance_col = meta.get("finance_column", "")
    if (
        finance_col
        and finance_col in columns_by_name
        and int(columns_by_name[finance_col]["non_null_count"]) > 0
    ):
        quoted_fin = _quote_ident(finance_col, columns_by_name)
        eff_col = meta.get("effective_key_column", "")
        key_col = meta.get("key_column", "")
        nu_col = meta.get("nu_ordem_column", "")

        ref_parts: list[str] = []
        if eff_col and eff_col in columns_by_name:
            ref_parts.append(
                f"NULLIF(CAST({_quote_ident(eff_col, columns_by_name)} AS TEXT), '')"
            )
        if key_col and key_col in columns_by_name:
            ref_parts.append(
                f"NULLIF(CAST({_quote_ident(key_col, columns_by_name)} AS TEXT), '')"
            )
        ref_parts.append("CAST(rowid AS TEXT)")
        ref_expr = f"COALESCE({', '.join(ref_parts)})"
        if nu_col and nu_col in columns_by_name:
            group_expr = (
                f"{ref_expr} || '|' || "
                f"COALESCE(CAST({_quote_ident(nu_col, columns_by_name)} AS TEXT), '')"
            )
        else:
            group_expr = ref_expr

        financed = connection.execute(
            "SELECT COALESCE(SUM(VAL_PARC), 0) FROM ("
            f"SELECT MAX(CAST({quoted_fin} AS REAL)) AS VAL_PARC "
            f"FROM dashboard_rows{where} GROUP BY {group_expr})",
            parameters,
        ).fetchone()[0]

        fin_spec = columns_by_name[finance_col]
        fin_format = str(fin_spec["data_type"])
        fin_label = (
            "Soma dos valores financiados relacionados"
            if fin_format == "currency"
            else f"Soma de {fin_spec['label']}"
        )
        cards.append(
            {"label": fin_label, "value": float(financed), "format": fin_format}
        )
    return cards


def _build_dashboard_facets(
    connection: sqlite3.Connection,
    columns_list: list[dict[str, Any]],
    columns_by_name: dict[str, dict[str, Any]],
    institutions: dict[str, str],
) -> list[dict[str, Any]]:
    facets: list[dict[str, Any]] = []
    base_names = [
        row[0]
        for row in connection.execute(
            "SELECT DISTINCT TI_NAME FROM map_lands "
            "WHERE TI_NAME IS NOT NULL AND TI_NAME != '' ORDER BY TI_NAME"
        ).fetchall()
    ]
    if base_names:
        facets.append(
            {
                "param": "camada_base",
                "column": "NOME_TERRA_INDIGENA",
                "label": "Camada Base",
                "role": "base_layer",
                "options": [{"value": name, "label": name} for name in base_names],
            }
        )

    role_order = {
        "municipio": 0,
        "uf": 1,
        "year": 2,
        "institution": 3,
        "generic": 4,
    }
    filter_specs = sorted(
        [
            col
            for col in columns_list
            if int(col["filterable"]) == 1 and col["semantic_role"] != "base_layer"
        ],
        key=lambda col: (
            role_order.get(str(col["semantic_role"]), 5),
            int(col["col_order"]),
        ),
    )
    for spec in filter_specs:
        col_name = str(spec["col_name"])
        quoted = _quote_ident(col_name, columns_by_name)
        values = [
            str(row[0])
            for row in connection.execute(
                f"SELECT DISTINCT {quoted} FROM dashboard_rows "
                f"WHERE {quoted} IS NOT NULL AND {quoted} != '' "
                f"ORDER BY {quoted} LIMIT 500"
            ).fetchall()
        ]
        if not values:
            continue
        is_inst = spec["semantic_role"] == "institution"
        options = [
            {
                "value": val,
                "label": institutions.get(val, val) if is_inst else val,
            }
            for val in values
        ]
        if is_inst:
            options.sort(key=lambda item: item["label"].casefold())
        facets.append(
            {
                "param": f"f_{col_name}",
                "column": col_name,
                "label": str(spec["label"]),
                "role": str(spec["semantic_role"]),
                "options": options,
            }
        )
    return facets


def _write_excel_from_sqlite(
    connection: sqlite3.Connection,
    export_specs: list[dict[str, Any]],
    columns_by_name: dict[str, dict[str, Any]],
    where: str,
    parameters: list[Any],
    order_clause: str,
    summary: list[dict[str, Any]],
    filter_display: list[tuple[str, str]],
    institutions: dict[str, str],
    destination_path: Path,
    camada_base: str = "",
    meta: dict[str, str] | None = None,
) -> None:
    """Exporta linhas filtradas/completas e detalhamento 1:N direto do SQLite (write_only)."""
    if not export_specs:
        raise ValueError("Nenhuma coluna selecionada para exportação.")
    select_sql = ", ".join(
        _quote_ident(str(spec["col_name"]), columns_by_name) for spec in export_specs
    )
    rows = connection.execute(
        f"SELECT {select_sql} FROM dashboard_rows{where} ORDER BY {order_clause}",
        parameters,
    )
    inst_indexes = [
        idx
        for idx, spec in enumerate(export_specs)
        if spec["semantic_role"] == "institution"
    ]

    workbook = Workbook(write_only=True)
    table_sheet = workbook.create_sheet("Tabela")
    table_sheet.freeze_panes = "A2"
    headers = [str(spec["label"]) for spec in export_specs]
    table_sheet.append(headers)
    written_rows = 0
    for row in rows:
        values = [_excel_cell_value(val) for val in row]
        for idx in inst_indexes:
            raw = values[idx]
            if raw is not None and str(raw).strip():
                values[idx] = institutions.get(str(raw).strip(), str(raw).strip())
        table_sheet.append(values)
        written_rows += 1
    if written_rows > 0:
        table_sheet.auto_filter.ref = (
            f"A1:{get_column_letter(len(headers))}{written_rows + 1}"
        )

    if _sqlite_table_exists(connection, "dashboard_intersections"):
        inter_clauses = [
            f"PLOT_UID IN (SELECT DISTINCT PLOT_UID FROM dashboard_rows{where})"
        ]
        inter_params = list(parameters)
        if camada_base:
            inter_clauses.append("CAMADA_BASE = ?")
            inter_params.append(camada_base)
        inter_where = f" WHERE {' AND '.join(inter_clauses)}"
        inter_total = int(
            connection.execute(
                f"SELECT COUNT(*) FROM dashboard_intersections{inter_where}",
                inter_params,
            ).fetchone()[0]
        )
        if inter_total + 1 > EXCEL_MAX_ROWS:
            raise ValueError(
                f"O detalhamento 1:N possui {inter_total:,} linhas e excede o limite do Excel "
                f"({EXCEL_MAX_ROWS - 1:,} linhas por aba)."
            )
        key_header = str((meta or {}).get("key_column") or "Chave")
        if (meta or {}).get("effective_key_column"):
            key_header = "Referência BACEN explícita"
        inter_sheet = workbook.create_sheet("Intersecoes_Detalhe")
        inter_sheet.freeze_panes = "A2"
        inter_headers = [
            key_header,
            "Arquivo de glebas",
            "Camada Base",
            "Área total (ha)",
            "Área da interseção (ha)",
            "Percentual da interseção (%)",
        ]
        inter_sheet.append(inter_headers)
        inter_rows = connection.execute(
            "SELECT KEY_VALUE, ARQUIVO_GLEBAS, CAMADA_BASE, AREA_TOTAL_HA, "
            "AREA_INTERSECAO_HA, PERCENTUAL_INTERSECAO "
            f"FROM dashboard_intersections{inter_where} "
            "ORDER BY PLOT_UID ASC, AREA_INTERSECAO_HA DESC, INTERSECTION_UID ASC",
            inter_params,
        )
        written_inter = 0
        for irow in inter_rows:
            inter_sheet.append([_excel_cell_value(v) for v in irow])
            written_inter += 1
        if written_inter > 0:
            inter_sheet.auto_filter.ref = (
                f"A1:{get_column_letter(len(inter_headers))}{written_inter + 1}"
            )

    summary_sheet = workbook.create_sheet("Resumo")
    summary_sheet.column_dimensions["A"].width = 48
    summary_sheet.column_dimensions["B"].width = 28
    summary_sheet.append(["Indicador", "Valor"])
    for item in summary:
        summary_sheet.append([item["label"], item["value"]])
    if filter_display:
        summary_sheet.append([])
        summary_sheet.append(["Filtros aplicados", "Valor"])
        for label, display_value in filter_display:
            summary_sheet.append([label, display_value])
    workbook.save(destination_path)


def _select_excel_columns(
    query: dict[str, list[str]],
    columns_list: list[dict[str, Any]],
    columns_by_name: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    scope = query.get("columns", ["visible"])[0].strip().lower()
    if scope == "all":
        return [
            spec for spec in columns_list if spec["semantic_role"] != "plot_uid"
        ]
    raw_visible = query.get("visible_cols", [""])[0].strip()
    if raw_visible:
        requested = [item.strip() for item in raw_visible.split(",") if item.strip()]
        selected = [
            columns_by_name[name]
            for name in requested
            if name in columns_by_name
            and columns_by_name[name]["semantic_role"] != "plot_uid"
        ]
        if selected:
            return selected
    return [
        spec
        for spec in columns_list
        if int(spec["default_visible"]) == 1 and spec["semantic_role"] != "plot_uid"
    ]


def _resolve_dashboard_order_clause(
    query: dict[str, list[str]],
    columns_by_name: dict[str, dict[str, Any]],
) -> tuple[str, str, str]:
    sort_by = query.get("sort_by", ["rowid"])[0]
    sort_direction = query.get("sort_direction", ["asc"])[0].lower()
    if sort_by != "rowid" and sort_by not in columns_by_name:
        sort_by = "rowid"
    if sort_direction not in {"asc", "desc"}:
        sort_direction = "asc"
    if sort_by == "rowid":
        return sort_by, sort_direction, f"rowid {sort_direction.upper()}"
    quoted_sort = _quote_ident(sort_by, columns_by_name)
    sort_type = columns_by_name[sort_by]["data_type"]
    sort_expr = (
        f"{quoted_sort} COLLATE NOCASE"
        if sort_type in {"text", "date"}
        else quoted_sort
    )
    return sort_by, sort_direction, f"{sort_expr} {sort_direction.upper()}, rowid ASC"


def _create_filtered_dashboard_excel(
    job_id: str, query: dict[str, list[str]]
) -> Path:
    output_dir = (JOB_DIR / job_id / "outputs").resolve()
    if not output_dir.is_dir():
        raise FileNotFoundError("Diretório de saída da análise não encontrado.")
    database_path = output_dir / "dashboard.sqlite"
    if not database_path.is_file():
        raise FileNotFoundError("Tabela da análise não encontrada.")

    institutions = _read_financial_institutions()
    fd, temporary_name = tempfile.mkstemp(
        prefix="download-", suffix=".xlsx", dir=output_dir
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
            where, parameters, _, filter_display, camada_base = _build_dashboard_where(
                query, columns_list, columns_by_name, meta, institutions
            )
            total = int(
                connection.execute(
                    f"SELECT COUNT(*) FROM dashboard_rows{where}", parameters
                ).fetchone()[0]
            )
            if total + 1 > EXCEL_MAX_ROWS:
                raise ValueError(
                    f"A tabela filtrada possui {total:,} linhas e excede o limite "
                    f"do Excel ({EXCEL_MAX_ROWS - 1:,} linhas por aba)."
                )
            _, _, order_clause = _resolve_dashboard_order_clause(query, columns_by_name)
            summary = _filtered_dashboard_summary(
                connection, where, parameters, total, columns_list, columns_by_name, meta
            )
            export_specs = _select_excel_columns(query, columns_list, columns_by_name)
            _write_excel_from_sqlite(
                connection=connection,
                export_specs=export_specs,
                columns_by_name=columns_by_name,
                where=where,
                parameters=parameters,
                order_clause=order_clause,
                summary=summary,
                filter_display=filter_display,
                institutions=institutions,
                destination_path=temporary_path,
                camada_base=camada_base,
                meta=meta,
            )
        return temporary_path
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _create_filtered_dashboard_csv(
    job_id: str, query: dict[str, list[str]]
) -> tuple[Path, bool]:
    output_dir = (JOB_DIR / job_id / "outputs").resolve()
    if not output_dir.is_dir():
        raise FileNotFoundError("Diretório de saída da análise não encontrado.")
    database_path = output_dir / "dashboard.sqlite"
    if not database_path.is_file():
        raise FileNotFoundError("Tabela da análise não encontrada.")

    institutions = _read_financial_institutions()
    fd, temporary_name = tempfile.mkstemp(
        prefix="download-", suffix=".csv", dir=output_dir
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
            where, parameters, active_filters, _, _ = _build_dashboard_where(
                query, columns_list, columns_by_name, meta, institutions
            )
            _, _, order_clause = _resolve_dashboard_order_clause(query, columns_by_name)
            export_specs = _select_excel_columns(query, columns_list, columns_by_name)
            if not export_specs:
                raise ValueError("Nenhuma coluna selecionada para exportação.")
            select_sql = ", ".join(
                _quote_ident(str(spec["col_name"]), columns_by_name)
                for spec in export_specs
            )
            rows = connection.execute(
                f"SELECT {select_sql} FROM dashboard_rows{where} ORDER BY {order_clause}",
                parameters,
            )
            inst_indexes = [
                idx
                for idx, spec in enumerate(export_specs)
                if spec["semantic_role"] == "institution"
            ]
            with temporary_path.open("w", encoding="utf-8-sig", newline="") as dest:
                writer = csv.writer(dest, delimiter=";")
                writer.writerow([str(spec["label"]) for spec in export_specs])
                for row in rows:
                    values = [_excel_cell_value(val) for val in row]
                    for idx in inst_indexes:
                        raw = values[idx]
                        if raw is not None and str(raw).strip():
                            values[idx] = institutions.get(
                                str(raw).strip(), str(raw).strip()
                            )
                    writer.writerow(values)
        return temporary_path, bool(active_filters)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _create_filtered_dashboard_parquet(
    job_id: str, query: dict[str, list[str]]
) -> tuple[Path, bool]:
    output_dir = (JOB_DIR / job_id / "outputs").resolve()
    if not output_dir.is_dir():
        raise FileNotFoundError("Diretório de saída da análise não encontrado.")
    database_path = output_dir / "dashboard.sqlite"
    if not database_path.is_file():
        raise FileNotFoundError("Tabela da análise não encontrada.")

    institutions = _read_financial_institutions()
    fd, temporary_name = tempfile.mkstemp(
        prefix="download-", suffix=".parquet", dir=output_dir
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
            where, parameters, active_filters, _, _ = _build_dashboard_where(
                query, columns_list, columns_by_name, meta, institutions
            )
            _, _, order_clause = _resolve_dashboard_order_clause(query, columns_by_name)
            export_specs = _select_excel_columns(query, columns_list, columns_by_name)
            if not export_specs:
                raise ValueError("Nenhuma coluna selecionada para exportação.")
            select_sql = ", ".join(
                _quote_ident(str(spec["col_name"]), columns_by_name)
                for spec in export_specs
            )
            rows = connection.execute(
                f"SELECT {select_sql} FROM dashboard_rows{where} ORDER BY {order_clause}",
                parameters,
            ).fetchall()

        float_types = {
            "area",
            "area_full",
            "percent_full",
            "number",
            "currency",
            "coord",
        }
        series_list: list[pl.Series] = []
        for idx, spec in enumerate(export_specs):
            col_name = str(spec["col_name"])
            dtype = str(spec["data_type"])
            is_inst = spec["semantic_role"] == "institution"
            col_vals = [row[idx] for row in rows]
            if dtype == "integer":
                parsed_ints: list[int | None] = []
                for val in col_vals:
                    if val is None or str(val).strip() == "":
                        parsed_ints.append(None)
                    else:
                        try:
                            parsed_ints.append(int(val))
                        except (TypeError, ValueError):
                            parsed_ints.append(None)
                series_list.append(pl.Series(col_name, parsed_ints, dtype=pl.Int64))
            elif dtype in float_types:
                parsed_floats: list[float | None] = []
                for val in col_vals:
                    if val is None or str(val).strip() == "":
                        parsed_floats.append(None)
                    else:
                        try:
                            parsed_floats.append(float(val))
                        except (TypeError, ValueError):
                            parsed_floats.append(None)
                series_list.append(pl.Series(col_name, parsed_floats, dtype=pl.Float64))
            else:
                parsed_strs: list[str | None] = []
                for val in col_vals:
                    if val is None:
                        parsed_strs.append(None)
                    else:
                        text = str(val).strip()
                        if not text:
                            parsed_strs.append(None)
                        elif is_inst:
                            parsed_strs.append(institutions.get(text, text))
                        else:
                            parsed_strs.append(text)
                series_list.append(pl.Series(col_name, parsed_strs, dtype=pl.String))

        pl.DataFrame(series_list).write_parquet(temporary_path, compression="zstd")
        return temporary_path, bool(active_filters)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _extract_filtered_geodataframe(
    connection: sqlite3.Connection,
    query: dict[str, list[str]],
    institutions: dict[str, str],
    include_intersections: bool = False,
) -> tuple[gpd.GeoDataFrame, gpd.GeoDataFrame | None, str, bool]:
    columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
    where, parameters, active_filters, _, camada_base = _build_dashboard_where(
        query, columns_list, columns_by_name, meta, institutions
    )
    area_crs = meta.get("area_crs") or "EPSG:4674"
    attr_specs = [
        spec for spec in columns_list if spec["semantic_role"] != "plot_uid"
    ]
    sub_cols = ", ".join(
        [
            "PLOT_UID",
            *(
                _quote_ident(str(spec["col_name"]), columns_by_name)
                for spec in attr_specs
            ),
        ]
    )
    outer_cols = ", ".join(
        f"d.{_quote_ident(str(spec['col_name']), columns_by_name)}"
        for spec in attr_specs
    )
    rows = connection.execute(
        f"SELECT {outer_cols}, mp.WKB AS __WKB "
        f"FROM (SELECT rowid AS __ROWID, {sub_cols} FROM dashboard_rows{where}) d "
        "JOIN map_plots mp ON mp.PLOT_UID = d.PLOT_UID "
        "ORDER BY d.__ROWID",
        parameters,
    ).fetchall()
    if not rows:
        raise ValueError("Nenhum registro no filtro atual para exportação geoespacial.")

    col_names = [str(spec["col_name"]) for spec in attr_specs]
    records: list[dict[str, Any]] = []
    geometries = []
    for row in rows:
        records.append({col: row[col] for col in col_names})
        geometries.append(shapely.from_wkb(row["__WKB"]))
    frame = gpd.GeoDataFrame(records, geometry=geometries, crs=area_crs)

    inter_gdf: gpd.GeoDataFrame | None = None
    if (
        include_intersections
        and _sqlite_table_exists(connection, "dashboard_intersections")
        and _sqlite_table_exists(connection, "map_intersections")
    ):
        filtered_plots = f"SELECT DISTINCT PLOT_UID FROM dashboard_rows{where}"
        inter_clauses = [f"di.PLOT_UID IN ({filtered_plots})"]
        inter_params = list(parameters)
        if camada_base:
            inter_clauses.append("di.CAMADA_BASE = ?")
            inter_params.append(camada_base)
        inter_rows = connection.execute(
            "SELECT di.KEY_VALUE, di.ARQUIVO_GLEBAS, di.CAMADA_BASE, "
            "di.AREA_TOTAL_HA, di.AREA_INTERSECAO_HA, di.PERCENTUAL_INTERSECAO, "
            "mi.WKB AS __WKB "
            "FROM dashboard_intersections di "
            "JOIN map_intersections mi ON mi.INTERSECTION_UID = di.INTERSECTION_UID "
            f"WHERE {' AND '.join(inter_clauses)} "
            "ORDER BY di.PLOT_UID ASC, di.AREA_INTERSECAO_HA DESC, di.INTERSECTION_UID ASC",
            inter_params,
        ).fetchall()
        if inter_rows:
            key_col_name = meta.get("key_column") or "REF_BACEN"
            inter_records = []
            inter_geoms = []
            for irow in inter_rows:
                inter_records.append(
                    {
                        key_col_name: irow["KEY_VALUE"],
                        "ARQUIVO_GLEBAS": irow["ARQUIVO_GLEBAS"],
                        "CAMADA_BASE": irow["CAMADA_BASE"],
                        "AREA_TOTAL_HA": irow["AREA_TOTAL_HA"],
                        "AREA_INTERSECAO_HA": irow["AREA_INTERSECAO_HA"],
                        "PERCENTUAL_INTERSECAO": irow["PERCENTUAL_INTERSECAO"],
                    }
                )
                inter_geoms.append(shapely.from_wkb(irow["__WKB"]))
            inter_gdf = gpd.GeoDataFrame(
                inter_records, geometry=inter_geoms, crs=area_crs
            )

    return frame, inter_gdf, area_crs, bool(active_filters)


def _create_filtered_dashboard_geopackage(
    job_id: str, query: dict[str, list[str]]
) -> tuple[Path, bool]:
    output_dir = (JOB_DIR / job_id / "outputs").resolve()
    if not output_dir.is_dir():
        raise FileNotFoundError("Diretório de saída da análise não encontrado.")
    database_path = output_dir / "dashboard.sqlite"
    if not database_path.is_file():
        raise FileNotFoundError("Tabela da análise não encontrada.")

    institutions = _read_financial_institutions()
    fd, temporary_name = tempfile.mkstemp(
        prefix="download-", suffix=".gpkg", dir=output_dir
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            frame, inter_gdf, area_crs, is_filtered = _extract_filtered_geodataframe(
                connection, query, institutions, include_intersections=True
            )
        write_geopackage(temporary_path, frame, area_crs, intersections=inter_gdf)
        return temporary_path, is_filtered
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _create_filtered_dashboard_geoparquet(
    job_id: str, query: dict[str, list[str]]
) -> tuple[Path, bool]:
    output_dir = (JOB_DIR / job_id / "outputs").resolve()
    if not output_dir.is_dir():
        raise FileNotFoundError("Diretório de saída da análise não encontrado.")
    database_path = output_dir / "dashboard.sqlite"
    if not database_path.is_file():
        raise FileNotFoundError("Tabela da análise não encontrada.")

    institutions = _read_financial_institutions()
    fd, temporary_name = tempfile.mkstemp(
        prefix="download-geo-", suffix=".parquet", dir=output_dir
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            frame, _, area_crs, is_filtered = _extract_filtered_geodataframe(
                connection, query, institutions, include_intersections=False
            )
        write_geoparquet(temporary_path, frame, area_crs)
        return temporary_path, is_filtered
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _generate_full_excel_background(job_id: str, output_dir: Path) -> None:
    """Gera o Excel completo com todas as colunas em segundo plano após liberar o dashboard."""
    database_path = output_dir / "dashboard.sqlite"
    final_path = output_dir / "resultado_completo.xlsx"
    temp_path = output_dir / "resultado_completo.tmp.xlsx"
    try:
        institutions = _read_financial_institutions()
        with closing(sqlite3.connect(database_path)) as connection:
            connection.row_factory = sqlite3.Row
            columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
            total = int(
                connection.execute("SELECT COUNT(*) FROM dashboard_rows").fetchone()[0]
            )
            if total + 1 > EXCEL_MAX_ROWS:
                raise ValueError(
                    f"O resultado possui {total:,} linhas e excede o limite do Excel "
                    f"({EXCEL_MAX_ROWS - 1:,} linhas por aba). Use filtros ou baixe em GeoPackage/GeoParquet."
                )
            summary = _filtered_dashboard_summary(
                connection, "", [], total, columns_list, columns_by_name, meta
            )
            export_specs = [
                spec for spec in columns_list if spec["semantic_role"] != "plot_uid"
            ]
            _write_excel_from_sqlite(
                connection=connection,
                export_specs=export_specs,
                columns_by_name=columns_by_name,
                where="",
                parameters=[],
                order_clause="rowid ASC",
                summary=summary,
                filter_display=[],
                institutions=institutions,
                destination_path=temp_path,
                camada_base="",
                meta=meta,
            )
        temp_path.replace(final_path)
        with _jobs_lock:
            job = _jobs.get(job_id)
            if job is not None:
                job["full_excel_status"] = "ready"
                job["full_excel_error"] = None
                job["full_excel_done"].set()
        LOGGER.info("Excel completo gerado em background para análise %s.", job_id)
    except Exception as error:
        temp_path.unlink(missing_ok=True)
        LOGGER.warning("Não foi possível gerar o Excel completo da análise %s: %s", job_id, error)
        with _jobs_lock:
            job = _jobs.get(job_id)
            if job is not None:
                job["full_excel_status"] = "failed"
                job["full_excel_error"] = str(error)
                job["full_excel_done"].set()


def _execute_job(job_id: str, args: SimpleNamespace, job_root: Path) -> None:
    global _active_job_id
    handler = JobLogHandler(job_id)
    analysis_logger = logging.getLogger("analise_sobreposicao")
    analysis_logger.addHandler(handler)
    succeeded = False
    try:
        with _jobs_lock:
            _jobs[job_id]["status"] = "running"
            _jobs[job_id]["message"] = "Análise geoespacial em execução."
        run(args)
        succeeded = True
        with _jobs_lock:
            _jobs[job_id]["status"] = "completed"
            _jobs[job_id]["message"] = "Análise concluída com sucesso."
            _jobs[job_id]["full_excel_status"] = "generating"
    except Exception as error:
        LOGGER.exception("Falha ao executar a análise %s.", job_id)
        with _jobs_lock:
            job = _jobs.get(job_id)
            if job is not None:
                job["status"] = "failed"
                job["message"] = str(error)
                job["full_excel_status"] = "failed"
                job["full_excel_error"] = str(error)
                job["full_excel_done"].set()
    finally:
        analysis_logger.removeHandler(handler)
        handler.close()
        shutil.rmtree(job_root / "uploads", ignore_errors=True)
        with _active_job_lock:
            if _active_job_id == job_id:
                _active_job_id = None

    if succeeded:
        bg_worker = threading.Thread(
            target=_generate_full_excel_background,
            args=(job_id, job_root / "outputs"),
            daemon=True,
        )
        bg_worker.start()


class RequestHandler(BaseHTTPRequestHandler):
    server_version = "GeoAnalysis/2.0"

    def log_message(self, format_string: str, *args: Any) -> None:
        LOGGER.info("%s - %s", self.address_string(), format_string % args)

    def _allowed_hosts(self) -> set[str]:
        port = getattr(self.server, "server_port", 8000)
        return {f"127.0.0.1:{port}", f"localhost:{port}"}

    def _validate_local_request(self) -> None:
        allowed_hosts = self._allowed_hosts()
        host = self.headers.get("Host", "").lower()
        if host not in allowed_hosts:
            raise ValueError("Cabeçalho Host não permitido.")
        origin = self.headers.get("Origin")
        if origin and origin.lower() not in {f"http://{h}" for h in allowed_hosts}:
            raise PermissionError("Origem da requisição não permitida.")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            _send_json(self, 200, {"status": "ok"})
            return
        job_match = re.fullmatch(r"/api/jobs/([a-f0-9]{32})", parsed.path)
        if job_match:
            with _jobs_lock:
                job = _jobs.get(job_match.group(1))
                if job is None:
                    _send_json(self, 404, {"detail": "Análise não encontrada."})
                    return
                response = {
                    "job_id": job_match.group(1),
                    "status": job["status"],
                    "message": job["message"],
                    "logs": list(job["logs"]),
                    "outputs": dict(job["outputs"]),
                    "full_excel_status": job["full_excel_status"],
                    "full_excel_error": job["full_excel_error"],
                }
            _send_json(self, 200, response)
            return
        plot_match = re.fullmatch(r"/api/jobs/([a-f0-9]{32})/plot/(\d+)", parsed.path)
        if plot_match:
            self._plot_detail(plot_match.group(1), plot_match.group(2))
            return
        map_match = re.fullmatch(r"/api/jobs/([a-f0-9]{32})/map", parsed.path)
        if map_match:
            self._map(map_match.group(1), parse_qs(parsed.query))
            return
        intersections_match = re.fullmatch(
            r"/api/jobs/([a-f0-9]{32})/intersections", parsed.path
        )
        if intersections_match:
            self._intersections(intersections_match.group(1), parse_qs(parsed.query))
            return
        dashboard_match = re.fullmatch(
            r"/api/jobs/([a-f0-9]{32})/dashboard", parsed.path
        )
        if dashboard_match:
            self._dashboard(dashboard_match.group(1), parse_qs(parsed.query))
            return
        download_match = re.fullmatch(
            r"/api/jobs/([a-f0-9]{32})/download/(xlsx|xlsx-full|csv|parquet|geoparquet|gpkg)",
            parsed.path,
        )
        if download_match:
            self._download(download_match.group(1), download_match.group(2))
            return
        static_path = STATIC_FILES.get(parsed.path)
        if static_path is None and parsed.path.startswith(("/fonts/", "/image/")):
            root = BASE_DIR / parsed.path.strip("/").split("/", 1)[0]
            relative = parsed.path.split("/", 2)[-1]
            candidate = (root / relative).resolve()
            if root.resolve() in candidate.parents and candidate.is_file():
                media_type = mimetypes.guess_type(candidate.name)[0] or "application/octet-stream"
                static_path = (candidate, media_type)
        if static_path is None or not static_path[0].is_file():
            _send_json(self, 404, {"detail": "Recurso não encontrado."})
            return
        file_path, media_type = static_path
        try:
            content_length = file_path.stat().st_size
            self.send_response(200)
            self.send_header("Content-Type", media_type)
            self.send_header("Content-Length", str(content_length))
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            with file_path.open("rb") as source:
                shutil.copyfileobj(source, self.wfile, length=UPLOAD_CHUNK_BYTES)
        except (BrokenPipeError, ConnectionResetError):
            LOGGER.info("Cliente encerrou a transferência de %s.", file_path.name)

    def _download(self, job_id: str, output_type: str) -> None:
        with _jobs_lock:
            job = _jobs.get(job_id)
            available = job is not None and job["status"] == "completed"
            full_done = job["full_excel_done"] if job is not None else None
        if not available:
            _send_json(self, 404, {"detail": "Análise ainda não concluída."})
            return

        output_dir = (JOB_DIR / job_id / "outputs").resolve()
        temporary_path: Path | None = None
        query = parse_qs(urlparse(self.path).query)

        if output_type == "xlsx":
            try:
                temporary_path = _create_filtered_dashboard_excel(job_id, query)
            except (FileNotFoundError, OSError, sqlite3.Error, ValueError, UnicodeError) as error:
                LOGGER.exception("Falha ao gerar a planilha filtrada da análise %s.", job_id)
                status = 400 if isinstance(error, ValueError) else 500
                _send_json(self, status, {"detail": f"Não foi possível gerar a planilha: {error}"})
                return
            file_path = temporary_path
            download_filename = "resultado_filtrado.xlsx"
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif output_type == "xlsx-full":
            if full_done is not None and not full_done.is_set():
                full_done.wait(timeout=300)
            with _jobs_lock:
                job = _jobs.get(job_id)
                full_status = job["full_excel_status"] if job is not None else "failed"
                full_error = job["full_excel_error"] if job is not None else "Análise indisponível."
            file_path = output_dir / "resultado_completo.xlsx"
            if full_status != "ready" or not file_path.is_file():
                _send_json(
                    self,
                    409 if full_status == "generating" else 400,
                    {"detail": full_error or "O Excel completo ainda não está disponível."},
                )
                return
            download_filename = "resultado_completo.xlsx"
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif output_type == "csv":
            try:
                temporary_path, is_filtered = _create_filtered_dashboard_csv(job_id, query)
            except (FileNotFoundError, OSError, sqlite3.Error, ValueError, UnicodeError) as error:
                LOGGER.exception("Falha ao gerar CSV da análise %s.", job_id)
                status = 400 if isinstance(error, ValueError) else 500
                _send_json(self, status, {"detail": f"Não foi possível gerar o CSV: {error}"})
                return
            file_path = temporary_path
            download_filename = (
                "resultado_filtrado.csv" if is_filtered else "resultado_completo.csv"
            )
            media_type = "text/csv; charset=utf-8"
        elif output_type == "parquet":
            try:
                temporary_path, is_filtered = _create_filtered_dashboard_parquet(job_id, query)
            except (FileNotFoundError, OSError, sqlite3.Error, ValueError) as error:
                LOGGER.exception("Falha ao gerar Parquet da análise %s.", job_id)
                status = 400 if isinstance(error, ValueError) else 500
                _send_json(self, status, {"detail": f"Não foi possível gerar o Parquet: {error}"})
                return
            file_path = temporary_path
            download_filename = (
                "resultado_filtrado.parquet" if is_filtered else "resultado_completo.parquet"
            )
            media_type = "application/vnd.apache.parquet"
        elif output_type == "geoparquet":
            try:
                temporary_path, is_filtered = _create_filtered_dashboard_geoparquet(job_id, query)
            except (FileNotFoundError, OSError, sqlite3.Error, ValueError) as error:
                LOGGER.exception("Falha ao gerar GeoParquet da análise %s.", job_id)
                status = 400 if isinstance(error, ValueError) else 500
                _send_json(self, status, {"detail": f"Não foi possível gerar o GeoParquet: {error}"})
                return
            file_path = temporary_path
            download_filename = (
                "glebas_sobrepostas_filtrado.parquet" if is_filtered else "glebas_sobrepostas.parquet"
            )
            media_type = "application/vnd.apache.parquet"
        else:
            try:
                temporary_path, is_filtered = _create_filtered_dashboard_geopackage(job_id, query)
            except (FileNotFoundError, OSError, sqlite3.Error, ValueError) as error:
                LOGGER.exception("Falha ao gerar o GeoPackage da análise %s.", job_id)
                status = 400 if isinstance(error, ValueError) else 500
                _send_json(self, status, {"detail": f"Não foi possível gerar o GeoPackage: {error}"})
                return
            file_path = temporary_path
            download_filename = (
                "glebas_sobrepostas_filtrado.gpkg" if is_filtered else "glebas_sobrepostas.gpkg"
            )
            media_type = "application/geopackage+sqlite3"

        try:
            self.send_response(200)
            self.send_header("Content-Type", media_type)
            self.send_header("Content-Length", str(file_path.stat().st_size))
            self.send_header(
                "Content-Disposition",
                f'attachment; filename="{download_filename.replace(chr(34), "")}"',
            )
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            with file_path.open("rb") as source:
                shutil.copyfileobj(source, self.wfile, length=UPLOAD_CHUNK_BYTES)
        except (BrokenPipeError, ConnectionResetError):
            LOGGER.info("Cliente encerrou o download de %s.", download_filename)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)

    def _plot_detail(self, job_id: str, plot_uid_text: str) -> None:
        with _jobs_lock:
            job = _jobs.get(job_id)
            available = job is not None and job["status"] == "completed"
        database_path = (JOB_DIR / job_id / "outputs" / "dashboard.sqlite").resolve()
        expected_parent = (JOB_DIR / job_id / "outputs").resolve()
        if (
            not available
            or database_path.parent != expected_parent
            or not database_path.is_file()
        ):
            _send_json(self, 404, {"detail": "Geometrias da análise ainda não disponíveis."})
            return

        plot_uid = int(plot_uid_text)
        try:
            with closing(sqlite3.connect(database_path)) as connection:
                plot_row = connection.execute(
                    "SELECT GEOJSON FROM map_plots WHERE PLOT_UID = ?", [plot_uid]
                ).fetchone()
                if plot_row is None:
                    _send_json(self, 404, {"detail": "Gleba não encontrada."})
                    return
                land_rows = connection.execute(
                    "SELECT ml.GEOJSON FROM map_lands ml "
                    "JOIN map_plot_lands mpl ON mpl.LAND_UID = ml.LAND_UID "
                    "WHERE mpl.PLOT_UID = ? ORDER BY ml.LAND_UID",
                    [plot_uid],
                ).fetchall()
                inter_rows = (
                    connection.execute(
                        "SELECT GEOJSON FROM map_intersections "
                        "WHERE PLOT_UID = ? ORDER BY INTERSECTION_UID",
                        [plot_uid],
                    ).fetchall()
                    if _sqlite_table_exists(connection, "map_intersections")
                    else []
                )
        except (OSError, sqlite3.Error) as error:
            LOGGER.exception("Falha ao buscar detalhe da gleba %s (%s).", plot_uid, job_id)
            _send_json(self, 500, {"detail": f"Não foi possível carregar a gleba: {error}"})
            return

        features_json = ",".join(
            [*(row[0] for row in land_rows), plot_row[0], *(row[0] for row in inter_rows)]
        )
        payload_str = f'{{"type":"FeatureCollection","features":[{features_json}]}}'
        _send_bytes(self, 200, payload_str.encode("utf-8"), "application/json; charset=utf-8")

    def _map(self, job_id: str, query: dict[str, list[str]]) -> None:
        with _jobs_lock:
            job = _jobs.get(job_id)
            available = job is not None and job["status"] == "completed"
        database_path = (JOB_DIR / job_id / "outputs" / "dashboard.sqlite").resolve()
        expected_parent = (JOB_DIR / job_id / "outputs").resolve()
        if (
            not available
            or database_path.parent != expected_parent
            or not database_path.is_file()
        ):
            _send_json(self, 404, {"detail": "Geometrias da análise ainda não disponíveis."})
            return

        focus_plot_uid = query.get("focus_plot_uid", [""])[0].strip()
        try:
            with closing(sqlite3.connect(database_path)) as connection:
                connection.row_factory = sqlite3.Row
                columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
                institutions = (
                    _read_financial_institutions()
                    if meta.get("institution_column")
                    else {}
                )
                where, parameters, _, _, camada_base = _build_dashboard_where(
                    query, columns_list, columns_by_name, meta, institutions
                )
                filtered_plots = (
                    f"SELECT DISTINCT PLOT_UID FROM dashboard_rows{where}"
                    f"{' AND' if where else ' WHERE'} PLOT_UID IS NOT NULL"
                )
                plot_total = int(
                    connection.execute(
                        f"SELECT COUNT(*) FROM ({filtered_plots})", parameters
                    ).fetchone()[0]
                )
                if camada_base:
                    land_total = int(
                        connection.execute(
                            "SELECT COUNT(DISTINCT mpl.LAND_UID) FROM map_plot_lands mpl "
                            "JOIN map_lands ml ON ml.LAND_UID = mpl.LAND_UID "
                            f"WHERE ml.TI_NAME = ? AND mpl.PLOT_UID IN ({filtered_plots})",
                            [camada_base, *parameters],
                        ).fetchone()[0]
                    )
                    land_rows = connection.execute(
                        "SELECT ml.GEOJSON FROM map_lands ml "
                        "WHERE ml.TI_NAME = ? AND ml.LAND_UID IN ("
                        f"SELECT DISTINCT LAND_UID FROM map_plot_lands WHERE PLOT_UID IN ({filtered_plots})"
                        ") ORDER BY ml.LAND_UID LIMIT ?",
                        [camada_base, *parameters, MAP_FEATURE_LIMIT],
                    ).fetchall()
                else:
                    land_total = int(
                        connection.execute(
                            "SELECT COUNT(DISTINCT LAND_UID) FROM map_plot_lands "
                            f"WHERE PLOT_UID IN ({filtered_plots})",
                            parameters,
                        ).fetchone()[0]
                    )
                    land_rows = connection.execute(
                        "SELECT GEOJSON FROM map_lands "
                        "WHERE LAND_UID IN (SELECT DISTINCT LAND_UID FROM map_plot_lands "
                        f"WHERE PLOT_UID IN ({filtered_plots})) "
                        "ORDER BY LAND_UID LIMIT ?",
                        [*parameters, MAP_FEATURE_LIMIT],
                    ).fetchall()

                plot_rows = connection.execute(
                    "SELECT PLOT_UID, GEOJSON FROM map_plots "
                    f"WHERE PLOT_UID IN ({filtered_plots}) ORDER BY PLOT_UID LIMIT ?",
                    [*parameters, MAP_FEATURE_LIMIT],
                ).fetchall()
                if focus_plot_uid:
                    focused_plot = connection.execute(
                        "SELECT PLOT_UID, GEOJSON FROM map_plots "
                        f"WHERE PLOT_UID = ? AND PLOT_UID IN ({filtered_plots})",
                        [focus_plot_uid, *parameters],
                    ).fetchone()
                    if focused_plot and all(
                        row["PLOT_UID"] != focused_plot["PLOT_UID"]
                        for row in plot_rows
                    ):
                        plot_rows.append(focused_plot)

                inter_total = 0
                inter_rows: list[Any] = []
                if _sqlite_table_exists(connection, "map_intersections"):
                    if camada_base:
                        inter_total = int(
                            connection.execute(
                                "SELECT COUNT(*) FROM map_intersections "
                                f"WHERE TI_NAME = ? AND PLOT_UID IN ({filtered_plots})",
                                [camada_base, *parameters],
                            ).fetchone()[0]
                        )
                        inter_rows = connection.execute(
                            "SELECT INTERSECTION_UID, GEOJSON FROM map_intersections "
                            f"WHERE TI_NAME = ? AND PLOT_UID IN ({filtered_plots}) "
                            "ORDER BY INTERSECTION_UID LIMIT ?",
                            [camada_base, *parameters, MAP_FEATURE_LIMIT],
                        ).fetchall()
                    else:
                        inter_total = int(
                            connection.execute(
                                "SELECT COUNT(*) FROM map_intersections "
                                f"WHERE PLOT_UID IN ({filtered_plots})",
                                parameters,
                            ).fetchone()[0]
                        )
                        inter_rows = connection.execute(
                            "SELECT INTERSECTION_UID, GEOJSON FROM map_intersections "
                            f"WHERE PLOT_UID IN ({filtered_plots}) "
                            "ORDER BY INTERSECTION_UID LIMIT ?",
                            [*parameters, MAP_FEATURE_LIMIT],
                        ).fetchall()
                    if focus_plot_uid:
                        focus_inter_sql = (
                            "SELECT INTERSECTION_UID, GEOJSON FROM map_intersections "
                            f"WHERE PLOT_UID = ? AND PLOT_UID IN ({filtered_plots})"
                            + (" AND TI_NAME = ?" if camada_base else "")
                            + " ORDER BY INTERSECTION_UID"
                        )
                        focus_inter_params = (
                            [focus_plot_uid, *parameters, camada_base]
                            if camada_base
                            else [focus_plot_uid, *parameters]
                        )
                        existing_inter_uids = {
                            row["INTERSECTION_UID"] for row in inter_rows
                        }
                        for frow in connection.execute(
                            focus_inter_sql, focus_inter_params
                        ).fetchall():
                            if frow["INTERSECTION_UID"] not in existing_inter_uids:
                                inter_rows.append(frow)
        except ValueError as error:
            _send_json(self, 400, {"detail": str(error)})
            return
        except (OSError, sqlite3.Error) as error:
            LOGGER.exception("Falha ao carregar o mapa da análise %s.", job_id)
            _send_json(self, 500, {"detail": f"Não foi possível carregar o mapa: {error}"})
            return

        features_json = ",".join(
            [
                *(row["GEOJSON"] for row in land_rows),
                *(row["GEOJSON"] for row in plot_rows),
                *(row["GEOJSON"] for row in inter_rows),
            ]
        )
        counts_json = json.dumps(
            {
                "glebas_total": plot_total,
                "glebas_preview": len(plot_rows),
                "camada_base_total": land_total,
                "camada_base_preview": len(land_rows),
                "intersecoes_total": inter_total,
                "intersecoes_preview": len(inter_rows),
            },
            ensure_ascii=False,
        )
        payload_str = (
            f'{{"type":"FeatureCollection","features":[{features_json}],"counts":{counts_json}}}'
        )
        _send_bytes(
            self, 200, payload_str.encode("utf-8"), "application/json; charset=utf-8"
        )

    def _intersections(self, job_id: str, query: dict[str, list[str]]) -> None:
        with _jobs_lock:
            job = _jobs.get(job_id)
            available = job is not None and job["status"] == "completed"
        database_path = (JOB_DIR / job_id / "outputs" / "dashboard.sqlite").resolve()
        expected_parent = (JOB_DIR / job_id / "outputs").resolve()
        if (
            not available
            or database_path.parent != expected_parent
            or not database_path.is_file()
        ):
            _send_json(self, 404, {"detail": "Tabela da análise ainda não disponível."})
            return

        try:
            page = int(query.get("page", ["1"])[0])
            page_size = int(query.get("page_size", ["50"])[0])
        except ValueError:
            _send_json(
                self, 400, {"detail": "A página e o tamanho devem ser números inteiros."}
            )
            return
        if page < 1 or page_size not in {25, 50, 100}:
            _send_json(self, 400, {"detail": "Página inválida ou tamanho não permitido."})
            return

        allowed_sort = {
            "rowid",
            "INTERSECTION_UID",
            "KEY_VALUE",
            "ARQUIVO_GLEBAS",
            "CAMADA_BASE",
            "AREA_TOTAL_HA",
            "AREA_INTERSECAO_HA",
            "PERCENTUAL_INTERSECAO",
        }
        sort_by = query.get("sort_by", ["INTERSECTION_UID"])[0]
        sort_direction = query.get("sort_direction", ["asc"])[0].lower()
        if sort_by not in allowed_sort or sort_direction not in {"asc", "desc"}:
            _send_json(self, 400, {"detail": "Ordenação inválida para interseções."})
            return

        focus_plot_uid_raw = query.get("focus_plot_uid", [""])[0].strip()
        try:
            focus_plot_uid = int(focus_plot_uid_raw) if focus_plot_uid_raw else None
        except ValueError:
            _send_json(self, 400, {"detail": "Identificador de gleba inválido."})
            return

        try:
            with closing(sqlite3.connect(database_path)) as connection:
                connection.row_factory = sqlite3.Row
                columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
                inst_col = meta.get("institution_column", "")
                institutions = (
                    _read_financial_institutions()
                    if inst_col and inst_col in columns_by_name
                    else {}
                )
                where, parameters, _, _, camada_base = _build_dashboard_where(
                    query, columns_list, columns_by_name, meta, institutions
                )
                dash_total = int(
                    connection.execute(
                        f"SELECT COUNT(*) FROM dashboard_rows{where}", parameters
                    ).fetchone()[0]
                )
                summary = _filtered_dashboard_summary(
                    connection, where, parameters, dash_total, columns_list, columns_by_name, meta
                )
                key_label = (
                    "Referência BACEN explícita"
                    if meta.get("effective_key_column")
                    else str(meta.get("key_column") or "Chave")
                )
                if not _sqlite_table_exists(connection, "dashboard_intersections"):
                    _send_json(
                        self,
                        200,
                        {
                            "rows": [],
                            "total": 0,
                            "page": 1,
                            "page_size": page_size,
                            "sort_by": sort_by,
                            "sort_direction": sort_direction,
                            "key_label": key_label,
                            "summary": summary,
                        },
                    )
                    return

                inter_clauses = [
                    f"PLOT_UID IN (SELECT DISTINCT PLOT_UID FROM dashboard_rows{where})"
                ]
                inter_params = list(parameters)
                if camada_base:
                    inter_clauses.append("CAMADA_BASE = ?")
                    inter_params.append(camada_base)
                inter_where = f" WHERE {' AND '.join(inter_clauses)}"

                if sort_by in {"rowid", "INTERSECTION_UID"}:
                    order_clause = f"INTERSECTION_UID {sort_direction.upper()}"
                elif sort_by in {"KEY_VALUE", "ARQUIVO_GLEBAS", "CAMADA_BASE"}:
                    order_clause = (
                        f"{sort_by} COLLATE NOCASE {sort_direction.upper()}, "
                        "INTERSECTION_UID ASC"
                    )
                else:
                    order_clause = f"{sort_by} {sort_direction.upper()}, INTERSECTION_UID ASC"

                total = int(
                    connection.execute(
                        f"SELECT COUNT(*) FROM dashboard_intersections{inter_where}",
                        inter_params,
                    ).fetchone()[0]
                )
                summary.insert(
                    2,
                    {
                        "label": "Interseções 1:N no filtro",
                        "value": int(total),
                        "format": "integer",
                    },
                )
                if focus_plot_uid is not None:
                    focused_page = connection.execute(
                        "SELECT page_number FROM ("
                        "SELECT PLOT_UID, "
                        f"CAST((ROW_NUMBER() OVER (ORDER BY {order_clause}) - 1) / ? AS INTEGER) + 1 AS page_number "
                        f"FROM dashboard_intersections{inter_where}) "
                        "WHERE PLOT_UID = ? LIMIT 1",
                        [page_size, *inter_params, focus_plot_uid],
                    ).fetchone()
                    if focused_page is not None:
                        page = int(focused_page["page_number"])

                rows = connection.execute(
                    "SELECT INTERSECTION_UID, PLOT_UID, LAND_UID, KEY_VALUE, ARQUIVO_GLEBAS, "
                    "CAMADA_BASE, AREA_TOTAL_HA, AREA_INTERSECAO_HA, PERCENTUAL_INTERSECAO "
                    f"FROM dashboard_intersections{inter_where} "
                    f"ORDER BY {order_clause} LIMIT ? OFFSET ?",
                    [*inter_params, page_size, (page - 1) * page_size],
                ).fetchall()
        except ValueError as error:
            _send_json(self, 400, {"detail": str(error)})
            return
        except (OSError, sqlite3.Error, UnicodeError) as error:
            LOGGER.exception("Falha ao carregar interseções 1:N da análise %s.", job_id)
            _send_json(
                self, 500, {"detail": f"Não foi possível carregar as interseções: {error}"}
            )
            return

        _send_json(
            self,
            200,
            {
                "rows": [dict(row) for row in rows],
                "total": total,
                "page": page,
                "page_size": page_size,
                "sort_by": sort_by,
                "sort_direction": sort_direction,
                "key_label": key_label,
                "summary": summary,
            },
        )

    def _dashboard(self, job_id: str, query: dict[str, list[str]]) -> None:
        with _jobs_lock:
            job = _jobs.get(job_id)
            available = job is not None and job["status"] == "completed"
        database_path = (JOB_DIR / job_id / "outputs" / "dashboard.sqlite").resolve()
        expected_parent = (JOB_DIR / job_id / "outputs").resolve()
        if (
            not available
            or database_path.parent != expected_parent
            or not database_path.is_file()
        ):
            _send_json(self, 404, {"detail": "Tabela da análise ainda não disponível."})
            return

        try:
            page = int(query.get("page", ["1"])[0])
            page_size = int(query.get("page_size", ["50"])[0])
        except ValueError:
            _send_json(
                self, 400, {"detail": "A página e o tamanho devem ser números inteiros."}
            )
            return
        if page < 1 or page_size not in {25, 50, 100}:
            _send_json(self, 400, {"detail": "Página inválida ou tamanho não permitido."})
            return

        sort_by = query.get("sort_by", ["rowid"])[0]
        sort_direction = query.get("sort_direction", ["asc"])[0].lower()
        if sort_direction not in {"asc", "desc"}:
            _send_json(self, 400, {"detail": "Ordenação inválida."})
            return

        focus_plot_uid_raw = query.get("focus_plot_uid", [""])[0].strip()
        try:
            focus_plot_uid = int(focus_plot_uid_raw) if focus_plot_uid_raw else None
        except ValueError:
            _send_json(self, 400, {"detail": "Identificador de gleba inválido."})
            return

        include_filters = query.get("include_filters", ["1"])[0] != "0"

        try:
            with closing(sqlite3.connect(database_path)) as connection:
                connection.row_factory = sqlite3.Row
                columns_list, columns_by_name, meta = _load_dashboard_schema(connection)
                if sort_by != "rowid" and sort_by not in columns_by_name:
                    _send_json(self, 400, {"detail": "Coluna de ordenação inválida."})
                    return

                if sort_by == "rowid":
                    order_clause = f"rowid {sort_direction.upper()}"
                else:
                    quoted_sort = _quote_ident(sort_by, columns_by_name)
                    sort_type = columns_by_name[sort_by]["data_type"]
                    sort_expression = (
                        f"{quoted_sort} COLLATE NOCASE"
                        if sort_type in {"text", "date"}
                        else quoted_sort
                    )
                    order_clause = f"{sort_expression} {sort_direction.upper()}, rowid ASC"

                inst_col = meta.get("institution_column", "")
                institutions = (
                    _read_financial_institutions()
                    if inst_col and inst_col in columns_by_name
                    else {}
                )
                where, parameters, active_filters, _, _ = _build_dashboard_where(
                    query, columns_list, columns_by_name, meta, institutions
                )

                total = int(
                    connection.execute(
                        f"SELECT COUNT(*) FROM dashboard_rows{where}", parameters
                    ).fetchone()[0]
                )
                if focus_plot_uid is not None:
                    focused_page = connection.execute(
                        "SELECT page_number FROM ("
                        "SELECT PLOT_UID, "
                        "CAST((ROW_NUMBER() OVER (ORDER BY "
                        f"{order_clause}) - 1) / ? AS INTEGER) + 1 AS page_number "
                        f"FROM dashboard_rows{where}) "
                        "WHERE PLOT_UID = ? LIMIT 1",
                        [page_size, *parameters, focus_plot_uid],
                    ).fetchone()
                    if focused_page is not None:
                        page = int(focused_page["page_number"])

                select_sql = ", ".join(
                    _quote_ident(str(col["col_name"]), columns_by_name)
                    for col in columns_list
                )
                rows = connection.execute(
                    f"SELECT {select_sql} FROM dashboard_rows{where} "
                    f"ORDER BY {order_clause} LIMIT ? OFFSET ?",
                    [*parameters, page_size, (page - 1) * page_size],
                ).fetchall()
                summary = _filtered_dashboard_summary(
                    connection,
                    where,
                    parameters,
                    total,
                    columns_list,
                    columns_by_name,
                    meta,
                )
                facets = (
                    _build_dashboard_facets(
                        connection, columns_list, columns_by_name, institutions
                    )
                    if include_filters
                    else None
                )
        except ValueError as error:
            _send_json(self, 400, {"detail": str(error)})
            return
        except (OSError, sqlite3.Error, UnicodeError) as error:
            LOGGER.exception("Falha ao carregar os resultados da análise %s.", job_id)
            _send_json(self, 500, {"detail": f"Não foi possível carregar a tabela: {error}"})
            return

        results: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            if inst_col and inst_col in item:
                raw_code = item[inst_col]
                item["NOME_IF"] = (
                    institutions.get(str(raw_code), str(raw_code))
                    if raw_code is not None and str(raw_code).strip()
                    else None
                )
            results.append(item)

        filters_payload: dict[str, Any] = {"selected": active_filters}
        if facets is not None:
            filters_payload["facets"] = facets

        _send_json(
            self,
            200,
            {
                "rows": results,
                "columns": columns_list,
                "meta": meta,
                "total": total,
                "page": page,
                "page_size": page_size,
                "sort_by": sort_by,
                "sort_direction": sort_direction,
                "summary": summary,
                "filters": filters_payload,
            },
        )

    def _inspect(self) -> None:
        temp_root: Path | None = None
        try:
            content_length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            _send_json(self, 400, {"detail": "Content-Length inválido ou ausente."})
            return
        if content_length <= 0:
            _send_json(self, 400, {"detail": "O corpo da requisição está vazio."})
            return
        if content_length > MAX_INSPECT_BYTES + MAX_REQUEST_OVERHEAD:
            _send_json(
                self,
                413,
                {
                    "detail": (
                        f"A amostra de inspeção excede o limite de "
                        f"{MAX_INSPECT_BYTES / 1024**2:.0f} MiB."
                    )
                },
            )
            return

        try:
            JOB_DIR.mkdir(parents=True, exist_ok=True)
            temp_root = Path(tempfile.mkdtemp(prefix="inspect-", dir=JOB_DIR))
            uploads, _ = _parse_multipart(
                self, temp_root, content_length, max_bytes=MAX_INSPECT_BYTES
            )

            inspected: dict[str, list[dict[str, Any]]] = {
                "plots": [],
                "tables": [],
                "indigenous": [],
            }
            for role in ("plots", "tables", "indigenous"):
                for file_path in uploads[role]:
                    try:
                        info = inspect_input_file(
                            file_path, role=role, original_name=file_path.name
                        )
                        inspected[role].append(info)
                    except Exception as error:
                        inspected[role].append({"name": file_path.name, "error": str(error)})

            plots_columns = list(
                dict.fromkeys(
                    col
                    for item in inspected["plots"]
                    for col in item.get("columns", [])
                )
            )
            tables_columns = list(
                dict.fromkeys(
                    col
                    for item in inspected["tables"]
                    for col in item.get("columns", [])
                )
            )
            indigenous_layers = list(
                dict.fromkeys(
                    layer["name"]
                    for item in inspected["indigenous"]
                    for layer in item.get("layers", [])
                    if layer.get("name")
                )
            )
            indigenous_columns = list(
                dict.fromkeys(
                    col
                    for item in inspected["indigenous"]
                    for col in item.get("columns", [])
                )
            )

            valid_plots = [item for item in inspected["plots"] if item.get("columns")]
            valid_tables = [item for item in inspected["tables"] if item.get("columns")]
            valid_indig = [
                item
                for item in inspected["indigenous"]
                if item.get("columns") or item.get("layers")
            ]

            suggested_geometry = next(
                (item["suggested_geometry"] for item in valid_plots if item.get("suggested_geometry")),
                None,
            )
            geom_cf_set = {"geometry", "geom", "wkt", "geometria"}
            if suggested_geometry:
                geom_cf_set.add(suggested_geometry.casefold())

            non_geom_plots = [
                col for col in plots_columns if col.casefold() not in geom_cf_set
            ]
            if valid_tables:
                table_cf_sets = [
                    {c.casefold() for c in item["columns"]} for item in valid_tables
                ]
                common_cf = {col.casefold() for col in non_geom_plots}
                for tset in table_cf_sets:
                    common_cf &= tset
                common_join_keys = [
                    col for col in non_geom_plots if col.casefold() in common_cf
                ]
            else:
                common_join_keys = list(non_geom_plots)

            ref_priority = {name.casefold(): i for i, name in enumerate(DEFAULT_REF_COLUMNS)}
            common_join_keys.sort(
                key=lambda col: (ref_priority.get(col.casefold(), len(ref_priority)), 0)
            )

            suggested_ref: str | None = None
            if valid_plots:
                first_plot = valid_plots[0]
                try:
                    _, suggested_ref = resolve_join_key_column(
                        first_plot["columns"],
                        [item["columns"] for item in valid_tables],
                        None,
                        first_plot.get("suggested_geometry") or "geometry",
                    )
                except ValueError:
                    suggested_ref = common_join_keys[0] if common_join_keys else None

            suggested_name = next(
                (item["suggested_name"] for item in valid_indig if item.get("suggested_name")),
                None,
            )
            suggested_finance = next(
                (
                    item["suggested_finance"]
                    for item in [*valid_tables, *valid_plots]
                    if item.get("suggested_finance")
                ),
                None,
            )
            consolidation_plan = plan_table_consolidation(
                [(str(item["name"]), list(item["columns"])) for item in valid_tables],
                suggested_ref,
            )

            _send_json(
                self,
                200,
                {
                    **inspected,
                    "suggestions": {
                        "plots_columns": plots_columns,
                        "tables_columns": tables_columns,
                        "indigenous_layers": indigenous_layers,
                        "indigenous_columns": indigenous_columns,
                        "common_join_keys": common_join_keys,
                        "suggested_ref": suggested_ref,
                        "suggested_geometry": suggested_geometry,
                        "suggested_name": suggested_name,
                        "suggested_finance": suggested_finance,
                    },
                    "consolidation_plan": consolidation_plan,
                },
            )
        except (ValueError, UnicodeError) as error:
            _send_json(self, 400, {"detail": str(error)})
        except Exception:
            LOGGER.exception("Falha ao inspecionar arquivos enviados.")
            _send_json(self, 500, {"detail": "Falha interna ao inspecionar arquivos."})
        finally:
            if temp_root is not None:
                shutil.rmtree(temp_root, ignore_errors=True)

    def do_POST(self) -> None:
        global _active_job_id
        job_id: str | None = None
        job_root: Path | None = None
        try:
            self._validate_local_request()
            req_path = urlparse(self.path).path
            if req_path == "/api/inspect":
                self._inspect()
                return
            if req_path != "/api/analyze":
                _send_json(self, 404, {"detail": "Endpoint não encontrado."})
                return
            try:
                content_length = int(self.headers.get("Content-Length", ""))
            except ValueError as error:
                raise ValueError("Content-Length inválido ou ausente.") from error
            if content_length <= 0:
                raise ValueError("O corpo da requisição está vazio.")
            if content_length > MAX_UPLOAD_BYTES + MAX_REQUEST_OVERHEAD:
                _send_json(
                    self,
                    413,
                    {"detail": f"O envio excede o limite de {MAX_UPLOAD_BYTES / 1024**3:.1f} GiB."},
                )
                return

            with _active_job_lock:
                if _active_job_id is not None:
                    _send_json(
                        self,
                        409,
                        {"detail": "Já existe uma análise em execução. Aguarde sua conclusão."},
                    )
                    return
                job_id = uuid.uuid4().hex
                _active_job_id = job_id

            job_root = JOB_DIR / job_id
            (job_root / "outputs").mkdir(parents=True, exist_ok=False)
            uploads, fields = _parse_multipart(self, job_root, content_length)
            if not uploads["plots"] or not uploads["indigenous"]:
                raise ValueError(
                    "Envie ao menos um arquivo de glebas (.csv, .gpkg, .geojson, .parquet) "
                    "e um arquivo da Camada Base."
                )

            crs_glebas = fields.get("crs_glebas", "EPSG:4674").strip()
            crs_area = fields.get("crs_area", "auto").strip()
            if not crs_glebas or not crs_area:
                raise ValueError("Informe ambos os CRS.")
            try:
                chunk_size = int(fields.get("chunk_size", "50000"))
            except ValueError as error:
                raise ValueError("O tamanho do bloco deve ser um inteiro positivo.") from error
            if chunk_size < 1:
                raise ValueError("O tamanho do bloco deve ser um inteiro positivo.")
            try:
                min_area_ha = float(
                    (fields.get("min_area_ha", "0") or "0").strip().replace(",", ".")
                )
            except ValueError as error:
                raise ValueError("Área mínima deve ser um número válido.") from error
            if min_area_ha < 0:
                raise ValueError("Área mínima deve ser maior ou igual a zero.")
            try:
                min_percent = float(
                    (fields.get("min_percent", "0") or "0").strip().replace(",", ".")
                )
            except ValueError as error:
                raise ValueError("Percentual mínimo deve ser um número válido.") from error
            if not 0.0 <= min_percent <= 100.0:
                raise ValueError("Percentual mínimo deve estar entre 0 e 100.")

            outputs = {
                "xlsx": "resultado_filtrado.xlsx",
                "xlsx_full": "resultado_completo.xlsx",
                "csv": "resultado_filtrado.csv",
                "parquet": "resultado_filtrado.parquet",
                "geoparquet": "glebas_sobrepostas.parquet",
                "gpkg": "glebas_sobrepostas.gpkg",
            }
            args = SimpleNamespace(
                glebas=uploads["plots"],
                operacoes=uploads["tables"],
                terras_indigenas=uploads["indigenous"],
                saida=None,
                preview_path=None,
                dashboard_path=job_root / "outputs" / "dashboard.sqlite",
                saida_gpkg=None,
                camada_ti=fields.get("camada_ti", "").strip() or None,
                coluna_ref=fields.get("coluna_ref", "").strip() or None,
                coluna_geometria=fields.get("coluna_geometria", "").strip() or None,
                coluna_nome_ti=fields.get("coluna_nome_ti", "").strip() or None,
                coluna_valor=fields.get("coluna_valor", "").strip() or None,
                crs_glebas=crs_glebas,
                crs_area=crs_area,
                area_minima_ha=min_area_ha,
                percentual_minimo=min_percent,
                codificacao="auto",
                delimitador=None,
                tamanho_bloco=chunk_size,
                verbose=False,
            )
            with _jobs_lock:
                _jobs[job_id] = {
                    "status": "queued",
                    "message": "Arquivos recebidos; análise aguardando execução.",
                    "logs": [],
                    "outputs": outputs,
                    "full_excel_status": "pending",
                    "full_excel_error": None,
                    "full_excel_done": threading.Event(),
                }
            worker = threading.Thread(
                target=_execute_job, args=(job_id, args, job_root), daemon=True
            )
            worker.start()
            _send_json(self, 202, {"job_id": job_id, "status": "queued"})
        except PermissionError as error:
            if job_id is not None:
                self._cleanup_failed_upload(job_id, job_root)
            _send_json(self, 403, {"detail": str(error)})
        except (ValueError, UnicodeError) as error:
            if job_id is not None:
                self._cleanup_failed_upload(job_id, job_root)
            _send_json(self, 400, {"detail": str(error)})
        except Exception:
            LOGGER.exception("Falha ao receber solicitação de análise.")
            if job_id is not None:
                self._cleanup_failed_upload(job_id, job_root)
            _send_json(self, 500, {"detail": "Falha interna ao iniciar a análise."})

    @staticmethod
    def _cleanup_failed_upload(job_id: str, job_root: Path | None) -> None:
        global _active_job_id
        if job_root is not None:
            shutil.rmtree(job_root, ignore_errors=True)
        with _jobs_lock:
            _jobs.pop(job_id, None)
        with _active_job_lock:
            if _active_job_id == job_id:
                _active_job_id = None


def create_server(host: str = "127.0.0.1", port: int = 8000) -> ThreadingHTTPServer:
    """Cria o servidor HTTP garantindo a existência do diretório de jobs."""
    JOB_DIR.mkdir(parents=True, exist_ok=True)
    return ThreadingHTTPServer((host, port), RequestHandler)


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )
    server = create_server("127.0.0.1", 8000)
    LOGGER.info("Servidor iniciado em http://127.0.0.1:%d", server.server_port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        LOGGER.info("Encerrando o servidor.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
