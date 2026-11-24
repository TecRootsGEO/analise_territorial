#!/usr/bin/env python
"""Cruza glebas agrícolas em WKT com a Camada Base e operações do SICOR."""

from __future__ import annotations

import argparse
import codecs
import csv
import json
import logging
import math
import re
import sqlite3
import sys
import tempfile
from collections import Counter
from contextlib import closing, contextmanager
from pathlib import Path
from typing import Iterable, Iterator

import geopandas as gpd
import numpy as np
import pandas as pd
import polars as pl
import pyogrio
import shapely
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from pyproj import CRS, Transformer
from pyproj.database import query_crs_info
from shapely.geometry import mapping
from shapely.geometry.base import BaseGeometry

LOGGER = logging.getLogger("analise_sobreposicao")
EXCEL_MAX_ROWS = 1_048_576
DEFAULT_AREA_CRS = "auto"
DEFAULT_NAME_COLUMNS = ("terrai_nom", "nome_ti", "nome", "name", "nm_ti", "nm_uc", "titulo")
DEFAULT_GEOMETRY_COLUMNS = (
    "GEOMETRIA_WKT",
    "GT_GEOMETRIA",
    "WKT_GEOMETRY",
    "WKT",
    "GEOMETRIA",
    "GEOMETRY",
    "THE_GEOM",
    "GEOM",
)
DEFAULT_REF_COLUMNS = (
    "REF_BACEN",
    "ID_GLEBA",
    "CD_CAR",
    "COD_IMOVEL",
    "CODIGO_IMOVEL",
    "ID_PARCELA",
    "CODIGO",
    "ID",
)
DEFAULT_FINANCE_COLUMNS = (
    "VL_PARC_CREDITO",
    "VALOR_FINANCIADO",
    "VL_FINANCIADO",
    "VALOR_CREDITO",
    "VALOR_TOTAL",
    "VALOR",
)
MAX_EXTRA_GLEBA_COLUMNS = 50
DEFAULT_ENCODING_SAMPLE_BYTES = 4 * 1024 * 1024
ALBERS_BRAZIL_EPSG = 10857
ALBERS_SOUTH_AMERICA_FALLBACK = "ESRI:102033"
WKT_PREFIX_PATTERN = re.compile(
    r"^\s*(?:SRID=\d+\s*;\s*)?(?:MULTI)?(?:POLYGON|GEOMETRYCOLLECTION)\b",
    re.IGNORECASE,
)
LEADING_ZERO_ID_PATTERN = re.compile(r"^0\d+$")


def is_csv_path(path: Path) -> bool:
    return path.suffix.lower() == ".csv"


def is_parquet_path(path: Path) -> bool:
    return path.suffix.lower() == ".parquet"


def is_vector_path(path: Path) -> bool:
    return path.suffix.lower() in {".gpkg", ".geojson", ".json"}


def normalize_column_name(value: object) -> str:
    """Padroniza cabeçalhos, removendo BOM, espaços e o # do SICOR."""
    return str(value).replace("\ufeff", "").strip().lstrip("#").strip()


def validate_unique_columns(columns: Iterable[str], source: Path) -> None:
    normalized = [normalize_column_name(column).casefold() for column in columns]
    duplicates = sorted(
        name for name, count in Counter(normalized).items() if count > 1
    )
    if duplicates:
        raise ValueError(
            f"A tabela {source} possui nomes de colunas duplicados: "
            f"{', '.join(duplicates)}"
        )


def unique_column_name(base: str, existing: set[str]) -> str:
    candidate = base
    suffix = 2
    existing_casefolded = {column.casefold() for column in existing}
    while candidate.casefold() in existing_casefolded:
        candidate = f"{base}_{suffix}"
        suffix += 1
    return candidate


def infer_delimiter(path: Path, encoding: str) -> str:
    """Detecta o delimitador pelo cabeçalho sem ler o arquivo inteiro."""
    normalized = encoding.lower().replace("_", "-")
    text_encoding = "utf-8" if normalized in {"utf8-lossy", "utf-8-lossy"} else encoding
    errors = "replace" if normalized in {"utf8-lossy", "utf-8-lossy"} else "strict"
    with path.open("r", encoding=text_encoding, errors=errors, newline="") as source:
        sample = source.read(64 * 1024)
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        LOGGER.warning("Não foi possível detectar o delimitador de %s; usando ';'.", path)
        return ";"


def is_utf8_encoding(encoding: str) -> bool:
    return encoding.lower().replace("_", "-") in {
        "utf-8",
        "utf8",
        "utf-8-sig",
        "utf-8-lossy",
        "utf8-lossy",
    }


def encoding_is_valid(
    path: Path, encoding: str, max_bytes: int | None = DEFAULT_ENCODING_SAMPLE_BYTES
) -> bool:
    """Valida a codificação em blocos até max_bytes (ou arquivo inteiro se None)."""
    decoder = codecs.getincrementaldecoder(encoding)(errors="strict")
    bytes_read = 0
    reached_eof = False
    try:
        with path.open("rb") as source:
            while True:
                to_read = 1024 * 1024
                if max_bytes is not None:
                    remaining = max_bytes - bytes_read
                    if remaining <= 0:
                        break
                    to_read = min(to_read, remaining)
                block = source.read(to_read)
                if not block:
                    reached_eof = True
                    break
                bytes_read += len(block)
                decoder.decode(block)
        if reached_eof:
            decoder.decode(b"", final=True)
    except UnicodeDecodeError:
        return False
    return True


def resolve_csv_encoding(
    path: Path,
    requested: str,
    sample_bytes: int | None = DEFAULT_ENCODING_SAMPLE_BYTES,
) -> str:
    """Detecta UTF-8 por amostragem; usa Windows-1252 ou Latin-1 quando incompatível."""
    requested = requested.strip()
    normalized_request = requested.lower().replace("_", "-")
    if normalized_request in {"utf8-lossy", "utf-8-lossy"}:
        return "utf-8-lossy"
    if requested.lower() != "auto":
        try:
            codecs.lookup(requested)
        except LookupError as error:
            raise ValueError(f"Codificação desconhecida: {requested}") from error
        return requested
    with path.open("rb") as source:
        signature = source.read(4)
    if signature.startswith(codecs.BOM_UTF32_LE) or signature.startswith(codecs.BOM_UTF32_BE):
        return "utf-32"
    if signature.startswith(codecs.BOM_UTF16_LE) or signature.startswith(codecs.BOM_UTF16_BE):
        return "utf-16"
    if encoding_is_valid(path, "utf-8-sig", max_bytes=sample_bytes):
        return "utf-8-sig"
    if encoding_is_valid(path, "cp1252", max_bytes=sample_bytes):
        return "cp1252"
    return "latin-1"


def polars_encoding(encoding: str) -> str:
    """Seleciona a opção Polars após codificações não UTF-8 terem sido transcodificadas."""
    normalized = encoding.lower().replace("_", "-")
    if normalized in {"utf-8-lossy", "utf8-lossy"}:
        return "utf8-lossy"
    return "utf8"


def read_csv_header(path: Path, delimiter: str, encoding: str) -> list[str]:
    """Lê somente a primeira linha do CSV com a codificação detectada."""
    normalized = encoding.lower().replace("_", "-")
    text_encoding = "utf-8-sig" if normalized in {"utf8-lossy", "utf-8-lossy"} else encoding
    errors = "replace" if normalized in {"utf8-lossy", "utf-8-lossy"} else "strict"
    with path.open("r", encoding=text_encoding, errors=errors, newline="") as source:
        try:
            return next(csv.reader(source, delimiter=delimiter))
        except StopIteration as error:
            raise ValueError(f"CSV vazio: {path}") from error


@contextmanager
def polars_compatible_csv(path: Path, encoding: str) -> Iterator[Path]:
    """Fornece um caminho UTF-8 ao Polars, transcodificando em blocos se necessário."""
    if is_utf8_encoding(encoding):
        yield path
        return

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", prefix="geospatial_csv_", suffix=".csv", delete=False
        ) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with path.open("r", encoding=encoding, newline="") as source:
                with temporary_path.open("w", encoding="utf-8", newline="") as destination:
                    while True:
                        block = source.read(1024 * 1024)
                        if not block:
                            break
                        destination.write(block)
        except UnicodeDecodeError:
            if encoding.lower() == "latin-1":
                raise
            LOGGER.warning(
                "Erro de decodificação %s em %s; aplicando fallback para latin-1.",
                encoding,
                path.name,
            )
            with path.open("r", encoding="latin-1", newline="") as source:
                with temporary_path.open("w", encoding="utf-8", newline="") as destination:
                    while True:
                        block = source.read(1024 * 1024)
                        if not block:
                            break
                        destination.write(block)
        yield temporary_path
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def iter_csv_batches(
    path: Path,
    delimiter: str,
    encoding: str,
    columns: list[str] | None,
    chunk_size: int,
):
    """Lê CSV como strings em blocos com fallback seguro caso a amostra UTF-8 falhe adiante."""
    rows_yielded = 0
    try:
        with polars_compatible_csv(path, encoding) as csv_path:
            query = pl.scan_csv(
                csv_path,
                separator=delimiter,
                encoding=polars_encoding(encoding),
                infer_schema_length=0,
                low_memory=True,
            )
            if columns is not None:
                query = query.select(columns)
            for batch in query.collect_batches(chunk_size=chunk_size, engine="streaming"):
                yield batch
                rows_yielded += len(batch)
    except pl.exceptions.ComputeError as error:
        if not is_utf8_encoding(encoding) or "lossy" in encoding.lower():
            raise
        if "utf" not in str(error).lower() and "encoding" not in str(error).lower():
            raise
        LOGGER.warning(
            "Caractere incompatível com UTF-8 detectado após a amostra em %s; "
            "reprocessando restante com fallback cp1252.",
            path.name,
        )
        with polars_compatible_csv(path, "cp1252") as csv_path:
            query = pl.scan_csv(
                csv_path,
                separator=delimiter,
                encoding="utf8",
                infer_schema_length=0,
                low_memory=True,
            )
            if columns is not None:
                query = query.select(columns)
            if rows_yielded > 0:
                query = query.slice(rows_yielded, None)
            yield from query.collect_batches(chunk_size=chunk_size, engine="streaming")


def resolve_column(
    columns: Iterable[object], requested: str | None, candidates: Iterable[str], role: str
) -> object:
    """Resolve uma coluna sem diferenciar maiúsculas/minúsculas."""
    by_normalized = {normalize_column_name(column).casefold(): column for column in columns}
    names = (requested,) if requested else tuple(candidates)
    for name in names:
        found = by_normalized.get(normalize_column_name(name).casefold())
        if found is not None:
            return found
    available = ", ".join(str(column) for column in columns)
    raise ValueError(f"Coluna de {role} não encontrada. Disponíveis: {available}")


def resolve_geometry_column(
    path: Path,
    columns: Iterable[object],
    delimiter: str,
    encoding: str,
    requested: str | None,
) -> object:
    """Resolve a coluna de geometria WKT por parâmetro, candidatos padrão ou amostra."""
    column_list = list(columns)
    if requested:
        return resolve_column(column_list, requested, DEFAULT_GEOMETRY_COLUMNS, "geometria WKT")
    try:
        return resolve_column(column_list, None, DEFAULT_GEOMETRY_COLUMNS, "geometria WKT")
    except ValueError:
        pass

    normalized_enc = encoding.lower().replace("_", "-")
    text_encoding = "utf-8-sig" if normalized_enc in {"utf8-lossy", "utf-8-lossy"} else encoding
    errors = "replace" if normalized_enc in {"utf8-lossy", "utf-8-lossy"} else "strict"
    match_counts = [0] * len(column_list)
    with path.open("r", encoding=text_encoding, errors=errors, newline="") as source:
        reader = csv.reader(source, delimiter=delimiter)
        next(reader, None)
        for row_index, row in enumerate(reader):
            if row_index >= 20:
                break
            for idx, value in enumerate(row[: len(column_list)]):
                if value and WKT_PREFIX_PATTERN.match(value):
                    match_counts[idx] += 1
    best_count = max(match_counts, default=0)
    if best_count > 0:
        return column_list[match_counts.index(best_count)]
    available = ", ".join(str(column) for column in column_list)
    raise ValueError(f"Coluna de geometria WKT não encontrada. Disponíveis: {available}")


def resolve_join_key_column(
    gleba_header: Iterable[object],
    operation_headers: list[list[str]],
    requested: str | None,
    geometry_column: object,
) -> tuple[object | None, str]:
    """Determina a coluna-chave de junção entre glebas e tabelas auxiliares."""
    gleba_cols = list(gleba_header)
    geom_cf = normalize_column_name(geometry_column).casefold()
    non_geom_cols = [
        col for col in gleba_cols if normalize_column_name(col).casefold() != geom_cf
    ]
    if requested:
        found = resolve_column(gleba_cols, requested, (requested,), "chave de junção")
        return found, normalize_column_name(found)

    gleba_by_cf = {normalize_column_name(col).casefold(): col for col in non_geom_cols}
    if operation_headers:
        op_sets = [
            {normalize_column_name(col).casefold() for col in op_header}
            for op_header in operation_headers
        ]
        common_cf = set(gleba_by_cf)
        for op_set in op_sets:
            common_cf &= op_set
        if not common_cf:
            raise ValueError(
                "Não há coluna-chave em comum entre os CSVs de glebas e as tabelas auxiliares. "
                "Informe a coluna-chave manualmente nas opções avançadas."
            )
        for candidate in DEFAULT_REF_COLUMNS:
            cand_cf = normalize_column_name(candidate).casefold()
            if cand_cf in common_cf:
                col = gleba_by_cf[cand_cf]
                return col, normalize_column_name(col)
        for col in non_geom_cols:
            if normalize_column_name(col).casefold() in common_cf:
                return col, normalize_column_name(col)

    for candidate in DEFAULT_REF_COLUMNS:
        cand_cf = normalize_column_name(candidate).casefold()
        if cand_cf in gleba_by_cf:
            col = gleba_by_cf[cand_cf]
            return col, normalize_column_name(col)
    if non_geom_cols:
        col = non_geom_cols[0]
        return col, normalize_column_name(col)
    return None, "ID_GLEBA"


def resolve_finance_column(
    columns: Iterable[object], requested: str | None
) -> object | None:
    """Resolve opcionalmente a coluna de valor monetário/numérico."""
    column_list = list(columns)
    by_cf = {normalize_column_name(col).casefold(): col for col in column_list}
    if requested:
        return by_cf.get(normalize_column_name(requested).casefold())
    for candidate in DEFAULT_FINANCE_COLUMNS:
        found = by_cf.get(normalize_column_name(candidate).casefold())
        if found is not None:
            return found
    for col in column_list:
        cf = normalize_column_name(col).casefold()
        if cf.startswith(("vl_", "valor_")) and not any(
            excluded in cf for excluded in ("area", "_ha", "perc", "taxa", "juros")
        ):
            return col
    return None


def polygonal_geometry(geometry: BaseGeometry | None) -> BaseGeometry | None:
    """Retém somente partes poligonais de geometrias corrigidas."""
    if geometry is None or geometry.is_empty:
        return None
    if geometry.geom_type in ("Polygon", "MultiPolygon"):
        return geometry
    if geometry.geom_type == "GeometryCollection":
        polygons = [
            part
            for part in geometry.geoms
            if part.geom_type in ("Polygon", "MultiPolygon") and not part.is_empty
        ]
        return shapely.union_all(polygons) if polygons else None
    return None


def repair_polygon(geometry: BaseGeometry | None) -> BaseGeometry | None:
    """Corrige geometria inválida, descartando feições que não sejam polígonos."""
    if geometry is None or geometry.is_empty:
        return None
    try:
        candidate = geometry if geometry.is_valid else shapely.make_valid(geometry)
        candidate = polygonal_geometry(candidate)
        if candidate is not None and not candidate.is_valid:
            candidate = polygonal_geometry(candidate.buffer(0))
        if candidate is None or candidate.is_empty or not candidate.is_valid:
            return None
        return candidate
    except Exception:
        LOGGER.debug("Falha ao corrigir geometria.", exc_info=True)
        return None


def resolve_indigenous_layer(path: Path, layer_name: str | None) -> str:
    if not path.is_file():
        raise FileNotFoundError(f"Arquivo espacial não encontrado: {path}")
    available_layers = gpd.list_layers(path)
    if available_layers.empty:
        raise ValueError(f"Nenhuma camada encontrada em {path}")
    selected_layer = layer_name or str(available_layers.iloc[0]["name"])
    if selected_layer not in set(available_layers["name"]):
        raise ValueError(
            f"A camada '{selected_layer}' não existe em {path}. "
            f"Camadas disponíveis: {', '.join(available_layers['name'])}"
        )
    return selected_layer


def _crs_from_geoparquet_meta(geom_meta: dict[str, object], default_crs: object) -> CRS:
    crs_info = geom_meta.get("crs")
    if isinstance(crs_info, dict):
        try:
            return CRS.from_json_dict(crs_info)
        except Exception:
            pass
    if isinstance(crs_info, str) and crs_info.strip():
        try:
            return CRS.from_user_input(crs_info)
        except Exception:
            pass
    return CRS.from_user_input(default_crs)


def read_geoparquet(
    path: Path,
    default_crs: object = "EPSG:4674",
    geometry_column: str | None = None,
    columns: list[str] | None = None,
) -> gpd.GeoDataFrame:
    """Lê um arquivo Parquet/GeoParquet usando Polars sem depender de pyarrow."""
    schema = pl.read_parquet_schema(path)
    meta = pl.read_parquet_metadata(path)
    geo_meta: dict[str, object] = {}
    if "geo" in meta:
        try:
            geo_meta = json.loads(meta["geo"])
        except Exception:
            geo_meta = {}

    raw_cols = list(schema.keys())
    primary_col = geo_meta.get("primary_column") if isinstance(geo_meta, dict) else None
    if geometry_column:
        geom_col = str(resolve_column(raw_cols, geometry_column, DEFAULT_GEOMETRY_COLUMNS, "geometria"))
    elif isinstance(primary_col, str) and primary_col in schema:
        geom_col = primary_col
    else:
        geom_col = str(resolve_column(raw_cols, None, DEFAULT_GEOMETRY_COLUMNS, "geometria"))

    geom_meta_dict = {}
    if isinstance(geo_meta.get("columns"), dict):
        geom_meta_dict = geo_meta["columns"].get(geom_col, {}) or {}
    resolved_crs = _crs_from_geoparquet_meta(geom_meta_dict, default_crs)

    selected_cols = list(dict.fromkeys([*(columns or [c for c in raw_cols if c != geom_col]), geom_col]))
    frame_pl = pl.read_parquet(path, columns=selected_cols)
    if schema[geom_col] == pl.Binary:
        raw_geoms = [val if val is not None else b"" for val in frame_pl.get_column(geom_col).to_list()]
        geoms = shapely.from_wkb(raw_geoms, on_invalid="ignore")
    else:
        raw_geoms = [str(val) if val is not None else "" for val in frame_pl.get_column(geom_col).to_list()]
        geoms = shapely.from_wkt(raw_geoms, on_invalid="ignore")

    data = {
        col: frame_pl.get_column(col).to_list()
        for col in frame_pl.columns
        if col != geom_col
    }
    return gpd.GeoDataFrame(data, geometry=geoms, crs=resolved_crs)


def _pandas_df_to_polars(frame: pd.DataFrame) -> pl.DataFrame:
    """Converte DataFrame Pandas para Polars sem depender de pyarrow."""
    series_list: list[pl.Series] = []
    for col in frame.columns:
        col_name = str(col)
        series = frame[col]
        if pd.api.types.is_bool_dtype(series.dtype):
            vals = [None if pd.isna(v) else bool(v) for v in series]
            series_list.append(pl.Series(col_name, vals, dtype=pl.Boolean))
        elif pd.api.types.is_integer_dtype(series.dtype):
            vals = [None if pd.isna(v) else int(v) for v in series]
            series_list.append(pl.Series(col_name, vals, dtype=pl.Int64))
        elif pd.api.types.is_float_dtype(series.dtype):
            vals = [None if pd.isna(v) or not math.isfinite(float(v)) else float(v) for v in series]
            series_list.append(pl.Series(col_name, vals, dtype=pl.Float64))
        else:
            vals = [None if pd.isna(v) else str(v) for v in series]
            series_list.append(pl.Series(col_name, vals, dtype=pl.String))
    return pl.DataFrame(series_list)


def write_geoparquet(
    output_path: Path, frame: gpd.GeoDataFrame, crs: object
) -> None:
    """Exporta GeoDataFrame para GeoParquet 1.1 nativamente via Polars (sem pyarrow)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    export = frame.drop(columns=["plot_uid", "land_uids"], errors="ignore")
    geom_col = export.geometry.name if hasattr(export, "geometry") else "geometry"
    attr_df = pd.DataFrame(export.drop(columns=[geom_col], errors="ignore"))
    pl_df = _pandas_df_to_polars(attr_df)

    geoms = list(export.geometry)
    wkb_values = [
        shapely.to_wkb(geom) if geom is not None and not geom.is_empty else None
        for geom in geoms
    ]
    pl_df = pl_df.with_columns(pl.Series("geometry", wkb_values, dtype=pl.Binary))

    geom_types = sorted(
        {geom.geom_type for geom in geoms if geom is not None and not geom.is_empty}
    )
    if not export.empty and any(g is not None and not g.is_empty for g in geoms):
        bounds = export.total_bounds
        bbox = [float(v) for v in bounds] if all(math.isfinite(float(v)) for v in bounds) else [0.0, 0.0, 0.0, 0.0]
    else:
        bbox = [0.0, 0.0, 0.0, 0.0]

    crs_obj = CRS.from_user_input(crs)
    geo_meta = {
        "version": "1.1.0",
        "primary_column": "geometry",
        "columns": {
            "geometry": {
                "encoding": "WKB",
                "geometry_types": geom_types,
                "crs": crs_obj.to_json_dict(),
                "bbox": bbox,
            }
        },
    }
    pl_df.write_parquet(
        output_path,
        compression="zstd",
        metadata={"geo": json.dumps(geo_meta, ensure_ascii=False)},
    )


def inspect_input_file(
    path: Path, role: str, original_name: str | None = None
) -> dict[str, object]:
    """Inspeciona cabeçalhos, camadas e sugestões sem carregar grandes arquivos em memória."""
    display_name = original_name or path.name
    ext = path.suffix.lower()
    if ext == ".csv":
        encoding = resolve_csv_encoding(path, "auto")
        delimiter = infer_delimiter(path, encoding)
        raw_header = read_csv_header(path, delimiter, encoding)
        validate_unique_columns(raw_header, path)
        columns = [normalize_column_name(col) for col in raw_header]
        suggested_geometry: str | None = None
        if role == "plots":
            try:
                geom_col = resolve_geometry_column(path, raw_header, delimiter, encoding, None)
                suggested_geometry = normalize_column_name(geom_col)
            except ValueError:
                suggested_geometry = None
        fin_col = resolve_finance_column(raw_header, None)
        return {
            "name": display_name,
            "format": "csv",
            "encoding": encoding,
            "delimiter": delimiter,
            "columns": columns,
            "layers": [],
            "suggested_geometry": suggested_geometry,
            "suggested_name": None,
            "suggested_finance": normalize_column_name(fin_col) if fin_col is not None else None,
        }

    if ext == ".parquet":
        schema = pl.read_parquet_schema(path)
        meta = pl.read_parquet_metadata(path)
        raw_cols = list(schema.keys())
        columns = [normalize_column_name(col) for col in raw_cols]
        geo_meta: dict[str, object] = {}
        if "geo" in meta:
            try:
                geo_meta = json.loads(meta["geo"])
            except Exception:
                geo_meta = {}
        suggested_geometry = None
        if role == "plots":
            primary = geo_meta.get("primary_column") if isinstance(geo_meta, dict) else None
            if isinstance(primary, str) and primary in schema:
                suggested_geometry = normalize_column_name(primary)
            else:
                try:
                    suggested_geometry = normalize_column_name(
                        resolve_column(raw_cols, None, DEFAULT_GEOMETRY_COLUMNS, "geometria")
                    )
                except ValueError:
                    suggested_geometry = None
        suggested_name = None
        if role == "indigenous":
            try:
                suggested_name = normalize_column_name(
                    resolve_column(raw_cols, None, DEFAULT_NAME_COLUMNS, "nome da Camada Base")
                )
            except ValueError:
                suggested_name = None
        fin_col = resolve_finance_column(raw_cols, None)
        return {
            "name": display_name,
            "format": "parquet",
            "encoding": None,
            "delimiter": None,
            "columns": columns,
            "layers": [],
            "suggested_geometry": suggested_geometry,
            "suggested_name": suggested_name,
            "suggested_finance": normalize_column_name(fin_col) if fin_col is not None else None,
        }

    if ext in {".gpkg", ".geojson", ".json"}:
        layers_df = gpd.list_layers(path)
        layers: list[dict[str, object]] = []
        for layer_name in layers_df["name"].tolist()[:20]:
            info = pyogrio.read_info(path, layer=str(layer_name))
            fields = [normalize_column_name(col) for col in info.get("fields", [])]
            layers.append(
                {
                    "name": str(layer_name),
                    "fields": fields,
                    "crs": str(info.get("crs") or ""),
                    "features": int(info.get("features") or 0),
                }
            )
        first_fields = layers[0]["fields"] if layers else []
        suggested_name = None
        if role == "indigenous" and first_fields:
            try:
                suggested_name = normalize_column_name(
                    resolve_column(first_fields, None, DEFAULT_NAME_COLUMNS, "nome da Camada Base")
                )
            except ValueError:
                suggested_name = None
        fin_col = resolve_finance_column(first_fields, None) if first_fields else None
        return {
            "name": display_name,
            "format": ext.lstrip("."),
            "encoding": None,
            "delimiter": None,
            "columns": first_fields,
            "layers": layers,
            "suggested_geometry": "geometry" if role == "plots" else None,
            "suggested_name": suggested_name,
            "suggested_finance": normalize_column_name(fin_col) if fin_col is not None else None,
        }

    raise ValueError(f"Formato de arquivo não suportado para inspeção: {display_name}")


def plan_table_consolidation(
    table_headers: list[tuple[str, list[str]]], key_column: str | None
) -> list[dict[str, object]]:
    """Descreve quais tabelas auxiliares serão empilhadas (UNION) ou cruzadas (JOIN)."""
    if not table_headers:
        return []
    effective_key = key_column or "REF_BACEN"
    groups: list[dict[str, object]] = []
    for idx, (filename, cols) in enumerate(table_headers):
        signature = _schema_signature(cols, key_column=effective_key)
        col_cf = {normalize_column_name(c).casefold() for c in cols}
        matched = next(
            (g for g in groups if _schemas_match(g["signature"], signature)),  # type: ignore[arg-type]
            None,
        )
        if matched is not None:
            matched["files"].append(filename)  # type: ignore[attr-defined]
            matched["signature"] = matched["signature"] | signature  # type: ignore[operator]
            matched["col_cf"] = matched["col_cf"] | col_cf  # type: ignore[operator]
        else:
            groups.append(
                {
                    "signature": signature,
                    "col_cf": col_cf,
                    "files": [filename],
                    "first_index": idx,
                }
            )

    primary_markers = {
        "vl_parc_credito",
        "valor_financiado",
        "vl_financiado",
        "valor_credito",
        "valor_total",
        "valor",
        "dt_emissao",
        "cnpj_if",
    }

    def group_priority(group: dict[str, object]) -> tuple[bool, int, int]:
        sig = group["signature"]  # type: ignore[assignment]
        return (bool(sig & primary_markers), len(sig), -int(group["first_index"]))  # type: ignore[arg-type]

    primary_group = max(groups, key=group_priority)
    primary_has_nu = "nu_ordem" in primary_group["col_cf"]  # type: ignore[operator]
    plan: list[dict[str, object]] = []

    ordered_groups = [primary_group, *[g for g in groups if g is not primary_group]]
    for group in ordered_groups:
        files = list(group["files"])  # type: ignore[arg-type]
        if group is primary_group:
            mode = "UNION" if len(files) > 1 else "PRIMARY"
            desc = (
                "Empilhadas verticalmente (mesmo esquema)"
                if len(files) > 1
                else "Tabela principal do cruzamento"
            )
            join_keys = [key_column] if key_column else []
        else:
            has_nu = primary_has_nu and ("nu_ordem" in group["col_cf"])  # type: ignore[operator]
            join_keys = (
                [key_column, "NU_ORDEM"]
                if has_nu and key_column
                else ([key_column] if key_column else [])
            )
            mode = "UNION + JOIN" if len(files) > 1 else "JOIN"
            prefix = "Empilhadas e cruzadas" if len(files) > 1 else "Cruzada horizontalmente"
            desc = f"{prefix} por {', '.join(join_keys)}" if join_keys else prefix
        plan.append(
            {
                "mode": mode,
                "files": files,
                "join_keys": join_keys,
                "description": desc,
            }
        )
    return plan


def load_indigenous_lands(
    paths: list[Path],
    area_crs: str,
    layer_name: str | None,
    ti_name_column: str | None,
) -> gpd.GeoDataFrame:
    """Lê, corrige e projeta uma ou mais camadas da Camada Base (.gpkg, .geojson, .parquet)."""
    layers: list[gpd.GeoDataFrame] = []
    for path in paths:
        ext = path.suffix.lower()
        if ext == ".parquet":
            selected_layer = path.stem
            frame = read_geoparquet(path, default_crs="EPSG:4674")
        elif ext in {".geojson", ".json"}:
            selected_layer = path.stem
            frame = gpd.read_file(path, engine="pyogrio")
            if frame.crs is None:
                frame = frame.set_crs("EPSG:4674", allow_override=True)
        else:
            selected_layer = resolve_indigenous_layer(path, layer_name)
            frame = gpd.read_file(path, layer=selected_layer, engine="pyogrio")

        if frame.crs is None:
            raise ValueError(f"A camada {selected_layer} em {path} não possui CRS definido.")
        if ti_name_column:
            name_column = resolve_column(
                frame.columns,
                ti_name_column,
                DEFAULT_NAME_COLUMNS,
                "nome da Camada Base",
            )
        else:
            try:
                name_column = resolve_column(
                    frame.columns, None, DEFAULT_NAME_COLUMNS, "nome da Camada Base"
                )
            except ValueError:
                name_column = None

        fallback_name = path.stem
        if name_column is not None:
            names = frame[name_column].astype("string").fillna("").str.strip()
            names = names.mask(names == "", fallback_name)
        else:
            names = pd.Series(fallback_name, index=frame.index, dtype="string")
        geometries = frame.geometry.map(repair_polygon)
        rejected = int(geometries.isna().sum())
        if rejected:
            LOGGER.warning("%s feições poligonais inválidas/sem geometria ignoradas em %s.", rejected, path)
        valid = geometries.notna()
        cleaned = gpd.GeoDataFrame(
            {"ti_name": names.loc[valid].to_numpy(), "geometry": geometries.loc[valid]},
            geometry="geometry",
            crs=frame.crs,
        )
        if cleaned.empty:
            raise ValueError(f"A camada {selected_layer} em {path} não contém polígonos válidos.")
        layers.append(cleaned.to_crs(area_crs))
        LOGGER.info(
            "Camada %s carregada: %d polígonos válidos (CRS de área %s).",
            selected_layer,
            len(cleaned),
            area_crs,
        )

    combined = pd.concat(layers, ignore_index=True)
    return gpd.GeoDataFrame(combined, geometry="geometry", crs=area_crs).reset_index(drop=True)


class SpatialBaseIndex:
    """Índice espacial reutilizável da Camada Base para estratégia filter-then-refine."""

    def __init__(
        self, lands: gpd.GeoDataFrame, source_crs: object, area_crs: CRS
    ) -> None:
        self.area_geoms = lands.geometry.to_numpy()
        shapely.prepare(self.area_geoms)
        self.area_tree = shapely.STRtree(self.area_geoms)
        self.land_names = lands["ti_name"].to_numpy()
        self.linear_unit_to_m = float(area_crs.axis_info[0].unit_conversion_factor)

        source_crs_obj = CRS.from_user_input(source_crs)
        lands_source = lands.geometry.to_crs(source_crs_obj).to_numpy()
        bounds = shapely.bounds(lands_source)
        eps = 0.01 if source_crs_obj.is_geographic else 1000.0
        expanded_boxes = shapely.box(
            bounds[:, 0] - eps,
            bounds[:, 1] - eps,
            bounds[:, 2] + eps,
            bounds[:, 3] + eps,
        )
        self.source_bbox_tree = shapely.STRtree(expanded_boxes)


def _evaluate_parsed_batch(
    parsed: np.ndarray,
    references: list[str],
    has_ref: np.ndarray,
    extract_extra: object | None,
    key_column: str,
    base_index: SpatialBaseIndex,
    source_crs: object,
    area_crs: CRS,
    row_offset: int,
    min_area_ha: float = 0.0,
    min_percent: float = 0.0,
) -> tuple[list[dict[str, object]], list[dict[str, object]], int, int, int, int]:
    """Executa filter-then-refine e retorna registros por gleba, interseções 1:N e métricas."""
    missing_mask = shapely.is_missing(parsed)
    invalid_wkt = int(missing_mask.sum())

    present_idx = np.flatnonzero((~missing_mask) & has_ref)
    if present_idx.size == 0:
        return [], [], 0, invalid_wkt, 0, 0

    present_geoms = parsed[present_idx]
    type_ids = shapely.get_type_id(present_geoms)
    polygonal = np.isin(type_ids, (3, 6, 7)) & (~shapely.is_empty(present_geoms))
    eligible_idx = present_idx[polygonal]
    polygonal_count = int(eligible_idx.size)
    if polygonal_count == 0:
        return [], [], 0, invalid_wkt, 0, 0

    # Consulta apenas por envelope no R-tree de origem (sem predicate=) evita TopologyException.
    hit_local, _ = base_index.source_bbox_tree.query(parsed[eligible_idx])
    if hit_local.size == 0:
        return [], [], polygonal_count, invalid_wkt, 0, 0

    unique_local = np.unique(hit_local)
    cand_idx = eligible_idx[unique_local]
    cand_geoms = parsed[cand_idx]

    valid_mask = shapely.is_valid(cand_geoms)
    is_gc_mask = shapely.get_type_id(cand_geoms) == 7
    surviving_geoms: list[BaseGeometry] = []
    surviving_chunk_rows: list[int] = []
    invalid_candidate_polygons = 0

    for chunk_row, geom, is_ok, is_gc in zip(
        cand_idx.tolist(), cand_geoms, valid_mask, is_gc_mask
    ):
        if is_ok and not is_gc:
            surviving_geoms.append(geom)
            surviving_chunk_rows.append(chunk_row)
        else:
            fixed = repair_polygon(geom)
            if fixed is None:
                invalid_candidate_polygons += 1
            else:
                surviving_geoms.append(fixed)
                surviving_chunk_rows.append(chunk_row)

    if not surviving_geoms:
        return [], [], polygonal_count, invalid_wkt, invalid_candidate_polygons, 0

    cand_area_geoms = (
        gpd.GeoSeries(surviving_geoms, crs=source_crs).to_crs(area_crs).to_numpy()
    )
    cand_hits, land_hits = base_index.area_tree.query(
        cand_area_geoms, predicate="intersects"
    )
    if cand_hits.size == 0:
        return [], [], polygonal_count, invalid_wkt, invalid_candidate_polygons, 0

    try:
        pair_intersections = shapely.intersection(
            cand_area_geoms[cand_hits], base_index.area_geoms[land_hits]
        )
    except Exception:
        repaired_cands = np.array(
            [repair_polygon(g) or g.buffer(0) for g in cand_area_geoms], dtype=object
        )
        pair_intersections = shapely.intersection(
            repaired_cands[cand_hits], base_index.area_geoms[land_hits]
        )
    pair_areas = shapely.area(pair_intersections)
    pos_mask = pair_areas > 0
    if not pos_mask.any():
        return [], [], polygonal_count, invalid_wkt, invalid_candidate_polygons, 0

    cand_pos = cand_hits[pos_mask]
    land_pos = land_hits[pos_mask]
    inter_pos = pair_intersections[pos_mask]
    areas_pos = pair_areas[pos_mask]

    order = np.argsort(cand_pos, kind="stable")
    cand_sorted = cand_pos[order]
    land_sorted = land_pos[order]
    inter_sorted = inter_pos[order]
    areas_sorted = areas_pos[order]

    area_factor = base_index.linear_unit_to_m**2
    cand_total_areas = shapely.area(cand_area_geoms) * area_factor
    unique_cands, start_idx = np.unique(cand_sorted, return_index=True)
    end_idx = np.append(start_idx[1:], cand_sorted.size)

    prelim_kept: list[tuple[int, int, int, float, float]] = []
    filtered_slivers = 0
    for c_idx, start, stop in zip(unique_cands.tolist(), start_idx.tolist(), end_idx.tolist()):
        total_area = float(cand_total_areas[c_idx])
        if total_area <= 0:
            continue
        if stop - start == 1:
            overlap_area = float(areas_sorted[start]) * area_factor
        else:
            overlap_area = (
                float(shapely.union_all(inter_sorted[start:stop]).area) * area_factor
            )
        if overlap_area <= 0:
            continue
        overlap_area = min(overlap_area, total_area)
        overlap_ha = overlap_area / 10_000
        overlap_pct = overlap_area / total_area * 100
        if (min_area_ha > 0 and overlap_ha < min_area_ha) or (
            min_percent > 0 and overlap_pct < min_percent
        ):
            filtered_slivers += 1
            continue
        prelim_kept.append((c_idx, start, stop, total_area, overlap_area))

    if not prelim_kept:
        return [], [], polygonal_count, invalid_wkt, invalid_candidate_polygons, filtered_slivers

    extra_values_by_cidx: dict[int, dict[str, object]] = {}
    if extract_extra is not None:
        matched_rows = [surviving_chunk_rows[item[0]] for item in prelim_kept]
        extracted_list = extract_extra(matched_rows)  # type: ignore[operator]
        for (c_idx, *_), extra_dict in zip(prelim_kept, extracted_list):
            extra_values_by_cidx[c_idx] = extra_dict

    results: list[dict[str, object]] = []
    intersections_out: list[dict[str, object]] = []
    for c_idx, start, stop, total_area, overlap_area in prelim_kept:
        land_slice = land_sorted[start:stop]
        names = sorted(
            {
                str(base_index.land_names[i]).strip()
                for i in land_slice.tolist()
                if pd.notna(base_index.land_names[i])
                and str(base_index.land_names[i]).strip()
            }
        )
        chunk_row = surviving_chunk_rows[c_idx]
        plot_uid = int(row_offset + chunk_row)
        ref_value = references[chunk_row]
        total_ha = total_area / 10_000
        record: dict[str, object] = {
            "plot_uid": plot_uid,
            key_column: ref_value,
            "AREA_TOTAL_HA": total_ha,
            "AREA_SOBREPOSTA_HA": overlap_area / 10_000,
            "PERCENTUAL_SOBREPOSICAO": overlap_area / total_area * 100,
            "NOME_TERRA_INDIGENA": " | ".join(names) if names else pd.NA,
            "land_uids": sorted({int(i) for i in land_slice.tolist()}),
            "geometry": cand_area_geoms[c_idx],
        }
        if extra_values_by_cidx:
            record.update(extra_values_by_cidx[c_idx])
        results.append(record)

        for idx in range(start, stop):
            part_geom = polygonal_geometry(inter_sorted[idx])
            if part_geom is None or part_geom.is_empty:
                continue
            part_area_m2 = min(float(areas_sorted[idx]) * area_factor, total_area)
            land_idx = int(land_sorted[idx])
            raw_land_name = base_index.land_names[land_idx]
            land_name = (
                str(raw_land_name).strip()
                if pd.notna(raw_land_name) and str(raw_land_name).strip()
                else "Camada Base"
            )
            intersections_out.append(
                {
                    "plot_uid": plot_uid,
                    "land_uid": land_idx,
                    key_column: ref_value,
                    "CAMADA_BASE": land_name,
                    "AREA_TOTAL_HA": total_ha,
                    "AREA_INTERSECAO_HA": part_area_m2 / 10_000,
                    "PERCENTUAL_INTERSECAO": part_area_m2 / total_area * 100,
                    "geometry": part_geom,
                }
            )

    return (
        results,
        intersections_out,
        polygonal_count,
        invalid_wkt,
        invalid_candidate_polygons,
        filtered_slivers,
    )


def _polars_extra_extractor(
    chunk: pl.DataFrame, extra_columns: list[object]
):
    if not extra_columns:
        return None
    col_names = [normalize_column_name(c) for c in extra_columns]
    raw_names = [str(c) for c in extra_columns]

    def _extract(matched_rows: list[int]) -> list[dict[str, object]]:
        matched_extra = chunk[matched_rows].select(raw_names)
        output: list[dict[str, object]] = []
        for local_idx in range(len(matched_rows)):
            row_dict = {}
            for raw_col, norm_col in zip(raw_names, col_names):
                row_dict[norm_col] = matched_extra[raw_col][local_idx]
            output.append(row_dict)
        return output

    return _extract


def process_wkt_chunk(
    chunk: pl.DataFrame,
    ref_column: object | None,
    key_column: str,
    geometry_column: object,
    extra_columns: list[object],
    base_index: SpatialBaseIndex,
    source_crs: object,
    area_crs: CRS,
    row_offset: int,
    min_area_ha: float = 0.0,
    min_percent: float = 0.0,
) -> tuple[list[dict[str, object]], list[dict[str, object]], int, int, int, int]:
    """Filtra por BBOX no CRS de origem antes de reparar, reprojetar e intersectar (WKT)."""
    try:
        wkt_series = chunk.get_column(str(geometry_column)).cast(pl.String).fill_null("")
        if wkt_series.str.contains(r"(?i)^\s*srid=\d+\s*;").any():
            wkt_series = wkt_series.str.replace(r"(?i)^\s*srid=\d+\s*;\s*", "")
        wkt_values = wkt_series.to_list()
        parsed = shapely.from_wkt(wkt_values, on_invalid="ignore")
    except Exception as error:
        raise ValueError(f"Falha ao interpretar WKT a partir da linha {row_offset + 2}.") from error

    if ref_column is not None:
        ref_series = (
            chunk.get_column(str(ref_column))
            .cast(pl.String)
            .fill_null("")
            .str.strip_chars()
        )
        references = ref_series.to_list()
        has_ref = ref_series.str.len_chars().to_numpy() > 0
    else:
        references = [f"GLEBA_{row_offset + index + 1}" for index in range(len(chunk))]
        has_ref = np.ones(len(chunk), dtype=bool)

    return _evaluate_parsed_batch(
        parsed=parsed,
        references=references,
        has_ref=has_ref,
        extract_extra=_polars_extra_extractor(chunk, extra_columns),
        key_column=key_column,
        base_index=base_index,
        source_crs=source_crs,
        area_crs=area_crs,
        row_offset=row_offset,
        min_area_ha=min_area_ha,
        min_percent=min_percent,
    )


def process_wkb_chunk(
    chunk: pl.DataFrame,
    ref_column: object | None,
    key_column: str,
    geometry_column: object,
    extra_columns: list[object],
    base_index: SpatialBaseIndex,
    source_crs: object,
    area_crs: CRS,
    row_offset: int,
    min_area_ha: float = 0.0,
    min_percent: float = 0.0,
) -> tuple[list[dict[str, object]], list[dict[str, object]], int, int, int, int]:
    """Filtra por BBOX no CRS de origem antes de reparar, reprojetar e intersectar (WKB)."""
    raw_wkb = [
        val if val is not None else b""
        for val in chunk.get_column(str(geometry_column)).to_list()
    ]
    parsed = shapely.from_wkb(raw_wkb, on_invalid="ignore")

    if ref_column is not None:
        ref_series = (
            chunk.get_column(str(ref_column))
            .cast(pl.String)
            .fill_null("")
            .str.strip_chars()
        )
        references = ref_series.to_list()
        has_ref = ref_series.str.len_chars().to_numpy() > 0
    else:
        references = [f"GLEBA_{row_offset + index + 1}" for index in range(len(chunk))]
        has_ref = np.ones(len(chunk), dtype=bool)

    return _evaluate_parsed_batch(
        parsed=parsed,
        references=references,
        has_ref=has_ref,
        extract_extra=_polars_extra_extractor(chunk, extra_columns),
        key_column=key_column,
        base_index=base_index,
        source_crs=source_crs,
        area_crs=area_crs,
        row_offset=row_offset,
        min_area_ha=min_area_ha,
        min_percent=min_percent,
    )


def process_geodataframe_plots(
    frame: gpd.GeoDataFrame,
    ref_column: object | None,
    key_column: str,
    extra_columns: list[object],
    base_index: SpatialBaseIndex,
    source_crs: object,
    area_crs: CRS,
    row_offset: int,
    min_area_ha: float = 0.0,
    min_percent: float = 0.0,
) -> tuple[list[dict[str, object]], list[dict[str, object]], int, int, int, int]:
    """Processa lote já geometricamente parseado (GPKG / GeoJSON) via filter-then-refine."""
    parsed = frame.geometry.to_numpy()
    if ref_column is not None:
        ref_series = (
            frame[str(ref_column)].astype("string").fillna("").str.strip()
        )
        references = ref_series.tolist()
        has_ref = (ref_series.str.len() > 0).to_numpy(dtype=bool)
    else:
        references = [f"GLEBA_{row_offset + index + 1}" for index in range(len(frame))]
        has_ref = np.ones(len(frame), dtype=bool)

    def _gdf_extractor(matched_rows: list[int]) -> list[dict[str, object]]:
        if not extra_columns:
            return [{} for _ in matched_rows]
        sub = frame.iloc[matched_rows]
        col_names = [normalize_column_name(c) for c in extra_columns]
        raw_names = [str(c) for c in extra_columns]
        output: list[dict[str, object]] = []
        for idx in range(len(matched_rows)):
            row = sub.iloc[idx]
            output.append(
                {
                    norm_col: (None if pd.isna(row[raw_col]) else row[raw_col])
                    for raw_col, norm_col in zip(raw_names, col_names)
                }
            )
        return output

    return _evaluate_parsed_batch(
        parsed=parsed,
        references=references,
        has_ref=has_ref,
        extract_extra=_gdf_extractor if extra_columns else None,
        key_column=key_column,
        base_index=base_index,
        source_crs=source_crs,
        area_crs=area_crs,
        row_offset=row_offset,
        min_area_ha=min_area_ha,
        min_percent=min_percent,
    )


def clean_wkt_chunk(
    chunk: pl.DataFrame,
    ref_column: object,
    geometry_column: object,
    row_offset: int,
) -> tuple[gpd.GeoDataFrame, int, int]:
    """Converte WKT e prepara polígonos válidos; retorna contagens de rejeição."""
    try:
        wkt_values = chunk.get_column(str(geometry_column)).to_list()
        parsed = shapely.from_wkt(
            [value if value is not None else "" for value in wkt_values],
            on_invalid="ignore",
        )
    except Exception as error:
        raise ValueError(f"Falha ao interpretar WKT a partir da linha {row_offset + 2}.") from error

    references = (
        chunk.get_column(str(ref_column))
        .cast(pl.String)
        .fill_null("")
        .str.strip_chars()
        .to_list()
    )
    repaired = [repair_polygon(geometry) for geometry in parsed]
    invalid_wkt = sum(geometry is None for geometry in parsed)
    invalid_polygon = sum(
        original is not None and fixed is None
        for original, fixed in zip(parsed, repaired)
    )
    valid_indexes = [
        index
        for index, (reference, geometry) in enumerate(zip(references, repaired))
        if reference and geometry is not None
    ]
    frame = gpd.GeoDataFrame(
        {
            "REF_BACEN": [references[index] for index in valid_indexes],
            "plot_uid": row_offset + np.asarray(valid_indexes, dtype=np.int64),
            "geometry": [repaired[index] for index in valid_indexes],
        },
        geometry="geometry",
    )
    return frame, invalid_wkt, invalid_polygon


def _select_albers_equal_area_crs() -> CRS:
    try:
        return CRS.from_epsg(ALBERS_BRAZIL_EPSG)
    except Exception:
        return CRS.from_user_input(ALBERS_SOUTH_AMERICA_FALLBACK)


def automatic_area_crs(
    base_paths: list[Path],
    layer_name: str | None,
) -> CRS:
    """Escolhe UTM SIRGAS 2000 até 6° de longitude e Albers Equal Area acima de 6°."""
    LOGGER.info("Lendo a extensão da Camada Base para selecionar o CRS de área automaticamente.")
    geographic_bounds: list[tuple[float, float, float, float]] = []
    for path in base_paths:
        ext = path.suffix.lower()
        if ext == ".parquet":
            gdf_pq = read_geoparquet(path, default_crs="EPSG:4674")
            if gdf_pq.empty or gdf_pq.crs is None:
                raise ValueError(f"O GeoParquet {path} não possui geometrias ou CRS válidos.")
            layer_crs = gdf_pq.crs
            layer_bounds = gdf_pq.total_bounds
            selected_layer = path.stem
        elif ext in {".geojson", ".json"}:
            selected_layer = path.stem
            layer_info = pyogrio.read_info(path)
            layer_crs = layer_info["crs"] or "EPSG:4674"
            layer_bounds = layer_info["total_bounds"]
        else:
            selected_layer = resolve_indigenous_layer(path, layer_name)
            layer_info = pyogrio.read_info(path, layer=selected_layer)
            if layer_info["crs"] is None:
                raise ValueError(f"A camada {selected_layer} em {path} não possui CRS definido.")
            layer_crs = layer_info["crs"]
            layer_bounds = layer_info["total_bounds"]

        if layer_bounds is None or not all(math.isfinite(float(value)) for value in layer_bounds):
            raise ValueError(
                f"Não foi possível determinar a extensão da camada {selected_layer} em {path}."
            )
        geographic_bounds.append(
            Transformer.from_crs(
                CRS.from_user_input(layer_crs), "EPSG:4326", always_xy=True
            ).transform_bounds(*[float(v) for v in layer_bounds], densify_pts=21)
        )

    if not geographic_bounds:
        raise ValueError(
            "Não foi possível determinar o CRS automaticamente: "
            "nenhuma extensão válida foi encontrada na Camada Base."
        )

    west = min(bound[0] for bound in geographic_bounds)
    south = min(bound[1] for bound in geographic_bounds)
    east = max(bound[2] for bound in geographic_bounds)
    north = max(bound[3] for bound in geographic_bounds)
    if not all(math.isfinite(value) for value in (west, south, east, north)):
        raise ValueError("Não foi possível transformar a extensão da Camada Base para EPSG:4326.")

    if east - west > 6.0:
        selected_crs = _select_albers_equal_area_crs()
        LOGGER.info(
            "A extensão da Camada Base cobre %.2f graus de longitude (> 6°); "
            "selecionado %s (%s) para cálculo equivalente de áreas.",
            east - west,
            selected_crs.to_string(),
            selected_crs.name,
        )
        return selected_crs

    center_longitude = (west + east) / 2
    center_latitude = (south + north) / 2
    if not -80 <= center_latitude <= 84:
        raise ValueError(
            "A extensão da Camada Base está fora da cobertura UTM; "
            "informe um CRS de área manualmente."
        )
    zone = min(60, max(1, math.floor((center_longitude + 180) / 6) + 1))
    hemisphere = "S" if center_latitude < 0 else "N"
    expected_name = f"SIRGAS 2000 / UTM zone {zone}{hemisphere}"
    matching_crs = next(
        (
            info
            for info in query_crs_info(auth_name="EPSG", pj_types="PROJECTED_CRS")
            if info.name == expected_name
        ),
        None,
    )
    if matching_crs is None:
        raise ValueError(
            f"Não há CRS EPSG SIRGAS 2000 / UTM disponível para a zona {zone}{hemisphere}; "
            "informe um CRS de área manualmente."
        )
    selected_crs = CRS.from_epsg(int(matching_crs.code))
    LOGGER.info(
        "CRS de área selecionado automaticamente: %s (%s).",
        selected_crs.to_string(),
        selected_crs.name,
    )
    return selected_crs


def find_overlaps(
    plots: gpd.GeoDataFrame, lands: gpd.GeoDataFrame
) -> list[dict[str, object]]:
    """Usa índice espacial para calcular área de interseção sem dupla contagem."""
    if plots.empty:
        return []
    candidates = gpd.sjoin(
        plots[["plot_uid", "REF_BACEN", "geometry"]],
        lands[["ti_name", "geometry"]],
        how="inner",
        predicate="intersects",
    )
    if candidates.empty:
        return []

    results: list[dict[str, object]] = []
    land_geometries = lands.geometry
    land_names = lands["ti_name"]
    plot_geometries = plots.set_index("plot_uid").geometry
    linear_unit_to_m = plots.crs.axis_info[0].unit_conversion_factor

    for plot_uid, matches in candidates.groupby("plot_uid", sort=False):
        land_indexes = matches["index_right"].to_numpy(dtype=np.int64)
        geometry = plot_geometries.loc[plot_uid]
        intersections = shapely.intersection(
            geometry, land_geometries.iloc[land_indexes].to_numpy()
        )
        areas = shapely.area(intersections)
        positive = areas > 0
        if not positive.any():
            continue

        # A união evita contar duas vezes áreas onde polígonos de TI se sobrepõem.
        overlap_area = float(shapely.union_all(intersections[positive]).area) * linear_unit_to_m**2
        total_area = float(geometry.area) * linear_unit_to_m**2
        if overlap_area <= 0 or total_area <= 0:
            continue
        overlap_area = min(overlap_area, total_area)
        names = sorted(
            {
                str(name).strip()
                for name in land_names.iloc[land_indexes[positive]].tolist()
                if pd.notna(name) and str(name).strip()
            }
        )
        results.append(
            {
                "plot_uid": int(plot_uid),
                "REF_BACEN": str(matches.iloc[0]["REF_BACEN"]),
                "AREA_TOTAL_HA": total_area / 10_000,
                "AREA_SOBREPOSTA_HA": overlap_area / 10_000,
                "PERCENTUAL_SOBREPOSICAO": overlap_area / total_area * 100,
                "NOME_TERRA_INDIGENA": " | ".join(names) if names else pd.NA,
                "geometry": geometry,
            }
        )
    return results


def parse_financial_values(values: pd.Series) -> pd.Series:
    """Converte valores monetários em número, aceitando decimal com ponto ou vírgula."""
    text = values.astype("string").str.strip()
    both_separators = text.str.contains(",", regex=False, na=False) & text.str.contains(
        ".", regex=False, na=False
    )
    comma_is_decimal = both_separators & (
        text.str.rfind(",") > text.str.rfind(".")
    )
    dot_is_decimal = both_separators & ~comma_is_decimal
    comma_only = text.str.contains(",", regex=False, na=False) & ~text.str.contains(
        ".", regex=False, na=False
    )
    text.loc[comma_is_decimal] = (
        text.loc[comma_is_decimal].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    )
    text.loc[dot_is_decimal] = text.loc[dot_is_decimal].str.replace(",", "", regex=False)
    text.loc[comma_only] = text.loc[comma_only].str.replace(",", ".", regex=False)
    return pd.to_numeric(text, errors="coerce")


def _polars_parse_financial_expr(col_name: str) -> pl.Expr:
    """Converte texto monetário (ponto ou vírgula decimal) para Float64 em Polars."""
    text = pl.col(col_name).cast(pl.String).str.strip_chars()
    has_comma = text.str.contains(",", literal=True)
    has_dot = text.str.contains(".", literal=True)
    comma_is_decimal = has_comma & has_dot & text.str.contains(r",[^.]*$")
    dot_is_decimal = has_comma & has_dot & ~comma_is_decimal
    comma_only = has_comma & ~has_dot
    normalized = (
        pl.when(comma_is_decimal)
        .then(
            text.str.replace_all(".", "", literal=True).str.replace_all(
                ",", ".", literal=True
            )
        )
        .when(dot_is_decimal)
        .then(text.str.replace_all(",", "", literal=True))
        .when(comma_only)
        .then(text.str.replace_all(",", ".", literal=True))
        .otherwise(text)
    )
    return normalized.replace("", None).cast(pl.Float64, strict=False)


def _polars_to_pandas(df: pl.DataFrame) -> pd.DataFrame:
    """Converte DataFrame Polars para Pandas sem depender de pyarrow."""
    if df.width == 0:
        return pd.DataFrame()
    if df.is_empty():
        return pd.DataFrame(columns=list(df.columns))
    return pd.DataFrame({col: df.get_column(col).to_list() for col in df.columns})


def read_matching_operations(
    path: Path,
    delimiter: str,
    encoding: str,
    reference_column: str | None,
    finance_column: str | None,
    references: set[str],
    chunk_size: int,
    key_column: str = "REF_BACEN",
) -> tuple[pd.DataFrame, float, int]:
    """Lê operações (.csv/.parquet) em Polars streaming e retém só referências sobrepostas."""
    if is_parquet_path(path):
        header = list(pl.read_parquet_schema(path).keys())
    else:
        header = read_csv_header(path, delimiter, encoding)
    validate_unique_columns(header, path)
    original_ref = resolve_column(
        header, reference_column, (key_column,), f"{key_column} das operações"
    )
    original_finance = resolve_finance_column(header, finance_column)
    if original_ref is None:
        raise ValueError(f"A tabela de operações deve conter a coluna {key_column}.")

    empty_columns = [normalize_column_name(column) for column in header]
    ref_norm = normalize_column_name(original_ref)
    if ref_norm != key_column:
        empty_columns = [
            key_column if column == ref_norm else column for column in empty_columns
        ]

    if not references:
        empty_df = pd.DataFrame(columns=empty_columns)
        if original_finance is not None:
            fin_norm = normalize_column_name(original_finance)
            if fin_norm in empty_df.columns:
                empty_df[fin_norm] = empty_df[fin_norm].astype("float64")
        return empty_df, 0.0, 0

    sorted_refs = sorted(references)
    finance_norm = (
        normalize_column_name(original_finance) if original_finance is not None else None
    )

    def _collect_batches(scan_path: Path, scan_encoding: str | None) -> tuple[list[pl.DataFrame], int]:
        if is_parquet_path(scan_path):
            query = pl.scan_parquet(scan_path)
        else:
            query = pl.scan_csv(
                scan_path,
                separator=delimiter,
                encoding=polars_encoding(scan_encoding or encoding),
                infer_schema_length=0,
                low_memory=True,
            )
        filtered = query.filter(
            pl.col(str(original_ref)).cast(pl.String).str.strip_chars().is_in(sorted_refs)
        )
        batches: list[pl.DataFrame] = []
        invalid_count = 0
        for batch in filtered.collect_batches(chunk_size=chunk_size, engine="streaming"):
            if batch.is_empty():
                continue
            rename_map = {col: normalize_column_name(col) for col in batch.columns}
            batch = batch.rename(rename_map)
            if ref_norm != key_column and ref_norm in batch.columns and key_column not in batch.columns:
                batch = batch.rename({ref_norm: key_column})
            batch = batch.with_columns(
                pl.col(key_column).cast(pl.String).str.strip_chars().replace("", None)
            )
            if finance_norm is not None and finance_norm in batch.columns:
                if batch.schema[finance_norm].is_numeric():
                    parsed_expr = pl.col(finance_norm).cast(pl.Float64, strict=False)
                else:
                    parsed_expr = _polars_parse_financial_expr(finance_norm)
                raw_text = pl.col(finance_norm).cast(pl.String).str.strip_chars()
                invalid_in_batch = batch.select(
                    (raw_text.is_not_null() & (raw_text != "") & parsed_expr.is_null()).sum()
                ).item()
                invalid_count += int(invalid_in_batch or 0)
                batch = batch.with_columns(parsed_expr.alias(finance_norm))
            batches.append(batch)
        return batches, invalid_count

    matched_batches: list[pl.DataFrame] = []
    invalid_financial_values = 0
    if is_parquet_path(path):
        matched_batches, invalid_financial_values = _collect_batches(path, None)
    else:
        try:
            with polars_compatible_csv(path, encoding) as csv_path:
                matched_batches, invalid_financial_values = _collect_batches(
                    csv_path, encoding
                )
        except pl.exceptions.ComputeError as error:
            if (
                not is_utf8_encoding(encoding)
                or "lossy" in encoding.lower()
                or ("utf" not in str(error).lower() and "encoding" not in str(error).lower())
            ):
                raise
            LOGGER.warning(
                "Caractere incompatível com UTF-8 após a amostra em %s; "
                "reprocessando operações com fallback cp1252.",
                path.name,
            )
            with polars_compatible_csv(path, "cp1252") as csv_path:
                matched_batches, invalid_financial_values = _collect_batches(
                    csv_path, "cp1252"
                )

    if matched_batches:
        combined_pl = pl.concat(matched_batches, how="vertical_relaxed")
        if finance_norm is not None and finance_norm in combined_pl.columns:
            total_val = combined_pl.select(pl.col(finance_norm).sum()).item()
            total_financed = (
                float(total_val)
                if total_val is not None and math.isfinite(float(total_val))
                else 0.0
            )
        else:
            total_financed = 0.0
        operations = _polars_to_pandas(combined_pl)
        operations[key_column] = _normalize_key_series(operations[key_column])
        if finance_norm is not None and finance_norm in operations.columns:
            operations[finance_norm] = pd.to_numeric(
                operations[finance_norm], errors="coerce"
            ).astype("float64")
    else:
        operations = pd.DataFrame(columns=empty_columns)
        if finance_norm is not None and finance_norm in operations.columns:
            operations[finance_norm] = operations[finance_norm].astype("float64")
        total_financed = 0.0
    return operations, total_financed, invalid_financial_values


def _normalize_key_series(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .replace({"": pd.NA, "<NA>": pd.NA})
    )


def _find_column_casefold(columns: Iterable[str], target_casefold: str) -> str | None:
    for column in columns:
        if normalize_column_name(column).casefold() == target_casefold:
            return column
    return None


def _schema_signature(columns: Iterable[str], key_column: str = "REF_BACEN") -> frozenset[str]:
    excluded = {normalize_column_name(key_column).casefold(), "nu_ordem"}
    return frozenset(
        normalize_column_name(column).casefold()
        for column in columns
        if normalize_column_name(column).casefold() not in excluded
    )


def _schemas_match(sig_a: frozenset[str], sig_b: frozenset[str]) -> bool:
    if not sig_a and not sig_b:
        return True
    if not sig_a or not sig_b:
        return False
    return len(sig_a & sig_b) / len(sig_a | sig_b) >= 0.5


def _aggregate_auxiliary_table(
    frame: pd.DataFrame, join_keys: list[str], source_column: str
) -> pd.DataFrame:
    """Agrega tabelas auxiliares por chave em Polars preservando ordem e tipos numéricos."""
    if frame.empty:
        return frame.copy()

    ordered = frame.copy()
    for key in join_keys:
        ordered[key] = _normalize_key_series(ordered[key])

    cd_primeiro_col = _find_column_casefold(ordered.columns, "cd_primeiro")
    value_columns = [column for column in ordered.columns if column not in join_keys]
    if source_column not in value_columns and source_column in ordered.columns:
        value_columns.append(source_column)
    numeric_cols = {
        col for col in value_columns if pd.api.types.is_numeric_dtype(ordered[col].dtype)
    }

    pl_df = _pandas_df_to_polars(ordered)
    if cd_primeiro_col is not None and cd_primeiro_col not in join_keys:
        pl_df = (
            pl_df.with_columns(
                (
                    pl.col(cd_primeiro_col)
                    .cast(pl.String)
                    .str.strip_chars()
                    .str.to_uppercase()
                    != "S"
                )
                .fill_null(True)
                .alias("__priority_sort")
            )
            .sort("__priority_sort", maintain_order=True)
            .drop("__priority_sort")
        )

    if not value_columns:
        agg_pl = pl_df.unique(subset=join_keys, maintain_order=True).select(join_keys)
        aggregated = _polars_to_pandas(agg_pl)
        for key in join_keys:
            aggregated[key] = _normalize_key_series(aggregated[key])
        return aggregated

    agg_exprs = [
        pl.col(col)
        .cast(pl.String)
        .str.strip_chars()
        .replace({"": None, "<NA>": None, "nan": None, "None": None})
        .drop_nulls()
        .unique(maintain_order=True)
        .str.join(" | ")
        .replace("", None)
        .alias(col)
        for col in value_columns
    ]
    agg_pl = pl_df.group_by(join_keys, maintain_order=True).agg(agg_exprs)
    aggregated = _polars_to_pandas(agg_pl)
    for key in join_keys:
        aggregated[key] = _normalize_key_series(aggregated[key])
    for col in numeric_cols:
        if col in aggregated.columns:
            non_null = aggregated[col].dropna().astype(str)
            if not non_null.str.contains(r" \| ", regex=True).any():
                aggregated[col] = pd.to_numeric(aggregated[col], errors="coerce")
    return aggregated


def _merge_source_labels(left: pd.Series, right: pd.Series) -> pd.Series:
    merged_labels: list[object] = []
    for left_val, right_val in zip(left, right):
        parts: list[str] = []
        seen: set[str] = set()
        for raw in (left_val, right_val):
            if pd.isna(raw):
                continue
            for piece in str(raw).split(" | "):
                token = piece.strip()
                if token and token != "<NA>" and token not in seen:
                    seen.add(token)
                    parts.append(token)
        merged_labels.append(" | ".join(parts) if parts else pd.NA)
    return pd.Series(merged_labels, index=left.index, dtype="string")


def consolidate_operation_tables(
    operation_frames: list[tuple[pd.DataFrame, str]],
    source_column: str,
    key_column: str = "REF_BACEN",
) -> pd.DataFrame:
    """Empilha tabelas de mesmo esquema e consolida tabelas auxiliares pela chave."""
    groups: list[dict[str, object]] = []
    for index, (operations, source_label) in enumerate(operation_frames):
        tagged = operations.copy()
        tagged[source_column] = source_label
        signature = _schema_signature(operations.columns, key_column=key_column)
        matched_group = next(
            (
                group
                for group in groups
                if _schemas_match(group["signature"], signature)  # type: ignore[arg-type]
            ),
            None,
        )
        if matched_group is not None:
            matched_group["frames"].append(tagged)  # type: ignore[attr-defined]
            matched_group["signature"] = (
                matched_group["signature"] | signature  # type: ignore[operator]
            )
        else:
            groups.append(
                {
                    "signature": signature,
                    "frames": [tagged],
                    "first_index": index,
                }
            )

    if len(groups) == 1:
        return pd.concat(groups[0]["frames"], ignore_index=True, sort=False)  # type: ignore[arg-type]

    primary_markers = {
        "vl_parc_credito",
        "valor_financiado",
        "vl_financiado",
        "valor_credito",
        "valor_total",
        "valor",
        "dt_emissao",
        "cnpj_if",
    }

    def group_priority(group: dict[str, object]) -> tuple[bool, int, int]:
        signature = group["signature"]  # type: ignore[assignment]
        has_primary = bool(signature & primary_markers)
        return (has_primary, len(signature), -int(group["first_index"]))  # type: ignore[arg-type]

    primary_group = max(groups, key=group_priority)
    group_dfs = [
        pd.concat(group["frames"], ignore_index=True, sort=False)  # type: ignore[arg-type]
        for group in groups
    ]
    primary_pos = groups.index(primary_group)

    merged = group_dfs[primary_pos].copy()
    merged[key_column] = _normalize_key_series(merged[key_column])

    for pos, aux_raw in enumerate(group_dfs):
        if pos == primary_pos:
            continue
        aux = aux_raw.copy()
        aux[key_column] = _normalize_key_series(aux[key_column])

        merged_nu_col = _find_column_casefold(merged.columns, "nu_ordem")
        aux_nu_col = _find_column_casefold(aux.columns, "nu_ordem")
        use_nu_ordem = False
        if merged_nu_col is not None and aux_nu_col is not None:
            aux[aux_nu_col] = _normalize_key_series(aux[aux_nu_col])
            if aux[aux_nu_col].notna().any():
                use_nu_ordem = True
                if aux_nu_col != merged_nu_col:
                    aux = aux.rename(columns={aux_nu_col: merged_nu_col})
                merged[merged_nu_col] = _normalize_key_series(merged[merged_nu_col])

        join_keys = [key_column, merged_nu_col] if use_nu_ordem and merged_nu_col else [key_column]
        aux_source_col = unique_column_name(
            "__aux_source", set(merged.columns) | set(aux.columns)
        )
        aux = aux.rename(columns={source_column: aux_source_col})

        # Evita sobrescrever ou duplicar colunas já existentes em merged.
        existing_cols = set(merged.columns) | {aux_source_col}
        join_key_set = set(join_keys)
        aux_renames: dict[str, str] = {}
        for column in aux.columns:
            if column in join_key_set or column == aux_source_col:
                continue
            if _find_column_casefold(merged.columns, normalize_column_name(column).casefold()) is not None:
                replacement = unique_column_name(column, existing_cols)
                aux_renames[column] = replacement
                existing_cols.add(replacement)
        if aux_renames:
            aux = aux.rename(columns=aux_renames)

        aux_value_cols = [
            column
            for column in aux.columns
            if column not in join_key_set and column != aux_source_col
        ]

        if aux.empty:
            for column in aux_value_cols:
                if column not in merged.columns:
                    merged[column] = pd.NA
            continue

        existing_refs = set(merged[key_column].dropna().tolist())
        in_merged = aux[key_column].isin(existing_refs)
        aux_existing = aux.loc[in_merged].copy()
        aux_new = aux.loc[~in_merged].copy()

        if not aux_existing.empty:
            agg_exact = _aggregate_auxiliary_table(aux_existing, join_keys, aux_source_col)
            merged = merged.merge(agg_exact, on=join_keys, how="left")

            if len(join_keys) > 1:
                unmatched_mask = merged[aux_source_col].isna() & merged[key_column].isin(
                    set(aux_existing[key_column].dropna().tolist())
                )
                if unmatched_mask.any():
                    aux_for_ref = aux_existing.drop(
                        columns=[key for key in join_keys if key != key_column],
                        errors="ignore",
                    )
                    agg_ref = _aggregate_auxiliary_table(
                        aux_for_ref, [key_column], aux_source_col
                    ).set_index(key_column)
                    fallback_refs = merged.loc[unmatched_mask, key_column]
                    for column in [*aux_value_cols, aux_source_col]:
                        mapped = fallback_refs.map(agg_ref[column])
                        if merged[column].dtype != mapped.dtype:
                            merged[column] = merged[column].astype(object)
                            mapped = mapped.astype(object)
                        merged.loc[unmatched_mask, column] = mapped

            merged[source_column] = _merge_source_labels(
                merged[source_column], merged[aux_source_col]
            )
            merged = merged.drop(columns=[aux_source_col])
        else:
            for column in aux_value_cols:
                if column not in merged.columns:
                    merged[column] = pd.NA

        if not aux_new.empty:
            aux_new_agg = _aggregate_auxiliary_table(aux_new, join_keys, aux_source_col)
            aux_new_agg = aux_new_agg.rename(columns={aux_source_col: source_column})
            merged = pd.concat([merged, aux_new_agg], ignore_index=True, sort=False)
            merged[key_column] = _normalize_key_series(merged[key_column])

    return merged


def _excel_cell_value(value: object) -> object:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        val = float(value)
        return val if math.isfinite(val) else None
    if isinstance(value, np.bool_):
        return bool(value)
    return value


def write_excel(
    output_path: Path,
    results: pd.DataFrame,
    summary: pd.DataFrame,
    intersections: pd.DataFrame | None = None,
) -> None:
    """Grava as planilhas em modo streaming (write_only) com cabeçalho congelado e filtros."""
    export = results.drop(columns=["plot_uid", "land_uids"], errors="ignore")
    if len(export) + 1 > EXCEL_MAX_ROWS:
        raise ValueError(
            f"O resultado possui {len(export):,} linhas e excede o limite do Excel "
            f"({EXCEL_MAX_ROWS - 1:,} linhas por aba). Reduza/divida a exportação."
        )
    inter_export: pd.DataFrame | None = None
    if intersections is not None:
        inter_export = intersections.drop(
            columns=["plot_uid", "land_uid", "intersection_uid", "geometry"],
            errors="ignore",
        )
        if len(inter_export) + 1 > EXCEL_MAX_ROWS:
            raise ValueError(
                f"O detalhamento 1:N possui {len(inter_export):,} linhas e excede o limite do Excel "
                f"({EXCEL_MAX_ROWS - 1:,} linhas por aba)."
            )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook(write_only=True)
    details_sheet = workbook.create_sheet("Sobreposicoes")
    details_sheet.freeze_panes = "A2"
    headers = [str(column) for column in export.columns]
    details_sheet.append(headers)
    for row in export.itertuples(index=False, name=None):
        details_sheet.append([_excel_cell_value(value) for value in row])
    if not export.empty and headers:
        details_sheet.auto_filter.ref = (
            f"A1:{get_column_letter(len(headers))}{len(export) + 1}"
        )

    if inter_export is not None:
        inter_sheet = workbook.create_sheet("Intersecoes_Detalhe")
        inter_sheet.freeze_panes = "A2"
        inter_headers = [str(column) for column in inter_export.columns]
        inter_sheet.append(inter_headers)
        for row in inter_export.itertuples(index=False, name=None):
            inter_sheet.append([_excel_cell_value(value) for value in row])
        if not inter_export.empty and inter_headers:
            inter_sheet.auto_filter.ref = (
                f"A1:{get_column_letter(len(inter_headers))}{len(inter_export) + 1}"
            )

    summary_sheet = workbook.create_sheet("Resumo")
    summary_sheet.freeze_panes = "A2"
    summary_sheet.column_dimensions["A"].width = 48
    summary_sheet.column_dimensions["B"].width = 24
    summary_sheet.append(["Indicador", "Valor"])
    for row in summary.itertuples(index=False, name=None):
        summary_sheet.append([_excel_cell_value(value) for value in row])
    workbook.save(output_path)


def _prepare_geopackage_frame(
    frame: gpd.GeoDataFrame,
    crs: object,
    rename_map: dict[str, str],
    drop_columns: tuple[str, ...],
) -> gpd.GeoDataFrame:
    """Normaliza colunas e tipos do GeoDataFrame para gravação segura em GeoPackage."""
    export = gpd.GeoDataFrame(frame.copy(), geometry="geometry", crs=crs)
    export = export.drop(columns=list(drop_columns), errors="ignore")

    geometry_column = export.geometry.name
    used_names: set[str] = {"fid", "geom", geometry_column.casefold()}
    safe_renames: dict[str, str] = {}
    for column in export.columns:
        if column == geometry_column:
            continue
        target = rename_map.get(column, str(column))
        safe_name = unique_column_name(target, used_names)
        used_names.add(safe_name.casefold())
        if safe_name != column:
            safe_renames[column] = safe_name
    if safe_renames:
        export = export.rename(columns=safe_renames)

    for column in export.columns:
        if column == geometry_column:
            continue
        series = export[column]
        if pd.api.types.is_bool_dtype(series.dtype):
            export[column] = pd.Series(
                [None if pd.isna(value) else bool(value) for value in series],
                index=export.index,
                dtype=object,
            )
        elif pd.api.types.is_integer_dtype(series.dtype):
            if isinstance(series.dtype, pd.api.extensions.ExtensionDtype) and series.isna().any():
                export[column] = pd.Series(
                    [None if pd.isna(value) else int(value) for value in series],
                    index=export.index,
                    dtype=object,
                )
            else:
                export[column] = series.astype("int64")
        elif pd.api.types.is_float_dtype(series.dtype):
            export[column] = pd.to_numeric(series, errors="coerce").astype("float64")
        else:
            export[column] = pd.Series(
                [None if pd.isna(value) else str(value) for value in series],
                index=export.index,
                dtype=object,
            )
    return export


def write_geopackage(
    output_path: Path,
    results: gpd.GeoDataFrame,
    crs: object,
    intersections: gpd.GeoDataFrame | None = None,
) -> None:
    """Exporta glebas sobrepostas e, quando houver, a camada 1:N de interseções recortadas."""
    if output_path.suffix.lower() != ".gpkg":
        raise ValueError(f"O arquivo de saída geoespacial deve ter extensão .gpkg: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    plot_rename_map = {
        "REF_BACEN": "Referência BACEN (mascarada)",
        "ARQUIVO_GLEBAS": "Arquivo de glebas",
        "AREA_TOTAL_HA": "Área total (ha)",
        "AREA_SOBREPOSTA_HA": "Área sobreposta (ha)",
        "PERCENTUAL_SOBREPOSICAO": "Percentual de sobreposição (%)",
        "NOME_TERRA_INDIGENA": "Camada Base",
    }
    export = _prepare_geopackage_frame(
        results,
        crs,
        plot_rename_map,
        drop_columns=("plot_uid", "land_uids"),
    )
    export.to_file(
        output_path,
        layer="glebas_sobrepostas",
        driver="GPKG",
        engine="pyogrio",
        index=False,
    )

    inter_count = 0
    if intersections is not None and not intersections.empty:
        inter_rename_map = {
            "REF_BACEN": "Referência BACEN (mascarada)",
            "ARQUIVO_GLEBAS": "Arquivo de glebas",
            "CAMADA_BASE": "Camada Base",
            "AREA_TOTAL_HA": "Área total (ha)",
            "AREA_INTERSECAO_HA": "Área da interseção (ha)",
            "PERCENTUAL_INTERSECAO": "Percentual da interseção (%)",
        }
        inter_export = _prepare_geopackage_frame(
            intersections,
            crs,
            inter_rename_map,
            drop_columns=("plot_uid", "land_uid", "intersection_uid"),
        )
        inter_export.to_file(
            output_path,
            layer="intersecoes",
            driver="GPKG",
            engine="pyogrio",
            mode="a",
            index=False,
        )
        inter_count = len(inter_export)

    LOGGER.info(
        "GeoPackage exportado: %s (%d glebas, %d interseções 1:N).",
        output_path,
        len(export),
        inter_count,
    )


def write_preview_geojson(
    output_path: Path,
    lands: gpd.GeoDataFrame,
    plots: gpd.GeoDataFrame,
    limit: int = 500,
) -> None:
    """Gera uma prévia leve das glebas sobrepostas e da Camada Base relacionada."""
    sample_plots = plots.head(limit)
    if sample_plots.empty:
        sample_lands = lands.head(limit)
    else:
        related = gpd.sjoin(
            lands[["geometry"]],
            sample_plots[["geometry"]],
            how="inner",
            predicate="intersects",
        )
        sample_lands = lands.loc[related.index.unique()[:limit]]

    def features_as_wgs84(frame: gpd.GeoDataFrame) -> list[dict[str, object]]:
        if frame.empty:
            return []
        preview = frame.to_crs("EPSG:4326").copy()
        collection = json.loads(preview.to_json(drop_id=True, to_wgs84=True))
        return collection["features"]

    features: list[dict[str, object]] = []
    for layer_name, frame in (
        ("terras_indigenas", sample_lands),
        ("glebas", sample_plots),
    ):
        layer_features = features_as_wgs84(frame)
        for feature in layer_features:
            feature["properties"]["layer"] = layer_name
        features.extend(layer_features)

    payload = {
        "type": "FeatureCollection",
        "features": features,
        "counts": {
            "glebas_total": len(plots),
            "glebas_preview": len(sample_plots),
            "terras_indigenas_total": len(lands),
            "terras_indigenas_preview": len(sample_lands),
        },
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as destination:
        json.dump(payload, destination, ensure_ascii=False, separators=(",", ":"))


def _load_optional_institutions() -> dict[str, str]:
    """Lê opcionalmente a tabela SICOR_LISTA_IFS.csv se existir localmente."""
    path = Path(__file__).resolve().parent / "dados" / "SICOR_LISTA_IFS.csv"
    if not path.is_file():
        return {}
    try:
        encoding = resolve_csv_encoding(path, "auto")
        delimiter = infer_delimiter(path, encoding)
        with path.open("r", encoding=encoding, newline="") as source:
            reader = csv.DictReader(source, delimiter=delimiter)
            headers = {
                normalize_column_name(name).upper(): name
                for name in reader.fieldnames or []
            }
            cnpj_col = headers.get("CNPJ_IF")
            name_col = headers.get("NOME_IF")
            if cnpj_col is None or name_col is None:
                return {}
            return {
                str(row[cnpj_col]).strip(): str(row[name_col]).strip()
                for row in reader
                if row.get(cnpj_col) and row.get(name_col)
            }
    except Exception:
        LOGGER.debug("Lookup opcional de instituições indisponível.", exc_info=True)
        return {}


def _is_money_column_name(column_name: str) -> bool:
    cf = normalize_column_name(column_name).casefold()
    return any(
        token in cf
        for token in ("vl_", "valor", "credito", "parcela", "multa", "preco", "brl", "r$")
    )


def _sqlite_quote(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def _sqlite_cell_value(value: object) -> object:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        val = float(value)
        return val if math.isfinite(val) else None
    if isinstance(value, np.bool_):
        return bool(value)
    return value


def write_dashboard_database(
    output_path: Path,
    results: pd.DataFrame,
    summary: pd.DataFrame,
    lands: gpd.GeoDataFrame,
    plots: gpd.GeoDataFrame,
    plot_land_relations: dict[int, list[int]] | None = None,
    plot_centroids: dict[int, tuple[float, float]] | None = None,
    key_column: str = "REF_BACEN",
    finance_column: str | None = None,
    area_crs: object | None = None,
    has_operations: bool = True,
    intersections: gpd.GeoDataFrame | None = None,
    operations_source_column: str | None = None,
    min_area_ha: float = 0.0,
    min_percent: float = 0.0,
) -> None:
    """Persiste colunas consolidadas, metadados, detalhamento 1:N e geometrias no SQLite."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()

    work = results.drop(columns=["land_uids"], errors="ignore").copy()
    if "plot_uid" in work.columns and "PLOT_UID" not in work.columns:
        work = work.rename(columns={"plot_uid": "PLOT_UID"})

    # Garante nomes de colunas únicos ignorando maiúsculas/minúsculas (exigência do SQLite).
    used_sql_names: set[str] = set()
    sql_renames: dict[str, str] = {}
    for col in work.columns:
        clean_name = re.sub(r"[\r\n\t\x00]+", " ", str(col)).strip() or "COLUNA"
        unique_name = unique_column_name(clean_name, used_sql_names)
        used_sql_names.add(unique_name.casefold())
        if unique_name != col:
            sql_renames[col] = unique_name
    if sql_renames:
        work = work.rename(columns=sql_renames)

    def map_name(name: str | None) -> str | None:
        if name is None:
            return None
        if name == "plot_uid":
            return "PLOT_UID"
        return sql_renames.get(name, name)

    actual_key_col = map_name(key_column) or key_column
    actual_finance_col = map_name(finance_column)
    actual_ops_src_col = (
        map_name(operations_source_column)
        if has_operations and operations_source_column
        else ""
    )

    if "PLOT_UID" in work.columns:
        work["PLOT_UID"] = pd.to_numeric(work["PLOT_UID"], errors="coerce").astype("Int64")
    else:
        work["PLOT_UID"] = pd.Series(pd.NA, index=work.index, dtype="Int64")

    if plot_centroids is not None:
        work["CENTROID_LAT"] = work["PLOT_UID"].map(
            lambda uid: plot_centroids.get(int(uid), (None, None))[0] if pd.notna(uid) else None
        )
        work["CENTROID_LON"] = work["PLOT_UID"].map(
            lambda uid: plot_centroids.get(int(uid), (None, None))[1] if pd.notna(uid) else None
        )
    else:
        work["CENTROID_LAT"] = pd.Series(np.nan, index=work.index, dtype="float64")
        work["CENTROID_LON"] = pd.Series(np.nan, index=work.index, dtype="float64")

    # Encontra colunas conhecidas caso existam no resultado consolidado.
    effective_key_col = _find_column_casefold(work.columns, "ref_bacen_efetivo")
    date_start_col = (
        _find_column_casefold(work.columns, "dt_emissao")
        or _find_column_casefold(work.columns, "dt_contratacao")
    )
    date_end_col = _find_column_casefold(work.columns, "dt_vencimento")
    borrower_col = (
        _find_column_casefold(work.columns, "cd_cpf_cnpj")
        or _find_column_casefold(work.columns, "cpf_cnpj")
        or _find_column_casefold(work.columns, "cpf_mutuario")
    )
    municipio_col = _find_column_casefold(work.columns, "cd_ibge_municipio")
    uf_col = _find_column_casefold(work.columns, "cd_estado")
    institution_col = _find_column_casefold(work.columns, "cnpj_if")
    nu_ordem_col = _find_column_casefold(work.columns, "nu_ordem")
    existing_ano_col = _find_column_casefold(work.columns, "ano")

    for code_col in (municipio_col, institution_col, nu_ordem_col):
        if code_col is not None and code_col in work.columns:
            cleaned = (
                work[code_col]
                .astype("string")
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)
            )
            work[code_col] = cleaned
    if uf_col is not None and uf_col in work.columns:
        work[uf_col] = work[uf_col].astype("string").str.strip().str.upper()

    ano_col = existing_ano_col
    if ano_col is None and date_start_col is not None:
        parsed_dates = pd.to_datetime(
            work[date_start_col], dayfirst=True, errors="coerce"
        )
        parsed_years = parsed_dates.dt.year.astype("Int64").astype("string")
        if int(parsed_years.notna().sum()) > 0:
            ano_col = unique_column_name("ANO", {c.casefold() for c in work.columns})
            work[ano_col] = parsed_years

    # Normaliza strings vazias para NA em colunas textuais.
    for col in work.columns:
        if pd.api.types.is_string_dtype(work[col].dtype) or work[col].dtype == object:
            series_str = work[col].astype("string").str.strip()
            work[col] = series_str.replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})

    role_labels = {
        "effective_key": "Referência BACEN explícita",
        "base_layer": "Camada Base",
        "area_total": "Área total da gleba (ha)",
        "area_overlap": "Área sobreposta (ha)",
        "overlap_pct": "Percentual de sobreposição (%)",
        "borrower": "CPF/CNPJ do mutuário",
        "date_start": "Data de contratação",
        "date_end": "Data de vencimento",
        "municipio": "Município (código IBGE)",
        "uf": "UF",
        "year": "Ano",
        "institution": "Instituição financeira",
        "plot_uid": "ID Gleba",
    }

    has_effective_key = (
        effective_key_col is not None and int(work[effective_key_col].notna().sum()) > 0
    )

    col_specs: list[dict[str, object]] = []
    for col in work.columns:
        cf = normalize_column_name(col).casefold()
        series = work[col]
        non_null = int(series.notna().sum())
        distinct = int(series.dropna().nunique()) if non_null > 0 else 0

        role = "generic"
        data_type = "text"
        label = str(col)
        default_visible = 0
        filterable = 0

        if col == "PLOT_UID":
            role = "plot_uid"
            data_type = "integer"
            label = role_labels["plot_uid"]
        elif col in {"CENTROID_LAT", "CENTROID_LON"}:
            role = "coord"
            data_type = "coord"
            label = "Latitude (centroide)" if col == "CENTROID_LAT" else "Longitude (centroide)"
        elif effective_key_col is not None and col == effective_key_col:
            role = "effective_key"
            data_type = "text"
            label = role_labels["effective_key"]
            default_visible = 1 if non_null > 0 else 0
        elif col == actual_key_col:
            role = "key"
            data_type = "text"
            label = (
                "Referência BACEN (mascarada)"
                if cf == "ref_bacen" and has_effective_key
                else str(col)
            )
            default_visible = 0 if has_effective_key else 1
        elif cf == "nome_terra_indigena":
            role = "base_layer"
            data_type = "text"
            label = role_labels["base_layer"]
            default_visible = 1
            filterable = 1
        elif cf == "area_total_ha":
            role = "area_total"
            data_type = "area"
            label = role_labels["area_total"]
            work[col] = pd.to_numeric(series, errors="coerce")
            default_visible = 1
        elif cf == "area_sobreposta_ha":
            role = "area_overlap"
            data_type = "area_full"
            label = role_labels["area_overlap"]
            work[col] = pd.to_numeric(series, errors="coerce")
            default_visible = 1
        elif cf == "percentual_sobreposicao":
            role = "overlap_pct"
            data_type = "percent_full"
            label = role_labels["overlap_pct"]
            work[col] = pd.to_numeric(series, errors="coerce")
            default_visible = 1
        elif actual_finance_col is not None and col == actual_finance_col:
            role = "finance"
            data_type = "currency" if _is_money_column_name(col) else "number"
            label = (
                "Valor da parcela de crédito (R$)"
                if cf in {"vl_parc_credito", "valor_financiado", "vl_financiado"}
                else str(col)
            )
            if not pd.api.types.is_numeric_dtype(series.dtype):
                work[col] = parse_financial_values(series)
            else:
                work[col] = pd.to_numeric(series, errors="coerce")
            default_visible = 1 if non_null > 0 else 0
        elif borrower_col is not None and col == borrower_col:
            role = "borrower"
            data_type = "text"
            label = role_labels["borrower"]
            default_visible = 1 if non_null > 0 else 0
        elif date_start_col is not None and col == date_start_col:
            role = "date_start"
            data_type = "date"
            label = role_labels["date_start"]
            default_visible = 1 if non_null > 0 else 0
        elif date_end_col is not None and col == date_end_col:
            role = "date_end"
            data_type = "date"
            label = role_labels["date_end"]
            default_visible = 1 if non_null > 0 else 0
        elif municipio_col is not None and col == municipio_col:
            role = "municipio"
            data_type = "text"
            label = role_labels["municipio"]
            filterable = 1 if non_null > 0 else 0
        elif uf_col is not None and col == uf_col:
            role = "uf"
            data_type = "text"
            label = role_labels["uf"]
            filterable = 1 if non_null > 0 else 0
        elif ano_col is not None and col == ano_col:
            role = "year"
            data_type = "text"
            label = role_labels["year"]
            filterable = 1 if non_null > 0 else 0
        elif institution_col is not None and col == institution_col:
            role = "institution"
            data_type = "text"
            label = role_labels["institution"]
            default_visible = 1 if non_null > 0 else 0
            filterable = 1 if non_null > 0 else 0
        elif cf == "arquivo_glebas" or cf.startswith("arquivo_operacoes"):
            role = "source_file"
            data_type = "text"
            label = "Arquivo de glebas" if cf == "arquivo_glebas" else "Tabelas relacionadas"
        else:
            if pd.api.types.is_integer_dtype(series.dtype):
                data_type = "integer"
            elif pd.api.types.is_float_dtype(series.dtype):
                data_type = "currency" if _is_money_column_name(col) else "number"
            elif non_null > 0:
                non_null_vals = series.dropna().astype(str).str.strip()
                has_leading_zero_codes = any(
                    LEADING_ZERO_ID_PATTERN.match(val) for val in non_null_vals.iloc[:200]
                )
                if not has_leading_zero_codes:
                    parsed_nums = parse_financial_values(non_null_vals)
                    if parsed_nums.notna().all():
                        work[col] = parse_financial_values(series)
                        data_type = "currency" if _is_money_column_name(col) else "number"

        col_specs.append(
            {
                "col_name": str(col),
                "label": label,
                "data_type": data_type,
                "semantic_role": role,
                "default_visible": default_visible,
                "filterable": filterable,
                "non_null_count": non_null,
                "distinct_count": distinct,
            }
        )

    # Para tabelas genéricas, promove colunas úteis com dados até completar 9 visíveis.
    visible_count = sum(1 for spec in col_specs if spec["default_visible"] == 1)
    if visible_count < 9:
        for spec in col_specs:
            if visible_count >= 9:
                break
            if (
                spec["default_visible"] == 0
                and spec["semantic_role"] in {"generic", "uf", "municipio", "year"}
                and int(spec["non_null_count"]) > 0
            ):
                spec["default_visible"] = 1
                visible_count += 1

    # Auto-seleciona até 4 filtros categóricos úteis fora do padrão SICOR.
    generic_filter_candidates = [
        spec
        for spec in col_specs
        if spec["semantic_role"] == "generic"
        and spec["data_type"] == "text"
        and int(spec["non_null_count"]) > 0
        and 2 <= int(spec["distinct_count"]) <= 200
    ]
    generic_filter_candidates.sort(key=lambda spec: (int(spec["distinct_count"]) > 50, 0))
    for spec in generic_filter_candidates[:4]:
        spec["filterable"] = 1

    role_priority = {
        "effective_key": 0,
        "key": 1,
        "borrower": 2,
        "finance": 3,
        "date_start": 4,
        "date_end": 5,
        "area_total": 6,
        "area_overlap": 7,
        "overlap_pct": 8,
        "base_layer": 9,
        "institution": 10,
    }

    def spec_sort_key(item: tuple[int, dict[str, object]]) -> tuple[int, int, int]:
        orig_idx, spec = item
        role = str(spec["semantic_role"])
        if role in role_priority and spec["default_visible"] == 1:
            return (0, role_priority[role], orig_idx)
        if spec["default_visible"] == 1:
            return (1, 0, orig_idx)
        if role not in {"coord", "plot_uid"}:
            return (2, role_priority.get(role, 99), orig_idx)
        return (3, 1 if role == "plot_uid" else 0, orig_idx)

    ordered_specs = [spec for _, spec in sorted(enumerate(col_specs), key=spec_sort_key)]
    for order_idx, spec in enumerate(ordered_specs):
        spec["col_order"] = order_idx

    ordered_col_names = [str(spec["col_name"]) for spec in ordered_specs]
    work = work[ordered_col_names]

    sqlite_type_map = {
        "integer": "INTEGER",
        "area": "REAL",
        "area_full": "REAL",
        "percent_full": "REAL",
        "number": "REAL",
        "currency": "REAL",
        "coord": "REAL",
    }
    column_defs = ", ".join(
        f"{_sqlite_quote(str(spec['col_name']))} {sqlite_type_map.get(str(spec['data_type']), 'TEXT')}"
        for spec in ordered_specs
    )

    institutions = _load_optional_institutions() if institution_col else {}

    # Prepara destaques do popup agrupados por PLOT_UID.
    excluded_highlight_roles = {
        "plot_uid",
        "key",
        "effective_key",
        "base_layer",
        "area_total",
        "area_overlap",
        "overlap_pct",
        "coord",
        "source_file",
    }
    highlight_candidates = sorted(
        [
            spec
            for spec in ordered_specs
            if spec["semantic_role"] not in excluded_highlight_roles
            and int(spec["non_null_count"]) > 0
        ],
        key=lambda spec: (
            {
                "finance": 0,
                "borrower": 1,
                "institution": 2,
                "uf": 3,
                "municipio": 4,
                "date_start": 5,
            }.get(str(spec["semantic_role"]), 10),
            int(spec["col_order"]),
        ),
    )

    plot_popup_meta: dict[int, dict[str, object]] = {}
    if not work.empty and "PLOT_UID" in work.columns:
        for uid_val, group in work.groupby("PLOT_UID", sort=False):
            if pd.isna(uid_val):
                continue
            uid_int = int(uid_val)
            first_row = group.iloc[0]
            eff_val = (
                first_row[effective_key_col]
                if effective_key_col is not None and pd.notna(first_row[effective_key_col])
                else None
            )
            key_val = (
                eff_val
                if eff_val is not None
                else (
                    first_row[actual_key_col]
                    if actual_key_col in group.columns and pd.notna(first_row[actual_key_col])
                    else None
                )
            )
            key_label = (
                role_labels["effective_key"]
                if eff_val is not None
                else str(actual_key_col)
            )
            highlights: list[dict[str, str]] = []
            for spec in highlight_candidates:
                if len(highlights) >= 5:
                    break
                cname = str(spec["col_name"])
                non_null_vals = group[cname].dropna().tolist()
                if not non_null_vals:
                    continue
                raw_val = non_null_vals[0]
                if spec["semantic_role"] == "institution":
                    display_val = institutions.get(str(raw_val), str(raw_val))
                else:
                    display_val = str(raw_val)
                highlights.append({"label": str(spec["label"]), "value": display_val})
            plot_popup_meta[uid_int] = {
                "key_label": key_label,
                "key_value": str(key_val) if key_val is not None else str(uid_int),
                "record_count": int(len(group)),
                "highlights": highlights,
            }

    def feature_json(
        layer: str, geometry: BaseGeometry, properties: dict[str, object]
    ) -> str:
        return json.dumps(
            {
                "type": "Feature",
                "geometry": mapping(geometry),
                "properties": {"layer": layer, **properties},
            },
            ensure_ascii=False,
            separators=(",", ":"),
        )

    with closing(sqlite3.connect(output_path)) as connection:
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA synchronous=OFF")
        connection.execute("PRAGMA temp_store=MEMORY")

        connection.execute(f"CREATE TABLE dashboard_rows ({column_defs})")
        if not work.empty:
            placeholders = ", ".join("?" for _ in ordered_col_names)
            insert_cols = ", ".join(_sqlite_quote(col) for col in ordered_col_names)
            insert_sql = f"INSERT INTO dashboard_rows ({insert_cols}) VALUES ({placeholders})"
            rows_iter = (
                tuple(_sqlite_cell_value(val) for val in row)
                for row in work.itertuples(index=False, name=None)
            )
            connection.executemany(insert_sql, rows_iter)

        connection.execute(
            "CREATE TABLE dashboard_columns ("
            "col_order INTEGER NOT NULL, "
            "col_name TEXT PRIMARY KEY, "
            "label TEXT NOT NULL, "
            "data_type TEXT NOT NULL, "
            "semantic_role TEXT NOT NULL, "
            "default_visible INTEGER NOT NULL, "
            "filterable INTEGER NOT NULL, "
            "non_null_count INTEGER NOT NULL, "
            "distinct_count INTEGER NOT NULL)"
        )
        connection.executemany(
            "INSERT INTO dashboard_columns ("
            "col_order, col_name, label, data_type, semantic_role, "
            "default_visible, filterable, non_null_count, distinct_count"
            ") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            [
                (
                    int(spec["col_order"]),
                    str(spec["col_name"]),
                    str(spec["label"]),
                    str(spec["data_type"]),
                    str(spec["semantic_role"]),
                    int(spec["default_visible"]),
                    int(spec["filterable"]),
                    int(spec["non_null_count"]),
                    int(spec["distinct_count"]),
                )
                for spec in ordered_specs
            ],
        )

        area_crs_text = ""
        if area_crs is not None:
            area_crs_text = (
                area_crs.to_string() if hasattr(area_crs, "to_string") else str(area_crs)
            )
        elif plots.crs is not None:
            area_crs_text = plots.crs.to_string()

        meta_entries = {
            "area_crs": area_crs_text,
            "key_column": actual_key_col,
            "effective_key_column": effective_key_col or "",
            "finance_column": actual_finance_col or "",
            "nu_ordem_column": nu_ordem_col or "",
            "institution_column": institution_col or "",
            "operations_source_column": actual_ops_src_col or "",
            "has_operations": "1" if has_operations else "0",
            "min_area_ha": str(min_area_ha),
            "min_percent": str(min_percent),
            "intersection_count": str(len(intersections) if intersections is not None else 0),
        }
        connection.execute(
            "CREATE TABLE dashboard_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)"
        )
        connection.executemany(
            "INSERT INTO dashboard_meta (key, value) VALUES (?, ?)",
            list(meta_entries.items()),
        )

        # Índices para busca, agrupamento e filtros sem full table scan.
        indexed_cols: list[str] = ["PLOT_UID", actual_key_col]
        if actual_ops_src_col:
            indexed_cols.append(actual_ops_src_col)
        for spec in ordered_specs:
            if spec["filterable"] == 1 or spec["semantic_role"] == "base_layer":
                indexed_cols.append(str(spec["col_name"]))
        for idx_num, col_name in enumerate(dict.fromkeys(indexed_cols)):
            if col_name in work.columns:
                connection.execute(
                    f"CREATE INDEX idx_dash_{idx_num} ON dashboard_rows ({_sqlite_quote(col_name)})"
                )

        connection.execute(
            "CREATE TABLE map_plots ("
            "PLOT_UID INTEGER PRIMARY KEY, GEOJSON TEXT NOT NULL, WKB BLOB NOT NULL)"
        )
        connection.execute(
            "CREATE TABLE map_lands ("
            "LAND_UID INTEGER PRIMARY KEY, TI_NAME TEXT NOT NULL, GEOJSON TEXT NOT NULL)"
        )
        connection.execute(
            "CREATE TABLE map_plot_lands ("
            "PLOT_UID INTEGER NOT NULL, LAND_UID INTEGER NOT NULL, "
            "PRIMARY KEY (PLOT_UID, LAND_UID))"
        )
        connection.execute(
            "CREATE TABLE dashboard_intersections ("
            "INTERSECTION_UID INTEGER PRIMARY KEY, "
            "PLOT_UID INTEGER NOT NULL, "
            "LAND_UID INTEGER NOT NULL, "
            "KEY_VALUE TEXT, "
            "ARQUIVO_GLEBAS TEXT, "
            "CAMADA_BASE TEXT NOT NULL, "
            "AREA_TOTAL_HA REAL, "
            "AREA_INTERSECAO_HA REAL, "
            "PERCENTUAL_INTERSECAO REAL)"
        )
        connection.execute(
            "CREATE TABLE map_intersections ("
            "INTERSECTION_UID INTEGER PRIMARY KEY, "
            "PLOT_UID INTEGER NOT NULL, "
            "LAND_UID INTEGER NOT NULL, "
            "TI_NAME TEXT NOT NULL, "
            "GEOJSON TEXT NOT NULL, "
            "WKB BLOB NOT NULL)"
        )

        if not plots.empty:
            plot_wgs84 = plots.to_crs("EPSG:4326")
            wkb_by_uid = {
                int(uid): shapely.to_wkb(geom)
                for uid, geom in zip(plots["plot_uid"], plots.geometry)
            }
            plot_rows_to_insert = []
            for row in plot_wgs84.itertuples(index=False):
                uid_int = int(row.plot_uid)
                ref_val = getattr(row, actual_key_col, None)
                if ref_val is None and key_column in plot_wgs84.columns:
                    ref_val = getattr(row, key_column, None)
                popup_meta = plot_popup_meta.get(uid_int, {})
                centroid = plot_centroids.get(uid_int) if plot_centroids else None
                props = {
                    "plot_uid": uid_int,
                    "REF_BACEN": str(ref_val) if ref_val is not None else str(uid_int),
                    "key_label": popup_meta.get("key_label", actual_key_col),
                    "key_value": popup_meta.get(
                        "key_value", str(ref_val) if ref_val is not None else str(uid_int)
                    ),
                    "camada_base": (
                        str(row.NOME_TERRA_INDIGENA)
                        if pd.notna(row.NOME_TERRA_INDIGENA)
                        else "Camada Base"
                    ),
                    "area_total_ha": float(row.AREA_TOTAL_HA),
                    "area_sobreposta_ha": float(row.AREA_SOBREPOSTA_HA),
                    "percentual_sobreposicao": float(row.PERCENTUAL_SOBREPOSICAO),
                    "centroid_lat": centroid[0] if centroid else None,
                    "centroid_lon": centroid[1] if centroid else None,
                    "record_count": popup_meta.get("record_count", 1),
                    "highlights": popup_meta.get("highlights", []),
                }
                plot_rows_to_insert.append(
                    (
                        uid_int,
                        feature_json("glebas", row.geometry, props),
                        wkb_by_uid[uid_int],
                    )
                )
            connection.executemany(
                "INSERT INTO map_plots (PLOT_UID, GEOJSON, WKB) VALUES (?, ?, ?)",
                plot_rows_to_insert,
            )

        if not lands.empty and not plots.empty:
            if plot_land_relations is not None:
                relations = sorted(
                    {
                        (int(plot_uid), int(land_uid))
                        for plot_uid, land_uids in plot_land_relations.items()
                        for land_uid in land_uids
                    }
                )
            else:
                land_frame = lands[["ti_name", "geometry"]].reset_index(drop=True)
                plot_frame = plots[["plot_uid", "geometry"]].reset_index(drop=True)
                related = gpd.sjoin(
                    land_frame[["geometry"]],
                    plot_frame[["plot_uid", "geometry"]],
                    how="inner",
                    predicate="intersects",
                )
                relations = sorted(
                    {
                        (int(row.plot_uid), int(land_uid))
                        for land_uid, row in related.iterrows()
                    }
                )

            if relations:
                connection.executemany(
                    "INSERT INTO map_plot_lands (PLOT_UID, LAND_UID) VALUES (?, ?)",
                    relations,
                )
                relevant_land_uids = sorted({land_uid for _, land_uid in relations})
                land_features = lands[["ti_name", "geometry"]].iloc[relevant_land_uids].to_crs("EPSG:4326")
                connection.executemany(
                    "INSERT INTO map_lands (LAND_UID, TI_NAME, GEOJSON) VALUES (?, ?, ?)",
                    (
                        (
                            int(land_uid),
                            str(name) if pd.notna(name) and str(name).strip() else "Camada Base",
                            feature_json(
                                "camada_base",
                                geometry,
                                {
                                    "ti_name": (
                                        str(name)
                                        if pd.notna(name) and str(name).strip()
                                        else "Camada Base"
                                    )
                                },
                            ),
                        )
                        for land_uid, name, geometry in land_features[
                            ["ti_name", "geometry"]
                        ].itertuples(index=True, name=None)
                    ),
                )

        if intersections is not None and not intersections.empty:
            inter_wgs84 = intersections.to_crs("EPSG:4326")
            inter_dash_rows = []
            inter_map_rows = []
            key_vals = (
                intersections[key_column].tolist()
                if key_column in intersections.columns
                else [None] * len(intersections)
            )
            source_vals = (
                intersections["ARQUIVO_GLEBAS"].tolist()
                if "ARQUIVO_GLEBAS" in intersections.columns
                else [None] * len(intersections)
            )
            for idx, (
                plot_uid_raw,
                land_uid_raw,
                raw_key,
                raw_source,
                land_name_raw,
                area_total_raw,
                area_inter_raw,
                pct_inter_raw,
                geom_area,
                geom_wgs84,
            ) in enumerate(
                zip(
                    intersections["plot_uid"].tolist(),
                    intersections["land_uid"].tolist(),
                    key_vals,
                    source_vals,
                    intersections["CAMADA_BASE"].tolist(),
                    intersections["AREA_TOTAL_HA"].tolist(),
                    intersections["AREA_INTERSECAO_HA"].tolist(),
                    intersections["PERCENTUAL_INTERSECAO"].tolist(),
                    intersections.geometry.tolist(),
                    inter_wgs84.geometry.tolist(),
                )
            ):
                inter_uid = idx + 1
                plot_uid_int = int(plot_uid_raw)
                land_uid_int = int(land_uid_raw)
                popup_meta = plot_popup_meta.get(plot_uid_int, {})
                fallback_key = (
                    str(raw_key)
                    if raw_key is not None and not pd.isna(raw_key) and str(raw_key).strip()
                    else str(plot_uid_int)
                )
                key_label = str(popup_meta.get("key_label", actual_key_col))
                key_value = str(popup_meta.get("key_value", fallback_key))
                land_name = (
                    str(land_name_raw).strip()
                    if land_name_raw is not None and not pd.isna(land_name_raw) and str(land_name_raw).strip()
                    else "Camada Base"
                )
                arquivo_gleba = (
                    str(raw_source)
                    if raw_source is not None and not pd.isna(raw_source) and str(raw_source).strip()
                    else None
                )
                area_total_f = float(area_total_raw) if pd.notna(area_total_raw) else 0.0
                area_inter_f = float(area_inter_raw) if pd.notna(area_inter_raw) else 0.0
                pct_inter_f = float(pct_inter_raw) if pd.notna(pct_inter_raw) else 0.0

                inter_dash_rows.append(
                    (
                        inter_uid,
                        plot_uid_int,
                        land_uid_int,
                        key_value,
                        arquivo_gleba,
                        land_name,
                        area_total_f,
                        area_inter_f,
                        pct_inter_f,
                    )
                )
                inter_props = {
                    "intersection_uid": inter_uid,
                    "plot_uid": plot_uid_int,
                    "land_uid": land_uid_int,
                    "key_label": key_label,
                    "key_value": key_value,
                    "camada_base": land_name,
                    "area_total_ha": area_total_f,
                    "area_intersecao_ha": area_inter_f,
                    "percentual_intersecao": pct_inter_f,
                }
                inter_map_rows.append(
                    (
                        inter_uid,
                        plot_uid_int,
                        land_uid_int,
                        land_name,
                        feature_json("intersecoes", geom_wgs84, inter_props),
                        shapely.to_wkb(geom_area),
                    )
                )

            connection.executemany(
                "INSERT INTO dashboard_intersections ("
                "INTERSECTION_UID, PLOT_UID, LAND_UID, KEY_VALUE, ARQUIVO_GLEBAS, "
                "CAMADA_BASE, AREA_TOTAL_HA, AREA_INTERSECAO_HA, PERCENTUAL_INTERSECAO"
                ") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                inter_dash_rows,
            )
            connection.executemany(
                "INSERT INTO map_intersections ("
                "INTERSECTION_UID, PLOT_UID, LAND_UID, TI_NAME, GEOJSON, WKB"
                ") VALUES (?, ?, ?, ?, ?, ?)",
                inter_map_rows,
            )

        connection.execute(
            "CREATE INDEX idx_map_plot_lands_land ON map_plot_lands (LAND_UID)"
        )
        connection.execute(
            "CREATE INDEX idx_map_lands_name ON map_lands (TI_NAME)"
        )
        connection.execute(
            "CREATE INDEX idx_dash_inter_plot ON dashboard_intersections (PLOT_UID)"
        )
        connection.execute(
            "CREATE INDEX idx_dash_inter_land ON dashboard_intersections (LAND_UID)"
        )
        connection.execute(
            "CREATE INDEX idx_dash_inter_base ON dashboard_intersections (CAMADA_BASE)"
        )
        connection.execute(
            "CREATE INDEX idx_map_inter_plot ON map_intersections (PLOT_UID)"
        )
        connection.execute(
            "CREATE INDEX idx_map_inter_name ON map_intersections (TI_NAME)"
        )
        summary.to_sql(
            "dashboard_summary",
            connection,
            if_exists="replace",
            index=False,
        )
        connection.commit()
        connection.execute("PRAGMA synchronous=NORMAL")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Identifica glebas que se sobrepõem à Camada Base e cruza operações."
    )
    parser.add_argument(
        "--glebas",
        required=True,
        action="append",
        type=Path,
        help="Arquivo de glebas (.csv, .gpkg, .geojson ou .parquet); repita para vários arquivos.",
    )
    parser.add_argument(
        "--operacoes",
        required=False,
        default=[],
        action="append",
        type=Path,
        help="Tabela auxiliar opcional (.csv ou .parquet) para junção com as glebas.",
    )
    parser.add_argument(
        "--terra-indigena",
        required=True,
        action="append",
        type=Path,
        dest="terras_indigenas",
        help="Camada Base (.gpkg, .geojson ou .parquet); repita a opção para múltiplos arquivos.",
    )
    parser.add_argument("--saida", required=True, type=Path, help="Arquivo XLSX de saída.")
    parser.add_argument(
        "--saida-gpkg",
        type=Path,
        help="Opcional: GeoPackage de glebas com sobreposição e interseções 1:N.",
    )
    parser.add_argument("--camada-ti", help="Nome da camada nos GeoPackages (padrão: primeira).")
    parser.add_argument("--coluna-ref", help="Coluna-chave de junção; padrão: detecta automaticamente.")
    parser.add_argument(
        "--coluna-geometria",
        help="Coluna de geometria WKT/WKB; detecta automaticamente por padrão.",
    )
    parser.add_argument("--coluna-nome-ti", help="Coluna com o nome da Camada Base.")
    parser.add_argument("--coluna-valor", help="Coluna com o valor numérico/financeiro.")
    parser.add_argument(
        "--crs-glebas",
        default="EPSG:4674",
        help="CRS das coordenadas WKT quando não houver CRS embutido (padrão: EPSG:4674).",
    )
    parser.add_argument(
        "--crs-area",
        default=DEFAULT_AREA_CRS,
        help=(
            "CRS projetado para áreas; padrão 'auto' usa SIRGAS 2000 UTM até 6° "
            "de longitude e Albers Equal Area (EPSG:10857) acima de 6°."
        ),
    )
    parser.add_argument(
        "--area-minima-ha",
        dest="area_minima_ha",
        type=float,
        default=0.0,
        help="Tolerância mínima de área sobreposta (ha) para descartar slivers (padrão: 0).",
    )
    parser.add_argument(
        "--percentual-minimo",
        dest="percentual_minimo",
        type=float,
        default=0.0,
        help="Tolerância mínima de percentual sobreposto (%) para descartar slivers (padrão: 0).",
    )
    parser.add_argument(
        "--codificacao",
        default="auto",
        help="Codificação dos CSVs; padrão: detecta UTF-8 e usa Windows-1252/Latin-1 como fallback.",
    )
    parser.add_argument("--delimitador", help="Delimitador dos CSVs; detectado se omitido.")
    parser.add_argument(
        "--tamanho-bloco", type=int, default=50_000, help="Linhas lidas por bloco."
    )
    parser.add_argument("--verbose", action="store_true", help="Ativa logs detalhados.")
    return parser.parse_args()


def _inspect_gleba_schema(
    path: Path,
    delimiter_override: str | None,
    encoding: str | None,
    requested_geometry: str | None,
) -> dict[str, object]:
    """Lê o esquema de um arquivo de glebas (.csv, .parquet, .gpkg, .geojson)."""
    if is_csv_path(path):
        effective_encoding = encoding or "auto"
        delimiter = delimiter_override or infer_delimiter(path, effective_encoding)
        header = read_csv_header(path, delimiter, effective_encoding)
        validate_unique_columns(header, path)
        geom_col = resolve_geometry_column(
            path, header, delimiter, effective_encoding, requested_geometry
        )
        return {
            "kind": "csv",
            "header": header,
            "delimiter": delimiter,
            "encoding": effective_encoding,
            "geometry_column": geom_col,
            "layer": None,
            "crs": None,
            "is_wkb": False,
        }

    if is_parquet_path(path):
        schema = pl.read_parquet_schema(path)
        meta = pl.read_parquet_metadata(path)
        header = list(schema.keys())
        validate_unique_columns(header, path)
        geo_meta: dict[str, object] = {}
        if "geo" in meta:
            try:
                geo_meta = json.loads(meta["geo"])
            except Exception:
                geo_meta = {}
        primary_col = geo_meta.get("primary_column") if isinstance(geo_meta, dict) else None
        if requested_geometry:
            geom_col = resolve_column(
                header, requested_geometry, DEFAULT_GEOMETRY_COLUMNS, "geometria"
            )
        elif isinstance(primary_col, str) and primary_col in schema:
            geom_col = primary_col
        else:
            geom_col = resolve_column(header, None, DEFAULT_GEOMETRY_COLUMNS, "geometria")
        geom_meta_dict = (
            geo_meta.get("columns", {}).get(str(geom_col), {}) or {}
            if isinstance(geo_meta.get("columns"), dict)
            else {}
        )
        resolved_crs = (
            _crs_from_geoparquet_meta(geom_meta_dict, "EPSG:4674")
            if geom_meta_dict.get("crs")
            else None
        )
        return {
            "kind": "parquet",
            "header": header,
            "delimiter": None,
            "encoding": None,
            "geometry_column": geom_col,
            "layer": None,
            "crs": resolved_crs,
            "is_wkb": schema[str(geom_col)] == pl.Binary,
        }

    if is_vector_path(path):
        if path.suffix.lower() == ".gpkg":
            available_layers = gpd.list_layers(path)
            if available_layers.empty:
                raise ValueError(f"Nenhuma camada encontrada em {path}")
            layer_name: str | None = str(available_layers.iloc[0]["name"])
            info = pyogrio.read_info(path, layer=layer_name)
        else:
            layer_name = None
            info = pyogrio.read_info(path)
        header = [str(col) for col in info.get("fields", [])]
        validate_unique_columns(header, path)
        raw_crs = info.get("crs")
        resolved_crs = CRS.from_user_input(raw_crs) if raw_crs else None
        return {
            "kind": "vector",
            "header": header,
            "delimiter": None,
            "encoding": None,
            "geometry_column": "geometry",
            "layer": layer_name,
            "crs": resolved_crs,
            "is_wkb": False,
        }

    raise ValueError(f"Formato de arquivo de glebas não suportado: {path}")


def run(args: argparse.Namespace) -> None:
    if args.tamanho_bloco < 1:
        raise ValueError("--tamanho-bloco deve ser maior que zero.")
    min_area_ha = float(getattr(args, "area_minima_ha", 0.0) or 0.0)
    min_percent = float(getattr(args, "percentual_minimo", 0.0) or 0.0)
    if min_area_ha < 0 or not math.isfinite(min_area_ha):
        raise ValueError("--area-minima-ha deve ser um número maior ou igual a zero.")
    if not 0.0 <= min_percent <= 100.0 or not math.isfinite(min_percent):
        raise ValueError("--percentual-minimo deve estar entre 0 e 100.")

    glebas_paths = (
        list(args.glebas)
        if isinstance(args.glebas, (list, tuple))
        else [args.glebas]
    )
    raw_operations = getattr(args, "operacoes", None)
    if not raw_operations:
        operation_paths: list[Path] = []
    elif isinstance(raw_operations, (list, tuple)):
        operation_paths = list(raw_operations)
    else:
        operation_paths = [raw_operations]

    if not glebas_paths:
        raise ValueError("Selecione ao menos um arquivo de glebas.")
    for input_path in [*glebas_paths, *operation_paths]:
        if not input_path.is_file():
            raise FileNotFoundError(f"Arquivo não encontrado: {input_path}")

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    csv_encodings = {
        input_path: resolve_csv_encoding(input_path, args.codificacao)
        for input_path in dict.fromkeys([*glebas_paths, *operation_paths])
        if is_csv_path(input_path)
    }
    if args.codificacao.lower() == "auto":
        for input_path, detected_encoding in csv_encodings.items():
            if detected_encoding not in {"utf-8", "utf-8-sig"}:
                LOGGER.info(
                    "Codificação detectada para %s: %s.",
                    input_path.name,
                    detected_encoding,
                )

    requested_area_crs = str(args.crs_area).strip()
    if requested_area_crs.casefold() == "auto":
        area_crs = automatic_area_crs(
            args.terras_indigenas,
            args.camada_ti,
        )
    else:
        area_crs = CRS.from_user_input(requested_area_crs)
    if not area_crs.is_projected:
        raise ValueError(f"O CRS de cálculo de área deve ser projetado, não {area_crs}.")
    if not area_crs.axis_info or area_crs.axis_info[0].unit_conversion_factor is None:
        raise ValueError(f"Não foi possível determinar a unidade linear do CRS {area_crs}.")

    operation_meta: list[tuple[Path, str, str, list[str]]] = []
    for operations_path in operation_paths:
        if is_parquet_path(operations_path):
            op_header = list(pl.read_parquet_schema(operations_path).keys())
            validate_unique_columns(op_header, operations_path)
            operation_meta.append((operations_path, "", "", op_header))
        elif is_csv_path(operations_path):
            op_encoding = csv_encodings[operations_path]
            op_delimiter = args.delimitador or infer_delimiter(operations_path, op_encoding)
            op_header = read_csv_header(operations_path, op_delimiter, op_encoding)
            validate_unique_columns(op_header, operations_path)
            operation_meta.append((operations_path, op_delimiter, op_encoding, op_header))
        else:
            raise ValueError(f"Formato de tabela auxiliar não suportado: {operations_path}")
    operation_headers = [item[3] for item in operation_meta]

    # Resolve a chave canônica a partir do primeiro arquivo de glebas e das tabelas auxiliares.
    first_gleba_path = glebas_paths[0]
    first_schema = _inspect_gleba_schema(
        first_gleba_path,
        args.delimitador,
        csv_encodings.get(first_gleba_path),
        args.coluna_geometria,
    )
    _, canonical_key_column = resolve_join_key_column(
        first_schema["header"],  # type: ignore[arg-type]
        operation_headers,
        args.coluna_ref,
        first_schema["geometry_column"],
    )
    LOGGER.info("Coluna-chave selecionada para cruzamento: %s.", canonical_key_column)

    lands = load_indigenous_lands(
        args.terras_indigenas, str(area_crs), args.camada_ti, args.coluna_nome_ti
    )
    base_indexes: dict[str, SpatialBaseIndex] = {}

    def get_base_index(src_crs: object) -> SpatialBaseIndex:
        crs_obj = CRS.from_user_input(src_crs)
        cache_key = crs_obj.to_wkt()
        cached = base_indexes.get(cache_key)
        if cached is None:
            cached = SpatialBaseIndex(lands, crs_obj, area_crs)
            base_indexes[cache_key] = cached
        return cached

    overlap_records: list[dict[str, object]] = []
    intersection_records: list[dict[str, object]] = []
    extra_gleba_col_names: list[str] = []
    seen_extra_cf: set[str] = set()
    total_rows = valid_plots = invalid_wkt = invalid_polygons = filtered_slivers = 0
    used_source_labels: set[str] = set()
    reserved_result_cf = {
        "plot_uid",
        canonical_key_column.casefold(),
        "arquivo_glebas",
        "area_total_ha",
        "area_sobreposta_ha",
        "percentual_sobreposicao",
        "nome_terra_indigena",
        "land_uids",
        "geometry",
    }

    for gleba_idx, glebas_path in enumerate(glebas_paths):
        schema_info = (
            first_schema
            if gleba_idx == 0
            else _inspect_gleba_schema(
                glebas_path,
                args.delimitador,
                csv_encodings.get(glebas_path),
                args.coluna_geometria,
            )
        )
        kind = str(schema_info["kind"])
        header: list[str] = schema_info["header"]  # type: ignore[assignment]
        geometry_column = schema_info["geometry_column"]

        if args.coluna_ref:
            reference_column: object | None = resolve_column(
                header, args.coluna_ref, (canonical_key_column,), "chave de junção das glebas"
            )
        elif operation_paths or _find_column_casefold(header, canonical_key_column.casefold()) is not None:
            reference_column = resolve_column(
                header,
                canonical_key_column,
                (canonical_key_column,),
                f"{canonical_key_column} das glebas",
            )
        else:
            reference_column = None

        raw_extra_cols = [
            col
            for col in header
            if col != geometry_column
            and (reference_column is None or col != reference_column)
        ]
        if len(raw_extra_cols) > MAX_EXTRA_GLEBA_COLUMNS:
            LOGGER.info(
                "%s possui %d colunas extras; mantendo apenas chave e geometria.",
                glebas_path.name,
                len(raw_extra_cols),
            )
            extra_cols: list[object] = []
        else:
            extra_cols = raw_extra_cols
            for col in extra_cols:
                norm = normalize_column_name(col)
                if norm.casefold() in reserved_result_cf:
                    norm = unique_column_name(f"GLEBAS_{norm}", reserved_result_cf | seen_extra_cf)
                if norm.casefold() not in seen_extra_cf:
                    seen_extra_cf.add(norm.casefold())
                    extra_gleba_col_names.append(norm)

        source_label = glebas_path.name
        suffix = 2
        while source_label in used_source_labels:
            source_label = f"{glebas_path.stem}_{suffix}{glebas_path.suffix}"
            suffix += 1
        used_source_labels.add(source_label)

        def _record_batch(
            matches: list[dict[str, object]],
            inter_matches: list[dict[str, object]],
            poly_count: int,
            bad_wkt: int,
            bad_cand_polygons: int,
            batch_slivers: int,
            batch_rows: int,
        ) -> None:
            nonlocal valid_plots, invalid_wkt, invalid_polygons, filtered_slivers, total_rows
            valid_plots += poly_count
            invalid_wkt += bad_wkt
            invalid_polygons += bad_cand_polygons
            filtered_slivers += batch_slivers
            for match in matches:
                match["ARQUIVO_GLEBAS"] = source_label
                for raw_col in extra_cols:
                    norm = normalize_column_name(raw_col)
                    if norm.casefold() in reserved_result_cf and norm in match:
                        safe_extra = next(
                            (
                                name
                                for name in extra_gleba_col_names
                                if name.casefold() == f"glebas_{norm.casefold()}"
                            ),
                            f"GLEBAS_{norm}",
                        )
                        match[safe_extra] = match.pop(norm)
            overlap_records.extend(matches)
            for inter in inter_matches:
                inter["ARQUIVO_GLEBAS"] = source_label
            intersection_records.extend(inter_matches)
            total_rows += batch_rows
            LOGGER.info(
                "Glebas lidas: %d; poligonais: %d; com sobreposição: %d.",
                total_rows,
                valid_plots,
                len(overlap_records),
            )

        if kind == "csv":
            delimiter = str(schema_info["delimiter"])
            glebas_encoding = str(schema_info["encoding"])
            columns_to_read: list[str] = []
            if reference_column is not None:
                columns_to_read.append(str(reference_column))
            columns_to_read.append(str(geometry_column))
            columns_to_read.extend(str(col) for col in extra_cols)
            LOGGER.info(
                "Processando %s em blocos de %d linhas (delimitador %r).",
                glebas_path.name,
                args.tamanho_bloco,
                delimiter,
            )
            batch_index = get_base_index(args.crs_glebas)
            for chunk in iter_csv_batches(
                glebas_path,
                delimiter,
                glebas_encoding,
                columns_to_read,
                args.tamanho_bloco,
            ):
                matches, inter_matches, poly_count, bad_wkt, bad_cand, batch_slivers = process_wkt_chunk(
                    chunk=chunk,
                    ref_column=reference_column,
                    key_column=canonical_key_column,
                    geometry_column=geometry_column,
                    extra_columns=extra_cols,
                    base_index=batch_index,
                    source_crs=args.crs_glebas,
                    area_crs=area_crs,
                    row_offset=total_rows,
                    min_area_ha=min_area_ha,
                    min_percent=min_percent,
                )
                _record_batch(
                    matches, inter_matches, poly_count, bad_wkt, bad_cand, batch_slivers, len(chunk)
                )
        elif kind == "parquet":
            columns_to_read = []
            if reference_column is not None:
                columns_to_read.append(str(reference_column))
            columns_to_read.append(str(geometry_column))
            columns_to_read.extend(str(col) for col in extra_cols)
            source_crs = schema_info["crs"] or args.crs_glebas
            batch_index = get_base_index(source_crs)
            is_wkb = bool(schema_info["is_wkb"])
            LOGGER.info(
                "Processando Parquet/GeoParquet %s (%s) em blocos de %d linhas.",
                glebas_path.name,
                "WKB" if is_wkb else "WKT",
                args.tamanho_bloco,
            )
            processor = process_wkb_chunk if is_wkb else process_wkt_chunk
            parquet_query = pl.scan_parquet(glebas_path).select(columns_to_read)
            for chunk in parquet_query.collect_batches(
                chunk_size=args.tamanho_bloco, engine="streaming"
            ):
                matches, inter_matches, poly_count, bad_geom, bad_cand, batch_slivers = processor(
                    chunk=chunk,
                    ref_column=reference_column,
                    key_column=canonical_key_column,
                    geometry_column=geometry_column,
                    extra_columns=extra_cols,
                    base_index=batch_index,
                    source_crs=source_crs,
                    area_crs=area_crs,
                    row_offset=total_rows,
                    min_area_ha=min_area_ha,
                    min_percent=min_percent,
                )
                _record_batch(
                    matches, inter_matches, poly_count, bad_geom, bad_cand, batch_slivers, len(chunk)
                )
        else:
            layer = schema_info["layer"]
            source_crs = schema_info["crs"] or args.crs_glebas
            batch_index = get_base_index(source_crs)
            cols_to_read_vec = (
                [str(reference_column)] if reference_column is not None else []
            ) + [str(col) for col in extra_cols]
            LOGGER.info(
                "Processando camada vetorial %s em fatias de %d registros.",
                glebas_path.name,
                args.tamanho_bloco,
            )
            offset = 0
            while True:
                read_kwargs: dict[str, object] = {
                    "rows": slice(offset, offset + args.tamanho_bloco),
                    "engine": "pyogrio",
                    "columns": cols_to_read_vec,
                }
                if layer is not None:
                    read_kwargs["layer"] = str(layer)
                frame = gpd.read_file(glebas_path, **read_kwargs)
                if frame.empty:
                    break
                frame_crs = frame.crs or source_crs
                frame_index = (
                    batch_index
                    if frame.crs is None or frame.crs == source_crs
                    else get_base_index(frame_crs)
                )
                matches, inter_matches, poly_count, bad_geom, bad_cand, batch_slivers = (
                    process_geodataframe_plots(
                        frame=frame,
                        ref_column=reference_column,
                        key_column=canonical_key_column,
                        extra_columns=extra_cols,
                        base_index=frame_index,
                        source_crs=frame_crs,
                        area_crs=area_crs,
                        row_offset=total_rows,
                        min_area_ha=min_area_ha,
                        min_percent=min_percent,
                    )
                )
                _record_batch(
                    matches, inter_matches, poly_count, bad_geom, bad_cand, batch_slivers, len(frame)
                )
                offset += len(frame)
                if len(frame) < args.tamanho_bloco:
                    break

    if invalid_wkt:
        LOGGER.warning("%d registros com geometria inválida/ausente foram ignorados.", invalid_wkt)
    if invalid_polygons:
        LOGGER.warning(
            "%d candidatas com geometria inválida não corrigível foram ignoradas.",
            invalid_polygons,
        )
    if filtered_slivers:
        LOGGER.info(
            "%d glebas descartadas pela tolerância mínima (área %.4f ha / %.4f%%).",
            filtered_slivers,
            min_area_ha,
            min_percent,
        )

    result_columns = [
        "plot_uid",
        canonical_key_column,
        "ARQUIVO_GLEBAS",
        "AREA_TOTAL_HA",
        "AREA_SOBREPOSTA_HA",
        "PERCENTUAL_SOBREPOSICAO",
        "NOME_TERRA_INDIGENA",
        *extra_gleba_col_names,
        "land_uids",
        "geometry",
    ]
    plot_results = pd.DataFrame(overlap_records, columns=result_columns)
    plot_land_relations: dict[int, list[int]] = (
        {
            int(uid): list(land_list)
            for uid, land_list in zip(plot_results["plot_uid"], plot_results["land_uids"])
        }
        if not plot_results.empty
        else {}
    )
    plot_results = plot_results.drop(columns=["land_uids"])
    plot_results = gpd.GeoDataFrame(
        plot_results, geometry="geometry", crs=area_crs
    )

    inter_columns = [
        "plot_uid",
        "land_uid",
        canonical_key_column,
        "ARQUIVO_GLEBAS",
        "CAMADA_BASE",
        "AREA_TOTAL_HA",
        "AREA_INTERSECAO_HA",
        "PERCENTUAL_INTERSECAO",
        "geometry",
    ]
    intersections_gdf = gpd.GeoDataFrame(
        pd.DataFrame(intersection_records, columns=inter_columns),
        geometry="geometry",
        crs=area_crs,
    )

    if not plot_results.empty:
        centroids_wgs84 = gpd.GeoSeries(
            plot_results.geometry.centroid, crs=area_crs
        ).to_crs("EPSG:4326")
        plot_centroids = {
            int(uid): (float(lat), float(lon))
            for uid, lon, lat in zip(
                plot_results["plot_uid"], centroids_wgs84.x, centroids_wgs84.y
            )
        }
    else:
        plot_centroids = {}

    references = set(plot_results[canonical_key_column].dropna().astype(str))
    base_results = plot_results.drop(columns=["geometry"])
    financed_sum = 0.0
    invalid_finance = 0
    references_with_operations: set[str] = set()
    operation_frames: list[tuple[pd.DataFrame, str]] = []
    canonical_columns: dict[str, str] = {}
    used_source_labels = set()
    resolved_finance_col: str | None = None

    for operations_path, operations_delimiter, operations_encoding, _ in operation_meta:
        operations, table_financed_sum, table_invalid_finance = read_matching_operations(
            operations_path,
            operations_delimiter,
            operations_encoding,
            args.coluna_ref,
            args.coluna_valor,
            references,
            args.tamanho_bloco,
            key_column=canonical_key_column,
        )
        financed_sum += table_financed_sum
        invalid_finance += table_invalid_finance
        if canonical_key_column in operations.columns:
            references_with_operations.update(
                operations[canonical_key_column].dropna().astype(str)
            )
        if resolved_finance_col is None:
            found_fin = resolve_finance_column(operations.columns, args.coluna_valor)
            if found_fin is not None:
                resolved_finance_col = str(found_fin)

        column_mapping: dict[str, str] = {}
        for column in operations.columns:
            if column == canonical_key_column:
                continue
            canonical = canonical_columns.setdefault(column.casefold(), column)
            if column != canonical:
                column_mapping[column] = canonical
        if column_mapping:
            operations = operations.rename(columns=column_mapping)

        source_label = unique_column_name(
            operations_path.stem or "tabela", used_source_labels
        )
        used_source_labels.add(source_label)
        operation_frames.append((operations, source_label))
        LOGGER.info(
            "Tabela %s relacionada: %d registros correspondentes.",
            operations_path.name,
            len(operations),
        )

    if invalid_finance:
        LOGGER.warning(
            "%d valores financiados inválidos nas operações relacionadas foram ignorados na soma.",
            invalid_finance,
        )

    operation_columns = {
        column for operations, _ in operation_frames for column in operations.columns
    }
    source_column = unique_column_name(
        "ARQUIVO_OPERACOES", set(base_results.columns) | operation_columns
    )
    if operation_frames:
        combined_operations = consolidate_operation_tables(
            operation_frames, source_column, key_column=canonical_key_column
        )
        # Se alguma coluna extra das glebas (ex.: NU_ORDEM) colidir com as operações, preserva a das operações.
        op_cols_cf = {
            normalize_column_name(col).casefold()
            for col in combined_operations.columns
            if col != canonical_key_column
        }
        extra_cf_set = {col.casefold() for col in extra_gleba_col_names}
        base_renames: dict[str, str] = {}
        existing_all = set(base_results.columns) | set(combined_operations.columns)
        for col in list(base_results.columns):
            if col.casefold() in extra_cf_set and col.casefold() in op_cols_cf:
                replacement = unique_column_name(f"GLEBAS_{col}", existing_all)
                base_renames[col] = replacement
                existing_all.add(replacement)
        if base_renames:
            base_results = base_results.rename(columns=base_renames)

        result_columns_set = set(base_results.columns) - {canonical_key_column}
        result_columns_by_casefold = {
            column.casefold(): column for column in result_columns_set
        }
        operation_collisions = {
            column
            for column in combined_operations.columns
            if column != canonical_key_column
            and column.casefold() in result_columns_by_casefold
        }
        existing_columns = result_columns_set | set(combined_operations.columns)
        renamed_collisions: dict[str, str] = {}
        for column in sorted(operation_collisions):
            result_column = result_columns_by_casefold[column.casefold()]
            replacement = unique_column_name(
                f"OPERACOES_{result_column}", existing_columns
            )
            renamed_collisions[column] = replacement
            existing_columns.add(replacement)
        if renamed_collisions:
            combined_operations = combined_operations.rename(columns=renamed_collisions)
        final = base_results.merge(
            combined_operations, on=canonical_key_column, how="left"
        )
    else:
        final = base_results
        found_fin = resolve_finance_column(final.columns, args.coluna_valor)
        if found_fin is not None:
            resolved_finance_col = str(found_fin)
            final[resolved_finance_col] = parse_financial_values(final[resolved_finance_col])
            financed_sum = float(final[resolved_finance_col].sum(skipna=True))

    if args.coluna_valor and resolved_finance_col is None:
        raise ValueError(
            f"Coluna de valor '{args.coluna_valor}' não encontrada nas tabelas enviadas."
        )

    summary_rows: list[tuple[str, object]] = [
        ("Glebas com geometria poligonal lida", valid_plots),
        ("Quantidade de glebas com sobreposição", len(plot_results)),
        ("Interseções detalhadas (gleba × Camada Base)", len(intersections_gdf)),
        ("Soma da área sobreposta (ha)", float(plot_results["AREA_SOBREPOSTA_HA"].sum())),
    ]
    if resolved_finance_col is not None:
        summary_rows.append(("Soma dos valores financiados relacionados", financed_sum))
    summary_rows.extend(
        [
            ("Registros de glebas lidos", total_rows),
            ("Geometrias inválidas/ausentes", invalid_wkt),
            ("Candidatas com geometria inválida não corrigível", invalid_polygons),
        ]
    )
    if min_area_ha > 0 or min_percent > 0 or filtered_slivers > 0:
        summary_rows.append(
            ("Sobreposições descartadas pela tolerância mínima", filtered_slivers)
        )
    if operation_paths:
        summary_rows.append(
            (
                "Glebas sem operação correspondente",
                int(
                    plot_results.loc[
                        ~plot_results[canonical_key_column].isin(references_with_operations)
                    ].shape[0]
                ),
            )
        )
    summary = pd.DataFrame(summary_rows, columns=["Indicador", "Valor"])

    preview_path = getattr(args, "preview_path", None)
    if preview_path is not None:
        write_preview_geojson(Path(preview_path), lands, plot_results)
    dashboard_path = getattr(args, "dashboard_path", None)
    if dashboard_path is not None:
        write_dashboard_database(
            Path(dashboard_path),
            final,
            summary,
            lands,
            plot_results,
            plot_land_relations=plot_land_relations,
            plot_centroids=plot_centroids,
            key_column=canonical_key_column,
            finance_column=resolved_finance_col,
            area_crs=area_crs,
            has_operations=bool(operation_paths),
            intersections=intersections_gdf,
            operations_source_column=source_column if operation_frames else None,
            min_area_ha=min_area_ha,
            min_percent=min_percent,
        )
    if args.saida_gpkg is not None:
        geometry_by_uid = plot_results.set_index("plot_uid")["geometry"]
        mapped_geometry = gpd.GeoSeries(
            final["plot_uid"].map(geometry_by_uid),
            index=final.index,
            crs=area_crs,
        )
        geopackage_results = gpd.GeoDataFrame(
            final.copy(),
            geometry=mapped_geometry,
            crs=area_crs,
        )
        write_geopackage(
            args.saida_gpkg,
            geopackage_results,
            area_crs,
            intersections=intersections_gdf,
        )
    if args.saida is not None:
        write_excel(args.saida, final, summary, intersections=intersections_gdf)
    LOGGER.info(
        "Concluído: %d glebas com sobreposição, %d interseções 1:N, %d linhas no resultado%s.",
        len(plot_results),
        len(intersections_gdf),
        len(final),
        f" e arquivo {args.saida}" if args.saida is not None else "",
    )


def main() -> int:
    args = parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    try:
        run(args)
    except Exception as error:
        LOGGER.exception("Processamento interrompido: %s", error)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
