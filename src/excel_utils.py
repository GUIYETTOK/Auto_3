from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import unicodedata


REQUIRED_HEADERS = ["품명", "규격", "단위", "수량", "단가", "금액"]
REQUEST_HEADERS = ["품명", "규격", "제조사", "단위", "구매량"]


@dataclass
class EstimateRow:
    source_file: str
    source_sheet: str
    file_datetime: datetime
    a7_text: str
    품명: str
    규격: str
    단위: str
    수량: Optional[float]
    단가: Optional[float]
    금액: Optional[float]


@dataclass
class RequestRow:
    source_file: str
    source_sheet: str
    품명: str
    규격: str
    제조사: str
    단위: str
    구매량: Optional[float]


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value)
    text = unicodedata.normalize("NFC", text)
    return "".join(text.split())


def is_data_sheet(df: pd.DataFrame) -> bool:
    return df.notna().any().any()


def find_header_row(df: pd.DataFrame) -> Optional[Tuple[int, Dict[str, int]]]:
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        normalized = [normalize_text(v) for v in row.tolist()]
        header_map: Dict[str, int] = {}
        for col_idx, cell in enumerate(normalized):
            if cell in REQUIRED_HEADERS and cell not in header_map:
                header_map[cell] = col_idx
        if all(h in header_map for h in REQUIRED_HEADERS):
            return row_idx, header_map
    return None


def find_request_header_row(df: pd.DataFrame) -> Optional[Tuple[int, Dict[str, int]]]:
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        normalized = [normalize_text(v) for v in row.tolist()]
        header_map: Dict[str, int] = {}
        for col_idx, cell in enumerate(normalized):
            if cell in REQUEST_HEADERS and cell not in header_map:
                header_map[cell] = col_idx
            if cell == "수량" and "구매량" not in header_map:
                header_map["구매량"] = col_idx
        if all(h in header_map for h in REQUEST_HEADERS if h != "구매량") and (
            "구매량" in header_map
        ):
            return row_idx, header_map
    return None


def to_float(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def get_file_datetime(path: Path) -> datetime:
    stat = path.stat()
    if hasattr(stat, "st_birthtime"):
        return datetime.fromtimestamp(stat.st_birthtime)
    return datetime.fromtimestamp(stat.st_mtime)


def extract_a7_text(df: pd.DataFrame) -> str:
    try:
        value = df.iat[6, 0]
    except IndexError:
        return ""
    return str(value).strip() if value is not None else ""


def iter_estimate_rows(
    df: pd.DataFrame,
    header_row: int,
    header_map: Dict[str, int],
) -> Iterable[Dict[str, object]]:
    data_rows = 0
    empty_streak = 0
    for row_idx in range(header_row + 1, len(df)):
        row = df.iloc[row_idx]
        values = {}
        for key in REQUIRED_HEADERS:
            col_idx = header_map[key]
            values[key] = row.iloc[col_idx] if col_idx < len(row) else None
        has_any = any(normalize_text(values[key]) for key in REQUIRED_HEADERS)
        if has_any:
            empty_streak = 0
            data_rows += 1
            yield values
        else:
            if data_rows > 0:
                empty_streak += 1
                if empty_streak >= 2:
                    break


def parse_estimate_sheet(
    path: Path,
    sheet_name: str,
    df: pd.DataFrame,
) -> List[EstimateRow]:
    header_info = find_header_row(df)
    if not header_info:
        return []
    header_row, header_map = header_info
    a7_text = extract_a7_text(df)
    file_dt = get_file_datetime(path)
    results: List[EstimateRow] = []
    for values in iter_estimate_rows(df, header_row, header_map):
        품명 = normalize_text(values["품명"])
        규격 = normalize_text(values["규격"])
        단위 = normalize_text(values["단위"])
        if not 품명:
            continue
        results.append(
            EstimateRow(
                source_file=str(path),
                source_sheet=sheet_name,
                file_datetime=file_dt,
                a7_text=a7_text,
                품명=품명,
                규격=규격,
                단위=단위,
                수량=to_float(values["수량"]),
                단가=to_float(values["단가"]),
                금액=to_float(values["금액"]),
            )
        )
    return results


def parse_estimate_file(path: Path) -> List[EstimateRow]:
    xl = pd.ExcelFile(path)
    results: List[EstimateRow] = []
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        if not is_data_sheet(df):
            continue
        results.extend(parse_estimate_sheet(path, sheet_name, df))
    return results


def parse_request_sheet(
    path: Path,
    sheet_name: str,
    df: pd.DataFrame,
) -> List[RequestRow]:
    header_info = find_request_header_row(df)
    if not header_info:
        return []
    header_row, header_map = header_info
    results: List[RequestRow] = []
    data_rows = 0
    empty_streak = 0
    for row_idx in range(header_row + 1, len(df)):
        row = df.iloc[row_idx]
        values = {}
        for key in REQUEST_HEADERS:
            col_idx = header_map.get(key)
            values[key] = row.iloc[col_idx] if col_idx is not None else None
        has_any = any(normalize_text(values[key]) for key in REQUEST_HEADERS)
        if has_any:
            empty_streak = 0
            data_rows += 1
            품명 = normalize_text(values["품명"])
            규격 = normalize_text(values["규격"])
            제조사 = normalize_text(values["제조사"])
            단위 = normalize_text(values["단위"])
            if not 품명:
                continue
            results.append(
                RequestRow(
                    source_file=str(path),
                    source_sheet=sheet_name,
                    품명=품명,
                    규격=규격,
                    제조사=제조사,
                    단위=단위,
                    구매량=to_float(values["구매량"]),
                )
            )
        else:
            if data_rows > 0:
                empty_streak += 1
                if empty_streak >= 2:
                    break
    return results


def parse_request_file(path: Path) -> List[RequestRow]:
    xl = pd.ExcelFile(path)
    results: List[RequestRow] = []
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        if not is_data_sheet(df):
            continue
        results.extend(parse_request_sheet(path, sheet_name, df))
    return results
