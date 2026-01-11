from __future__ import annotations

import argparse
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional
import re

from excel_utils import RequestRow
from level2_parse_request import parse_request


@dataclass
class MatchedRow:
    request: RequestRow
    matched: bool
    match_method: Optional[str]
    matched_source_file: Optional[str]
    matched_source_sheet: Optional[str]
    matched_file_datetime: Optional[str]
    matched_a7_text: Optional[str]
    단가: Optional[float]
    금액: Optional[float]
    status: str
    candidates: Optional[List[dict]]


def lookup_latest_price(
    conn: sqlite3.Connection, 품명: str, 규격: str
) -> Optional[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT source_file, source_sheet, file_datetime, a7_text, 단가
        FROM estimate_items
        WHERE 품명 = ? AND 규격 = ?
        ORDER BY file_datetime DESC
        LIMIT 1
        """,
        (품명, 규격),
    )
    return cursor.fetchone()


def lookup_candidates(
    conn: sqlite3.Connection, 품명: str, 규격: str
) -> List[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT source_file, source_sheet, file_datetime, a7_text, 단가
        FROM estimate_items
        WHERE 품명 = ? AND 규격 = ? AND 단가 IS NOT NULL
        ORDER BY file_datetime DESC
        """,
        (품명, 규격),
    )
    return cursor.fetchall()


def lookup_latest_price_by_spec(
    conn: sqlite3.Connection, 규격: str
) -> Optional[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT source_file, source_sheet, file_datetime, a7_text, 단가
        FROM estimate_items
        WHERE 규격 = ?
        ORDER BY file_datetime DESC
        LIMIT 1
        """,
        (규격,),
    )
    return cursor.fetchone()


def lookup_latest_price_by_name(
    conn: sqlite3.Connection, 품명: str
) -> Optional[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT source_file, source_sheet, file_datetime, a7_text, 단가
        FROM estimate_items
        WHERE 품명 = ?
        ORDER BY file_datetime DESC
        LIMIT 1
        """,
        (품명,),
    )
    return cursor.fetchone()


def normalize_code(value: str) -> str:
    if not value:
        return ""
    cleaned = re.sub(r"[^0-9a-zA-Z]", "", value.upper())
    return cleaned


def lookup_latest_price_fuzzy_spec(
    conn: sqlite3.Connection, 규격: str
) -> Optional[sqlite3.Row]:
    target = normalize_code(규격)
    if not target:
        return None
    cursor = conn.execute(
        """
        SELECT source_file, source_sheet, file_datetime, a7_text, 단가, 규격
        FROM estimate_items
        WHERE 규격 IS NOT NULL AND 규격 != ''
        ORDER BY file_datetime DESC
        """
    )
    for row in cursor.fetchall():
        candidate = normalize_code(row["규격"])
        if not candidate:
            continue
        if target.endswith(candidate) or candidate.endswith(target):
            return row
    return None


def match_prices(db_path: Path, requests: List[RequestRow]) -> List[MatchedRow]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        results: List[MatchedRow] = []
        for row in requests:
            candidates_rows = (
                lookup_candidates(conn, row.품명, row.규격)
                if row.품명 and row.규격
                else []
            )
            candidates = [
                {
                    "단가": float(c["단가"]) if c["단가"] is not None else None,
                    "견적서파일": c["source_file"],
                    "시트": c["source_sheet"],
                    "견적날짜": c["file_datetime"],
                    "A7": c["a7_text"],
                }
                for c in candidates_rows
            ]
            matched_row = lookup_latest_price(conn, row.품명, row.규격)
            match_method = "품명+규격" if matched_row else None
            if not matched_row and row.규격:
                matched_row = lookup_latest_price_by_spec(conn, row.규격)
                match_method = "규격" if matched_row else None
            if not matched_row and row.규격:
                matched_row = lookup_latest_price_fuzzy_spec(conn, row.규격)
                match_method = "규격-유사" if matched_row else None
            if not matched_row and row.품명:
                matched_row = lookup_latest_price_by_name(conn, row.품명)
                match_method = "품명" if matched_row else None
            if matched_row and matched_row["단가"] is not None:
                단가 = float(matched_row["단가"])
                금액 = 단가 * row.구매량 if row.구매량 is not None else None
                results.append(
                    MatchedRow(
                        request=row,
                        matched=True,
                        match_method=match_method,
                        matched_source_file=matched_row["source_file"],
                        matched_source_sheet=matched_row["source_sheet"],
                        matched_file_datetime=matched_row["file_datetime"],
                        matched_a7_text=matched_row["a7_text"],
                        단가=단가,
                        금액=금액,
                        status="매칭",
                        candidates=candidates,
                    )
                )
            else:
                results.append(
                    MatchedRow(
                        request=row,
                        matched=False,
                        match_method=match_method,
                        matched_source_file=matched_row["source_file"]
                        if matched_row
                        else None,
                        matched_source_sheet=matched_row["source_sheet"]
                        if matched_row
                        else None,
                        matched_file_datetime=matched_row["file_datetime"]
                        if matched_row
                        else None,
                        matched_a7_text=matched_row["a7_text"] if matched_row else None,
                        단가=None,
                        금액=None,
                        status="단가 없음",
                        candidates=candidates,
                    )
                )
        return results
    finally:
        conn.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Level 3: DB 기반 단가 매칭")
    parser.add_argument("db_path", help="Level 1에서 생성한 SQLite DB 경로")
    parser.add_argument("input_path", help="견적의뢰서 파일 또는 DB 폴더 경로")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    db_path = Path(args.db_path).expanduser().resolve()
    input_path = Path(args.input_path).expanduser().resolve()
    requests = parse_request(input_path)
    matched = match_prices(db_path, requests)

    print(f"완료: {len(matched)}건 매칭")
    for row in matched[:5]:
        print(
            f"{row.request.품명} | {row.request.규격} | {row.단가} | {row.금액} | {row.status}"
        )


if __name__ == "__main__":
    main()
