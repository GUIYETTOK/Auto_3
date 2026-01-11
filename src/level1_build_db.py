from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path
from typing import Iterable, List

import unicodedata

from excel_utils import EstimateRow, parse_estimate_file


def iter_estimate_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if path.is_dir():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".xls", ".xlsx"}:
            continue
        name = unicodedata.normalize("NFC", path.name)
        if "견적서" not in name:
            continue
        if "견적의뢰" in name or "견적요청" in name:
            continue
        yield path


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS estimate_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT NOT NULL,
            source_sheet TEXT NOT NULL,
            file_datetime TEXT NOT NULL,
            a7_text TEXT,
            품명 TEXT NOT NULL,
            규격 TEXT,
            단위 TEXT,
            수량 REAL,
            단가 REAL,
            금액 REAL
        )
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_estimate_lookup
        ON estimate_items (품명, 규격, file_datetime)
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_estimate_spec
        ON estimate_items (규격, file_datetime)
        """
    )


def insert_rows(conn: sqlite3.Connection, rows: List[EstimateRow]) -> int:
    if not rows:
        return 0
    conn.executemany(
        """
        INSERT INTO estimate_items (
            source_file, source_sheet, file_datetime, a7_text,
            품명, 규격, 단위, 수량, 단가, 금액
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                row.source_file,
                row.source_sheet,
                row.file_datetime.isoformat(),
                row.a7_text,
                row.품명,
                row.규격,
                row.단위,
                row.수량,
                row.단가,
                row.금액,
            )
            for row in rows
        ],
    )
    return len(rows)


def build_db(db_folder: Path, output_db: Path) -> tuple[int, int]:
    estimate_files = list(iter_estimate_files(db_folder))
    if not estimate_files:
        raise SystemExit("견적서 파일을 찾지 못했습니다.")

    output_db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(output_db)
    try:
        init_db(conn)
        total_rows = 0
        for path in estimate_files:
            rows = parse_estimate_file(path)
            total_rows += insert_rows(conn, rows)
        conn.commit()
    finally:
        conn.close()
    print(f"완료: {len(estimate_files)}개 파일, {total_rows}건 적재 -> {output_db}")
    return len(estimate_files), total_rows


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Level 1: 견적서 DB 구축")
    parser.add_argument("db_folder", help="견적서 엑셀 파일이 있는 DB 폴더")
    parser.add_argument(
        "--output-db",
        default="estimate.sqlite3",
        help="출력 SQLite DB 경로 (기본: estimate.sqlite3)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    db_folder = Path(args.db_folder).expanduser().resolve()
    output_db = Path(args.output_db).expanduser().resolve()
    build_db(db_folder, output_db)


if __name__ == "__main__":
    main()
