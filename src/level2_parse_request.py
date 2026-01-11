from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, List

import unicodedata

from excel_utils import RequestRow, parse_request_file


def iter_request_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if path.is_dir():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".xls", ".xlsx"}:
            continue
        name = unicodedata.normalize("NFC", path.name)
        if "견적의뢰" in name or "견적요청" in name:
            yield path


def parse_request(input_path: Path) -> List[RequestRow]:
    if input_path.is_dir():
        files = list(iter_request_files(input_path))
        if not files:
            raise SystemExit("견적의뢰서 파일을 찾지 못했습니다.")
        rows: List[RequestRow] = []
        for path in files:
            rows.extend(parse_request_file(path))
        return rows
    return parse_request_file(input_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Level 2: 견적의뢰서 파싱")
    parser.add_argument("input_path", help="견적의뢰서 파일 또는 DB 폴더 경로")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input_path).expanduser().resolve()
    rows = parse_request(input_path)
    print(f"완료: {len(rows)}건 파싱")
    for row in rows[:5]:
        print(
            f"{row.품명} | {row.규격} | {row.제조사} | {row.단위} | {row.구매량}"
        )


if __name__ == "__main__":
    main()
