from __future__ import annotations

import json
import unicodedata
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

# Ensure local modules import correctly when running via uvicorn
sys.path.append(str(Path(__file__).resolve().parent))

from excel_utils import RequestRow
from level1_build_db import build_db
from level2_parse_request import parse_request
from level3_match_prices import MatchedRow, match_prices
from level4_generate_estimate import generate_estimate

app = FastAPI(title="견적서 자동화 API")
app_root = Path(__file__).resolve().parent.parent


def is_within_root(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
    except ValueError:
        return False
    return True


def request_to_dict(row: RequestRow) -> Dict[str, Any]:
    return {
        "품명": row.품명,
        "규격": row.규격,
        "제조사": row.제조사,
        "단위": row.단위,
        "구매량": row.구매량,
        "source_file": row.source_file,
        "source_sheet": row.source_sheet,
    }


def matched_to_dict(row: MatchedRow) -> Dict[str, Any]:
    date_only = None
    if row.matched_file_datetime:
        date_only = row.matched_file_datetime.split("T")[0]
    candidates = []
    for candidate in row.candidates or []:
        cand_date = None
        if candidate.get("견적날짜"):
            cand_date = str(candidate["견적날짜"]).split("T")[0]
        candidates.append(
            {
                "단가": candidate.get("단가"),
                "견적서파일": candidate.get("견적서파일"),
                "시트": candidate.get("시트"),
                "A7": candidate.get("A7"),
                "견적날짜": cand_date,
            }
        )
    return {
        "품명": row.request.품명,
        "규격": row.request.규격,
        "단위": row.request.단위,
        "수량": row.request.구매량,
        "단가": row.단가,
        "금액": row.금액,
        "상태": row.status,
        "매칭방식": row.match_method,
        "후보": candidates,
        "출처": {
            "견적서파일": row.matched_source_file,
            "시트": row.matched_source_sheet,
            "A7": row.matched_a7_text,
            "견적날짜": date_only,
        },
    }


def save_upload(upload: UploadFile) -> Path:
    suffix = Path(upload.filename or "upload.xlsx").suffix
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp.write(upload.file.read())
    temp.flush()
    return Path(temp.name)


def derive_request_label(name: Optional[str]) -> str:
    if not name:
        return ""
    base = Path(name).stem
    base = unicodedata.normalize("NFC", base)
    compact = "".join(base.split())
    return compact[:4]


@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/db/build")
def build_database(db_folder: str = Form(...), output_db: Optional[str] = Form(None)) -> Dict[str, Any]:
    db_folder_path = Path(db_folder).expanduser().resolve()
    output_db_path = (
        Path(output_db).expanduser().resolve()
        if output_db
        else db_folder_path / "estimate.sqlite3"
    )
    file_count, row_count = build_db(db_folder_path, output_db_path)
    return {
        "db_path": str(output_db_path),
        "file_count": file_count,
        "row_count": row_count,
    }


@app.post("/requests/parse")
def parse_request_api(
    input_path: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
) -> Dict[str, Any]:
    if file:
        temp_path = save_upload(file)
        rows = parse_request(temp_path)
    else:
        if not input_path:
            raise HTTPException(status_code=400, detail="input_path 또는 file이 필요합니다.")
        rows = parse_request(Path(input_path).expanduser().resolve())
    return {"count": len(rows), "items": [request_to_dict(r) for r in rows]}


@app.post("/requests/match")
def match_request_api(
    db_path: str = Form(...),
    input_path: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
) -> Dict[str, Any]:
    db_path_value = Path(db_path).expanduser().resolve()
    if file:
        temp_path = save_upload(file)
        requests = parse_request(temp_path)
    else:
        if not input_path:
            raise HTTPException(status_code=400, detail="input_path 또는 file이 필요합니다.")
        requests = parse_request(Path(input_path).expanduser().resolve())
    matched = match_prices(db_path_value, requests)
    return {"count": len(matched), "items": [matched_to_dict(r) for r in matched]}


@app.post("/estimates/generate")
def generate_estimate_api(
    db_path: str = Form(...),
    output_path: str = Form(...),
    input_path: Optional[str] = Form(None),
    template_path: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
    overrides: Optional[str] = Form(None),
):
    db_path_value = Path(db_path).expanduser().resolve()
    output_path_value = Path(output_path).expanduser().resolve()
    template_value = Path(template_path).expanduser().resolve() if template_path else None

    if file:
        temp_path = save_upload(file)
        input_value = temp_path
        request_label = derive_request_label(file.filename)
    else:
        if not input_path:
            raise HTTPException(status_code=400, detail="input_path 또는 file이 필요합니다.")
        input_value = Path(input_path).expanduser().resolve()
        request_label = None

    override_map: Dict[int, Optional[float]] = {}
    if overrides:
        try:
            parsed = json.loads(overrides)
            if isinstance(parsed, dict):
                for key, value in parsed.items():
                    idx = int(key)
                    override_map[idx] = None if value in (None, "") else float(value)
        except (ValueError, json.JSONDecodeError):
            raise HTTPException(status_code=400, detail="overrides 형식이 올바르지 않습니다.")

    result = generate_estimate(
        db_path=db_path_value,
        input_path=input_value,
        output_path=output_path_value,
        template_path=template_value,
        overrides=override_map or None,
        request_label=request_label,
    )
    return FileResponse(path=str(result), filename=result.name)


@app.get("/files")
def download_file(path: str):
    target = Path(path).expanduser().resolve()
    if not is_within_root(target, app_root):
        raise HTTPException(status_code=403, detail="허용되지 않은 경로입니다.")
    if not target.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    return FileResponse(path=str(target), filename=target.name)


ui_dir = Path(__file__).resolve().parent.parent / "web"
if ui_dir.exists():
    app.mount("/ui", StaticFiles(directory=str(ui_dir), html=True), name="ui")
