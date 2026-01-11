from __future__ import annotations

import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist_bundle"
EXE = ROOT / "dist" / "견적서자동화.exe"


def main() -> None:
    DIST.mkdir(exist_ok=True)
    if EXE.exists():
        shutil.copy2(EXE, DIST / EXE.name)
    (DIST / "DB").mkdir(exist_ok=True)
    (DIST / "output").mkdir(exist_ok=True)
    print(f"완료: {DIST}")


if __name__ == "__main__":
    main()
