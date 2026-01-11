from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser

import uvicorn


def open_browser(url: str) -> None:
    for _ in range(30):
        try:
            webbrowser.open(url)
            return
        except Exception:
            time.sleep(0.2)


def main() -> None:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    url = "http://127.0.0.1:8000/ui"
    threading.Thread(target=open_browser, args=(url,), daemon=True).start()
    uvicorn.run("src.web_app:app", host="127.0.0.1", port=8000)


if __name__ == "__main__":
    main()
