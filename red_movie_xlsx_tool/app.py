from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any
import webbrowser

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

try:
    from .generator import DEFAULT_QUESTIONS, MULTI_SEPARATOR, SURVEY_TITLE, generate_workbook
except ImportError:
    from generator import DEFAULT_QUESTIONS, MULTI_SEPARATOR, SURVEY_TITLE, generate_workbook

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATE_DIR = BASE_DIR / "templates"
GENERATED_DIR = BASE_DIR / "generated_samples"
GENERATED_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="red-movie-xlsx-tool")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class GenerateRequest(BaseModel):
    count: int = Field(default=50, ge=1, le=5000)
    seed: str | None = None
    questions: list[dict[str, Any]]


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (TEMPLATE_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/api/schema")
def schema() -> dict[str, Any]:
    return {
        "title": SURVEY_TITLE,
        "questions": DEFAULT_QUESTIONS,
        "multi_separator": MULTI_SEPARATOR,
        "multi_separator_codepoint": f"U+{ord(MULTI_SEPARATOR):04X}",
    }


@app.get("/api/files")
def files() -> dict[str, Any]:
    items = []
    for path in sorted(GENERATED_DIR.glob("*.xlsx"), key=lambda item: item.stat().st_mtime, reverse=True):
        stat = path.stat()
        items.append(
            {
                "filename": path.name,
                "path": str(path),
                "size": stat.st_size,
                "download_url": f"/api/download/{path.name}",
            }
        )
    return {"files": items[:20]}


@app.post("/api/generate")
def generate(request: GenerateRequest) -> dict[str, Any]:
    try:
        result = generate_workbook(
            request.questions,
            count=request.count,
            seed=request.seed,
            output_dir=GENERATED_DIR,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "ok": True,
        "filename": result.path.name,
        "path": str(result.path),
        "download_url": f"/api/download/{result.path.name}",
        "row_count": result.row_count,
        "question_count": result.question_count,
        "multi_separator": MULTI_SEPARATOR,
        "multi_separator_codepoint": f"U+{ord(MULTI_SEPARATOR):04X}",
        "summary": result.summary,
    }


@app.get("/api/download/{filename}")
def download(filename: str) -> FileResponse:
    target = safe_generated_file(filename)
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=target.name,
    )


def safe_generated_file(filename: str) -> Path:
    if not filename.endswith(".xlsx") or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=404, detail="文件不存在")
    target = (GENERATED_DIR / filename).resolve()
    generated_root = GENERATED_DIR.resolve()
    if target.parent != generated_root or not target.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return target


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--open-browser", action="store_true")
    args = parser.parse_args()

    if args.open_browser:
        webbrowser.open(f"http://{args.host}:{args.port}")

    import uvicorn

    uvicorn.run("app:app", host=args.host, port=args.port, reload=False)


if __name__ == "__main__":
    main()
