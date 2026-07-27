"""FastAPI application for AI Document Formatter."""
from __future__ import annotations

import os
import re
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from backend.formatting import DocumentKind, OutputFormat, build_document, clean_text, export_document, extract_upload

ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / os.getenv("OUTPUT_DIR", "outputs")
OUTPUTS.mkdir(exist_ok=True)
MAX_INPUT_CHARS = int(os.getenv("MAX_INPUT_CHARS", "50000"))

app = FastAPI(title="AI Document Formatter", version="1.0.0", docs_url="/api/docs", redoc_url=None)
app.mount("/static", StaticFiles(directory=ROOT / "frontend"), name="static")


class FormatRequest(BaseModel):
    text: str = Field(min_length=1, max_length=50000)
    output_format: OutputFormat = "pdf"
    document_kind: DocumentKind = "auto"
    use_ai: bool = True
    accent: str = "#6d5dfc"


def slug(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")[:48] or "formatted-document"


def save_export(text: str, output_format: OutputFormat, document_kind: DocumentKind, use_ai: bool, accent: str) -> dict:
    try:
        text = clean_text(text, MAX_INPUT_CHARS)
        document = build_document(text, document_kind, use_ai)
        content, content_type = export_document(document, output_format, accent)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    filename = f"{slug(document.title)}-{uuid.uuid4().hex[:8]}.{output_format}"
    (OUTPUTS / filename).write_bytes(content)
    return {"title": document.title, "kind": document.kind, "summary": document.summary, "enhanced": document.enhanced,
            "filename": filename, "download_url": f"/api/download/{filename}", "content_type": content_type}


@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def home() -> HTMLResponse:
    return HTMLResponse((ROOT / "frontend" / "index.html").read_text(encoding="utf-8"))


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "ai_configured": bool(os.getenv("OPENAI_API_KEY")), "max_input_chars": MAX_INPUT_CHARS}


@app.post("/api/documents", status_code=201)
def create_document(request: FormatRequest) -> dict:
    return save_export(**request.model_dump())


@app.post("/api/upload", status_code=201)
async def upload_document(file: UploadFile = File(...), output_format: OutputFormat = Form("pdf"), document_kind: DocumentKind = Form("auto"), use_ai: bool = Form(True), accent: str = Form("#6d5dfc")) -> dict:
    return save_export(extract_upload(file.filename or "upload.txt", await file.read()), output_format, document_kind, use_ai, accent)


@app.post("/api/batch", status_code=201)
async def batch_documents(files: list[UploadFile] = File(...), output_format: OutputFormat = Form("pdf"), document_kind: DocumentKind = Form("auto"), use_ai: bool = Form(True), accent: str = Form("#6d5dfc")) -> dict:
    if len(files) > 10:
        raise HTTPException(status_code=422, detail="Batch processing is limited to 10 files.")
    results = []
    for file in files:
        try:
            results.append(save_export(extract_upload(file.filename or "upload.txt", await file.read()), output_format, document_kind, use_ai, accent))
        except (HTTPException, ValueError) as exc:
            results.append({"filename": file.filename, "error": getattr(exc, "detail", str(exc))})
    return {"results": results}


@app.get("/api/download/{filename}")
def download(filename: str) -> FileResponse:
    path = (OUTPUTS / filename).resolve()
    if path.parent != OUTPUTS.resolve() or not path.is_file():
        raise HTTPException(status_code=404, detail="Document not found.")
    return FileResponse(path, filename=path.name)
