"""Local-only web interface for a generic proforma invoice workflow."""

from __future__ import annotations

import os
import sqlite3
import uuid
from datetime import date
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from pypdf import PdfReader

from invoice_core import next_number, render_pdf, render_xlsx, validate_invoice

ROOT = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("PI_DATA_DIR", str(ROOT / "data"))).resolve()
SPEC_DIR = DATA_DIR / "specs"
OUTPUT_DIR = DATA_DIR / "output"
DB_FILE = DATA_DIR / "numbers.sqlite3"
SELLER = os.environ.get("PI_SELLER_NAME", "Example Supplier").strip() or "Example Supplier"
DETAILS = os.environ.get("PI_SELLER_DETAILS", "").strip()
MAX_SPEC_BYTES = 10 * 1024 * 1024
app = FastAPI(title="PI / Proforma Invoice Automation")


@app.get("/", response_class=HTMLResponse)
def index():
    return (ROOT / "index.html").read_text(encoding="utf-8")


@app.get("/api/specs")
def list_specs():
    SPEC_DIR.mkdir(parents=True, exist_ok=True)
    return [{"id": p.stem, "name": p.stem + ".pdf"} for p in sorted(SPEC_DIR.glob("*.pdf"))
            if len(p.stem) == 32 and all(c in "0123456789abcdef" for c in p.stem)]


@app.post("/api/specs")
async def add_spec(file: UploadFile):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Choose a PDF file")
    contents = await file.read(MAX_SPEC_BYTES + 1)
    if len(contents) > MAX_SPEC_BYTES:
        raise HTTPException(400, "PDF must be 10 MB or smaller")
    if not contents.startswith(b"%PDF-"):
        raise HTTPException(400, "File is not a PDF")
    import io
    try:
        pdf = PdfReader(io.BytesIO(contents), strict=True)
        if pdf.is_encrypted or not 1 <= len(pdf.pages) <= 30:
            raise ValueError("Encrypted, empty, or over 30 pages")
        for page in pdf.pages:
            _ = page.mediabox
    except Exception:
        raise HTTPException(400, "PDF is invalid, encrypted, or exceeds 30 pages") from None
    SPEC_DIR.mkdir(parents=True, exist_ok=True)
    spec_id = uuid.uuid4().hex
    target = SPEC_DIR / f"{spec_id}.pdf"
    target.write_bytes(contents)
    return {"id": spec_id, "name": target.name}


@app.post("/api/invoices")
async def create_invoice(request: Request):
    try:
        invoice = validate_invoice(await request.json())
        for line in invoice["lines"]:
            if line["spec_id"] and not (SPEC_DIR / f"{line['spec_id']}.pdf").is_file():
                raise ValueError("A selected spec PDF is missing")
        number = next_number(DB_FILE, invoice["date"])
        pdf_path = OUTPUT_DIR / f"{number}.pdf"
        xlsx_path = OUTPUT_DIR / f"{number}.xlsx"
        render_pdf(invoice, number, SELLER, DETAILS, pdf_path, SPEC_DIR)
        render_xlsx(invoice, number, SELLER, DETAILS, xlsx_path)
    except (ValueError, TypeError, KeyError, sqlite3.Error) as exc:
        raise HTTPException(400, str(exc)) from None
    return {"number": number, "pdf": f"/api/invoices/{number}.pdf",
            "xlsx": f"/api/invoices/{number}.xlsx", "total": str(invoice["total"]),
            "vat_included": str(invoice["vat"])}


@app.get("/api/invoices/{filename}")
def download(filename: str):
    if not filename.endswith((".pdf", ".xlsx")):
        raise HTTPException(404)
    stem, ext = filename.rsplit(".", 1)
    import re
    if not re.fullmatch(r"PFI-\d{4}-\d{4,}", stem):
        raise HTTPException(404)
    path = OUTPUT_DIR / filename
    if not path.is_file():
        raise HTTPException(404)
    return FileResponse(path, filename=filename, media_type="application/pdf" if ext == "pdf" else
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=os.environ.get("PI_HOST", "127.0.0.1"), port=int(os.environ.get("PI_PORT", "8000")))

