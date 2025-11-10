# main.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
import fitz  # PyMuPDF
import io
import json
import re
from typing import List, Dict, Any

app = FastAPI(title="PDF Annexure Extract & Fill")

# Helper: decide if a text block looks like a question/label
QUESTION_PATTERNS = [
    re.compile(r"\bANNEXURE\b", re.IGNORECASE),
    re.compile(r"\bANNEXURE\s*[-:]?", re.IGNORECASE),
    re.compile(r"\bQUESTION\b", re.IGNORECASE),
    re.compile(r"\?$"),                      # ends with question mark
    re.compile(r"^\s*\d+\.\s*"),             # numbered lines like "1. Name:"
    re.compile(r":\s*$"),                    # ends with colon e.g. "Name:"
]


def looks_like_question(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    # Check patterns
    for p in QUESTION_PATTERNS:
        if p.search(t):
            return True
    # Some heuristics: short labels with colon or short lines with label words
    if len(t) <= 120 and (t.endswith(":") or ":" in t and len(t.split()) <= 6):
        return True
    return False


@app.post("/extract-questions")
async def extract_questions(pdf: UploadFile = File(...)):
    """
    Extract likely question/label fields from PDF and return:
    [{ "label": "...", "page": 1, "x": 123.4, "y": 456.7, "w": 50.0, "h": 12.0 }]
    x,y represent the suggested insertion point for the answer (right of label).
    Coordinates are PDF points (origin top-left).
    """
    if pdf.content_type != "application/pdf":
        raise HTTPException(status_code=400, detail="File must be a PDF")

    data = await pdf.read()
    doc = fitz.open(stream=data, filetype="pdf")

    fields: List[Dict[str, Any]] = []

    for page_index in range(len(doc)):
        page = doc[page_index]
        # Get text blocks: each block is (x0, y0, x1, y1, "text", block_no, block_type)
        blocks = page.get_text("blocks")
        for b in blocks:
            x0, y0, x1, y1, btext = b[0], b[1], b[2], b[3], b[4]
            text = str(btext).strip()
            # Split block into lines and evaluate each line separately
            for line_offset, line in enumerate(text.splitlines()):
                line = line.strip()
                if not line:
                    continue
                if looks_like_question(line):
                    # Approximate the line's vertical position within the block
                    block_height = y1 - y0 if (y1 - y0) > 0 else 1
                    # distribute lines equally to estimate y for this line
                    total_lines = max(1, len(text.splitlines()))
                    line_y = y0 + (line_offset / total_lines) * block_height
                    # Suggest insertion point just to the right of the block
                    insert_x = x1 + 5  # 5pt padding
                    insert_y = line_y
                    fields.append({
                        "label": line,
                        "page": page_index + 1,
                        "x": float(round(insert_x, 2)),
                        "y": float(round(insert_y, 2)),
                        "w": float(round(x1 - x0, 2)),
                        "h": float(round(block_height / total_lines, 2))
                    })

        # Additional heuristic: search for question mark occurrences (word-level)
        # page.search_for returns rectangles for matched text
        try:
            # find any question marks specifically (rare), and include their bbox
            matches = page.search_for("?")
            for r in matches:
                # r is Rect(x0, y0, x1, y1)
                insert_x = r.x1 + 5
                insert_y = r.y0
                fields.append({
                    "label": "?",
                    "page": page_index + 1,
                    "x": float(round(insert_x, 2)),
                    "y": float(round(insert_y, 2)),
                    "w": float(round(r.x1 - r.x0, 2)),
                    "h": float(round(r.y1 - r.y0, 2))
                })
        except Exception:
            pass

    return {"success": True, "fields": fields}


@app.post("/fill-pdf")
async def fill_pdf(
    pdf: UploadFile = File(...),
    answers_json: str = Form(...)
):
    """
    Receive original PDF and a JSON string of answers.
    answers_json should be either:
    1) List of objects: [{"page":1, "x":123.4, "y":456.7, "answer":"text"} ...]
    OR
    2) List of objects that reference label: [{"page":1, "label":"Name", "answer":"Baskaran"} ...]
       When label is provided, we attempt to locate that label on the page and place answer to its right.
    Returns filled PDF for download.
    """
    # validate PDF
    if pdf.content_type != "application/pdf":
        raise HTTPException(status_code=400, detail="File must be a PDF")

    try:
        answers = json.loads(answers_json)
        if not isinstance(answers, list):
            raise ValueError("answers_json must be a JSON list")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid answers_json: {e}")

    raw = await pdf.read()
    doc = fitz.open(stream=raw, filetype="pdf")

    for item in answers:
        # validate fields
        if "page" not in item:
            raise HTTPException(status_code=400, detail="Each answer must include 'page'")
        page_idx = int(item["page"]) - 1
        if page_idx < 0 or page_idx >= len(doc):
            raise HTTPException(status_code=400, detail=f"Page {item['page']} out of range")

        page = doc[page_idx]
        answer_text = str(item.get("answer", ""))

        # If x,y provided, use directly
        if "x" in item and "y" in item:
            x = float(item["x"])
            y = float(item["y"])
            # Insert text; try to place inside a textbox to handle wrapping
            # We'll use a small width area to the right of x (e.g., 300pt)
            rect = fitz.Rect(x, y, x + 350, y + 50)
            page.insert_textbox(rect, answer_text, fontsize=11, fontname="helv", align=0)
            continue

        # If label provided, try to find it on the page and place answer to its right
        label = item.get("label")
        if label:
            # use page.search_for to find occurrences of the label text
            found = page.search_for(label, hit_max=16)
            if not found:
                # fallback: try searching by lowercase / trimmed
                found = page.search_for(label.strip(), hit_max=16)
            if not found:
                # fallback: attempt to find by partial token (first few words)
                tokens = label.split()
                if tokens:
                    substring = " ".join(tokens[:3])
                    found = page.search_for(substring, hit_max=16)

            if found:
                # choose the first occurrence for now
                r = found[0]
                insert_x = r.x1 + 5
                insert_y = r.y0
                rect = fitz.Rect(insert_x, insert_y, insert_x + 350, insert_y + 50)
                page.insert_textbox(rect, answer_text, fontsize=11, fontname="helv", align=0)
                continue
            else:
                # if not found, place answer top-left of page as fallback
                rect = fitz.Rect(50, 50, 400, 100)
                page.insert_textbox(rect, answer_text, fontsize=11, fontname="helv", align=0)
                continue

        # If neither x/y nor label given, throw error
        raise HTTPException(status_code=400, detail="Answer item must include either x & y or label")

    # Save to bytes
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    headers = {"Content-Disposition": "attachment; filename=filled.pdf"}
    return StreamingResponse(out, media_type="application/pdf", headers=headers)


# Simple root
@app.get("/")
async def root():
    return {"message": "PDF Annexure Extract & Fill API. Use /extract-questions and /fill-pdf endpoints."}
