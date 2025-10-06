# pdf_prefill.py
# Robust prefill for AcroForm & non-form PDFs using PyMuPDF,
# with Section / Page / Index control, stable widget order,
# field-based occurrence counting, Yes/No option handling,
# and checkbox-square detection for bullet lists.

import os, re, csv, math, json, argparse, string, unicodedata
from difflib import SequenceMatcher
from typing import List, Dict, Any, Tuple, Optional, Set

import pandas as pd
import fitz  # PyMuPDF

# ---------------------------
# Config / globals
# ---------------------------
PUNCT = str.maketrans("", "", string.punctuation)

# Optional global defaults (lowest priority)
SUBSCRIPTION_DEFAULTS = {
    # "Name of Investor": "John Doe",
}

# Optional label aliases (normalize both sides)
FIELD_ALIASES = {
    "investor name": "name of investor",
    "name of investor printed or typed": "name of investor",
    # add more if needed...
}

# ---------------------------
# String / matching utils
# ---------------------------
def _normalize(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = s.translate(PUNCT)
    return s

def _sim(a: str, b: str) -> float:
    return SequenceMatcher(None, _normalize(a), _normalize(b)).ratio()

def _best_match_scored(label: str, candidates: List[str]) -> Tuple[Optional[str], float]:
    if not candidates:
        return None, 0.0
    best, best_score = None, 0.0
    for c in candidates:
        score = _sim(label, c)
        if score > best_score:
            best, best_score = c, score
    return best, best_score

def alias_normal(norm_label: str) -> str:
    return FIELD_ALIASES.get(norm_label, norm_label)

# truth helpers for Yes/No
def _truthy(val: str) -> Optional[bool]:
    s = (str(val or "")).strip().lower()
    if s in {"y", "yes", "true", "1", "x", "✓", "check", "checked"}:
        return True
    if s in {"n", "no", "false", "0", "uncheck", "unchecked"}:
        return False
    return None

def _rect_key(x0, y0, x1, y1):
    # coarse rounding to avoid tiny numeric jitter
    return (int(round(x0)), int(round(y0)), int(round(x1)), int(round(y1)))



# ---------------------------
# Lookup loader (Excel/CSV) with Section/Page/Index
# ---------------------------
def read_lookup_rows(path: str) -> List[Dict[str, Any]]:
    """
    Returns rows with keys:
      Field, Value, Section (optional), Page (optional, int), Index (optional, int)
    Adds: field_norm, section_norm
    """
    if not os.path.exists(path):
        print(f"⚠️  Lookup file not found: {path}")
        return []

    # Load
    try:
        if path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path)
    except Exception as e:
        print(f"⚠️  Could not load {path}: {e}")
        return []

    # Validate minimal columns
    if {"Field", "Value"} - set(df.columns):
        raise ValueError("Lookup must have at least columns: Field, Value. Optional: Section, Page, Index")

    # Clean text cols
    for col in ["Field", "Value", "Section"]:
        if col in df.columns:
            df[col] = df[col].astype(str).fillna("").map(lambda x: x.strip())

    # Page → Int
    if "Page" in df.columns:
        df["Page"] = pd.to_numeric(df["Page"], errors="coerce").astype("Int64")

    # Index → Int (after stripping junk like '#3')
    if "Index" in df.columns:
        idx_series = (
            df["Index"]
            .astype(str)
            .str.replace(r"[^\d\-]+", "", regex=True)
            .replace({"": None})
        )
        df["Index"] = pd.to_numeric(idx_series, errors="coerce").astype("Int64")

    # Drop empty values
    df = df[(df["Field"].astype(str).str.strip() != "") &
            (df["Value"].astype(str).str.strip() != "") &
            (df["Value"].astype(str).str.lower() != "nan")]

    rows: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        field = str(r.get("Field", "")).strip()
        value = str(r.get("Value", "")).strip()
        section = str(r.get("Section", "")).strip() if "Section" in df.columns else ""
        page = int(r["Page"]) if ("Page" in df.columns and pd.notna(r["Page"])) else None
        index = int(r["Index"]) if ("Index" in df.columns and pd.notna(r["Index"])) else None

        rows.append({
            "Field": field,
            "Value": value,
            "Section": section,
            "Page": page,
            "Index": index,
            "field_norm": alias_normal(_normalize(field)),
            "section_norm": _normalize(section) if section else "",
        })

    # Append defaults (lowest priority)
    for k, v in SUBSCRIPTION_DEFAULTS.items():
        rows.append({
            "Field": k,
            "Value": v,
            "Section": "",
            "Page": None,
            "Index": None,
            "field_norm": alias_normal(_normalize(k)),
            "section_norm": "",
        })
    return rows

def expected_section_set(lookup_rows: List[Dict[str, Any]]) -> Set[str]:
    return {r["section_norm"] for r in lookup_rows if r.get("section_norm")}

# ---------------------------
# Text / section helpers
# ---------------------------
_COLON_CHARS = {":", "：", "﹕", "꞉", "˸", "፡", "︓"}

def _strip_colon_like(s: str) -> str:
    s = s.strip()
    while s and s[-1] in _COLON_CHARS:
        s = s[:-1].rstrip()
    return s

def _text_blocks(page: fitz.Page) -> List[Dict[str, Any]]:
    blocks = []
    for b in page.get_text("blocks"):
        x0, y0, x1, y1, text, *_ = b
        blocks.append({"x0": float(x0), "y0": float(y0),
                       "x1": float(x1), "y1": float(y1),
                       "text": (text or "").strip()})
    blocks.sort(key=lambda r: (round(r["y0"], 1), round(r["x0"], 1)))
    return blocks

def _page_lines_with_fonts(page: fitz.Page) -> List[Dict[str, Any]]:
    out = []
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        for l in b.get("lines", []):
            texts, sizes, y_vals = [], [], []
            x0 = None; x1 = None
            for s in l.get("spans", []):
                t = (s.get("text") or "").strip()
                if not t:
                    continue
                texts.append(t)
                sizes.append(float(s.get("size", 0.0)))
                x0 = s["bbox"][0] if x0 is None else min(x0, s["bbox"][0])
                x1 = s["bbox"][2] if x1 is None else max(x1, s["bbox"][2])
                y_vals.append((s["bbox"][1] + s["bbox"][3]) / 2.0)
            if texts:
                out.append({
                    "text": " ".join(texts).strip(),
                    "y_mid": sum(y_vals) / len(y_vals) if y_vals else 0.0,
                    "x0": x0 or 0.0,
                    "x1": x1 or 0.0,
                    "max_size": max(sizes) if sizes else 0.0
                })
    return out

def _is_section_header_relaxed(text: str) -> bool:
    if not text:
        return False
    t = unicodedata.normalize("NFKC", text).strip()
    if len(t) > 180:
        return False
    ends_like_colon = (len(t) > 5 and t[-1] in _COLON_CHARS)
    lower_t = t.lower()
    contains_keywords = (
        "for completion by subscribers" in lower_t
        or "investor information" in lower_t
        or "subscription details" in lower_t
        or "wire information" in lower_t
        or "beneficiary bank" in lower_t
        or "intermediary bank" in lower_t
        or "erisa status" in lower_t
        or "signature" in lower_t
        or "for all subscribers" in lower_t
    )
    letters = [c for c in t if c.isalpha()]
    caps_ratio = (sum(1 for c in letters if c.isupper()) / max(1, len(letters))) if letters else 0.0
    shouty = caps_ratio >= 0.45
    return ends_like_colon or contains_keywords or (shouty and len(t.split()) <= 15)

def find_sections_on_page(page: fitz.Page,
                          expected_sections_norm: Optional[Set[str]] = None,
                          dry_run: bool = False) -> List[Dict[str, Any]]:
    expected_sections_norm = expected_sections_norm or set()
    lines = _page_lines_with_fonts(page)
    candidates = []

    # 1) relaxed headers
    for ln in lines:
        t = ln["text"].strip()
        if not t:
            continue
        if "\n" in t:
            t = t.splitlines()[0].strip()
        if _is_section_header_relaxed(t):
            name = _strip_colon_like(unicodedata.normalize("NFKC", t))
            candidates.append({"name": name, "name_norm": _normalize(name), "y1": ln["y_mid"]})

    # 2) fuzzy vs expected sections (from Excel)
    if expected_sections_norm:
        expected_list = list(expected_sections_norm)
        for ln in lines:
            norm_line = _normalize(_strip_colon_like(unicodedata.normalize("NFKC", ln["text"])))
            if len(norm_line) > 220:
                continue
            bm, sc = _best_match_scored(norm_line, expected_list)
            if bm and sc >= 0.82:
                name = _strip_colon_like(ln["text"])
                candidates.append({"name": name, "name_norm": norm_line, "y1": ln["y_mid"]})

    # 3) size-based catch (top 20% lines)
    if lines:
        sizes = sorted([ln["max_size"] for ln in lines if ln["max_size"] > 0])
        if sizes:
            thresh = sizes[int(0.80 * (len(sizes) - 1))]
            for ln in lines:
                if ln["max_size"] >= thresh and len(ln["text"]) <= 150:
                    t = _strip_colon_like(ln["text"])
                    if t and _is_section_header_relaxed(t):
                        candidates.append({"name": t, "name_norm": _normalize(t), "y1": ln["y_mid"]})

    candidates.sort(key=lambda c: (round(c["y1"], 1)))
    dedup, seen = [], set()
    for c in candidates:
        key = (c["name_norm"], round(c["y1"], 1))
        if key in seen:
            continue
        seen.add(key)
        dedup.append(c)

    if dry_run and dedup:
        print("  sections found:")
        for s in dedup:
            print("   -", s["name"])
    return dedup

def nearest_section_name(sections: List[Dict[str, Any]], y_mid: float) -> Tuple[str, str]:
    above = [s for s in sections if s["y1"] <= y_mid]
    if not above:
        return "", ""
    return above[-1]["name"], above[-1]["name_norm"]

# ---------------------------
# Yes/No detection around a point
# ---------------------------
def _nearby_yes_no_option(page: fitz.Page, x: float, y: float) -> Optional[str]:
    """Return 'yes' or 'no' if a token is detected near (x,y)."""
    XRANGE = 140.0
    YRANGE = 24.0
    nearest, best_dx = None, 1e9
    for b in page.get_text("blocks"):
        x0, y0, x1, y1, text, *_ = b
        t = (text or "").strip()
        if not t or len(t) > 6:  # very short tokens only
            continue
        ymid = (y0 + y1) / 2.0
        if abs(ymid - y) > YRANGE:
            continue
        if (x0 - XRANGE) <= x <= (x1 + XRANGE):
            token = t.lower().strip("._:-)()]}[({")
            if token in {"yes", "no"}:
                dx = min(abs(x - x0), abs(x - x1))
                if dx < best_dx:
                    best_dx = dx
                    nearest = token
    return nearest

# ---------------------------
# Drawn underline detection (for non-form PDFs)
# ---------------------------
def _line_like_segments(
    page: fitz.Page,
    min_len=60,
    max_len=1200,
    max_slope=0.02,
    max_thick=2.5
):
    """
    Return horizontal-ish line segments on the page, tolerating variations in how
    PyMuPDF encodes drawing items. Handles:
      - ('l', (x0, y0), (x1, y1))
      - ('re', x, y, w, h, ...)  OR  ('re', (x, y, w, h))
      - ignores other commands safely
    """
    def _as_float_tuple(obj, n):
        """Try to coerce obj to an n-length tuple of floats; else None."""
        try:
            if isinstance(obj, (list, tuple)) and len(obj) >= n:
                return tuple(float(obj[i]) for i in range(n))
        except Exception:
            pass
        return None

    segs = []
    drawings = page.get_drawings() or []
    for d in drawings:
        items = d.get("items", []) or []
        for it in items:
            if not it:
                continue

            cmd = it[0]

            # ----- line segments -----
            if cmd == "l":
                # expected: ('l', (x0, y0), (x1, y1))
                if len(it) >= 3:
                    p0 = _as_float_tuple(it[1], 2)
                    p1 = _as_float_tuple(it[2], 2)
                    if p0 and p1:
                        x0, y0 = p0
                        x1, y1 = p1
                        dx, dy = (x1 - x0), (y1 - y0)
                        length = math.hypot(dx, dy)
                        # keep “horizontal-ish” lines within size bounds
                        if (min_len <= length <= max_len) and (abs(dy) / (abs(dx) + 1e-6) <= max_slope):
                            xL, xR = (x0, x1) if x0 <= x1 else (x1, x0)
                            yMid = (y0 + y1) / 2.0
                            segs.append({"x0": xL, "x1": xR, "y0": yMid, "y1": yMid, "len": length})
                continue  # move to next item

            # ----- rectangle strokes (underlines often drawn as very thin rects) -----
            if cmd == "re":
                # Accept either ('re', x, y, w, h, ...)  OR  ('re', (x, y, w, h), ...)
                x = y = w = h = None

                if len(it) >= 5:
                    # ('re', x, y, w, h, ...)
                    try:
                        x = float(it[1]); y = float(it[2]); w = float(it[3]); h = float(it[4])
                    except Exception:
                        x = y = w = h = None

                if x is None:
                    # maybe ('re', (x, y, w, h), ...)
                    if len(it) >= 2:
                        rect4 = _as_float_tuple(it[1], 4)
                        if rect4:
                            x, y, w, h = rect4

                if x is None or w is None or h is None:
                    # unsupported/odd shape – skip safely
                    continue

                if h <= max_thick and (min_len <= w <= max_len):
                    xL, xR = x, x + w
                    yMid = y + h / 2.0
                    segs.append({"x0": xL, "x1": xR, "y0": yMid, "y1": yMid, "len": w})
                continue  # move to next item

            # else: ignore other drawing commands safely and continue
            continue

    # sort + de-dup near-identical segments
    segs.sort(key=lambda s: (round(s["y0"], 1), round(s["x0"], 1)))
    merged = []
    for s in segs:
        if merged:
            m = merged[-1]
            if abs(s["y0"] - m["y0"]) < 0.8 and abs(s["x0"] - m["x0"]) < 2 and abs(s["x1"] - m["x1"]) < 2:
                continue
        merged.append(s)
    return merged


# ---------------------------
# Checkbox square detection (new)
# ---------------------------
def _square_checkboxes(page: fitz.Page,
                       min_side: float = 8.0,
                       max_side: float = 20.0,
                       max_aspect: float = 1.35) -> List[Dict[str, float]]:
    """
    Return small, nearly-square rectangles likely to be checkboxes.
    Handles PyMuPDF variations where 're' items can be:
      - ('re', x, y, w, h, ...)
      - ('re', (x, y, w, h), ...)
      - ('re', fitz.Rect(...), ...)
    """
    def _as_float_tuple(obj, n):
        try:
            if isinstance(obj, (list, tuple)) and len(obj) >= n:
                return tuple(float(obj[i]) for i in range(n))
        except Exception:
            pass
        return None

    boxes: List[Dict[str, float]] = []
    for d in page.get_drawings() or []:
        for it in d.get("items", []) or []:
            if not it or it[0] != "re":
                continue

            x = y = w = h = None

            # Case A: flat scalars ('re', x, y, w, h, ...)
            if len(it) >= 5:
                try:
                    x = float(it[1]); y = float(it[2]); w = float(it[3]); h = float(it[4])
                except Exception:
                    x = y = w = h = None

            # Case B: ('re', Rect, ...)
            if x is None and len(it) >= 2 and isinstance(it[1], fitz.Rect):
                r: fitz.Rect = it[1]  # type: ignore
                x, y, w, h = float(r.x0), float(r.y0), float(r.width), float(r.height)

            # Case C: ('re', (x,y,w,h), ...)
            if x is None and len(it) >= 2:
                rect4 = _as_float_tuple(it[1], 4)
                if rect4:
                    x, y, w, h = rect4

            if x is None or w is None or h is None:
                continue  # unrecognized shape -> skip safely

            # small-ish and nearly square
            side_min = min(w, h)
            side_max = max(w, h)
            if side_min <= 0:
                continue
            aspect = side_max / side_min
            if (min_side <= side_min <= max_side) and aspect <= max_aspect:
                boxes.append({
                    "x0": x, "y0": y,
                    "x1": x + w, "y1": y + h,
                    "cx": x + w / 2.0, "cy": y + h / 2.0,
                    "w": w, "h": h
                })

    # de-dupe near-identical boxes (PDFs sometimes draw twice)
    boxes.sort(key=lambda b: (round(b["y0"], 1), round(b["x0"], 1)))
    dedup: List[Dict[str, float]] = []
    for b in boxes:
        if dedup:
            m = dedup[-1]
            if (abs(b["x0"] - m["x0"]) < 1.0 and abs(b["y0"] - m["y0"]) < 1.0 and
                abs(b["x1"] - m["x1"]) < 1.0 and abs(b["y1"] - m["y1"]) < 1.0):
                continue
        dedup.append(b)
    return dedup

# alias to tolerate either helper name in code
_checkbox_squares = _square_checkboxes

# ---------------------------
# Template builder (non-form PDFs)
# ---------------------------
def _nearest_label_for_block(ul_block: Dict[str, Any], neighbor_blocks: List[Dict[str, Any]], y_tol=22.0, back_scan=8) -> str:
    same_line = [b for b in neighbor_blocks
                 if abs(b["y0"] - ul_block["y0"]) < y_tol and b["x1"] <= ul_block["x0"] and b["text"]]
    if same_line:
        same_line.sort(key=lambda b: ul_block["x0"] - b["x1"])
        return same_line[0]["text"]

    idx = neighbor_blocks.index(ul_block) if ul_block in neighbor_blocks else len(neighbor_blocks) - 1
    cands = []
    for j in range(max(0, idx - back_scan), idx):
        b = neighbor_blocks[j]
        if not b["text"]:
            continue
        if 0 <= (ul_block["y0"] - b["y0"]) < (y_tol * 2.0):
            cands.append(b)
    if cands:
        cands.sort(key=lambda b: (abs(ul_block["y0"] - b["y0"]), abs(ul_block["x0"] - b["x1"])))
        return cands[0]["text"]
    return ""

def build_pdf_template(pdf_path: str,
                       template_json: str = "template_fields.json",
                       lookup_rows: Optional[List[Dict[str, Any]]] = None,
                       dry_run: bool = False):
    import re

    # --- helpers -------------------------------------------------------------
    def _section_core(name: str) -> str:
        """
        Make section names comparable to lookup rows:
        - strip trailing parentheticals: 'For All Subscribers (U.S. and Non-U.S.)' -> 'For All Subscribers'
        - squeeze spaces
        """
        s = (name or "").strip()
        s = re.sub(r"\s*\(.*?\)\s*$", "", s).strip()
        return s

    def _short_label(section_name: str, ordinal: int) -> str:
        """Concise key like 'For All Subscribers – item 1'."""
        core = _section_core(section_name) or "Checklist"
        return f"{core} – item {ordinal}"

    expected = expected_section_set(lookup_rows or [])
    doc = fitz.open(pdf_path)
    tpl = {"pdf": pdf_path, "fields": []}
    underline_re = re.compile(r'_{3,}')
    counters: Dict[Tuple[int, str, str], int] = {}

    for pno, page in enumerate(doc, start=1):
        if dry_run:
            print(f"p{pno}:")
        sections = find_sections_on_page(page, expected_sections_norm=expected, dry_run=dry_run)
        blocks = _text_blocks(page)
        y_tol = 18.0

        # === NEW: collect checkbox / radio widget rects right away for this page
        try:
            _all_widgets = list(page.widgets() or [])
        except TypeError:
            _all_widgets = list(page.widgets or [])
        cb_rects: List[fitz.Rect] = []
        for _w in _all_widgets or []:
            try:
                if _w.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
                    cb_rects.append(fitz.Rect(_w.rect))
            except Exception:
                # best-effort fallback by name
                nm = (getattr(_w, "field_name", "") or "").lower()
                if "check" in nm or "box" in nm or "radio" in nm:
                    cb_rects.append(fitz.Rect(_w.rect))

        # ---------- 1) typed underscores ----------
        for i, b in enumerate(blocks):
            txt = b["text"]
            if not txt:
                continue
            m = underline_re.search(txt)
            if not m:
                continue

            left_txt = txt.split('_')[0] if '_' in txt else ""
            label_text = (left_txt or _nearest_label_for_block(b, blocks) or f"unknown_{pno}_{i}").strip()

            insert_x   = (b["x0"] + m.start() * 5.0) if m else (b["x1"] + 10.0)
            baseline_y = (b["y0"] + b["y1"]) / 2.0
            y_mid      = baseline_y

            section_name, section_norm = nearest_section_name(sections, y_mid)
            label_norm = alias_normal(_normalize(label_text))

            key = (pno, section_norm, label_norm)
            counters[key] = counters.get(key, 0) + 1
            idx = counters[key]

            entry = {
                "page": pno,
                "label": label_text,
                "label_short": label_text,
                "label_full": label_text,
                "label_norm": label_norm,
                "anchor_x": insert_x,
                "anchor_y": baseline_y,
                "line_box": [b["x0"], b["y0"], b["x1"], b["y1"]],
                "placement": "start",
                "section": section_name,
                "section_norm": section_norm,
                "index": idx,
            }
            tpl["fields"].append(entry)
            if dry_run:
                print(f"   • field[{idx}] (typed underscores) → '{entry['label']}' @y≈{y_mid:.1f} (sec: {section_name})")

        # ---------- 2) drawn underlines ----------
        segs = _line_like_segments(page)
        for j, seg in enumerate(segs):
            same_row = [blk for blk in blocks
                        if blk["text"]
                        and abs(((blk["y0"] + blk["y1"]) / 2.0) - seg["y0"]) < y_tol
                        and blk["x1"] <= seg["x0"] + 4]
            if same_row:
                same_row.sort(key=lambda blk: seg["x0"] - blk["x1"])
                label_text = same_row[0]["text"]
            else:
                above = [blk for blk in blocks
                         if blk["text"] and (0 <= (seg["y0"] - blk["y1"]) < 2 * y_tol) and blk["x0"] <= seg["x0"]]
                if above:
                    above.sort(key=lambda blk: (seg["y0"] - blk["y1"], seg["x0"] - blk["x1"]))
                    label_text = above[0]["text"]
                else:
                    label_text = f"unknown_drawn_{pno}_{j}"

            insert_x = seg["x0"] + 6
            y_mid    = seg["y0"]

            section_name, section_norm = nearest_section_name(sections, y_mid)
            label_norm = alias_normal(_normalize(label_text))
            key = (pno, section_norm, label_norm)
            counters[key] = counters.get(key, 0) + 1
            idx = counters[key]

            entry = {
                "page": pno,
                "label": label_text.strip(),
                "label_short": label_text.strip(),
                "label_full": label_text.strip(),
                "label_norm": label_norm,
                "anchor_x": insert_x,
                "anchor_y": y_mid,
                "line_box": [seg["x0"], seg["y0"], seg["x1"], seg["y0"]],
                "placement": "center",
                "section": section_name,
                "section_norm": section_norm,
                "index": idx,
            }
            tpl["fields"].append(entry)
            if dry_run:
                print(f"   • field[{idx}] (drawn line)        → '{entry['label']}' @y≈{y_mid:.1f} (sec: {section_name})")

        # ---------- 3) drawn square checkboxes (vectors) ----------
        vector_boxes = _square_checkboxes(page) if '_square_checkboxes' in globals() else []

        # === NEW: drop vector boxes that overlap real checkbox / radio widgets
        pruned_vector_boxes = []
        for bx in vector_boxes:
            B = fitz.Rect(bx["x0"], bx["y0"], bx["x1"], bx["y1"])
            if any(B.intersects(R) for R in cb_rects):
                continue
            pruned_vector_boxes.append(bx)

        # (loop uses pruned_vector_boxes)
        for k, bx in enumerate(pruned_vector_boxes, start=1):

            cx = bx.get("cx", (bx["x0"] + bx["x1"]) / 2.0)
            cy = bx.get("cy", (bx["y0"] + bx["y1"]) / 2.0)

            right_cands = [
                (
                    abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) + max(0.0, blk["x0"] - bx["x1"]),
                    blk
                )
                for blk in blocks
                if blk["text"] and abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) < y_tol and blk["x0"] >= bx["x1"] - 2
            ]
            if right_cands:
                right_cands.sort(key=lambda t: t[0])
                bullet_text = right_cands[0][1]["text"].strip()
            else:
                left_cands = [
                    (
                        abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) + max(0.0, bx["x0"] - blk["x1"]),
                        blk
                    )
                    for blk in blocks
                    if blk["text"] and abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) < y_tol and blk["x1"] <= bx["x0"] + 2
                ]
                if left_cands:
                    left_cands.sort(key=lambda t: t[0])
                    bullet_text = left_cands[0][1]["text"].strip()
                else:
                    bullet_text = ""

            section_name, section_norm = nearest_section_name(sections, cy)

            short_key  = _short_label(section_name, k)          # e.g., "For All Subscribers – item 1"
            pretty     = f"{short_key}: {bullet_text}".strip(": ")
            label_norm = alias_normal(_normalize(pretty))
            key = (pno, section_norm, label_norm)
            counters[key] = counters.get(key, 0) + 1
            idx = counters[key]

            entry = {
                "page": pno,
                "label": pretty,                   # readable label
                "label_short": short_key,          # concise canonical key for lookup
                "label_full": bullet_text,         # raw bullet text
                "label_norm": label_norm,
                "anchor_x": cx,
                "anchor_y": cy,
                "box_rect": [bx["x0"], bx["y0"], bx["x1"], bx["y1"]],
                "line_box": [bx["x0"], bx["y0"], bx["x1"], bx["y1"]],
                "placement": "checkbox",
                "section": section_name,
                "section_norm": section_norm,
                "index": idx,
            }
            tpl["fields"].append(entry)
            if dry_run:
                print(f"   • checkbox[{idx}] (vector)         → '{entry['label']}' "
                      f"[short='{entry['label_short']}'] @y≈{cy:.1f} (sec: {section_name})")

        # ---------- 4) REAL AcroForm checkbox/radio widgets ----------
        try:
            widgets = list(page.widgets() or [])
        except TypeError:
            widgets = list(page.widgets or [])
        cb_widgets = []
        for w in widgets:
            try:
                if w.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
                    cb_widgets.append(w)
            except Exception:
                nm = (w.field_name or "").lower()
                if "check" in nm or "box" in nm or "radio" in nm:
                    cb_widgets.append(w)

        cb_widgets.sort(key=lambda ww: (round(ww.rect.y0, 2), round(ww.rect.x0, 2)))
        for m, w in enumerate(cb_widgets, start=1):
            r  = w.rect
            cx = (r.x0 + r.x1) / 2.0
            cy = (r.y0 + r.y1) / 2.0

            right_cands = [
                (
                    abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) + max(0.0, blk["x0"] - r.x1),
                    blk
                )
                for blk in blocks
                if blk["text"]
                and abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) < y_tol
                and blk["x0"] >= r.x1 - 2
            ]
            if right_cands:
                right_cands.sort(key=lambda t: t[0])
                bullet_text = right_cands[0][1]["text"].strip()
            else:
                left_cands = [
                    (
                        abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) + max(0.0, r.x0 - blk["x1"]),
                        blk
                    )
                    for blk in blocks
                    if blk["text"]
                    and abs(((blk["y0"] + blk["y1"]) / 2.0) - cy) < y_tol
                    and blk["x1"] <= r.x0 + 2
                ]
                if left_cands:
                    left_cands.sort(key=lambda t: t[0])
                    bullet_text = left_cands[0][1]["text"].strip()
                else:
                    bullet_text = ""

            section_name, section_norm = nearest_section_name(sections, cy)
            short_key  = _short_label(section_name, m)
            pretty     = f"{short_key}: {bullet_text}".strip(": ")
            label_norm = alias_normal(_normalize(pretty))
            key = (pno, section_norm, label_norm)
            counters[key] = counters.get(key, 0) + 1
            idx = counters[key]

            entry = {
                "page": pno,
                "label": pretty,
                "label_short": short_key,          # concise canonical key
                "label_full": bullet_text,
                "label_norm": label_norm,
                "anchor_x": cx,
                "anchor_y": cy,
                "box_rect": [float(r.x0), float(r.y0), float(r.x1), float(r.y1)],
                "line_box": [float(r.x0), float(r.y0), float(r.x1), float(r.y1)],
                "placement": "checkbox",
                "section": section_name,
                "section_norm": section_norm,
                "index": idx,
            }
            tpl["fields"].append(entry)
            if dry_run:
                print(
                    f"   • checkbox[{idx}] (AcroForm)       → '{entry['label']}' "
                    f"[short='{entry['label_short']}'] @y≈{cy:.1f} (sec: {section_name})"
                )

        # ---------- 5) REAL AcroForm TEXT widgets (adds things like "Subscriber's Name") ----------
        try:
            widgets_all = list(page.widgets() or [])
        except TypeError:
            widgets_all = list(page.widgets or [])

        text_widgets = []
        for w in widgets_all:
            try:
                if w.field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    text_widgets.append(w)
            except Exception:
                # heuristic fallback by name
                nm = (w.field_name or "").lower()
                if any(k in nm for k in ["text", "name", "address", "city", "state", "zip"]):
                    text_widgets.append(w)

        text_widgets.sort(key=lambda ww: (round(ww.rect.y0, 2), round(ww.rect.x0, 2)))

        for t_idx, w in enumerate(text_widgets, start=1):
            r = w.rect
            y_mid = (r.y0 + r.y1) / 2.0

            same_row = [blk for blk in blocks
                        if blk["text"]
                        and abs(((blk["y0"] + blk["y1"]) / 2.0) - y_mid) < y_tol
                        and blk["x1"] <= r.x0 + 4]
            if same_row:
                same_row.sort(key=lambda blk: r.x0 - blk["x1"])
                label_text = same_row[0]["text"]
            else:
                above = [blk for blk in blocks
                         if blk["text"]
                         and (0 <= (y_mid - blk["y1"]) < 2 * y_tol)
                         and blk["x0"] <= r.x0]
                if above:
                    above.sort(key=lambda blk: (y_mid - blk["y1"], r.x0 - blk["x1"]))
                    label_text = above[0]["text"]
                else:
                    label_text = (w.field_name or f"acro_text_{pno}_{t_idx}")

            section_name, section_norm = nearest_section_name(sections, y_mid)
            label_norm = alias_normal(_normalize(label_text))

            key = (pno, section_norm, label_norm)
            counters[key] = counters.get(key, 0) + 1
            idx = counters[key]

            entry = {
                "page": pno,
                "label": label_text.strip(),
                "label_short": label_text.strip(),
                "label_full": label_text.strip(),
                "label_norm": label_norm,
                "anchor_x": float((r.x0 + r.x1) / 2.0),
                "anchor_y": float((r.y0 + r.y1) / 2.0),
                "box_rect": [float(r.x0), float(r.y0), float(r.x1), float(r.y1)],
                "line_box": [float(r.x0), float(r.y0), float(r.x1), float(r.y1)],
                "placement": "acro_text",
                "section": section_name,
                "section_norm": section_norm,
                "index": idx,
            }
            tpl["fields"].append(entry)
            if dry_run:
                print(f"   • field[{idx}] (AcroForm text)     → '{entry['label']}' @y≈{y_mid:.1f} (sec: {section_name})")

    # <<< end for pno, page loop

    with open(template_json, "w", encoding="utf-8") as f:
        json.dump(tpl, f, indent=2, ensure_ascii=False)
    print(f"🧩 Template saved to {template_json} with {len(tpl['fields'])} fields.")
    doc.close()
    return template_json





# ---------------------------
# Resolver: choose best row by Page / Section / Index / Field
# ---------------------------
def resolve_value(rows: List[Dict[str, Any]],
                  field_label: str,
                  page: Optional[int],
                  section_norm: str,
                  occurrence_index: int,
                  min_field_fuzzy: float = 0.82,
                  return_row: bool = False,
                  strict_index: bool = True,
                  require_page_match: bool = False,       # NEW
                  require_section_match: bool = False):    # NEW
    """
    Select value for a given (Field label, Page, Section, occurrence_index).

    strict_index=True:
      - If ANY candidate rows for this Field specify an Index, require Index == occurrence_index.
      - If NO candidate rows have Index, ignore Index and select by Section/Page.

    require_page_match=True:
      - If 'page' is provided, only accept rows whose Page == page. If none exist, return None.

    require_section_match=True:
      - If 'section_norm' is provided, only accept rows whose section_norm == section_norm. If none exist, return None.
    """
    field_norm = alias_normal(_normalize(field_label))

    # 1) Candidate by exact field; else fuzzy
    field_candidates = [r for r in rows if r["field_norm"] == field_norm]
    if not field_candidates:
        all_fields = [r["field_norm"] for r in rows]
        bm, sc = _best_match_scored(field_norm, all_fields)
        if bm and sc >= min_field_fuzzy:
            field_candidates = [r for r in rows if r["field_norm"] == bm]
        else:
            return (None, None) if return_row else None

    candidates = list(field_candidates)

    # 2) Section handling
    if section_norm:
        sec_exact = [r for r in candidates if r.get("section_norm") == section_norm]
        if require_section_match:
            # hard require section match
            candidates = sec_exact
            if not candidates:
                return (None, None) if return_row else None
        elif sec_exact:
            # prefer (soft)
            candidates = sec_exact

    # 3) Page handling
    if page is not None:
        pg_exact = [r for r in candidates if r.get("Page") is not None and int(r["Page"]) == int(page)]
        if require_page_match:
            # hard require page match
            candidates = pg_exact
            if not candidates:
                return (None, None) if return_row else None
        elif pg_exact:
            # prefer (soft)
            candidates = pg_exact

    # 4) Index handling
    if strict_index:
        with_idx = [r for r in candidates if r.get("Index") is not None]
        if with_idx:
            exact_idx = [r for r in with_idx if int(r["Index"]) == int(occurrence_index)]
            if not exact_idx:
                return (None, None) if return_row else None
            candidates = exact_idx

    # 5) Tie-breaker
    def score_row(r: Dict[str, Any]) -> int:
        s = 0
        if r.get("Page") is not None and page is not None and int(r["Page"]) == int(page):
            s += 8
        if r.get("section_norm") and section_norm and r["section_norm"] == section_norm:
            s += 6
        if r.get("Index") is not None and int(r["Index"]) == int(occurrence_index):
            s += 3
        return s

    best = max(candidates, key=score_row, default=None)
    if return_row:
        return (best["Value"], best) if best else (None, None)
    return best["Value"] if best else None

# ---------------------------
# AcroForm fill (stable order + field-based Index + Yes/No)
# ---------------------------
def _get_widgets(page: fitz.Page):
    try:
        it = page.widgets() or []
    except TypeError:
        it = page.widgets or []
    return list(it)

# --- helper: collect explicit indices for a given field/page/(section) ---
def _explicit_indices_for(rows, field_norm: str, page: int, section_norm: str) -> List[int]:
    """Return sorted list of explicit Index values available for this (field, page, section?)."""
    cand = [r for r in rows if r.get("field_norm") == field_norm]
    if page is not None:
        cand = [r for r in cand if r.get("Page") is not None and int(r["Page"]) == int(page)]
    if section_norm:
        cand = [r for r in cand if r.get("section_norm") == section_norm]
    cand = [r for r in cand if r.get("Index") is not None]
    idxs = sorted({int(r["Index"]) for r in cand})
    return idxs

def fill_acroform_with_context(input_pdf: str,
                               output_pdf: str,
                               lookup_rows: List[Dict[str, Any]],
                               dry_run: bool = False):
    expected = expected_section_set(lookup_rows)
    doc = fitz.open(input_pdf)
    changed: List[Tuple[str, str, str, int, str, int]] = []
    widgets_exist = False

    for pno in range(len(doc)):
        page = doc[pno]
        if dry_run:
            print(f"p{pno+1}:")
        sections = find_sections_on_page(page, expected_sections_norm=expected, dry_run=dry_run)
        blocks = _text_blocks(page)

        widgets = list(page.widgets() or [])
        widgets.sort(key=lambda w: (round(w.rect.y0, 2), round(w.rect.x0, 2)))
        widgets_exist = widgets_exist or bool(widgets)

        # counters / bookkeeping
        occ_counters: Dict[Tuple[int, str, str], int] = {}
        used_indices: Dict[Tuple[int, str, str], Set[int]] = {}

        for w in widgets:
            try:
                name = (w.field_name or "").strip()
                # do NOT skip pre-filled – we may intentionally overwrite
                try:
                    cur_existing = w.field_value if w.field_value is not None else ""
                except Exception:
                    cur_existing = ""

                r    = w.rect
                midy = (r.y0 + r.y1) / 2.0

                # nearest-left label guess
                y_tol = 18.0
                lefts = [blk for blk in blocks
                         if blk["text"]
                         and abs(((blk["y0"] + blk["y1"]) / 2.0) - midy) < y_tol
                         and blk["x1"] <= r.x0 + 4]
                if lefts:
                    lefts.sort(key=lambda blk: r.x0 - blk["x1"])
                    label_guess = lefts[0]["text"]
                else:
                    above = [blk for blk in blocks
                             if blk["text"]
                             and (0 <= (midy - blk["y1"]) < 2 * y_tol)
                             and blk["x0"] <= r.x0]
                    if above:
                        above.sort(key=lambda blk: (midy - blk["y1"], r.x0 - blk["x1"]))
                        label_guess = above[0]["text"]
                    else:
                        label_guess = name or ""

                # section nearest to the widget
                section_name, detected_section_norm = nearest_section_name(sections, midy)

                # ---- Phase 1: discover FIELD bucket (ignore Index) ----
                _, picked_row = resolve_value(
                    lookup_rows,
                    label_guess or name,
                    page=pno + 1,
                    section_norm=detected_section_norm,
                    occurrence_index=1,
                    return_row=True,
                    strict_index=False,
                )
                if not picked_row:
                    continue

                field_norm          = picked_row["field_norm"]
                excel_section_norm  = picked_row.get("section_norm") or ""
                effective_section   = excel_section_norm or detected_section_norm

                key = (pno + 1, effective_section, field_norm)

                # ---- Choose occurrence index: explicit first, else fallback counter ----
                explicit = _explicit_indices_for(lookup_rows, field_norm, pno + 1, effective_section)
                if explicit:
                    used = used_indices.setdefault(key, set())
                    next_idx = None
                    for idx_val in explicit:
                        if idx_val not in used:
                            next_idx = idx_val
                            break
                    if next_idx is None:
                        occ_counters[key] = occ_counters.get(key, 0) + 1
                        next_idx = occ_counters[key]
                    used.add(next_idx)
                    idx = next_idx
                else:
                    occ_counters[key] = occ_counters.get(key, 0) + 1
                    idx = occ_counters[key]

                # ---- Phase 2: resolve with progressive fallback ----
                value = resolve_value(
                    lookup_rows,
                    label_guess or name,
                    page=pno + 1,
                    section_norm=effective_section,
                    occurrence_index=idx,
                    strict_index=True,
                )
                if value is None:
                    # 2a) ignore Index
                    value = resolve_value(
                        lookup_rows, label_guess or name,
                        page=pno + 1, section_norm=effective_section,
                        occurrence_index=idx, strict_index=False
                    )
                if value is None and effective_section:
                    # 2b) ignore Section too
                    value = resolve_value(
                        lookup_rows, label_guess or name,
                        page=pno + 1, section_norm="",  # relax section
                        occurrence_index=idx, strict_index=False
                    )
                if value is None:
                    # 2c) last resort: ignore Page & Section
                    value = resolve_value(
                        lookup_rows, label_guess or name,
                        page=None, section_norm="",
                        occurrence_index=idx, strict_index=False
                    )
                if value is None:
                    continue

                # ---- Yes/No handling (ERISA etc.) ----
                option = _nearby_yes_no_option(page, r.x0, (r.y0 + r.y1) / 2.0)
                if not option:
                    nm = (name or "").lower()
                    if "yes" in nm:
                        option = "yes"
                    elif "no" in nm:
                        option = "no"

                to_write = str(value)
                if option in {"yes", "no"}:
                    yn = _truthy(value)
                    if yn is None:
                        if dry_run:
                            print(f"[DRY] skip unclear yes/no for '{name}'")
                        continue
                    match = (yn and option == "yes") or ((yn is False) and option == "no")
                    if not match:
                        if dry_run:
                            print(f"[DRY] skip non-matching option '{name}' for value '{value}'")
                        continue

                    # Acrobat-proof visual X inside the widget — do not write the word "Yes"
                    if dry_run:
                        print(f"[DRY] X inside widget '{name or '(unnamed)'}' at "
                              f"({r.x0:.1f},{r.y0:.1f},{r.x1:.1f},{r.y1:.1f})")
                    else:
                        pad = max(0.8, 0.18 * min(r.width, r.height))
                        width = max(1.6, 0.22 * min(r.width, r.height))
                        p1 = fitz.Point(r.x0 + pad, r.y0 + pad)
                        p2 = fitz.Point(r.x1 - pad, r.y1 - pad)
                        p3 = fitz.Point(r.x0 + pad, r.y1 - pad)
                        p4 = fitz.Point(r.x1 - pad, r.y0 + pad)
                        page.draw_line(p1, p2, width=width, color=(0, 0, 0), overlay=True)
                        page.draw_line(p3, p4, width=width, color=(0, 0, 0), overlay=True)
                        changed.append((name or "(unnamed)", label_guess, "X", pno + 1, section_name, idx))
                    continue

                # Non yes/no flow: write value (allow overwrite unless identical)
                if dry_run:
                    print(f"[DRY] fill '{name or '(unnamed)'}' ← '{label_guess or name}' "
                          f"(sec='{section_name}' idx={idx}): '{to_write}'")
                else:
                    try:
                        cur_text = cur_existing
                        if isinstance(cur_text, bytes):
                            cur_text = cur_text.decode("utf-8", "ignore")
                    except Exception:
                        cur_text = ""
                    if str(cur_text) != to_write:
                        w.field_value = to_write
                        w.update()
                        changed.append((name or "(unnamed)", label_guess, to_write, pno + 1, section_name, idx))

            except Exception as e:
                print(f"⚠️ Widget update error on page {pno+1}: {e}")

        if widgets:
            try:
                page.apply_redactions()
            except Exception:
                pass

    if dry_run:
        print("Dry-run complete (AcroForm). No file written.")
        doc.close()
        return changed, widgets_exist

    doc.save(output_pdf, incremental=False)
    doc.close()
    if changed:
        print(f"📝 AcroForm write complete → {output_pdf}")
    else:
        print(f"ℹ️ AcroForm path wrote a copy → {output_pdf}")
    return changed, widgets_exist

def _why_skip(label: str, idx: int, reason: str, page_no: int, dry_run: bool):
    if dry_run:
        print(f"[DRY][SKIP] p{page_no} '{label}' (idx={idx}) -> {reason}")

def _find_any_widget_overlapping(
    page: fitz.Page,
    box: Tuple[float, float, float, float]
):
    """
    Return the first widget whose rect intersects `box`.
    Prefers checkbox / radio widgets if multiple overlap.
    """
    # Get widgets robustly across PyMuPDF versions
    try:
        widgets = list(page.widgets() or [])
    except TypeError:
        widgets = list(page.widgets or [])
    except Exception:
        widgets = []

    if not widgets:
        return None

    target = fitz.Rect(*box)

    # 1) Prefer checkbox / radio widgets
    for w in widgets:
        try:
            if getattr(w, "field_type", None) in (
                getattr(fitz, "PDF_WIDGET_TYPE_CHECKBOX", None),
                getattr(fitz, "PDF_WIDGET_TYPE_RADIOBUTTON", None),
            ):
                if fitz.Rect(w.rect).intersects(target):
                    return w
        except Exception:
            pass

    # 2) Otherwise return any intersecting widget
    for w in widgets:
        try:
            if fitz.Rect(w.rect).intersects(target):
                return w
        except Exception:
            pass

    return None


# ---------------------------
# Overlay fill (template) with field-based Index + Yes/No + checkboxes
# ---------------------------

def fill_from_template(pdf_path: str,
                       template_json: str,
                       lookup_rows: List[Dict[str, Any]],
                       out_pdf: str,
                       center_on_line: bool = True,
                       font_size: float = 10.5,
                       min_field_fuzzy: float = 0.82,
                       dry_run: bool = False):
    import json
    with open(template_json, "r", encoding="utf-8") as f:
        tpl = json.load(f)

    if dry_run:
        from collections import Counter
        page_counts = Counter(fd.get("page") for fd in tpl.get("fields", []))
        total = len(tpl.get("fields", []))
        print(f"🧩 Template loaded: {total} fields")
        for p in sorted(page_counts):
            print(f"   • page {p}: {page_counts[p]} fields")

    doc = fitz.open(pdf_path)
    filled = 0

    # ---------- tiny helpers ----------
    def _rect_from_fdef(fdef):
        if fdef.get("box_rect"):
            x0, y0, x1, y1 = map(float, fdef["box_rect"])
        elif fdef.get("line_box"):
            x0, y0, x1, y1 = map(float, fdef["line_box"])
        else:
            x = float(fdef["anchor_x"]); y = float(fdef["anchor_y"])
            x0, y0, x1, y1 = x - 6, y - 6, x + 6, y + 6
        return fitz.Rect(x0, y0, x1, y1)

    def _find_any_widget_overlapping(page: fitz.Page, box: Tuple[float, float, float, float]):
        try:
            widgets = list(page.widgets() or [])
        except TypeError:
            widgets = list(page.widgets or [])
        if not widgets:
            return None
        target = fitz.Rect(*box)
        for w in widgets:
            try:
                if fitz.Rect(w.rect).intersects(target):
                    return w
            except Exception:
                pass
        return None

    def _ensure_checkbox_widget(page: fitz.Page,
                                rect: Tuple[float, float, float, float],
                                field_name: str,
                                tooltip: str = ""):
        try:
            widget_dict = {
                "type": fitz.PDF_WIDGET_TYPE_CHECKBOX,
                "rect": fitz.Rect(*rect),
                "field_name": field_name,
                "field_value": "Off",
                "tooltip": tooltip or field_name,
                "text_color": (0, 0, 0),
                "border_color": None,
                "fill_color": None,
                "readonly": False,
                "required": False,
                "rotate": 0,
            }
            w = page.add_widget(widget_dict)
            try:
                w.set_flags(fitz.ANNOT_PRINT)
            except Exception:
                pass
            return w
        except Exception:
            return None

    def _rect_key(x0, y0, x1, y1, grid: float = 2.0):
        return (int(round(x0 / grid)), int(round(y0 / grid)),
                int(round(x1 / grid)), int(round(y1 / grid)))

    def _draw_center_X(page: fitz.Page, rect: fitz.Rect):
        pad = max(0.8, 0.18 * min(rect.width, rect.height))
        width = max(1.6, 0.22 * min(rect.width, rect.height))
        p1 = fitz.Point(rect.x0 + pad, rect.y0 + pad)
        p2 = fitz.Point(rect.x1 - pad, rect.y1 - pad)
        p3 = fitz.Point(rect.x0 + pad, rect.y1 - pad)
        p4 = fitz.Point(rect.x1 - pad, rect.y0 + pad)
        page.draw_line(p1, p2, width=width, color=(0, 0, 0), overlay=True)
        page.draw_line(p3, p4, width=width, color=(0, 0, 0), overlay=True)

    # STRICT lookup just for checkboxes: match EXACTLY on label_short
    def _strict_checkbox_value(label_short: str, page_no: int, section_norm: str, occurrence_index: int):
        norm = alias_normal(_normalize(label_short))
        cands = [r for r in lookup_rows if r.get("field_norm") == norm]
        if not cands:
            return None
        # prefer section, then page
        if section_norm:
            sec = [r for r in cands if r.get("section_norm") == section_norm]
            if sec:
                cands = sec
        if page_no is not None:
            pg = [r for r in cands if r.get("Page") is not None and int(r["Page"]) == int(page_no)]
            if pg:
                cands = pg
        # if any candidate rows specify Index, require exact Index
        with_idx = [r for r in cands if r.get("Index") is not None]
        if with_idx:
            exact = [r for r in with_idx if int(r["Index"]) == int(occurrence_index)]
            if exact:
                cands = exact
            else:
                return None
        # simple tie-breaker: prefer those that match both page & section
        def _score(r):
            s = 0
            if r.get("section_norm") == section_norm and section_norm:
                s += 5
            if r.get("Page") is not None and int(r["Page"]) == int(page_no):
                s += 4
            if r.get("Index") is not None and int(r["Index"]) == int(occurrence_index):
                s += 2
            return s
        best = max(cands, key=_score, default=None)
        return best["Value"] if best else None

    # book-keeping
    ticked_regions: Dict[int, Set[Tuple[int, int, int, int]]] = {}
    occ_counters: Dict[Tuple[int, str, str], int] = {}
    used_indices: Dict[Tuple[int, str, str], Set[int]] = {}

    for fdef in tpl.get("fields", []):
        label = fdef.get("label", "") or ""
        if not label or label.startswith("unknown_"):
            continue

        page_idx = max(0, int(fdef["page"]) - 1)
        detected_section_norm = fdef.get("section_norm", "")
        placement = (fdef.get("placement") or "").lower()
        page = doc[page_idx]

        # ---- Phase 1: discover bucket (ignore Index) ----
        raw_label   = label
        label_short = fdef.get("label_short", "")
        # For CHECKBOXES: only use label_short to avoid “item 1” vs “item 2” fuzzy slips
        if placement == "checkbox" and label_short:
            search_keys = [label_short]
            fuzzy_for_this = 0.999  # effectively disables fuzzy
        else:
            search_keys: List[str] = []
            if label_short:
                search_keys.append(label_short)
            if ":" in raw_label:
                search_keys.append(raw_label.split(":", 1)[0].strip())
            search_keys.append(raw_label)
            fuzzy_for_this = (0.68 if placement == "checkbox" else min_field_fuzzy)

        picked_row = None
        picked_key = None
        for key_try in search_keys:
            _, pr = resolve_value(
                lookup_rows, key_try,
                page=page_idx + 1, section_norm=detected_section_norm,
                occurrence_index=1, min_field_fuzzy=fuzzy_for_this,
                return_row=True, strict_index=False
            )
            if pr:
                picked_row = pr
                picked_key = key_try
                break
        if not picked_row:
            if dry_run:
                kind = "checkbox" if placement == "checkbox" else "field"
                print(f"[DRY] p{page_idx+1} {kind} '{label}' → no match for {search_keys}")
            continue

        field_norm         = picked_row["field_norm"]
        excel_section_norm = picked_row.get("section_norm") or ""
        effective_section  = excel_section_norm or detected_section_norm
        bucket_key         = (page_idx + 1, effective_section, field_norm)

        # Occurrence index (explicit first)
        explicit = _explicit_indices_for(lookup_rows, field_norm, page_idx + 1, effective_section)
        if explicit:
            used = used_indices.setdefault(bucket_key, set())
            next_idx = None
            for idx_val in explicit:
                if idx_val not in used:
                    next_idx = idx_val
                    break
            if next_idx is None:
                occ_counters[bucket_key] = occ_counters.get(bucket_key, 0) + 1
                next_idx = occ_counters[bucket_key]
            used.add(next_idx); idx = next_idx
        else:
            occ_counters[bucket_key] = occ_counters.get(bucket_key, 0) + 1
            idx = occ_counters[bucket_key]

        # ---- Phase 2: get value
        if placement == "checkbox" and label_short:
            # STRICT path: exact match on label_short
            value = _strict_checkbox_value(label_short, page_idx + 1, effective_section, idx)
        else:
            # normal path
            value = resolve_value(
                lookup_rows, picked_key or label,
                page=page_idx + 1,
                section_norm=effective_section,
                occurrence_index=idx,
                min_field_fuzzy=fuzzy_for_this,
                strict_index=True,
                require_page_match=True,          # NEW: do not cross-fill from other pages
                require_section_match=True        # NEW: do not cross-fill from other sections
            )


        if value is None:
            if dry_run:
                print(f"[DRY] p{page_idx+1} '{label}' (idx={idx}) → no value")
            continue

        # =========================
        # CHECKBOX
        # =========================
        if placement == "checkbox":
            yn = _truthy(value)
            if yn is not True and yn is not False:
                if dry_run:
                    print(f"[DRY] p{page_idx+1} checkbox '{label_short or label}' -> skip (unclear '{value}')")
                continue

            r = _rect_from_fdef(fdef)
            rk = _rect_key(r.x0, r.y0, r.x1, r.y1)
            pgset = ticked_regions.setdefault(page_idx, set())
            if rk in pgset:
                # already handled this rectangle
                continue

            # If yes/no tokens are around this rectangle, ensure we only tick the *matching* one.
            rcx = (r.x0 + r.x1) / 2.0
            rcy = (r.y0 + r.y1) / 2.0
            opt_here = _nearby_yes_no_option(page, rcx, rcy)  # 'yes'/'no'/None
            if opt_here in {"yes", "no"}:
                if (yn and opt_here != "yes") or ((yn is False) and opt_here != "no"):
                    if dry_run:
                        print(f"[DRY] p{page_idx+1} checkbox '{label_short or label}' -> skip (token mismatch)")
                    continue

            if dry_run:
                print(f"[DRY] p{page_idx+1} checkbox '{label_short or label}' -> TICK @ ({rcx:.1f},{rcy:.1f})")
                pgset.add(rk)
                continue

            # prefer widget
            w = _find_any_widget_overlapping(page, (r.x0, r.y0, r.x1, r.y1))
            if w is not None:
                try:
                    if getattr(w, "field_value", "") != "Yes" and yn is True:
                        w.field_value = "Yes"; w.update()
                    elif getattr(w, "field_value", "") == "Yes" and yn is False:
                        # If needed, you could set to "Off". We leave it as-is to avoid surprises.
                        pass
                    filled += 1
                    pgset.add(rk)
                    continue
                except Exception:
                    pass

            # else draw centered X if truthy
            if yn is True:
                _draw_center_X(page, r)
                filled += 1
                pgset.add(rk)
            continue

        # =========================
        # NORMAL TEXT
        # =========================
        x = float(fdef["anchor_x"])
        y = float(fdef["anchor_y"])
        line_box = fdef.get("line_box")
        if center_on_line and line_box:
            x0, y0, x1, y1 = map(float, line_box)
            ux0 = max(x0, x); ux1 = x1
            approx_char_w = 4.8
            text_w = max(1.0, len(str(value)) * approx_char_w)
            draw_x = ux0 + max(0.0, (ux1 - ux0 - text_w)) / 2.0
            draw_y = (y0 + y1) / 2.0
        else:
            draw_x, draw_y = x, y

        if dry_run:
            print(f"[DRY] p{page_idx+1} '{label}' (idx={idx}) → '{value}' at ({draw_x:.1f},{draw_y:.1f})")
        else:
            rect = fitz.Rect(draw_x, draw_y - font_size, draw_x + 1200, draw_y + 2 * font_size)
            page.insert_textbox(rect, str(value), fontsize=font_size, align=fitz.TEXT_ALIGN_LEFT)
            filled += 1

    if dry_run:
        print("Dry-run complete (overlay). No file written.")
        doc.close()
        return filled

    try:
        doc.set_need_appearances(True)
    except Exception:
        pass

    doc.save(out_pdf)
    doc.close()
    print(f"🎉 Overlay filled PDF saved to {out_pdf} (values placed: {filled})")
    return filled


# ---------------------------
# Coordinates exporter (debug)
# ---------------------------
def export_pdf_coordinates(pdf_path: str, csv_path: str = "pdf_coordinates.csv"):
    doc = fitz.open(pdf_path)
    rows = []
    for page_num, page in enumerate(doc, start=1):
        for block in page.get_text("blocks"):
            x, y, x1, y1, text, *_ = block
            rows.append([page_num, int(x), int(y), int(x1), int(y1), (text or "").strip()])
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Page", "X0", "Y0", "X1", "Y1", "Text"])
        writer.writerows(rows)
    print(f"📄 Export complete: {csv_path}")
    doc.close()

# ---------------------------
# Orchestrator
# ---------------------------
def prefill_pdf(input_pdf: str,
                output_pdf: str,
                lookup_path: str,
                template_json: str = "template_fields.json",
                build_template_if_missing: bool = True,
                dry_run: bool = False,
                rebuild_template: bool = False):
    lookup_rows = read_lookup_rows(lookup_path)

    # 1) AcroForms
    print("🔎 Inspecting AcroForm widgets…")
    changed, widgets_exist = fill_acroform_with_context(
        input_pdf=input_pdf,
        output_pdf=output_pdf,
        lookup_rows=lookup_rows,
        dry_run=dry_run
    )

    # In DRY mode we *also* run overlay for visibility.
    if not dry_run and widgets_exist and changed:
        print("✅ AcroForm path succeeded; overlay will ALSO run to draw vector checks.")
        # no return: we still run the overlay so vector checkboxes (like p4) get drawn

    else:
        if dry_run:
            if widgets_exist and changed:
                print("ℹ️ DRY-RUN: AcroForm had matches; overlay will also run for reporting…")
            else:
                print("ℹ️ DRY-RUN: No AcroForm changes; proceeding to overlay…")
        else:
            print("ℹ️ AcroForm did not fill everything; proceeding to overlay…")

    # 2) Overlay (template)
    print("ℹ️ Proceeding with underline/vector overlay template…")
    if rebuild_template or (not os.path.exists(template_json) and build_template_if_missing):
        print("🧩 Building template…(forced)" if rebuild_template else "🧩 Building template…")
        build_pdf_template(input_pdf, template_json, lookup_rows=lookup_rows, dry_run=dry_run)
    else:
        print(f"📄 Using existing template: {template_json}")

    fill_from_template(
        pdf_path=input_pdf,
        template_json=template_json,
        lookup_rows=lookup_rows,
        out_pdf=output_pdf,
        center_on_line=True,
        font_size=10.5,
        min_field_fuzzy=0.82,
        dry_run=dry_run
    )



# ---------------------------
# CLI
# ---------------------------
def main():
    ap = argparse.ArgumentParser(description="PDF prefill (PyMuPDF) with Section / Page / Index + Yes/No + checkbox squares")
    ap.add_argument("--input", required=True, help="Input PDF path")
    ap.add_argument("--output", required=True, help="Output PDF path")
    ap.add_argument("--lookup", default="lookup_table.xlsx", help="Excel/CSV with Field,Value[,Section,Page,Index]")
    ap.add_argument("--template", default="template_fields.json", help="Template JSON (built if missing)")
    ap.add_argument("--dry-run", action="store_true", help="Print what would be filled; no write")
    ap.add_argument("--export-coords", action="store_true", help="Also export pdf_coordinates.csv (debug)")
    ap.add_argument("--rebuild-template", action="store_true", help="Force rebuild of the template JSON even if it already exists")
    args = ap.parse_args()

    if args.export_coords:
        export_pdf_coordinates(args.input)

    prefill_pdf(
        input_pdf=args.input,
        output_pdf=args.output,
        lookup_path=args.lookup,
        template_json=args.template,
        build_template_if_missing=True,
        dry_run=args.dry_run,
        rebuild_template=args.rebuild_template,  # <-- pass through
    )


if __name__ == "__main__":
    main()
