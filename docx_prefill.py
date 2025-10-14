# docx_prefill.py
# DOCX prefill: placeholders, checkboxes, section-aware table filling, underline patterns

import os, re, math, string, unicodedata
from typing import List, Dict, Any, Optional, Tuple, Iterable, Set
from difflib import SequenceMatcher
from numbers import Number
import pandas as pd

# ---------------------------
# Config / globals
# ---------------------------
PUNCT = str.maketrans("", "", string.punctuation)

SUBSCRIPTION_DEFAULTS = {
    # "Name of Investor": "John Doe",
}

FIELD_ALIASES = {
    "investor name": "name of investor",
    "name of investor printed or typed": "name of investor",
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

def _truthy(val: str) -> Optional[bool]:
    s = (str(val or "")).strip().lower()
    if s in {"y","yes","true","1","x","✓","check","checked"}:
        return True
    if s in {"n","no","false","0","uncheck","unchecked"}:
        return False
    return None

# ---------------------------
# Lookup loader (Excel/CSV)
# ---------------------------
def to_text_value(v) -> str:
    try:
        import pandas as _pd
        if _pd.isna(v):
            return ""
    except Exception:
        pass
    if v is None:
        return ""
    if isinstance(v, Number):
        try:
            if isinstance(v, float) and not math.isfinite(v):
                return ""
        except Exception:
            pass
        try:
            if float(v).is_integer():
                return str(int(v))
        except Exception:
            pass
        s = f"{float(v):.12f}".rstrip("0").rstrip(".")
        return s or "0"
    s = str(v).strip()
    if not s:
        return ""
    m = re.fullmatch(r"([+-]?\d+)\.0+\b", s)
    if m:
        return m.group(1)
    m = re.fullmatch(r"([+-]?\d+\.\d*?[1-9])0+\b", s)
    if m:
        return m.group(1)
    return s

def read_lookup_rows(path: str) -> List[Dict[str, Any]]:
    if not os.path.exists(path):
        print(f"⚠️  Lookup file not found: {path}")
        return []
    try:
        if path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path)
    except Exception as e:
        print(f"⚠️  Could not load {path}: {e}")
        return []
    if {"Field", "Value"} - set(df.columns):
        raise ValueError("Lookup must have columns: Field, Value. Optional: Section, Page, Index")
    for col in ["Field", "Value", "Section"]:
        if col in df.columns:
            df[col] = df[col].astype(str).fillna("").map(lambda x: x.strip())
    if "Page" in df.columns:
        df["Page"] = pd.to_numeric(df["Page"], errors="coerce").astype("Int64")
    if "Index" in df.columns:
        idx_series = (
            df["Index"].astype(str).str.replace(r"[^\d\-]+", "", regex=True).replace({"": None})
        )
        df["Index"] = pd.to_numeric(idx_series, errors="coerce").astype("Int64")

    df = df[(df["Field"].astype(str).str.strip() != "") &
            (df["Value"].astype(str).str.strip() != "") &
            (df["Value"].astype(str).str.lower() != "nan")]

    rows: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        field = str(r.get("Field", "")).strip()
        value = to_text_value(r.get("Value", ""))
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

# ---------------------------
# Resolver (Field/Page/Section/Index)
# ---------------------------
def resolve_value(rows: List[Dict[str, Any]],
                  field_label: str,
                  page: Optional[int],
                  section_norm: str,
                  occurrence_index: int,
                  min_field_fuzzy: float = 0.82,
                  return_row: bool = False,
                  strict_index: bool = True,
                  require_page_match: bool = False,
                  require_section_match: bool = False):

    field_norm = alias_normal(_normalize(field_label))
    field_candidates = [r for r in rows if r["field_norm"] == field_norm]
    if not field_candidates:
        all_fields = [r["field_norm"] for r in rows]
        bm, sc = _best_match_scored(field_norm, all_fields)
        if bm and sc >= min_field_fuzzy:
            field_candidates = [r for r in rows if r["field_norm"] == bm]
        else:
            return (None, None) if return_row else None

    candidates = list(field_candidates)

    if section_norm:
        sec_exact = [r for r in candidates if r.get("section_norm") == section_norm]
        if require_section_match:
            candidates = sec_exact
            if not candidates:
                return (None, None) if return_row else None
        elif sec_exact:
            candidates = sec_exact

    if page is not None:
        pg_exact = [r for r in candidates if r.get("Page") is not None and int(r["Page"]) == int(page)]
        if require_page_match:
            candidates = pg_exact
            if not candidates:
                return (None, None) if return_row else None
        elif pg_exact:
            candidates = pg_exact

    if strict_index:
        with_idx = [r for r in candidates if r.get("Index") is not None]
        if with_idx:
            exact_idx = [r for r in with_idx if int(r["Index"]) == int(occurrence_index)]
            if not exact_idx:
                return (None, None) if return_row else None
            candidates = exact_idx

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
# DOCX helpers & filling
# ---------------------------
def _iter_all_paragraphs_and_cells(doc) -> Iterable:
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

def _field_value_for_name(lookup_rows, name: str):
    v = resolve_value(lookup_rows, name, page=None, section_norm="", occurrence_index=1,
                      min_field_fuzzy=0.82, return_row=False, strict_index=False,
                      require_page_match=False, require_section_match=False)
    return v

def _replace_placeholders_in_text(text: str, lookup_rows) -> str:
    pats = [
        re.compile(r"\{\{\s*(.*?)\s*\}\}"),
        re.compile(r"\[\[\s*(.*?)\s*\]\]"),
        re.compile(r"\$\{\s*(.*?)\s*\}"),
    ]
    def _sub(m):
        key = (m.group(1) or "").strip()
        val = _field_value_for_name(lookup_rows, key)
        return str(val) if val is not None else m.group(0)
    for pat in pats:
        text = pat.sub(_sub, text)
    return text

def _set_paragraph_text_keep_simple_format(p, new_text: str):
    while p.runs:
        r = p.runs[0]
        r.clear()
        r.text = ""
        r.element.getparent().remove(r.element)
    p.add_run(new_text)

def _replace_placeholders_everywhere(doc, lookup_rows, dry_run=False):
    for p in _iter_all_paragraphs_and_cells(doc):
        old = p.text
        new = _replace_placeholders_in_text(old, lookup_rows)
        if new != old:
            if dry_run:
                print(f"[DRY][DOCX] placeholder: '{old}' → '{new}'")
            else:
                _set_paragraph_text_keep_simple_format(p, new)

def _fill_colon_underscore_lines(doc, lookup_rows, dry_run=False):
    us_pat = re.compile(r"^(.*?[:：﹕꞉˸፡︓])\s*_+\s*$")
    for p in _iter_all_paragraphs_and_cells(doc):
        txt = p.text.strip()
        if not txt:
            continue
        m = us_pat.match(txt)
        if not m:
            continue
        label_full = m.group(1)
        lab = re.sub(r"[:：﹕꞉˸፡︓]\s*$", "", label_full).strip()
        if not lab:
            continue
        val = _field_value_for_name(lookup_rows, lab)
        if val is None:
            continue
        new_text = f"{label_full} {val}"
        if dry_run:
            print(f"[DRY][DOCX] underline: '{lab}' → '{val}'")
        else:
            _set_paragraph_text_keep_simple_format(p, new_text)

def _fill_checkboxes(doc, lookup_rows, dry_run=False):
    box_patterns = [
        re.compile(r"^\s*[□☐]\s+(.*)$"),
        re.compile(r"^\s*\[\s?\]\s+(.*)$"),
        re.compile(r"^\s*\[\s?[xX✓]\s?\]\s+(.*)$"),
    ]
    for p in _iter_all_paragraphs_and_cells(doc):
        t = p.text.strip()
        if not t:
            continue
        for pat in box_patterns:
            m = pat.match(t)
            if not m:
                continue
            lab = m.group(1).strip()
            if not lab:
                break
            val = _field_value_for_name(lookup_rows, lab)
            yn = _truthy(val)
            if yn is None:
                continue
            new = (f"☒ {lab}" if yn else f"☐ {lab}")
            if dry_run:
                print(f"[DRY][DOCX] checkbox: '{lab}' → {'CHECK' if yn else 'UNCHECK'}")
            else:
                _set_paragraph_text_keep_simple_format(p, new)
            break

# --- DOCX table helpers (section-aware, write “inside the box”) ---
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

def _iter_block_items(parent):
    parent_elm = parent._element
    for child in parent_elm.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def _first_cell_text(cell: _Cell) -> str:
    return " ".join(p.text for p in cell.paragraphs).strip()

def _strip_trailing_paren(s: str) -> str:
    return re.sub(r"\s*\([^()]*\)\s*$", "", s or "").strip()

def _looks_like_section_title(text: str) -> bool:
    if not text:
        return False
    t = unicodedata.normalize("NFKC", text).strip()
    if len(t) > 180:
        return False
    letters = [c for c in t if c.isalpha()]
    caps_ratio = (sum(1 for c in letters if c.isupper()) / max(1, len(letters))) if letters else 0.0
    return t.endswith(":") or (caps_ratio >= 0.45 and len(t.split()) <= 15)

def _norm_key_for_match(s: str) -> str:
    t = unicodedata.normalize("NFKC", str(s or "")).strip().lower()
    t = re.sub(r"\s+", " ", t)
    return t.translate(str.maketrans("", "", string.punctuation))

def _looks_like_placeholder_only(text: str) -> bool:
    # True when a cell has only lines/dashes/whitespace (no alphanumerics)
    if not text:
        return True
    vis = re.sub(r"\s+", "", text)
    return bool(vis) and not re.search(r"[A-Za-z0-9]", vis)

def _nearest_section_for_table(doc, tbl, known_sections_norm: Set[str]) -> str:
    # Walk the document blocks to find tbl, then scan upward for the closest paragraph
    # that either looks like a section title or matches a known section.
    last_para_text = ""
    for blk in _iter_block_items(doc):
        if blk is tbl:
            break
        if hasattr(blk, "text"):
            last_para_text = blk.text or last_para_text
    # Walk upward again from the table element to previous siblings
    # (best effort; python-docx doesn't expose easy prev-sibling iteration)
    # Use the already captured last_para_text as a pragmatic nearest paragraph.
    cand = (last_para_text or "").strip()
    cand_norm = _norm_key_for_match(cand)
    if cand_norm in known_sections_norm:
        return cand_norm
    if _looks_like_section_title(cand) and len(cand_norm) >= 4:
        return cand_norm
    return ""  # unknown


def _set_cell_text(cell: _Cell, text: str):
    for p in list(cell.paragraphs):
        p.clear()
    cell.text = str(text)

def _write_value_inside_box(cell: _Cell, value: str):
    """
    Prefer writing into the result text of legacy FORMTEXT fields so the grey box stays.
    Fallbacks cover placeholder lines, shaded paragraphs, etc.
    """
    import re
    from docx.oxml.shared import OxmlElement, qn

    val = value or ""

    PLACEHOLDER_RE = re.compile(r"[_\-\u2014\.\s\u2002\u2003\u2007\u2009\u00A0]+")
    ALNUM_RE = re.compile(r"[A-Za-z0-9]")

    def _para_is_shaded(para) -> bool:
        try:
            return bool(para._element.xpath('.//w:pPr/w:shd'))
        except Exception:
            return False

    def _cell_is_shaded(c) -> bool:
        try:
            return bool(c._tc.xpath('.//w:tcPr/w:shd'))
        except Exception:
            return False

    def _set_in_para(para, text: str):
        # add or reuse an empty run without clearing structure
        empties = [r for r in para.runs if not (r.text or "").strip()]
        if empties:
            empties[-1].text = text
        else:
            para.add_run(text)

    # ---- 1) Handle legacy FORMTEXT: write into result between 'separate' and 'end' ----
    for para in cell.paragraphs:
        p = para._element
        instr_nodes = p.xpath(".//w:instrText")
        has_formtext = any(("FORMTEXT" in (n.text or "").upper()) for n in instr_nodes)
        if not has_formtext:
            continue

        result_runs = []
        got_separate = False
        for r in p.xpath("./w:r"):
            if r.xpath("./w:fldChar[@w:fldCharType='begin']"):
                got_separate = False
                result_runs = []
                continue
            if r.xpath("./w:fldChar[@w:fldCharType='separate']"):
                got_separate = True
                continue
            if r.xpath("./w:fldChar[@w:fldCharType='end']"):
                break
            if got_separate:
                result_runs.append(r)

        if result_runs:
            # clear any existing w:t nodes in the result and set a fresh one on the first run
            first = result_runs[0]
            # remove existing text nodes
            for rr in result_runs:
                for t in rr.xpath("./w:t"):
                    t.getparent().remove(t)
            # add a fresh text node (preserve spaces)
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = val
            first.append(t)
            return  # done for this cell

    # ---- 2) Fallbacks (for non-field boxes/placeholders/shading) ----

    def glen(group):
        return sum(len((r.text or "").replace(" ", "")) for r in group)

    best_group = []
    shaded_paras = []
    box_like_paras = []

    for para in cell.paragraphs:
        if _para_is_shaded(para):
            shaded_paras.append(para)

        vis = "".join((r.text or "") for r in para.runs)
        if vis and not ALNUM_RE.search(vis):
            box_like_paras.append(para)

        cur, best_here = [], []
        for run in para.runs:
            t = run.text or ""
            if t and PLACEHOLDER_RE.fullmatch(t):
                cur.append(run)
            else:
                if glen(cur) > glen(best_here):
                    best_here = cur[:]
                cur = []
        if glen(cur) > glen(best_here):
            best_here = cur[:]
        if glen(best_here) > glen(best_group):
            best_group = best_here

    if best_group:
        L = max(1, glen(best_group))
        best_group[0].text = val[:L]
        for r in best_group[1:]:
            r.text = ""
        return

    if shaded_paras:
        _set_in_para(shaded_paras[0], val)
        return

    if _cell_is_shaded(cell):
        p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        _set_in_para(p, val)
        return

    if box_like_paras:
        _set_in_para(box_like_paras[0], val)
        return

    p = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
    _set_in_para(p, val)



def _cell_has_box_target(cell: _Cell) -> bool:
    try:
        if cell._tc.xpath('.//w:tcPr/w:shd', namespaces=cell._tc.nsmap):
            return True
    except Exception:
        pass
    for p in cell.paragraphs:
        for r in p.runs:
            if (r.text or "").strip().upper() == "FORMTEXT":
                return True
        for r in p.runs:
            t = r.text or ""
            if t and re.fullmatch(r"[_\-\u2014\.\s\u2002\u2003\u2007\u2009\u00A0]+", t):
                return True
        vis = "".join((r.text or "") for r in p.runs)
        if vis and not re.search(r"[A-Za-z0-9]", vis):
            return True
    return False

def _value_cell_candidates(tbl: Table, ri: int, ci: int) -> List[_Cell]:
    cands = []
    ncols = len(tbl.rows[ri].cells)
    if ci + 1 < ncols:
        cands.append(tbl.rows[ri].cells[ci + 1])
    if ri + 1 < len(tbl.rows):
        c2 = tbl.rows[ri + 1].cells[ci]
        if _cell_has_box_target(c2):
            cands.append(c2)
    if (ri + 1 < len(tbl.rows)) and (ci + 1 < ncols):
        c3 = tbl.rows[ri + 1].cells[ci + 1]
        if _cell_has_box_target(c3):
            cands.append(c3)
    return sorted(cands, key=lambda c: (0 if _cell_has_box_target(c) else 1))

def _guess_pair_headers(tbl: Table) -> Dict[int, str]:
    hdr = {}
    try:
        row0 = tbl.rows[0]
        for ci in range(0, len(row0.cells), 2):
            hdr[ci] = _first_cell_text(row0.cells[ci]).strip()
    except Exception:
        pass
    return hdr

def _scan_up_label(tbl: Table, ri: int, ci: int) -> str:
    try:
        hdr = _first_cell_text(tbl.rows[0].cells[ci]).strip()
        if hdr:
            return hdr
    except Exception:
        pass
    for rj in range(ri - 1, -1, -1):
        try:
            t = _first_cell_text(tbl.rows[rj].cells[ci]).strip()
        except Exception:
            t = ""
        if t:
            return t
    return ""

def _fill_docx_tables_section_aware(doc, lookup_rows, dry_run=False):
    from collections import defaultdict
    import re as _re

    def _norm_key(s: str) -> str:
        t = unicodedata.normalize("NFKC", str(s or "")).strip().lower()
        t = re.sub(r"\s+", " ", t)
        return t.translate(str.maketrans("", "", string.punctuation))

    # Index lookup values three ways
    values_exact = {}
    values_by_field = defaultdict(dict)         # fld -> {idx: val} (first-seen per idx)
    values_sectionless = defaultdict(dict)      # fld (no section) -> {idx: val}
    known_sections_norm: Set[str] = set()
    known_fields_norm: Set[str] = set()

    def _to_text_value(v):
        try:
            if isinstance(v, float):
                if v.is_integer():
                    return str(int(v))
                s = f"{v:.6f}".rstrip("0").rstrip(".")
                return s if s else "0"
        except Exception:
            pass
        return str(v).strip()

    for r in lookup_rows:
        sec = _norm_key(r.get("Section", ""))
        fld = _norm_key(r.get("Field", ""))
        try:
            idx = int(r.get("Index") or 1)
        except Exception:
            idx = 1
        val = _to_text_value(r.get("Value", ""))
        if not fld:
            continue
        if sec:
            known_sections_norm.add(sec)
        known_fields_norm.add(fld)

        if val != "":
            values_exact[(sec, fld, idx)] = val
            if not sec:
                values_sectionless[fld][idx] = val
            if idx not in values_by_field[fld]:
                values_by_field[fld][idx] = val

    # Helper: try to extract a field label from row text when the label cell is blank/placeholder
    def _extract_row_label_fallback(label_cell_text: str, value_cell_text: str) -> str:
        # Search for any known field token inside either cell's visible text.
        blob = f"{label_cell_text} {value_cell_text}".strip()
        blob_norm = _norm_key(blob)
        if not blob_norm:
            return ""
        # look for best match by longest field name contained in the blob
        best = ""
        best_len = 0
        for fld in known_fields_norm:
            if fld and fld in blob_norm and len(fld) > best_len:
                best, best_len = fld, len(fld)
        return best

    # Precompute the nearest section for each table
    table_to_section_norm = {}
    for tbl in doc.tables:
        table_to_section_norm[tbl] = _nearest_section_for_table(doc, tbl, known_sections_norm)

    for tbl in doc.tables:
        try:
            ncols = len(tbl.rows[0].cells)
        except Exception:
            continue
        if ncols < 2:
            continue

        # If row 0 has headers in the label column, remember them (kept from your original heuristic)
        pair_headers = {}
        try:
            row0 = tbl.rows[0]
            for ci in range(0, len(row0.cells), 2):
                pair_headers[ci] = _first_cell_text(row0.cells[ci]).strip()
        except Exception:
            pass

        header_has_any = any(_first_cell_text(tbl.rows[0].cells[c]).strip()
                             for c in range(min(ncols, 2)))
        start_row = 1 if header_has_any else 0

        # Section guess for this table
        sec_norm_guess = table_to_section_norm.get(tbl, "") or ""

        for ri in range(start_row, len(tbl.rows)):
            row = tbl.rows[ri]
            for ci in range(0, ncols, 2):
                try:
                    label_cell = row.cells[ci]
                except Exception:
                    continue

                label_text_raw = _first_cell_text(label_cell).strip()
                label_text = label_text_raw
                # If the label cell is effectively placeholder-only, try to infer the label from row text.
                if not label_text or _looks_like_placeholder_only(label_text):
                    # peek the adjacent value cell text for clues
                    val_cell_text = ""
                    if ci + 1 < ncols:
                        try:
                            val_cell_text = _first_cell_text(row.cells[ci + 1]).strip()
                        except Exception:
                            val_cell_text = ""
                    inferred = _extract_row_label_fallback(label_text, val_cell_text)
                    if inferred:
                        fld_norm = inferred
                    else:
                        # last resort: scan upward in same column (original behavior)
                        label_text = _scan_up_label(tbl, ri, ci).strip()
                        if not label_text:
                            continue
                        fld_norm = _norm_key(label_text)
                else:
                    fld_norm = _norm_key(label_text)

                if not fld_norm:
                    continue

                # Prefer the nearest section guess for this table;
                # if not found, keep your original header-based hint as a weak signal
                sec_display = (pair_headers.get(ci, "") or "").strip()
                sec_norm = sec_norm_guess or _norm_key(_strip_trailing_paren(sec_display))

                # Index hint from header numbers like "... (1)" etc.
                pair_index_hint = (ci // 2) + 1
                m = _re.search(r'(\d+)\b', sec_display)
                if m:
                    try:
                        pair_index_hint = int(m.group(1))
                    except Exception:
                        pass

                # Selection: strict section match first, then sectionless, then any-by-field (same as your order)
                chosen = ""
                for idx_try in (pair_index_hint, 1, 2, 3, 4, 5):
                    if sec_norm:
                        chosen = values_exact.get((sec_norm, fld_norm, idx_try), "")
                        if chosen:
                            break
                if not chosen:
                    for idx_try in (pair_index_hint, 1, 2, 3, 4, 5):
                        chosen = values_sectionless.get(fld_norm, {}).get(idx_try, "")
                        if chosen:
                            break
                if not chosen:
                    for idx_try in (pair_index_hint, 1, 2, 3, 4, 5):
                        chosen = values_by_field.get(fld_norm, {}).get(idx_try, "")
                        if chosen:
                            break

                if not chosen:
                    # no value, skip writing
                    continue

                vcands = _value_cell_candidates(tbl, ri, ci)
                if not vcands:
                    continue
                target_cell = vcands[0]
                if dry_run:
                    print(f"[DRY][DOCX] (sec='{sec_norm or 'No Section'}') '{label_text or fld_norm}' → '{chosen}' (row {ri}, col {ci})")
                else:
                    _write_value_inside_box(target_cell, chosen)



def _fill_table_right_cells_simple(doc, lookup_rows, dry_run=False):
    import re
    for tbl in doc.tables:
        for row in tbl.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            lab_raw = _first_cell_text(cells[0])
            if not lab_raw:
                continue
            lab = re.sub(r"[:：﹕꞉˸፡︓]\s*$", "", lab_raw).strip()
            if not lab or len(lab) > 120:
                continue
            val = _field_value_for_name(lookup_rows, lab)
            if val is None:
                continue
            right_cell = cells[1]
            if dry_run:
                print(f"[DRY][DOCX] table2col: '{lab}' → '{val}'")
            else:
                _write_value_inside_box(right_cell, str(val))

# --- add to docx_prefill.py ---
# ---------------------------
# DOCX → template builder
# ---------------------------
# --- Template builder for DOCX (backward-compatible) --------------------------
def build_docs_template(doc_or_path,
                        template_json: str,
                        lookup_rows=None,
                        dry_run: bool = False) -> dict:
    """
    Build a template JSON with fields discovered from a DOCX.

    Backward-compatible promises:
      • Keep 'page' at 0
      • Keep placement keys: "table_pair" and "para_underline"
      • Keep labels as extracted (no aggressive renaming)
      • If no clear section is found, leave section=""
      • Indexing is stable, now done per (section, label_short)

    Enhancements:
      • Detect section headers from table headers and "Section X:" style paragraphs
      • Assign section to fields that previously had section=""
      • Avoid cross-section duplicates by indexing within each section
    """
    try:
        from docx import Document
        from docx.table import Table, _Cell
        from docx.text.paragraph import Paragraph
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl
    except Exception as e:
        raise RuntimeError("python-docx is required. Install with: pip install python-docx") from e

    # -------------------- small local helpers (match existing semantics) -------
    def _iter_block_items(parent):
        parent_elm = parent._element
        for child in parent_elm.body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def _first_cell_text(cell: _Cell) -> str:
        return " ".join(p.text for p in cell.paragraphs).strip()

    def _strip_trailing_paren(s: str) -> str:
        return re.sub(r"\s*\([^()]*\)\s*$", "", s or "").strip()

    def _looks_like_section_title(text: str) -> bool:
        if not text:
            return False
        t = unicodedata.normalize("NFKC", text).strip()
        if len(t) > 180:
            return False
        # Accept explicit and common header styles
        if t.lower().startswith("section ") or t.endswith(":"):
            return True
        letters = [c for c in t if c.isalpha()]
        if not letters:
            return False
        caps_ratio = sum(1 for c in letters if c.isupper()) / max(1, len(letters))
        return caps_ratio >= 0.45 and len(t.split()) <= 15

    def _is_table_header_row(tbl: Table, ri: int) -> bool:
        """Header if first row looks merged or shaded with concise text."""
        try:
            row = tbl.rows[ri]
        except Exception:
            return False
        merged_like = len(row.cells) == 1
        shaded_like = False
        try:
            for c in row.cells:
                if c._tc.xpath('.//w:tcPr/w:shd', namespaces=c._tc.nsmap):
                    shaded_like = True
                    break
        except Exception:
            pass
        row_text = " | ".join(_first_cell_text(c) for c in row.cells).strip()
        simple = (":" not in row_text) and (len(row_text) <= 160)
        return merged_like or (shaded_like and simple)

    def _table_section_title(tbl: Table) -> str:
        try:
            if _is_table_header_row(tbl, 0):
                return _strip_trailing_paren(_first_cell_text(tbl.rows[0].cells[0]))
        except Exception:
            pass
        return ""

    def _label_cols_for_table(tbl: Table) -> list:
        """Conservative: even columns tend to be labels in label/value grids."""
        try:
            ncols = len(tbl.rows[0].cells)
        except Exception:
            return [0]
        if ncols <= 1:
            return [0]
        return [c for c in range(0, ncols, 2)]

    underline_pat = re.compile(r"^(.*?[:：﹕꞉˸፡︓])\s*[_\u2014\-.\s]{3,}$")

    # -------------------- open document ----------------------------------------
    if isinstance(doc_or_path, Document):
        doc = doc_or_path
    else:
        doc = Document(doc_or_path)

    page = 0  # unchanged; python-docx has no pagination
    current_section = ""  # default section remains empty, like before

    raw_fields = []  # collect first, then re-index by (section, label_short)

    def _push(label: str, section: str, placement: str):
        lab = (label or "").strip()
        if not lab:
            return
        sec = (section or "").strip()  # may be ""
        lab_short = lab.split("\n", 1)[0].strip()
        raw_fields.append({
            "label": lab,
            "label_short": lab_short,
            "section": sec,   # may be "", preserving old behavior
            "page": page,     # keep 0
            "placement": placement,
        })

    # -------------------- scan in natural order --------------------------------
    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            txt = (block.text or "").strip()
            if not txt:
                continue
            # Section titles (non-destructive: only updates current_section)
            if _looks_like_section_title(txt):
                current_section = _strip_trailing_paren(txt)
                continue

            # Paragraph underline fields (same as previous behavior)
            m = underline_pat.match(txt)
            if m:
                label = re.sub(r"[:：﹕꞉˸፡︓]\s*$", "", m.group(1)).strip()
                if label:
                    _push(label, current_section, "para_underline")
                continue

        if isinstance(block, Table):
            # Prefer a header title if present; otherwise keep current_section
            table_sec = _table_section_title(block) or current_section
            if _table_section_title(block):
                current_section = table_sec  # advance section only on strong header

            start_row = 1 if _is_table_header_row(block, 0) else 0
            label_cols = _label_cols_for_table(block)

            for ri in range(start_row, len(block.rows)):
                row = block.rows[ri]
                for ci in label_cols:
                    try:
                        label_cell = row.cells[ci]
                    except Exception:
                        continue
                    label = _first_cell_text(label_cell).strip()
                    if not label:
                        continue
                    label_clean = _strip_trailing_paren(re.sub(r"[:：﹕꞉˸፡︓]\s*$", "", label).strip())
                    if not label_clean:
                        continue
                    _push(label_clean, table_sec, "table_pair")

    # -------------------- stable per-section indexing (backward compatible) ----
    # Previous behavior: indexing across whole doc; we keep that when section==""
    # New: when sections are present, we index within each section to avoid clashes.
    fields = []
    counters: Dict[Tuple[str, str], int] = {}
    for f in raw_fields:
        key = (f["section"].lower(), f["label_short"].lower())
        counters[key] = counters.get(key, 0) + 1
        fields.append({
            **f,
            "index": counters[key],  # starts at 1, stable in document order
        })

    tpl = {"fields": fields}

    if dry_run:
        print(f"[DRY] Detected {len(fields)} fields")
        for f in fields:
            sec = f["section"] or "—"
            print(f"  • {f['label_short']}  | Sec: {sec}  | Idx: {f['index']}  | {f['placement']}")
        return tpl

    with open(template_json, "w", encoding="utf-8") as f:
        json.dump(tpl, f, indent=2, ensure_ascii=False)
    print(f"🧩 Template written → {template_json}  (fields: {len(fields)})")
    return tpl


# Maintain compatibility with callers that expect build_pdf_template(...)
def build_pdf_template(doc_or_path, template_json: str, lookup_rows=None, dry_run: bool = False):
    return build_docs_template(doc_or_path, template_json, lookup_rows=lookup_rows, dry_run=dry_run)





# ---------------------------
# Public entry point
# ---------------------------
def prefill_docx(input_docx: str,
                 output_docx: str,
                 lookup_rows: List[Dict[str, Any]],
                 dry_run: bool = False):
    try:
        from docx import Document
    except Exception as e:
        raise RuntimeError("python-docx is required. Install with: pip install python-docx") from e

    # Canonicalize lookup rows
    def _norm(s: str) -> str:
        s = unicodedata.normalize("NFKC", str(s or ""))
        s = re.sub(r"\s+", " ", s).strip().lower()
        return s.translate(str.maketrans("", "", string.punctuation))

    def _val_to_text(v):
        if isinstance(v, float):
            if v.is_integer():
                return str(int(v))
            s = f"{v:.6f}".rstrip("0").rstrip(".")
            return s if s else "0"
        return str(v).strip()

    rows_in = len(lookup_rows or [])
    seen, conflicts, rows_use = {}, {}, []
    for r in (lookup_rows or []):
        sec_raw = r.get("Section") or r.get("section") or r.get("section_norm") or ""
        fld_raw = r.get("Field")   or r.get("label")   or r.get("label_norm")   or ""
        idx_raw = r.get("Index")   or r.get("index")   or 1
        try:
            idx = int(idx_raw)
        except Exception:
            idx = 1
        sec_norm = _norm(sec_raw)
        fld_norm = _norm(fld_raw)
        if not fld_norm:
            continue
        val = _val_to_text(r.get("Value", ""))
        key = (sec_norm, fld_norm, idx)
        if key in seen:
            prior = seen[key]
            if val and prior and val != prior:
                conflicts.setdefault(key, set()).update([prior, val])
            continue
        seen[key] = val
        rows_use.append({
            "Section": sec_raw,
            "section_norm": sec_norm,
            "Field":   fld_raw,
            "field_norm": fld_norm,
            "Index": idx,
            "Value": val,
        })
    lookup_rows = rows_use
    print(f"📋 DOCX prefill: rows in -> {rows_in}, kept -> {len(rows_use)}, deduped -> {rows_in - len(rows_use)}")
    if conflicts:
        print(f"   • {len(conflicts)} conflicting key(s); first value kept.")

    doc = Document(input_docx)

    _replace_placeholders_everywhere(doc, lookup_rows, dry_run=dry_run)
    _fill_checkboxes(doc, lookup_rows, dry_run=dry_run)
    _fill_docx_tables_section_aware(doc, lookup_rows, dry_run=dry_run)
    _fill_table_right_cells_simple(doc, lookup_rows, dry_run=dry_run)
    _fill_colon_underscore_lines(doc, lookup_rows, dry_run=dry_run)

    if dry_run:
        print("Dry-run complete (DOCX). No file written.")
        return 0

    doc.save(output_docx)
    print(f"📝 DOCX write complete → {output_docx}")
    return 1

# ---------------------------
# CLI
# ---------------------------
if __name__ == "__main__":
    import argparse, sys
    ap = argparse.ArgumentParser(description="Prefill DOCX from lookup data")
    ap.add_argument("--input", required=True, help="Input .docx file")
    ap.add_argument("--output", required=True, help="Output .docx file")
    ap.add_argument("--lookup", required=True, help="CSV/XLSX with Field,Value[,Section,Page,Index]")
    ap.add_argument("--dry-run", action="store_true", help="Print actions without writing")
    args = ap.parse_args()

    rows = read_lookup_rows(args.lookup)
    ok = prefill_docx(args.input, args.output, rows, dry_run=args.dry_run)
    sys.exit(0 if ok else 1)
