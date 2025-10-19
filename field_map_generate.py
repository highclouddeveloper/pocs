import os
import sys
import json
import argparse
import traceback
import inspect
import pandas as pd
import re  # --- NEW

# --- Inline checkbox splitting helpers (generic, pattern-based) ---
_BOX_TOKEN_RE = re.compile(r'(\[\s*[xX✓]?\s*\]|[□☐☒])')
_INLINE_SEP_RE = re.compile(r'(?:\t+|\s{2,}|\s+/\s+|\s+\|\s+|\s+OR\s+)', re.IGNORECASE)

def _split_compound_checkboxish(label: str) -> list:
    """
    Split lines like 'Initial Subscription Additional Subscription' into
    ['Initial Subscription', 'Additional Subscription'].

    Very defensive:
      - No colon in label (not a normal 'Label: ____' line)
      - Label short (<= 60 chars), words only (no digits)
      - Each chunk is 1–4 TitleCase words (or ALLCAPS), 2–4 chunks total
    Returns [] when we decide it's NOT a compound checkbox line.
    """
    t = (label or "").strip()
    if not t:
        return []
    if ":" in t:
        return []
    if len(t) > 60:
        return []
    if re.search(r"\d", t):  # avoid account numbers etc.
        return []

    words = t.split()
    if len(words) < 3:  # need at least two chunks of 1–4 words each
        return []

    def _is_titleish(w: str) -> bool:
        # Accept TitleCase (Bank, State) or ALLCAPS (USA)
        return (w[:1].isupper() and (w[1:].islower() or w[1:] == "")) or w.isupper()

    if not all(_is_titleish(w) for w in words):
        return []

    # Greedy grouping into 1–4-word chunks
    chunks = []
    cur = []
    for w in words:
        cur.append(w)
        if len(cur) >= 2 and len(cur) <= 4:
            # Heuristic: end a chunk if next word would make it too long,
            # or if we've already collected a 2+ word phrase.
            chunks.append(" ".join(cur))
            cur = []
    if cur:
        # attach any remainder to the last chunk if short, else bail
        if chunks and len(cur) <= 2:
            chunks[-1] = f"{chunks[-1]} {' '.join(cur)}"
            cur = []
        else:
            return []

    # Need 2–4 chunks, and each chunk 1–4 words
    if len(chunks) < 2 or len(chunks) > 4:
        return []

    if not all(1 <= len(c.split()) <= 4 for c in chunks):
        return []

    # Extra guard: combined chunks must be exactly the original text (ignoring single spaces)
    if " ".join(chunks) != t:
        return []

    return chunks


def _looks_short_option(s: str, max_words: int = 5) -> bool:
    return len((s or "").split()) <= max_words

def _is_title_case_phrase(s: str) -> bool:
    # Accept phrases where most words start uppercase (Title-Case-ish).
    words = (s or "").split()
    if not words:
        return False
    caps = sum(1 for w in words if w[:1].isupper())
    return caps >= max(1, int(0.6 * len(words)))

def _extract_inline_checkbox_options(text: str):
    """
    Extract inline options from a single-line label such as:
      '□ Initial Subscription    □ Additional Subscription'
      '[ ] Initial Subscription   [ ] Additional Subscription'
      'Initial Subscription    Additional Subscription' (2+ spaces)
      'Initial Subscription OR Additional Subscription'
      'Initial Subscription Additional Subscription'  <-- (title-case split)
    Returns list[str] or None.
    """
    if not text:
        return None

    # 0) Normalize a cheap test string
    t_norm = re.sub(r'\s+', ' ', text).strip()

    # 1) Split on explicit checkbox tokens (□, ☐, ☒, [ ], [x], etc.)
    t = _BOX_TOKEN_RE.sub(' ::: ', text)
    t = re.sub(r'\s+', ' ', t).strip()
    parts = [p.strip(' -:') for p in t.split(':::') if p and p.strip()]
    parts = [p for p in parts if _looks_short_option(p)]
    if len(parts) >= 2 and len(set(p.lower() for p in parts)) >= 2:
        return parts

    # 2) Fallback: split on tabs, 2+ spaces, '/', '|', or ' OR '
    parts = [p.strip(' -:') for p in _INLINE_SEP_RE.split(text) if p and p.strip()]
    parts = [p for p in parts if _looks_short_option(p)]
    if 2 <= len(parts) <= 8 and len(set(p.lower() for p in parts)) >= 2:
        return parts

    # 3) Title-case chunk heuristic (handles "Initial Subscription Additional Subscription")
    #    Only attempt when the whole thing is short, title-case-ish, with no colon.
    if ':' not in t_norm:
        words = t_norm.split()
        if 2 <= len(words) <= 8 and _is_title_case_phrase(t_norm):
            # Try equal halves first (common for 2 options of equal length)
            if len(words) % 2 == 0:
                mid = len(words) // 2
                left = " ".join(words[:mid]).strip(' -:')
                right = " ".join(words[mid:]).strip(' -:')
                if _looks_short_option(left) and _looks_short_option(right) and left.lower() != right.lower():
                    return [left, right]

            # Otherwise, look for a reasonable split point before a capitalized word
            # (e.g., "... Subscription | Additional ...")
            for i in range(2, len(words)-1):
                if words[i][0].isupper():
                    left = " ".join(words[:i]).strip(' -:')
                    right = " ".join(words[i:]).strip(' -:')
                    if (_looks_short_option(left) and _looks_short_option(right)
                            and _is_title_case_phrase(left) and _is_title_case_phrase(right)
                            and left.lower() != right.lower()):
                        return [left, right]

    return None



# Try to find build template no matter where you place this helper.
# Priority:
# 1) Local builder defined in this file (globals()).
# 2) DOCX builder from docx_prefill (if available when input is .docx).
# 3) PDF builder from pdf_prefill (fallback/default).
_build_pdf_template_external = None
_build_pdf_template_import_error = None
_build_pdf_template_import_tb = None

# Optional DOCX builder import
_docx_builder = None
try:
    from docx_prefill import build_docx_template as _docx_builder  # type: ignore
except Exception:
    _docx_builder = None

# Optional PDF builder import (legacy/default)
try:
    from pdf_prefill import build_pdf_template as _build_pdf_template_external  # type: ignore
except Exception as _e:
    _build_pdf_template_external = None
    _build_pdf_template_import_error = _e
    _build_pdf_template_import_tb = traceback.format_exc()


def _describe_builder(fn) -> str:
    try:
        mod = getattr(fn, "__module__", None)
        src = inspect.getsourcefile(fn) or inspect.getfile(fn)
        return f"{mod or '<unknown module>'} :: {src}"
    except Exception:
        return "<uninspectable builder>"


def _pick_builder_by_input(input_path: str, verbose: bool = False):
    """
    Decide which builder to call based on the input file extension.
      - *.docx -> prefer docx_prefill.build_docx_template if available
      - otherwise -> pdf_prefill.build_pdf_template
    Also allows a local `build_pdf_template` or `build_docx_template` to override via globals().
    """
    # Local overrides (if user defined directly in this file)
    local_build_pdf = globals().get("build_pdf_template")
    local_build_docx = globals().get("build_docx_template")

    ext = (os.path.splitext(input_path)[1] or "").lower()

    if ext == ".docx":
        # Prefer local DOCX builder, then imported docx_prefill version
        if callable(local_build_docx):
            if verbose:
                print(f"🔧 Using local DOCX builder: {_describe_builder(local_build_docx)}")
            return local_build_docx, "docx"
        if callable(_docx_builder):
            if verbose:
                print(f"🔧 Using docx builder: {_describe_builder(_docx_builder)}")
            return _docx_builder, "docx"
        # Fallback to any local/pdf builder if DOCX not present
        if callable(local_build_pdf):
            if verbose:
                print(f"🔧 Using local PDF builder (fallback for DOCX): {_describe_builder(local_build_pdf)}")
            return local_build_pdf, "pdf"
        if callable(_build_pdf_template_external):
            if verbose:
                print(f"🔧 Using imported PDF builder (fallback for DOCX): {_describe_builder(_build_pdf_template_external)}")
            return _build_pdf_template_external, "pdf"

    # Non-DOCX -> default to PDF path
    if callable(local_build_pdf):
        if verbose:
            print(f"🔧 Using local PDF builder: {_describe_builder(local_build_pdf)}")
        return local_build_pdf, "pdf"
    if callable(_build_pdf_template_external):
        if verbose:
            print(f"🔧 Using imported PDF builder: {_describe_builder(_build_pdf_template_external)}")
        return _build_pdf_template_external, "pdf"

    # Nothing found; show helpful error with import traceback if any
    detail = ""
    if _build_pdf_template_import_error is not None:
        sys.stderr.write("----- pdf_prefill import traceback -----\n")
        if _build_pdf_template_import_tb:
            sys.stderr.write(_build_pdf_template_import_tb + "\n")
        else:
            sys.stderr.write(repr(_build_pdf_template_import_error) + "\n")
        sys.stderr.write("----- end traceback -----\n")
        detail = (
            f"\n(Import error: {type(_build_pdf_template_import_error).__name__}: "
            f"{_build_pdf_template_import_error})"
        )
    raise RuntimeError(
        "No builder found. Define build_docx_template/build_pdf_template in this file "
        "OR ensure docx_prefill/pdf_prefill are importable." + detail
    )


def _ensure_parent_dir(path: str):
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)


def _field_title_from_fdef(fdef: dict) -> str:
    """Prefer a concise key when available, otherwise fallback to label."""
    title = (fdef.get("label_short") or fdef.get("label") or "").strip()
    # If there's a colon, keep the left side *unless* it's extremely short (keep context).
    if ":" in title:
        left = title.split(":", 1)[0].strip()
        if len(left) >= 6:
            title = left
    return title or (fdef.get("label_full") or "").strip()

# --- NEW: helpers to detect inline checkbox groups (no hard-coding of names)
_INLINE_SPLIT_RE = re.compile(r"(?:\t+|\s{2,}|(?:\s+OR\s+))", re.IGNORECASE)
_BOX_PREFIX_RE   = re.compile(r"^\s*(?:[□☐☒]\s*|\[\s?[xX✓]?\s?\]\s*)")

def _looks_short_option(s: str, max_words: int = 4) -> bool:
    return len((s or "").split()) <= max_words

def _split_inline_checkbox_options(display_text: str):
    """
    If display_text looks like a single line containing multiple short options
    separated by tabs / 2+ spaces / ' OR ', split into options.
    Returns list[str] or None if not applicable.
    """
    if not display_text or not display_text.strip():
        return None
    t = _BOX_PREFIX_RE.sub("", display_text).strip()
    parts = [p.strip() for p in _INLINE_SPLIT_RE.split(t) if p and p.strip()]
    parts = [p for p in parts if _looks_short_option(p)]
    if 2 <= len(parts) <= 6 and len(set(p.lower() for p in parts)) >= 2:
        return parts
    return None


def export_lookup_template_from_json(template_json: str,
                                     out_path: str = "lookup_template.xlsx") -> str:
    """
    Produce an Excel (or CSV if path ends with .csv) with columns:
    Section | Page | Field | Index | Value | Choices

    Notes:
    - We NEVER drop fields anymore, even if the label is 'unknown_*'.
    - Page is left blank if not provided (DOCX layouts typically don't expose page #).
    - Field will always be non-empty and unique (we synthesize a stable name when needed).
    """
    if not os.path.exists(template_json):
        raise FileNotFoundError(f"Template JSON not found: {template_json}")

    with open(template_json, "r", encoding="utf-8") as f:
        tpl = json.load(f)

    rows = []
    seq = 0  # used to make a stable unique fallback field name when needed

    for fdef in (tpl.get("fields") or []):
        seq += 1

        # Raw properties from template (if present)
        raw_label = (fdef.get("label") or "").strip()
        section = (fdef.get("section") or "").strip()

        # Page handling: keep it if it's a real positive int, else leave blank (for DOCX it’s often unknown)
        page_val = fdef.get("page", None)
        try:
            page = int(page_val) if page_val not in (None, "") else ""
            if isinstance(page, int) and page <= 0:
                page = ""  # don't render '0' which is meaningless for DOCX
        except Exception:
            page = ""

        # Index handling
        try:
            index = int(fdef.get("index", 1) or 1)
        except Exception:
            index = 1

        # The human-facing field title: prefer label_short/label; if that’s empty or looks 'unknown', synthesize.
        title_from_def = _field_title_from_fdef(fdef).strip()
        looks_unknown = (raw_label.lower().startswith("unknown_") if raw_label else True)
        if not title_from_def or title_from_def.lower().startswith("unknown_") or looks_unknown:
            # Synthesize a readable, unique name that still gives context
            sec_display = section if section else "No Section"
            title_from_def = f"Field #{seq} (Sec: {sec_display}, Idx: {index})"

        # Choices for dropdowns (if any)
        choices = ""
        if (fdef.get("placement") or "").lower() == "acro_choice":
            ch = fdef.get("choices") or []
            if isinstance(ch, list) and ch:
                choices = " | ".join(str(x) for x in ch)

        # --- NEW: inline checkbox splitting (pattern-based, no hardcoding)
        # Apply ONLY at export, so builders/fillers remain untouched.
        source_for_split = (fdef.get("label_short") or fdef.get("label") or "").strip()
        inline_opts = _split_inline_checkbox_options(title_from_def)

        if inline_opts:
            # Emit one row per option. Use Index 1..N *for this inline group* to keep order.
            for j, opt in enumerate(inline_opts, start=1):
                rows.append({
                    "Section": section,
                    "Page": page,          # blank when unknown
                    "Field": opt,
                    "Index": j,            # within the inline options
                    "Value": "",
                    "Choices": choices,
                })
                # Done for this fdef; continue to next field
                continue
        # (fallback) normal single field export (unchanged)
        # Try to split compound checkbox lines into multiple rows (e.g., "Initial Subscription Additional Subscription")
        split_items = _split_compound_checkboxish(title_from_def)

        if split_items:
            # Emit one row per detected chunk. Keep the same Section/Page.
            for i, item in enumerate(split_items, start=1):
                rows.append({
                    "Section": section,
                    "Page": page,
                    "Field": item,
                    "Index": i,        # natural 1..N across the chunks on that line
                    "Value": "",
                    "Choices": "",     # these are checkboxes; leave Choices blank
                })
        else:
            rows.append({
                "Section": section,
                "Page": page,
                "Field": title_from_def,
                "Index": index,
                "Value": "",
                "Choices": choices,
            })


    # Stable ordering for human-friendly editing
    def _sort_key(r):
        # Treat blank page as +inf so numbered pages appear first
        page_sort = (999999 if r["Page"] == "" else int(r["Page"]))
        return (page_sort, r["Section"].lower(), r["Field"].lower(), int(r["Index"]))

    rows.sort(key=_sort_key)

    # Ensure the Choices column is present even when empty
    cols = ["Section", "Page", "Field", "Index", "Value", "Choices"]
    df = pd.DataFrame(rows, columns=cols)

    # Ensure NaNs don't show up
    df = df.fillna("")

    _ensure_parent_dir(out_path)

    # Write Excel or CSV (fallback to CSV if openpyxl not installed)
    if out_path.lower().endswith(".csv"):
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:
        try:
            df.to_excel(out_path, index=False)
        except Exception as e:
            fallback = os.path.splitext(out_path)[0] + ".csv"
            df.to_csv(fallback, index=False, encoding="utf-8-sig")
            print(f"⚠️  Could not write Excel ({e}). Wrote CSV instead: {fallback}")
            return fallback

    print(f"📤 Lookup template written to {out_path} with {len(df)} rows.")
    return out_path



def _call_builder_with_compat(builder, input_path: str, template_json: str):
    """
    Call a builder with broad signature compatibility:
    - build_XXX_template(path, template_json, lookup_rows=None, dry_run=False)
    - build_XXX_template(path, template_json=..., lookup_rows=None, dry_run=False)
    - build_XXX_template(doc_or_path=path, template_json=..., ...)
    """
    _ensure_parent_dir(template_json)
    try:
        return builder(input_path, template_json, lookup_rows=None, dry_run=False)
    except TypeError:
        pass
    try:
        return builder(input_path, template_json=template_json, lookup_rows=None, dry_run=False)
    except TypeError:
        pass
    try:
        return builder(doc_or_path=input_path, template_json=template_json, lookup_rows=None, dry_run=False)
    except TypeError:
        pass
    return builder(input_path, template_json)


def _load_template_field_count(template_json: str) -> int:
    try:
        with open(template_json, "r", encoding="utf-8") as f:
            tpl = json.load(f)
        return len(tpl.get("fields", []) or [])
    except Exception:
        return -1


def export_lookup_template(input_path: str,
                           template_json: str = "template_fields.json",
                           out_path: str = "lookup_template.xlsx",
                           rebuild_template: bool = False,
                           debug_import: bool = False) -> str:
    """
    Ensures a template JSON exists (building it if needed), then exports the lookup sheet.
    Adds diagnostics so you can see which builder ran and how many fields were detected.
    """
    builder, kind = _pick_builder_by_input(input_path, verbose=True or debug_import)

    must_build = rebuild_template or not os.path.exists(template_json)
    if must_build:
        print(f"🧩 Building template → {template_json}")
        _call_builder_with_compat(builder, input_path, template_json)
        cnt = _load_template_field_count(template_json)
        print(f"🧩 Template saved to {template_json} with {cnt} fields.")
        if cnt == 0:
            print("⚠️  No fields were detected.")
            if kind == "pdf":
                print("   • If using PDF, confirm you imported the correct pdf_prefill.py.")
                print("     → Run:  python -c \"import pdf_prefill,inspect; print(pdf_prefill.__file__)\"")
                print("   • If it’s a dynamic/XFA PDF or lacks detectable lines/widgets, try flattening (Print to PDF).")
                print("   • Or run the builder in DRY mode to see detection logs.")
            else:
                print("   • For DOCX, verify the document has tables/boxes/underlines the builder can detect.")
    else:
        print(f"📄 Using existing template: {template_json} ({_load_template_field_count(template_json)} fields)")

    return export_lookup_template_from_json(template_json, out_path)


# ---------------------------
# CLI
# ---------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Export a blank lookup sheet (Section, Page, Field, Index, Value, Choices) from a template JSON"
    )
    ap.add_argument("--input", required=True, help="Input file (PDF or DOCX)")
    ap.add_argument("--template", default="template_fields.json", help="Template JSON (rebuilt if missing or --rebuild-template)")
    ap.add_argument("--make-lookup", metavar="OUT.xlsx",
                    help="Path to write the blank lookup sheet (xlsx or csv).")
    ap.add_argument("--rebuild-template", action="store_true",
                    help="Force rebuild of the template JSON even if it already exists")
    ap.add_argument("--debug-import", action="store_true",
                    help="Print extra import info for the builder used.")
    args = ap.parse_args()

    if args.make_lookup:
        export_lookup_template(
            input_path=args.input,
            template_json=args.template,
            out_path=args.make_lookup,
            rebuild_template=args.rebuild_template,
            debug_import=args.debug_import,
        )
        return

    print("Nothing to do. Pass --make-lookup OUT.xlsx (or .csv) to export a blank lookup sheet.")


if __name__ == "__main__":
    main()
