import os
import sys
import json
import argparse
import traceback
import inspect
import pandas as pd

# Try to find build_pdf_template no matter where you place this helper.
# 1) Prefer a function already defined in the same module (globals()).
# 2) Else try to import it from pdf_prefill (common case in your project).
_build_pdf_template_external = None
_build_pdf_template_import_error = None
_build_pdf_template_import_tb = None
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


def _get_builder(verbose: bool = False):
    """Return a callable build_pdf_template or raise a helpful error with root cause."""
    fn = globals().get("build_pdf_template", None)
    if callable(fn):
        if verbose:
            print(f"🔧 Using local build_pdf_template: {_describe_builder(fn)}")
        return fn
    if callable(_build_pdf_template_external):
        if verbose:
            print(f"🔧 Using imported build_pdf_template: {_describe_builder(_build_pdf_template_external)}")
        return _build_pdf_template_external

    detail = ""
    if _build_pdf_template_import_error is not None:
        # Print full traceback to stderr so the real cause (syntax/indent) is visible immediately.
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
        "build_pdf_template(...) was not found. "
        "Define it above this code OR ensure `from pdf_prefill import build_pdf_template` works."
        + detail
    )


def _ensure_parent_dir(path: str):
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)


def _field_title_from_fdef(fdef: dict) -> str:
    """Prefer a concise key when available, otherwise fallback to label."""
    title = (fdef.get("label_short") or fdef.get("label") or "").strip()
    if ":" in title:
        left = title.split(":", 1)[0].strip()
        if len(left) >= 10:
            title = left
    return title


def export_lookup_template_from_json(template_json: str,
                                     out_path: str = "lookup_template.xlsx") -> str:
    """
    Produce an Excel (or CSV if path ends with .csv) with columns:
    Section | Page | Field | Index | Value | Choices (optional)
    """
    if not os.path.exists(template_json):
        raise FileNotFoundError(f"Template JSON not found: {template_json}")

    with open(template_json, "r", encoding="utf-8") as f:
        tpl = json.load(f)

    rows = []
    for fdef in tpl.get("fields", []) or []:
        raw_label = (fdef.get("label") or "").strip()
        if not raw_label or raw_label.startswith("unknown_"):
            continue

        section = (fdef.get("section") or "").strip()
        try:
            page = int(fdef.get("page", 0) or 0)
        except Exception:
            page = 0
        try:
            index = int(fdef.get("index", 1) or 1)
        except Exception:
            index = 1

        field = _field_title_from_fdef(fdef)

        # NEW: include dropdown choices if present so you can pick one later in Excel/CSV
        choices = ""
        if (fdef.get("placement") or "").lower() == "acro_choice":
            ch = fdef.get("choices") or []
            if isinstance(ch, list) and ch:
                # join using ' | ' for readability (won’t break CSV)
                choices = " | ".join(str(x) for x in ch)

        rows.append({
            "Section": section,
            "Page": page,
            "Field": field,
            "Index": index,
            "Value": "",          # you fill this later
            "Choices": choices,   # optional; blank for non-dropdowns
        })

    # Stable ordering for human-friendly editing
    rows.sort(key=lambda r: (r["Page"], r["Section"].lower(), r["Field"].lower(), r["Index"]))

    # Ensure the Choices column is present even when empty
    cols = ["Section", "Page", "Field", "Index", "Value", "Choices"]
    df = pd.DataFrame(rows, columns=cols)

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


def _call_builder_with_compat(builder, input_pdf: str, template_json: str):
    """
    Call build_pdf_template with broad signature compatibility:
    - build_pdf_template(path, template_json, lookup_rows=None, dry_run=False)
    - build_pdf_template(path, template_json=..., lookup_rows=None, dry_run=False)
    - build_pdf_template(doc_or_path=path, template_json=..., ...)
    """
    _ensure_parent_dir(template_json)

    # Try the most common positional signature first
    try:
        return builder(input_pdf, template_json, lookup_rows=None, dry_run=False)
    except TypeError:
        pass

    # Try keyword style with template_json=
    try:
        return builder(input_pdf, template_json=template_json, lookup_rows=None, dry_run=False)
    except TypeError:
        pass

    # Try fully keyworded with doc_or_path=
    try:
        return builder(doc_or_path=input_pdf, template_json=template_json, lookup_rows=None, dry_run=False)
    except TypeError:
        pass

    # Last attempt: just pass (path, template_json)
    return builder(input_pdf, template_json)


def _load_template_field_count(template_json: str) -> int:
    try:
        with open(template_json, "r", encoding="utf-8") as f:
            tpl = json.load(f)
        return len(tpl.get("fields", []) or [])
    except Exception:
        return -1


def export_lookup_template_from_pdf(input_pdf: str,
                                    template_json: str = "template_fields.json",
                                    out_path: str = "lookup_template.xlsx",
                                    rebuild_template: bool = False,
                                    debug_import: bool = False) -> str:
    """
    Ensures a template JSON exists (building it if needed), then exports the lookup sheet.
    Adds diagnostics so you can see which builder ran and how many fields were detected.
    """
    must_build = rebuild_template or not os.path.exists(template_json)
    if must_build:
        builder = _get_builder(verbose=True or debug_import)
        print(f"🧩 Building template → {template_json}")
        _call_builder_with_compat(builder, input_pdf, template_json)
        cnt = _load_template_field_count(template_json)
        print(f"🧩 Template saved to {template_json} with {cnt} fields.")
        if cnt == 0:
            print("⚠️  No fields were detected.")
            print("   • Most common cause: Python imported a DIFFERENT 'pdf_prefill' than your edited file.")
            print("     → Run:  python -c \"import pdf_prefill,inspect; print(pdf_prefill.__file__)\"")
            print("       and confirm it points to your project’s pdf_prefill.py.")
            print("   • If it is the right file, your PDF may be dynamic/XFA or lacks detectable lines/widgets.")
            print("     → Try flattening the PDF (Print to PDF) and rebuild the template.")
            print("     → Or run the builder in DRY mode to see detection logs:")
            print("           python -c \"from pdf_prefill import build_pdf_template; "
                  "build_pdf_template(r'%s', r'%s', dry_run=True)\"" % (input_pdf, template_json))
            # We still proceed to export (will produce 0-row sheet) so the command succeeds,
            # but the diagnostics above should make the root cause obvious.
    else:
        print(f"📄 Using existing template: {template_json} "
              f"({ _load_template_field_count(template_json) } fields)")

    return export_lookup_template_from_json(template_json, out_path)


# ---------------------------
# CLI
# ---------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Export a blank lookup sheet (Section, Page, Field, Index, Value, Choices) from template_fields.json"
    )
    ap.add_argument("--input", required=True, help="Input PDF path")
    ap.add_argument("--template", default="template_fields.json", help="Template JSON (built if missing)")
    ap.add_argument("--make-lookup", metavar="OUT.xlsx",
                    help="Path to write the blank lookup sheet (xlsx or csv).")
    ap.add_argument("--rebuild-template", action="store_true",
                    help="Force rebuild of the template JSON even if it already exists")
    ap.add_argument("--debug-import", action="store_true",
                    help="Print extra import info for the builder used.")
    args = ap.parse_args()

    if args.make_lookup:
        export_lookup_template_from_pdf(
            input_pdf=args.input,
            template_json=args.template,
            out_path=args.make_lookup,
            rebuild_template=args.rebuild_template,
            debug_import=args.debug_import,
        )
        return

    print("Nothing to do. Pass --make-lookup OUT.xlsx (or .csv) to export a blank lookup sheet.")


if __name__ == "__main__":
    main()
