import os
import json
import argparse
import pandas as pd

# Try to find build_pdf_template no matter where you place this helper.
# 1) Prefer a function already defined in the same module (globals()).
# 2) Else try to import it from pdf_prefill (common case in your project).
try:
    from pdf_prefill import build_pdf_template as _build_pdf_template_external  # type: ignore
except Exception:
    _build_pdf_template_external = None


def _get_builder():
    """Return a callable build_pdf_template or raise a helpful error."""
    fn = globals().get("build_pdf_template", None)
    if callable(fn):
        return fn
    if callable(_build_pdf_template_external):
        return _build_pdf_template_external
    raise RuntimeError(
        "build_pdf_template(...) was not found. "
        "Define it above this code OR ensure `from pdf_prefill import build_pdf_template` works."
    )


def _field_title_from_fdef(fdef: dict) -> str:
    """Prefer a concise key when available, otherwise fallback to label."""
    # For checkbox bullets we created `label_short` like 'For All Subscribers – item 1'
    title = (fdef.get("label_short") or fdef.get("label") or "").strip()

    # If extremely long, keep the text before the first colon (common pattern).
    if ":" in title:
        left = title.split(":", 1)[0].strip()
        # Only shorten when the left side is meaningful
        if len(left) >= 10:
            title = left

    return title


def export_lookup_template_from_json(template_json: str,
                                     out_path: str = "lookup_template.xlsx") -> str:
    """
    Produce an Excel (or CSV if path ends with .csv) with columns:
    Section | Page | Field | Index | Value

    Returns the output path.
    """
    if not os.path.exists(template_json):
        raise FileNotFoundError(f"Template JSON not found: {template_json}")

    with open(template_json, "r", encoding="utf-8") as f:
        tpl = json.load(f)

    rows = []
    for fdef in tpl.get("fields", []):
        # Skip junk placeholders
        raw_label = (fdef.get("label") or "").strip()
        if not raw_label or raw_label.startswith("unknown_"):
            continue

        section = (fdef.get("section") or "").strip()
        page    = int(fdef.get("page", 0) or 0)
        index   = int(fdef.get("index", 1) or 1)
        field   = _field_title_from_fdef(fdef)

        rows.append({
            "Section": section,
            "Page": page,
            "Field": field,
            "Index": index,
            "Value": "",  # you fill this in later
        })

    # Stable ordering for human-friendly editing
    rows.sort(key=lambda r: (r["Page"], r["Section"].lower(), r["Field"].lower(), r["Index"]))

    df = pd.DataFrame(rows, columns=["Section", "Page", "Field", "Index", "Value"])

    # Write Excel or CSV
    if out_path.lower().endswith(".csv"):
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(out_path, index=False)  # requires openpyxl installed

    print(f"📤 Lookup template written to {out_path} with {len(df)} rows.")
    return out_path


def export_lookup_template_from_pdf(input_pdf: str,
                                    template_json: str = "template_fields.json",
                                    out_path: str = "lookup_template.xlsx",
                                    rebuild_template: bool = False) -> str:
    """
    Ensures a template JSON exists (building it if needed), then exports the lookup sheet.
    """
    must_build = rebuild_template or not os.path.exists(template_json)
    if must_build:
        builder = _get_builder()
        # You can pass lookup_rows=None here (sections will still be auto-detected)
        builder(input_pdf, template_json, lookup_rows=None, dry_run=False)

    return export_lookup_template_from_json(template_json, out_path)


# ---------------------------
# CLI
# ---------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Export a blank lookup sheet (Section, Page, Field, Index, Value) from template_fields.json"
    )
    ap.add_argument("--input", required=True, help="Input PDF path")
    ap.add_argument("--template", default="template_fields.json", help="Template JSON (built if missing)")
    ap.add_argument("--make-lookup", metavar="OUT.xlsx",
                    help="Path to write the blank lookup sheet (xlsx or csv).")
    ap.add_argument("--rebuild-template", action="store_true",
                    help="Force rebuild of the template JSON even if it already exists")
    args = ap.parse_args()

    if args.make_lookup:
        export_lookup_template_from_pdf(
            input_pdf=args.input,
            template_json=args.template,
            out_path=args.make_lookup,
            rebuild_template=args.rebuild_template,
        )
        return

    # If the user didn’t pass --make-lookup, just print a short hint
    print("Nothing to do. Pass --make-lookup OUT.xlsx (or .csv) to export a blank lookup sheet.")


if __name__ == "__main__":
    main()
