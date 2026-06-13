"""
main.py — CLI entry point for local PDF generation (no Telegram required).

Usage
─────
  python main.py <excel_file.xlsx> [options]

Options
  --entry-no   NUMBER     18-digit customs pre-entry number
  --date       YYYYMMDD   Declaration date
  --output     PATH       Output PDF path  (default: output/<contract>.pdf)

Examples
  python main.py manba.xlsx
  python main.py manba.xlsx --entry-no 941520260000090882 --date 20260317
  python main.py manba.xlsx --output result.pdf
"""

from __future__ import annotations

import argparse
import logging
import os
import sys

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Generate a Chinese customs declaration PDF from an Excel file."
    )
    ap.add_argument("excel",       help="Path to the Excel (.xlsx) source file")
    ap.add_argument("--entry-no",  default="", metavar="NUMBER",
                    help="18-digit customs pre-entry number")
    ap.add_argument("--date",      default="", metavar="YYYYMMDD",
                    help="Declaration date in YYYYMMDD format")
    ap.add_argument("--output",    default="", metavar="PATH",
                    help="Output PDF path (default: output/<contract_no>.pdf)")
    args = ap.parse_args()

    if not os.path.exists(args.excel):
        log.error("Excel file not found: %s", args.excel)
        sys.exit(1)

    # ── Parse Excel ───────────────────────────────────────────────────────
    from excel_parser import ExcelParser
    parser = ExcelParser()
    log.info("Parsing %s …", args.excel)
    try:
        data = parser.parse(args.excel)
    except Exception as exc:
        log.error("Failed to parse Excel: %s", exc)
        sys.exit(1)

    log.info(
        "Parsed: contract=%s  consignee=%s  items=%d  packages=%d  "
        "gross=%.1f kg  net=%.1f kg  total=$%.2f",
        data.contract_no, data.consignee, len(data.items),
        data.packages, data.gross_weight, data.net_weight,
        data.total_amount_usd,
    )

    if not data.items:
        log.error("No items found in the Excel file. Check sheet names and column headers.")
        sys.exit(1)

    # ── Generate PDF ──────────────────────────────────────────────────────
    from pdf_editor import PDFEditor
    editor = PDFEditor()
    log.info("Generating PDF …")
    try:
        pdf_bytes = editor.edit(
            data,
            pre_entry_no=args.entry_no,
            declaration_date=args.date,
        )
    except FileNotFoundError as exc:
        log.error("%s", exc)
        sys.exit(1)
    except Exception as exc:
        log.error("PDF generation failed: %s", exc)
        sys.exit(1)

    # ── Write output ──────────────────────────────────────────────────────
    from config import OUTPUT_DIR
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if args.output:
        out_path = args.output
    else:
        safe = data.contract_no.replace("/", "-").replace("\\", "-") or "declaration"
        out_path = os.path.join(OUTPUT_DIR, f"declaration_{safe}.pdf")

    with open(out_path, "wb") as f:
        f.write(pdf_bytes)

    log.info("PDF saved to: %s  (%d bytes)", out_path, len(pdf_bytes))
    print(f"\n✅  Done!  →  {out_path}")

    # ── Write missing-translations report ─────────────────────────────────
    from translation import write_missing_translations_report
    report_dir = os.path.dirname(os.path.abspath(out_path))
    report_path = os.path.join(report_dir, "missing_translations.txt")
    write_missing_translations_report(report_path)


if __name__ == "__main__":
    main()
