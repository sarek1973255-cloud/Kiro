"""
PDFEditor — edits the customs declaration PDF template.

Architecture:
  - Template PDF is the immutable background (never destroyed).
  - Old template text is erased via a three-phase approach that guarantees
    table borders and separator lines are always restored:
      Phase 1 — CAPTURE: page.get_drawings() records all vector paths in the
                clearing area BEFORE any modification.
      Phase 2 — ERASE: per-span redact annotations at exact span bboxes are
                applied with apply_redactions(graphics=0) to remove text.
      Phase 3 — RESTORE: all captured drawings are redrawn ON TOP of the white
                fill, ensuring separator lines are always visible.
  - After text-only erasure, new content is painted with insert_text.

Coordinate space (physical portrait, x: 0-595, y: 0-842, origin at TOP):
  - The template has /Rotate 90, so it displays as landscape in PDF viewers.
  - All PyMuPDF operations use physical coords.
  - insert_text(point, rotate=90): chars advance in +y (= rightward landscape).

FONT RULES (per spec)
─────────────────────
  • ALL dynamically inserted text uses SimSun at 9 pt.
  • If SimSun is not found on the system, NotoSansCJKsc is used as a fallback.
  • Font size is NEVER reduced for long text — the text area expands instead.
  • Original template text properties (bold, spacing, rendering) are untouched.

TABLE LINE RULES
────────────────
  • Table lines are NEVER redrawn, repainted, or overwritten.
  • The three-phase erase+restore cycle keeps all vector paths intact.
  • Text is nudged slightly if it would overlap a cell border.
"""

from __future__ import annotations

import io
import math
import os
import logging
import unicodedata
import urllib.request
from typing import Optional, List, Tuple

import fitz  # PyMuPDF

from config import (
    TEMPLATE_PDF, CJK_FONT_FILE, CJK_FONT_URL, SIMSUN_FONT_FILE,
    FONT_SIZE_HEADER, FONT_SIZE_ITEM, FONT_SIZE_SMALL,
    PAGE1_CAPACITY, TEMPLATE_CAPACITY,
    HEADER_RECTS, PAGE_HEADER_RECTS, ERASE_ONLY_RECTS,
    PAGE1_ITEM_X, TEMPLATE_ITEM_X, ITEM_STRIP_WIDTH,
    SC1_X, SC2_X, SC3_X,
    ITEM_Y, STATIC_ITEM,
    ITEM_Y_MIN, ITEM_Y_MAX,
    SENDER_WRITE_X, PRODUCER_WRITE_X,
)

from excel_parser import DeclarationData, ProductItem
from translation import get_chinese_company_name, normalize_text as _normalize_text

# Minimum inset from cell edge (physical x) to text start.
_BORDER_INSET = 1.0           # pt — x gap from physical cell edge to glyph start

# Safety clamp: used in right_align=True (header labels) only.
# Item rows use the center formula instead — see _insert_centered below.
_CELL_PADDING = 1.0   # pt — minimum gap from a cell border to a header text start

# Internal font alias — registered once per page via insert_text's fontfile arg.
_FONT_ALIAS = "simsun_custom"

log = logging.getLogger(__name__)

Rect4 = Tuple[float, float, float, float]


class PDFEditor:
    def __init__(self, template_path: str = TEMPLATE_PDF):
        self.template_path = template_path
        self._font_path: Optional[str] = self._resolve_font()
        # Cache for text width measurements: (text, fontsize) -> width_pt
        self._width_cache: dict = {}
        # Scratch document used for exact font metric measurements
        self._scratch_doc: Optional[fitz.Document] = None

        # Measure actual font ascent (x_insert - bbox_x0 for rotate=90 text).
        # This varies by font: NotoSansCJKsc ≈ 10.44 pt, SimSun ≈ 7.74 pt.
        # All x-position constants are derived from this value so that output
        # matches the reference declaration regardless of which font is loaded.
        self._font_ascent: float = self._measure_font_ascent()
        # _text_in_rect offset: places bbox_x0 at template reference x0.
        # Formula: x_off = font_ascent - 8.50  (tuned to match namuna.pdf template)
        # Gives bbox_x0 = rect_x0 + x_off + BORDER_INSET - ascent = rect_x0 - 7.5
        # e.g. packages rect_x0=202.3 → bbox_x0=194.8 ✓ (template: 194.8)
        self._x_off: float = self._font_ascent - 8.50
        # Item sub-column offsets (added to strip base ix before _insert_centered).
        self._sc1_x: float = 1.70 + self._font_ascent
        self._sc2_x: float = 11.60 + self._font_ascent
        self._sc3_x: float = 22.70 + self._font_ascent
        # Sender / producer company name x position.
        # bbox_x0 = sender_x + BORDER_INSET - ascent = 100.1 + 1.0 - ascent + ascent = 101.1
        # which matches template namuna.pdf sender x0=101.1
        self._sender_x: float   = 100.1 + self._font_ascent
        self._producer_x: float = 146.6 + self._font_ascent

        log.info(
            "PDFEditor ready.  Font: %s  ascent=%.2f  x_off=%.2f",
            self._font_path or "unavailable — text insertion will be skipped",
            self._font_ascent, self._x_off,
        )

    def __del__(self):
        if self._scratch_doc is not None:
            try:
                self._scratch_doc.close()
            except Exception:
                pass

    # ── Font ascent measurement ────────────────────────────────────────────

    def _measure_font_ascent(self) -> float:
        """
        Measure the effective font ascent for x-positioning of rotate=90 text.

        For rotate=90 text inserted at (x_insert, y):
            bbox_x0 = x_insert - font_ascent

        We insert a test CJK glyph at a known x and read back its bbox_x0.
        This must be called AFTER _resolve_font() so self._font_path is set.
        """
        if not self._font_path:
            return 10.44  # NotoSansCJKsc fallback
        try:
            tmp = fitz.open()
            pg = tmp.new_page(width=400, height=600)
            test_x = 200.0
            pg.insert_text(
                (test_x, 300.0), "中",
                fontname="asc_probe",
                fontfile=self._font_path,
                fontsize=FONT_SIZE_HEADER,
                rotate=90,
                color=(1, 1, 1),   # invisible — measurement only
            )
            words = pg.get_text("words")
            tmp.close()
            if words:
                ascent = test_x - words[0][0]
                log.info("Font ascent measured: %.4f pt (font: %s)",
                         ascent, os.path.basename(self._font_path))
                return ascent
        except Exception as exc:
            log.debug("Font ascent measurement failed: %s — using default 10.44", exc)
        return 10.44  # safe default (NotoSansCJKsc)

    # ── Font resolution ────────────────────────────────────────────────────

    @staticmethod
    def _resolve_font() -> Optional[str]:
        """
        Return the path to the font file to use for ALL text insertion.

        Priority:
          1. SimSun (simsun.ttc / simsun.ttf) — required per spec.
          2. NotoSansCJKsc-Regular.otf         — fallback if SimSun absent.

        If neither exists locally, NotoSansCJKsc is downloaded from GitHub.
        """
        if SIMSUN_FONT_FILE and os.path.exists(SIMSUN_FONT_FILE):
            log.info("Using SimSun font: %s", SIMSUN_FONT_FILE)
            return SIMSUN_FONT_FILE

        # SimSun not found — fall back to NotoSansCJKsc
        if os.path.exists(CJK_FONT_FILE):
            log.warning(
                "SimSun not found — falling back to NotoSansCJKsc.  "
                "Place simsun.ttc in the fonts/ directory for the required font."
            )
            return CJK_FONT_FILE

        # Neither exists — try to download NotoSansCJKsc
        os.makedirs(os.path.dirname(CJK_FONT_FILE) or ".", exist_ok=True)
        try:
            log.info("Downloading fallback CJK font from GitHub …")
            urllib.request.urlretrieve(CJK_FONT_URL, CJK_FONT_FILE)
            log.info("CJK font saved to %s", CJK_FONT_FILE)
            return CJK_FONT_FILE
        except Exception as exc:
            log.warning("Font download failed: %s", exc)
            return None

    # ── Text width measurement ─────────────────────────────────────────────

    def _get_scratch_doc(self) -> fitz.Document:
        """Return a cached scratch document for text measurement."""
        if self._scratch_doc is None:
            self._scratch_doc = fitz.open()
        return self._scratch_doc

    def _measure_text_width(self, text: str, fontsize: float) -> float:
        """
        Return the exact advance width of `text` at `fontsize` using the
        loaded font.  Text is rendered to a single persistent scratch page
        whose content is cleared between uses; the bbox of the rendered words
        gives the exact advance width from real glyph metrics.

        Results are cached — each unique (text, fontsize) pair is measured
        only once.  Texts with spaces are handled by taking the full y-span
        across all word bboxes rather than just the first word.
        """
        if not self._font_path or not text:
            return 0.0

        cache_key = (text, fontsize)
        if cache_key in self._width_cache:
            return self._width_cache[cache_key]

        try:
            sd = self._get_scratch_doc()
            # Replace the scratch page on every call so no previous text
            # pollutes get_text("words").  clean_contents() does NOT erase
            # previously painted glyphs — they accumulate and make y_max-y_min
            # grow with each call, causing all widths to be overestimated.
            while len(sd) > 0:
                sd.delete_page(0)
            sd.new_page(width=300, height=1200)
            page = sd[0]

            y_insert = 1100.0
            page.insert_text(
                (50.0, y_insert),
                text,
                fontname=_FONT_ALIAS,
                fontfile=self._font_path,
                fontsize=fontsize,
                rotate=90,
                color=(1, 1, 1),   # white — invisible, we only read bbox
            )
            words = page.get_text("words")
            if words:
                # With rotate=90, text advances upward (toward lower y).
                # y_max ≈ y_insert (start of text), y_min ≈ y_insert - advance.
                y_min = min(w[1] for w in words)
                y_max = max(w[3] for w in words)
                width = y_max - y_min
            else:
                width = self._estimate_text_width(text, fontsize)
        except Exception as exc:
            log.debug("Text measurement failed for %r: %s — using estimate",
                      text[:20], exc)
            width = self._estimate_text_width(text, fontsize)

        self._width_cache[cache_key] = width
        return width

    @staticmethod
    def _estimate_text_width(text: str, fontsize: float) -> float:
        """
        Fallback glyph-width estimate when direct measurement fails.

        Calibrated against measured NotoSansCJKsc / SimSun advance widths
        at 9 pt (see measurement log in module docstring):
          CJK ideograph:        1.00 × fontsize  (9 pt → 9 pt)
          ASCII digit (0-9):    0.55 × fontsize  (9 pt → 4.95 pt)
          ASCII letter:         0.55 × fontsize
          Brackets/slash/pipe:  0.32 × fontsize
          Period/comma/colon:   0.28 × fontsize
          Space:                0.35 × fontsize
          Other Unicode:        0.75 × fontsize
        """
        width = 0.0
        for ch in text:
            cp = ord(ch)
            if (0x4E00 <= cp <= 0x9FFF   # CJK Unified Ideographs
                    or 0x3400 <= cp <= 0x4DBF   # CJK Extension A
                    or 0x3000 <= cp <= 0x303F   # CJK Symbols & Punctuation
                    or 0xFF01 <= cp <= 0xFF60   # Fullwidth Latin
                    or 0xF900 <= cp <= 0xFAFF): # CJK Compatibility
                width += fontsize * 1.00
            elif 0x30 <= cp <= 0x39:                        # digits
                width += fontsize * 0.55
            elif 0x41 <= cp <= 0x5A or 0x61 <= cp <= 0x7A: # ASCII letters
                width += fontsize * 0.55
            elif ch in '|/\\!':
                width += fontsize * 0.32
            elif ch in '.,;:':
                width += fontsize * 0.28
            elif ch in '()[]{}':
                width += fontsize * 0.35
            elif ch == ' ':
                width += fontsize * 0.35
            else:
                width += fontsize * 0.75
        return width

    # ── Per-page font registration ─────────────────────────────────────────

    def _ensure_font_on_page(self, page: fitz.Page) -> bool:
        """
        Register the CJK font on *page* once so every subsequent insert_text
        call can reference it by alias alone (no fontfile= on each call).
        This avoids PyMuPDF re-embedding the font data for every text
        insertion, which is the main cause of bloated output files.

        Returns True if the font is now available, False otherwise.
        """
        if not self._font_path:
            return False
        try:
            page.insert_font(fontname=_FONT_ALIAS, fontfile=self._font_path)
            return True
        except Exception as exc:
            log.warning("insert_font failed: %s", exc)
            return False

    # ── Core left-aligned text insertion ──────────────────────────────────

    # Header fields (right_align=True) anchor at y_hi - _LEFT_ALIGN_MARGIN.
    _LEFT_ALIGN_MARGIN = 2.5   # pt — header right-align gap from y_hi border

    def _insert_centered(
        self,
        page: fitz.Page,
        x: float,
        y_lo: float,
        y_hi: float,
        text: str,
        fontsize: float = FONT_SIZE_ITEM,
        right_align: bool = False,
    ) -> None:
        """
        Insert `text` within the physical y-range [y_lo, y_hi].

        Coordinate note: rotate=90 text advances toward LOWER physical y
        (= rightward in landscape).  y_hi is the landscape-LEFT edge of the
        cell (high physical y); y_lo is the landscape-RIGHT edge.

        Two alignment modes:

          right_align=False (default — CENTER for item rows):
            y_insert = (y_lo + y_hi + text_width) / 2
            Centers the text between y_lo and y_hi so the top margin equals
            the bottom margin.  When text_width > cell_height the surplus is
            split equally above and below — minimising any border overlap.
            Requires measuring the actual glyph advance width.

          right_align=True (header labels):
            y_insert = y_hi - _LEFT_ALIGN_MARGIN
            Anchors at y_hi (landscape-left / label side).  Used for header
            fields (e.g. sender 德力西, contract number).

        Font size is NEVER reduced for overflow — text extends beyond y_lo.
        A border inset of _BORDER_INSET pt is applied to x.
        """
        if not text:
            return
        if not self._font_path:
            log.warning("No font available — skipping: %r", text[:20])
            return

        text = unicodedata.normalize("NFKC", text)

        if right_align:
            y_insert = y_hi - self._LEFT_ALIGN_MARGIN
        else:
            # CENTER: text is placed so the top margin equals the bottom margin.
            #   y_insert = (y_lo + y_hi + text_width) / 2
            # For text that fits in the cell, both margins = (cell - width) / 2.
            # For text wider than the cell, the overflow is split equally so the
            # text crosses BOTH borders by the same amount — minimising how far
            # it intrudes into neighbouring cells vs. a one-sided anchor.
            text_width = self._measure_text_width(text, fontsize)
            y_insert = (y_lo + y_hi + text_width) / 2

            # CLAMP: prevent text from touching cell borders.
            # y_insert must not exceed y_hi - padding (top border).
            # text end (y_insert - text_width) must not go below y_lo + padding.
            _Y_PADDING = 1.5   # pt — minimum gap from text to cell border
            y_max_allowed = y_hi - _Y_PADDING
            y_min_end     = y_lo + _Y_PADDING
            # Clamp: if text would start beyond y_hi border, pull it down
            if y_insert > y_max_allowed:
                y_insert = y_max_allowed
            # If text end would go below y_lo border, push start up
            # (only if it doesn't violate the y_hi clamp)
            if y_insert - text_width < y_min_end:
                y_insert = y_min_end + text_width
                # Re-apply y_hi clamp (text is just too long — prefer y_hi)
                if y_insert > y_max_allowed:
                    y_insert = y_max_allowed

        x_insert = x + _BORDER_INSET

        try:
            page.insert_text(
                (x_insert, y_insert),
                text,
                fontname=_FONT_ALIAS,
                fontsize=fontsize,
                rotate=90,
                color=(0, 0, 0),
            )
        except Exception as exc:
            log.warning(
                "insert_text failed at (%.1f, %.1f) text=%r: %s",
                x_insert, y_insert, text[:30], exc,
            )

    # ── Public API ────────────────────────────────────────────────────────

    def edit(
        self,
        data: DeclarationData,
        pre_entry_no: str = "",
        declaration_date: str = "",
    ) -> bytes:
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found: {self.template_path}")

        if pre_entry_no:
            data.pre_entry_no = pre_entry_no
        if declaration_date:
            data.declaration_date = declaration_date

        n            = len(data.items)
        pages_needed = self._calc_pages(n)

        expected_slots = (
            PAGE1_CAPACITY
            if pages_needed == 1
            else PAGE1_CAPACITY + (pages_needed - 1) * TEMPLATE_CAPACITY
        )
        assert expected_slots >= n, (
            f"Slot count mismatch: {expected_slots} slots for {n} items"
        )

        # ── Build output document from DEEP-INDEPENDENT page copies ────
        tmpl = fitz.open(self.template_path)
        tmpl_count    = len(tmpl)
        cont_page_idx = tmpl_count - 1  # last template page = continuation layout

        doc = fitz.open()
        for pidx in range(pages_needed):
            src_idx = 0 if pidx == 0 else cont_page_idx
            doc.insert_pdf(tmpl, from_page=src_idx, to_page=src_idx)
        tmpl.close()

        total = len(doc)

        for pidx in range(total):
            page     = doc[pidx]
            is_first = (pidx == 0)
            item_xs  = PAGE1_ITEM_X if is_first else TEMPLATE_ITEM_X
            item_off = (
                0 if is_first
                else PAGE1_CAPACITY + (pidx - 1) * TEMPLATE_CAPACITY
            )

            # ── STEP 1: Collect all logical regions that need clearing ──
            clear_rects: List[Rect4] = []

            for r in PAGE_HEADER_RECTS.values():
                clear_rects.append(r)

            if is_first:
                for r in HEADER_RECTS.values():
                    clear_rects.append(r)
                # Always erase fields that must remain empty (user fills manually)
                for r in ERASE_ONLY_RECTS.values():
                    clear_rects.append(r)

            for ix in item_xs:
                clear_rects.append((
                    ix + 0.5,
                    ITEM_Y_MIN,
                    ix + ITEM_STRIP_WIDTH - 0.5,
                    ITEM_Y_MAX,
                ))

            # ── STEP 2: Text-only erasure (table lines preserved) ───────
            self._erase_text_in_regions(page, clear_rects)

            # Register the CJK font AFTER erasure (apply_redactions can
            # rebuild the page resource dict, losing earlier registrations).
            self._ensure_font_on_page(page)

            # ── STEP 3: Write per-page content ─────────────────────────
            self._write_page_number(page, pidx + 1, total)

            if data.pre_entry_no:
                self._write_customs_number(page, data.pre_entry_no,
                                           rects=PAGE_HEADER_RECTS)

            if is_first:
                self._write_header(page, data)

            for slot, ix in enumerate(item_xs):
                gi = item_off + slot
                if gi < n:
                    self._write_item(page, ix, data.items[gi])

        buf = io.BytesIO()
        # Subset fonts — only embed the glyphs actually used on each page.
        # This dramatically reduces file size when using large CJK fonts
        # (SimSun ≈ 10 MB full → typically < 200 KB subset per page).
        try:
            doc.subset_fonts()
        except AttributeError:
            # PyMuPDF < 1.23.0: subset_fonts not available
            log.debug("subset_fonts() not available — skipping font subsetting.")
        except Exception as exc:
            log.debug("subset_fonts() failed: %s — continuing without subsetting.", exc)

        # garbage=4: maximum object compaction; deflate=True: stream compression.
        # clean=True: remove unused objects and duplicates (esp. embedded fonts).
        # Together these keep the output file close in size to the template.
        doc.save(buf, garbage=4, deflate=True, clean=True)
        doc.close()
        log.info("PDF generated: %d bytes, %d page(s), %d item(s).",
                 buf.tell(), total, n)
        return buf.getvalue()

    # ── Text-only erasure ─────────────────────────────────────────────────

    @staticmethod
    def _erase_text_in_regions(page: fitz.Page, regions: List[Rect4]) -> None:
        """
        Erase template text inside the given regions WITHOUT any fill or
        white rectangle overlay — the cell background, borders, and grid
        lines are left completely untouched, exactly as in the original
        template.

        Method: per-span redact annotations with fill=None so that
        apply_redactions() removes only the text-rendering operators from
        the content stream and leaves every other visual element (lines,
        fills, images) intact.  No Phase-1 drawing capture or Phase-3
        redraw is required because nothing is ever painted over.
        """
        try:
            raw = page.get_text("rawdict", flags=0)
        except Exception as exc:
            log.warning("get_text failed — text erasure skipped: %s", exc)
            return

        region_rects = [fitz.Rect(*r) for r in regions]
        matching_spans: List[fitz.Rect] = []

        for block in raw.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    span_rect = fitz.Rect(span["bbox"])
                    for rr in region_rects:
                        if rr.intersects(span_rect):
                            matching_spans.append(span_rect)
                            break

        if not matching_spans:
            return

        # Add redact annotations with NO fill — text is removed, background
        # colour, cell shading, and all vector paths remain untouched.
        for sr in matching_spans:
            page.add_redact_annot(sr, fill=None)

        try:
            page.apply_redactions(images=0, graphics=0)
        except TypeError:
            # PyMuPDF < 1.18: graphics flag not supported — still no fill.
            page.apply_redactions(images=0)
        except Exception as exc:
            log.warning("apply_redactions failed (%s) — erasure skipped.", exc)
            _delete_redact_annotations(page)
            return

        log.debug("Erase: %d span(s) removed (no fill).", len(matching_spans))

    # ── Page-number & customs-number helpers ─────────────────────────────

    def _write_page_number(self, page: fitz.Page, cur: int, tot: int) -> None:
        self._text_in_rect(
            page, PAGE_HEADER_RECTS["page_no"],
            f"{cur}/{tot}", fontsize=FONT_SIZE_SMALL,
        )

    def _write_customs_number(
        self,
        page: fitz.Page,
        number: str,
        rects: dict,
    ) -> None:
        for key in ("customs_no_top", "customs_no_bottom"):
            if key in rects:
                self._text_in_rect(
                    page, rects[key], number,
                    fontsize=FONT_SIZE_HEADER,
                )
        if "barcode_text" in rects:
            self._write_barcode_text(page, number, rects["barcode_text"])

    # ── Header (page 1 only) ──────────────────────────────────────────────

    def _write_header(self, page: fitz.Page, data: DeclarationData) -> None:
        def wl(key: str, val: str, right_align: bool = False) -> None:
            if val:
                self._text_in_rect(
                    page, HEADER_RECTS[key], val,
                    fontsize=FONT_SIZE_HEADER,
                    right_align=right_align,
                )

        # Fields anchored to y_hi (label side) — RIGHT-ALIGN:
        # consignee: foreign (non-Chinese) company — keep original name as-is.
        # Only the domestic sender (境内发货人) is translated to Chinese below.
        wl("contract_no",  data.contract_no,                   right_align=True)
        wl("consignee",    data.consignee or "",                right_align=True)
        wl("gross_weight", self._fmt_weight(data.gross_weight), right_align=True)
        wl("net_weight",   self._fmt_weight(data.net_weight),   right_align=True)
        wl("packages",     str(data.packages),                  right_align=True)
        # declaration_date — bo'sh qoldiriladi (foydalanuvchi ruchnoy to'ldiradi)

        if data.sender_name:
            sender_cn = get_chinese_company_name(data.sender_name)
            _, y0, _, y1 = HEADER_RECTS["sender"]
            self._insert_centered(page, self._sender_x, y0, y1,
                                   sender_cn, FONT_SIZE_HEADER,
                                   right_align=True)
            _, y0, _, y1 = HEADER_RECTS["producer"]
            self._insert_centered(page, self._producer_x, y0, y1,
                                   sender_cn, FONT_SIZE_HEADER,
                                   right_align=True)

        if data.pre_entry_no:
            self._write_customs_number(page, data.pre_entry_no,
                                       rects=HEADER_RECTS)

    # ── Item row ──────────────────────────────────────────────────────────

    def _write_item(self, page: fitz.Page, ix: float, item: ProductItem) -> None:
        """
        Write all data fields for one product item into the strip at physical
        x = ix.  Every field is centered within its assigned y-range, which
        is the horizontal span of the column in landscape display.

        Sub-row layout (physical x within the strip):
          SC1 (ix + 8 pt)  — main product info: item_no, hs_code, ch_name, qty,
                             unit_price, origin_cn, dest_cn, source_tax
          SC2 (ix + 19 pt) — secondary:  brand_row, total_price, origin_code,
                             dest_code, tax_code
          SC3 (ix + 29 pt) — repeat:  qty, currency
        """
        fs = FONT_SIZE_ITEM

        def wc(y_key: str, sc_off: float, text: str,
               fontsize: float = FONT_SIZE_ITEM) -> None:
            y_lo, y_hi = ITEM_Y[y_key]
            self._insert_centered(page, ix + sc_off, y_lo, y_hi, text, fontsize)

        # ── Row 1 (SC1) ────────────────────────────────────────────────────
        # item_no: use same x offset as SC1 fields (ix + SC1_X).
        # For rotate=90 text, get_text reports x0 ≈ x_insert - font_ascent.
        # Template has item_no at x0=ix; SC1_X+_BORDER_INSET gives x_insert≈ix+9
        # so get_text x0 ≈ ix+9-9 = ix. Matches template exactly.
        y_lo, y_hi = ITEM_Y["item_no"]
        wc("item_no", self._sc1_x, str(item.item_no))

        wc("hs_code",    self._sc1_x, item.hs_code)
        wc("ch_name",    self._sc1_x, item.chinese_name)
        wc("qty",        self._sc1_x, item.qty_str)
        wc("unit_price", self._sc1_x, f"{item.unit_price:.4f}")
        wc("origin_cn",  self._sc1_x, STATIC_ITEM["origin_cn"])
        wc("dest_cn",    self._sc1_x, STATIC_ITEM["dest_cn"])
        self._write_source_tax(page, ix, fs)

        # ── Row 2 (SC2) ────────────────────────────────────────────────────
        wc("brand_row",   self._sc2_x, STATIC_ITEM["brand_row"])
        wc("total_price", self._sc2_x, item.total_price_str)
        wc("origin_code", self._sc2_x, STATIC_ITEM["origin_code"])
        wc("dest_code",   self._sc2_x, STATIC_ITEM["dest_code"])
        wc("tax_code",    self._sc2_x, STATIC_ITEM["tax_code"])

        # ── Row 3 (SC3) ────────────────────────────────────────────────────
        wc("qty",      self._sc3_x, item.qty_str)
        wc("currency", self._sc3_x, STATIC_ITEM["currency"])

    def _write_source_tax(self, page: fitz.Page, ix: float, fs: float) -> None:
        """
        Write the 境内货源地 + 征免 combined field, splitting at the space so
        each part is centered in its own landscape column:
          Part 1 — "(33019)杭州其他"  in 境内货源地  (physical y ≈ 93–152)
          Part 2 — "照章征税"         in 征免         (physical y ≈ 43–93)

        The boundary y ≈ 93 is the separator line between the two columns
        as measured from namuna.pdf.
        """
        combined  = STATIC_ITEM["source_tax"]   # "(33019)杭州其他 照章征税"
        parts     = combined.split(" ", 1)
        y_lo_full, y_hi_full = ITEM_Y["source_tax"]   # (43.0, 152.0)
        y_boundary = 93.0                               # measured separator

        if len(parts) == 2:
            part1, part2 = parts
            # "(33019)杭州其他" — upper half (closer to y_hi = 152 in physical)
            self._insert_centered(page, ix + self._sc1_x,
                                   y_boundary, y_hi_full, part1, fs)
            # "照章征税" — lower half (closer to y_lo = 43 in physical)
            self._insert_centered(page, ix + self._sc1_x,
                                   y_lo_full, y_boundary, part2, fs)
        else:
            # Fallback: single string in full range
            self._insert_centered(page, ix + self._sc1_x,
                                   y_lo_full, y_hi_full, combined, fs)

    # ── Low-level text writers ────────────────────────────────────────────

    def _text_in_rect(
        self,
        page: fitz.Page,
        rect_tuple: Rect4,
        text: str,
        fontsize: float = FONT_SIZE_ITEM,
        right_align: bool = False,
    ) -> None:
        """Insert text within a rect given as (x0, y0, x1, y1)."""
        x0, y0, x1, y1 = rect_tuple
        self._insert_centered(page, x0 + self._x_off, y0, y1, text,
                               fontsize=fontsize, right_align=right_align)

    def _write_barcode_text(
        self,
        page: fitz.Page,
        number: str,
        rect: Rect4,
    ) -> None:
        x0, y0, x1, y1 = rect
        try:
            page.insert_text(
                (x0 + 2.0, y1),
                f"*{number}*",
                fontname="cour",
                fontsize=5.5,
                rotate=90,
                color=(0, 0, 0),
            )
        except Exception as exc:
            log.warning("Barcode text write failed: %s", exc)

    # ── Pagination ────────────────────────────────────────────────────────

    @staticmethod
    def _calc_pages(n: int) -> int:
        if n <= PAGE1_CAPACITY:
            return 1
        return 1 + math.ceil((n - PAGE1_CAPACITY) / TEMPLATE_CAPACITY)

    # ── Formatting ────────────────────────────────────────────────────────

    @staticmethod
    def _fmt_weight(val: float) -> str:
        return str(int(val)) if val == int(val) else str(val)


# ── Module-level helpers ──────────────────────────────────────────────────

def _delete_redact_annotations(page: fitz.Page) -> None:
    """
    Delete all unapplied redact annotations from the page.

    Called by the fallback path in _erase_text_in_regions when
    apply_redactions raises before it could process the annotations.
    Unapplied redact annotations must be removed: PDF viewers render them
    as visible purple rectangles, corrupting the document appearance.

    PDF_ANNOT_REDACT = 12 per the PDF specification.
    """
    redact_type = getattr(fitz, "PDF_ANNOT_REDACT", 12)
    to_delete = [a for a in page.annots() if a.type[0] == redact_type]
    for annot in to_delete:
        page.delete_annot(annot)
    if to_delete:
        log.debug("Deleted %d unapplied redact annotation(s).", len(to_delete))


def _redraw_path_drawings(page: fitz.Page, drawings: List[dict]) -> None:
    """
    Redraw a list of path drawings (as returned by page.get_drawings()) onto
    the page as new content placed ABOVE the current content stream.

    In PDF rendering, content added later is painted on top of content that
    was already there.  After apply_redactions() has added white fill over the
    cleared areas, calling this function repaints the border lines and separator
    lines ON TOP of that white fill — making them visible again.

    Supported drawing primitives:
      'l'  — straight line
      'r'  — rectangle
      'c'  — cubic Bézier curve
      'qu' — quadrilateral

    All other primitive types are skipped without error.
    """
    for d in drawings:
        items = d.get("items", [])
        if not items:
            continue
        try:
            shape = page.new_shape()
            drew = False
            for item in items:
                kind = item[0]
                try:
                    if kind == "l":
                        shape.draw_line(item[1], item[2])
                        drew = True
                    elif kind == "r":
                        shape.draw_rect(item[1])
                        drew = True
                    elif kind == "c":
                        shape.draw_bezier(item[1], item[2], item[3], item[4])
                        drew = True
                    elif kind == "qu":
                        shape.draw_quad(item[1])
                        drew = True
                except Exception:
                    pass

            if not drew:
                continue

            lc = d.get("lineCap", 0)
            if isinstance(lc, (tuple, list)):
                lc = lc[0] if lc else 0

            shape.finish(
                color=d.get("color"),
                fill=d.get("fill"),
                width=d.get("width", 1),
                lineCap=int(lc),
                lineJoin=int(d.get("lineJoin") or 0),
                dashes=d.get("dashes") or "",
                even_odd=bool(d.get("even_odd", False)),
                closePath=bool(d.get("closePath", False)),
            )
            shape.commit()
        except Exception as exc:
            log.debug("_redraw_path_drawings: skipped one path — %s", exc)
