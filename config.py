"""
Coordinate configuration for the customs PDF editor.

All coordinates are in PyMuPDF's get_text() / insert_text() physical space,
measured directly from namuna.pdf using page.get_text("rawdict").

The template PDF has /Rotate 90.  PyMuPDF exposes the PHYSICAL (pre-rotation)
coordinate system where:
  x: 0 → 595  (portrait top → bottom = landscape left-column → right-column)
  y: 0 → 842  (portrait top → bottom = landscape right → left)

For insert_text((x, y_start), text, rotate=90):
  • Characters advance from y_start TOWARD LOWER y  (= rightward in landscape).
  • y_start is therefore the LEFT edge of the text in landscape view.
  • y_end   is the RIGHT edge (lower y value).

Visual column layout (landscape left→right) maps to physical y (high→low):
  项号 ← y≈800-806   ···   征免 ← y≈43-60
"""

from __future__ import annotations

import os

# ── Paths ──────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PDF  = os.path.join(BASE_DIR, "template", "namuna.pdf")
OUTPUT_DIR    = os.path.join(BASE_DIR, "output")
FONT_DIR      = os.path.join(BASE_DIR, "fonts")

# ── SimSun font resolution ─────────────────────────────────────────────────
# SimSun (宋体) is the required font per spec.  On Windows it ships with the OS.
# On Linux/macOS, place simsun.ttc or simsun.ttf in the fonts/ directory.
# Candidates are checked in order; the first existing file is used.
_SIMSUN_CANDIDATES = [
    os.path.join(FONT_DIR, "simsun.ttc"),
    os.path.join(FONT_DIR, "simsun.ttf"),
    os.path.join(FONT_DIR, "SimSun.ttc"),
    os.path.join(FONT_DIR, "SimSun.ttf"),
    r"C:\Windows\Fonts\simsun.ttc",           # Windows system path
    r"C:\Windows\Fonts\simsun.ttf",
    "/usr/share/fonts/truetype/simsun.ttc",   # some Linux distros
    "/usr/share/fonts/simsun.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",  # macOS Songti (similar)
]

def _find_simsun() -> str | None:
    for p in _SIMSUN_CANDIDATES:
        if os.path.exists(p):
            return p
    return None

SIMSUN_FONT_FILE = _find_simsun()   # None if SimSun is not found on this system

# CJK fallback — used only when SimSun is not available.
CJK_FONT_FILE = os.path.join(FONT_DIR, "NotoSansCJKsc-Regular.otf")

CJK_FONT_URL  = (
    "https://github.com/googlefonts/noto-cjk/raw/main"
    "/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf"
)

# ── Font sizes ─────────────────────────────────────────────────────────────
# Per spec: ALL dynamically inserted text must use SimSun at 9 pt.
# Long text should NOT be shrunk — the text area is expanded instead.
FONT_SIZE_HEADER = 9.0
FONT_SIZE_ITEM   = 9.0
FONT_SIZE_SMALL  = 9.0

# ── Page capacity ─────────────────────────────────────────────────────────
PAGE1_CAPACITY    = 6     # item slots on first page
TEMPLATE_CAPACITY = 14    # item slots on each continuation page

# ── Header fields (page 1 only) ────────────────────────────────────────────
# Rect = (x0, y0, x1, y1)  — used for TEXT ERASURE (wide enough to cover the
# original template glyphs).  Writing uses SENDER_WRITE_X / PRODUCER_WRITE_X
# below to compensate for NotoSansCJKsc vs SimSun ascent difference.
#
# Measured ground-truth bboxes (for reference):
#   DL0405              → [171.9, 774.3, 180.9, 801.3]
#   Elektro Fayz...     → [124.7, 702.5, 133.7, 801.5]
#   德力西(sender)       → [101.1, 774.5, 110.1, 801.5]  template SimSun bbox
#   德力西(producer)     → [147.6, 774.5, 156.6, 801.5]  template SimSun bbox
#   25308.8             → [194.4, 505.7, 203.4, 537.2]
#   24350               → [194.8, 432.5, 203.8, 455.0]
#   978                 → [194.8, 569.8, 203.8, 583.2]
#   20260317            → [101.1, 286.2, 110.1, 322.2]
#   customs no (top)    → [75.5,  466.3,  84.5, 547.2]
#   customs no (bot)    → [75.5,  676.2,  84.5, 757.2]
HEADER_RECTS = {
    # NotoSansCJKsc font metric offset vs SimSun: noto_ascent ≈ 10.44 pt, SimSun ≈ 2.12 pt.
    # For _text_in_rect: x_insert = rect_x0 + 3.0.  bbox_x0 = x_insert - noto_ascent.
    # To align bbox_x0 with template: rect_x0 = template_bbox_x0 + 10.44 - 3.0 = template_bbox_x0 + 7.44
    "contract_no":       (179.3, 772.0, 183.5, 804.0),   # x0 was 171.0; raised to 179.3 so NotoSansCJKsc bbox_x0 ≈ 171.88 (matches template DL0405)
    "consignee":         (132.2, 700.0, 136.5, 804.0),   # x0 was 131.3; raised to 132.2 so bbox_x0 ≈ 124.71 (matches template Elektro position)
    # sender/producer x-range must:
    #   • start AFTER the adjacent label so the label glyph is not erased:
    #       境内发货人  label x1 = 98.66  → sender  x0 must be > 98.66  → use 99.5
    #       生产销售单位 label x1 = 145.16 → producer x0 must be > 145.16 → use 146.0
    #   • cover the template 德力西 glyphs (x: 101-110 sender, 147.6-156.6 producer)
    #   • x1 must NOT reach 境外收货人 label which starts at x=114.4 → use 114.3
    # Actual write position is controlled by SENDER_WRITE_X / PRODUCER_WRITE_X.
    "sender":            ( 99.5, 672.0, 114.3, 804.0),   # x1 was 115.0; narrowed to 114.3 to preserve 境外收货人 label (x0=114.4)
    "producer":          (146.0, 672.0, 159.0, 804.0),   # x1 was 162.0; narrowed to 159.0 to preserve 合同协议号 label (x0=159.6)
    "gross_weight":      (201.8, 503.0, 211.0, 539.0),   # x0 was 201.0; raised to 201.8 so NotoSansCJKsc bbox_x0 ≈ 194.35 (matches template 25308.8)
    "net_weight":        (202.3, 430.0, 211.4, 457.0),   # x0 was 201.4; raised to 202.3 so NotoSansCJKsc bbox_x0 ≈ 194.79 (matches template 24350)
    "packages":          (202.3, 567.0, 211.4, 585.0),   # x0 was 201.4; raised to 202.3 so NotoSansCJKsc bbox_x0 ≈ 194.79 (matches template 978)
    "declaration_date":  (108.6, 284.0, 112.5, 324.0),   # x0 was 107.7; raised to 108.6 so bbox_x0 ≈ 101.12. x1 was 117.7; narrowed to 112.5 to preserve 提运单号 label (x0=113.1)
    "customs_no_top":    ( 83.0, 464.0,  89.5, 549.0),   # x0 was 82.1 (erased 海关编号 label); x1 was 92.1 (erased (9415) code at x=89.9)
    "customs_no_bottom": ( 83.0, 674.0,  92.1, 759.0),   # x0 was 82.1; raised to 83.0 to preserve 预录入编号 label (x1=82.9)
    "barcode_text":      ( 24.0,  74.0,  57.0, 270.0),
}

# ── Fields that must ALWAYS be erased and left EMPTY (user fills manually) ─
# These rects define areas where template text is removed but NO new text is
# written.  The user fills these fields by hand after printing.
ERASE_ONLY_RECTS = {
    # 申报日期 (Declaration date) — "20260317" in template — user fills manually
    "declaration_date":   (100.0, 284.0, 112.0, 324.0),
    # 提运单号 (Transport/Waybill number) — "Y261063799" in template
    "transport_no":       (123.0, 275.0, 135.0, 324.0),
    # 运输工具名称及航次号 (Vehicle name & voyage) — "新Q49076/MA7GT4NE40A350"
    "vehicle_voyage":     (123.0, 349.0, 135.0, 457.0),
    # 集装箱标箱数及号码 (Container number) — "2;TCNU7843764;"
    "container_no":       (239.0, 600.0, 251.0, 757.0),
}

# ── Sender / Producer write x-positions ────────────────────────────────────
# NotoSansCJKsc CJK ascent ≈ 10.4 pt (SimSun ≈ 9 pt).
# For insert_text(rotate=90): bbox_x0 = x_insert - ascent.
# Target bbox_x0 matches template SimSun positions:
#   sender   template bbox_x0 = 101.12  →  x_insert = 101.12 + 10.4 = 111.5
#   producer template bbox_x0 = 147.62  →  x_insert = 147.62 + 10.4 = 158.0
# _insert_centered adds _BORDER_INSET (1.0) to the x param it receives, so:
#   SENDER_WRITE_X   = x_insert − 1.0 = 110.5
#   PRODUCER_WRITE_X = x_insert − 1.0 = 157.0
SENDER_WRITE_X   = 113.24  # calibrated: reference bbox_x0=103.8, NotoSansCJKsc ascent=10.44 → x_insert=114.24 → param=113.24
PRODUCER_WRITE_X = 159.74  # calibrated: reference bbox_x0=150.3, NotoSansCJKsc ascent=10.44 → x_insert=160.74 → param=159.74

# ── Per-page header (written on ALL pages) ─────────────────────────────────
PAGE_HEADER_RECTS = {
    "page_no":           ( 83.6,  36.0,  92.0,  52.0),   # x0 raised 80.5→83.6 so NotoSansCJKsc bbox_x0≈76.16≈template 76.14; y1 narrowed 54→52 to preserve 页码/页数: label
    "customs_no_top":    ( 83.0, 464.0,  89.5, 549.0),   # x0 was 82.1 (erased 海关编号 label); x1 was 92.1 (erased (9415) code at x=89.9)
    "customs_no_bottom": ( 83.0, 674.0,  92.1, 759.0),   # x0 was 82.1; raised to 83.0 to preserve 预录入编号 label (x1=82.9)
    "barcode_text":      ( 24.0,  74.0,  57.0, 270.0),
}

# ── Product item column x-positions (physical x, left edge of strip) ───────
# Verified against item_no bboxes extracted from namuna.pdf.
PAGE1_ITEM_X = [288.6, 320.9, 353.8, 386.0, 417.9, 450.2]

TEMPLATE_ITEM_X = [
    103.4, 136.7, 166.7, 200.0, 232.9, 266.2,
    295.9, 329.2, 361.7, 395.0, 424.7, 459.9,
    490.9, 524.2,
]

ITEM_STRIP_WIDTH = 33.0  # physical x-width of each item strip

# ── Sub-column x offsets within one item strip ─────────────────────────────
SC1_X = 12.14  # calibrated: reference SC1 bbox_x0 = ix+2.74, NotoSansCJKsc ascent=10.44
SC2_X = 22.04  # calibrated: reference SC2 bbox_x0 = ix+SC2_gap, same +2.74 shift
SC3_X = 33.14  # calibrated: reference SC3 bbox_x0 = ix+SC3_gap, same +2.74 shift

# ── Y-span bounds for strip clearing ──────────────────────────────────────
ITEM_Y_MIN =  43.0   # minimum y across all item fields (tax_code low end)
ITEM_Y_MAX = 806.0   # maximum y across all item fields (item_no high end)

# ── Y-ranges for each data field ──────────────────────────────────────────
# Each (y0, y1): text is inserted at y_start=y1, advances toward y0.
# y1 ≈ bbox_y_max + 1.5 pt  (small descender / baseline margin)
# y0 ≈ bbox_y_min - 1   pt
#
# ALL VALUES measured from namuna.pdf (ground truth):
#
#   FIELD        | MEASURED BBOX (y range)  | (y0,   y1 )
#   -------------|--------------------------|-------------
#   item_no      | 800.0 – 804.5            | (800,  806)
#   hs_code      | 744.0 – 781.5  ← SC1    | (743,  783)
#   ch_name      | 687.0 – 742.0  ← SC1    | (684,  743)
#   qty          | 439.0 – 461.5  ← SC1,SC3| (437,  463)
#   unit_price   | 353.1 – 380.1            | (351,  382)
#   origin_cn    | 275.1 – 293.1            | (273,  295)
#   dest_cn      | 198.3 – 252.3            | (196,  254)
#   source_tax   | 44.9  – 150.8  (combined)| (43,   152)
#   brand_row    | 658.0 – 734.5            | (655,  736)
#   total_price  | 357.6 – 384.6            | (355,  386)
#   origin_code  | 275.1 – 297.6  ← SC2    | (273,  299)
#   dest_code    | 198.3 – 220.8  ← SC2    | (196,  222)
#   tax_code     | 44.9  – 58.4   ← SC2    | (43,    60)
#   qty_repeat   | 439.0 – 461.5  ← SC3    | (437,  463)
#   currency     | 344.1 – 362.1  ← SC3    | (342,  363)
ITEM_Y = {
    # ─── Row 1 (SC1) ────────────────────────────────────────────────────────
    "item_no":     (800.0, 804.5),  # template: y=(800.0,804.5)
    "hs_code":     (736.5, 786.5),  # template: y_hi=781.5; +5pt for NotoSansCJKsc wider digits (50pt vs 45pt)
    "ch_name":     (687.0, 732.0),  # template: y=(687.0,732.0)
    "qty":         (439.0, 461.5),  # template: y=(439.0,461.5)
    "unit_price":  (353.0, 380.0),  # template: y=(353.1,380.1)
    "origin_cn":   (275.0, 293.0),  # template: y=(275.1,293.1)
    "dest_cn":     (198.0, 252.0),  # template: y=(198.3,252.3)
    "source_tax":  ( 44.9, 150.8),  # template: y=(44.9,150.8)

    # ─── Row 2 (SC2) ────────────────────────────────────────────────────────
    "brand_row":   (658.0, 734.5),  # template: y=(658.0,734.5)
    "total_price": (357.5, 384.5),  # template: y=(357.6,384.6)
    "origin_code": (275.0, 297.5),  # template: y=(275.1,297.6)
    "dest_code":   (198.0, 220.5),  # template: y=(198.3,220.8)
    "tax_code":    ( 44.9,  58.4),  # template: y=(44.9,58.4)

    # ─── Row 3 (SC3) ────────────────────────────────────────────────────────
    "qty_repeat":  (439.0, 461.5),  # template: y=(439.0,461.5)
    "currency":    (344.0, 362.0),  # template: y=(344.1,362.1)
}

# ── Static values written into every occupied item slot ───────────────────
STATIC_ITEM = {
    "brand_row":   "0|0|无品牌|无型号",
    "origin_cn":   "中国",
    "origin_code": "(CHN)",
    "dest_cn":     "乌兹别克斯坦",
    "dest_code":   "(UZB)",
    "source_tax":  "(33019)杭州其他 照章征税",
    "tax_code":    "(1)",
    "currency":    "美元",
}
