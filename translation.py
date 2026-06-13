"""
ChineseTranslation — maps (HS code, Russian product description) → Chinese
customs name.

Lookup stages
─────────────
1. (hs_code, keyword) — normalized keyword contained in normalized Russian
   description (longest match wins)
2. hs_code only       — default name for the code
3. Generic fallback   — online translation via deep-translator, or "商品"

Normalization
─────────────
All lookups use normalize_text() before comparison:
  - Unicode NFC normalization
  - Lowercase
  - Remove punctuation (.,;:!?«»"'()[]{}-)
  - Collapse whitespace to single spaces
  - Strip leading/trailing whitespace

This ensures "Контактор", "контактор,", "Контактор AC", "контактор  "
all match the keyword "контактор".

HS code normalization
─────────────────────
HS codes are stripped of non-digits and padded/trimmed to 10 digits.
If the raw Excel HS code has fewer digits (e.g. 9), it is padded with a
trailing zero so "853649000" → "8536490000" and matches the dictionary.

Missing translations
────────────────────
Every value that reaches Stage 3 (online/fallback) is recorded internally.
Call write_missing_translations_report(path) at end of processing to write
missing_translations.txt with all untranslated values.

HS codes covered
────────────────
Taken from namuna.pdf (reference) and declaration_DL0905.pdf (test output),
plus codes found in declaration_ATAF1305.pdf and declaration_DBT3004.pdf:
  8504312909  电流互感器           Current Transformers
  8504409100  变频器               Frequency Inverters / VFDs
  8535210000  真空断路器           Vacuum Circuit Breakers
  8536201007  断路器               MCBs / Moulded-Case Breakers
  8536209007  断路器/电动驱动       ACB variants / Motor Drives
  8536490000  接触器/继电器         Contactors / Relays
  8536508000  负荷隔离开关         Load Isolation Switches
  8536901000  分支夹               Branch Clamps / Connectors
  8546901000  绝缘件               Insulating Parts
  8526920000  无线电控制套件       Radio Control Kits
  9030339100  数字面板电压表       Digital Panel Voltmeters
  9107000000  时间开关             Time Switches
  9403700008  电气密封面板         Sealed Electrical Panels
  3925902000  穿孔线槽             Perforated Cable Ducts
  8302100000  铰链                 Hinges / Hardware
  7308909809  钢铁结构件           Steel Structural Parts
  7326909807  其他钢铁制品         Other Steel Products

  --- From declaration_ATAF1305.pdf ---
  3924100000  塑料餐具             Plastic Tableware / Kitchen Goods
  3923100009  塑料箱               Plastic Boxes / Cases
  3923301090  塑料瓶               Plastic Bottles / Flasks
  7323930000  铁制厨房用具         Iron/Steel Kitchen Utensils
  6702900000  人造花卉             Artificial Flowers / Decorative Plants
  9505900000  节日用品             Festive / Seasonal Articles
  3406000000  蜡烛                 Candles
  4823908597  纸制品               Paper Articles / Products

  --- From declaration_DBT3004.pdf ---
  8421310009  空气过滤器           Air Filters for Internal Combustion Engines
"""

from __future__ import annotations

import logging
import re
import unicodedata
from typing import Optional

log = logging.getLogger(__name__)

# ── Missing-translation tracking ──────────────────────────────────────────
# Populated by _record_missing_translation(); flushed by
# write_missing_translations_report().
_missing_translations: list[dict] = []


# ── Text normalization ────────────────────────────────────────────────────

def normalize_text(text: str) -> str:
    """
    Normalize Russian/Latin text for fuzzy dictionary matching.

    Steps applied in order:
      1. Unicode NFC normalization  (handles decomposed Cyrillic, etc.)
      2. Lowercase
      3. Replace punctuation and dashes with spaces
         (comma, period, colon, semicolon, !, ?, guillemets, ASCII quotes,
          parentheses, brackets, braces, hyphen, en-dash, em-dash)
      4. Collapse multiple spaces to one
      5. Strip leading/trailing whitespace

    Examples:
      "Контактор"   → "контактор"
      "контактор,"  → "контактор"
      "Контактор "  → "контактор"
      "контактор,"  → "контактор"
      "Контактор AC"→ "контактор ac"
      " реле  тока" → "реле тока"
    """
    if not text:
        return ""
    text = unicodedata.normalize("NFC", text)
    text = text.lower()
    text = re.sub(r'[.,;:!?«»"""\'()\[\]{}\-–—]+', " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ── HS code normalization ──────────────────────────────────────────────────

def normalize_hs(hs_code: str | int) -> str:
    """
    Normalize an HS code to a canonical 10-digit string.

    Strips all non-digit characters, then:
      - If result is < 10 digits, pad with trailing zeros → 10 digits.
      - If result is > 10 digits, truncate to 10 digits.

    Examples:
      "8536490000"  → "8536490000"
      " 8536490000" → "8536490000"
      "853649000"   → "8536490000"   (9-digit → padded)
      "8536.49.0000"→ "8536490000"
    """
    digits = re.sub(r"\D", "", str(hs_code).strip())
    if len(digits) < 10:
        digits = digits.ljust(10, "0")
    return digits[:10]


# ── (hs_code, lowercase_keyword) → Chinese customs name ───────────────────
# Keywords are stored normalized (already lowercase, no extra punctuation).
# The lookup also normalizes the runtime Russian description before comparison.
_KW_MAP: dict[tuple[str, str], str] = {

    # ── 8504312909 Current Transformers ───────────────────────────────────
    ("8504312909", "трансформатор тока"):       "电流互感器",
    ("8504312909", "тт"):                       "电流互感器",

    # ── 8504409100 Frequency Inverters / VFDs / Inverters ────────────────
    # 变频器 = Variable Frequency Drive (VFD) / Frequency converter
    # 逆变器 = Inverter (DC→AC), e.g. DELIXI INVERTER products
    ("8504409100", "преобразователь частоты"):  "变频器",
    ("8504409100", "частотный"):                "变频器",
    ("8504409100", "частотн"):                  "变频器",
    ("8504409100", "vfd"):                      "变频器",
    ("8504409100", "инвертор"):                 "逆变器",
    ("8504409100", "inverter"):                 "逆变器",

    # ── 8535210000 Vacuum Circuit Breakers ────────────────────────────────
    ("8535210000", "вакуумный выключатель"):    "真空断路器",
    ("8535210000", "вакуумн"):                  "真空断路器",

    # ── 8536201007 Standard MCBs ──────────────────────────────────────────
    ("8536201007", "автоматический выключатель"): "断路器",
    ("8536201007", "автомат"):                  "断路器",
    ("8536201007", "nmm"):                      "断路器",
    ("8536201007", "mcb"):                      "断路器",

    # ── 8536209007 ACB variants / Motor Drives ────────────────────────────
    ("8536209007", "электропривод"):            "电动驱动",
    ("8536209007", "электро привод"):           "电动驱动",
    ("8536209007", "привод"):                   "电动驱动",
    ("8536209007", "4p avr"):                   "自动断路器",
    ("8536209007", "avr"):                      "自动断路器",
    ("8536209007", "авр"):                      "自动断路器",
    ("8536209007", "автоматический выключатель"): "断路器",
    ("8536209007", "автомат"):                  "断路器",
    ("8536209007", "nmm"):                      "断路器",

    # ── 8536490000 Contactors / Relays ────────────────────────────────────
    ("8536490000", "контактор"):                "接触器",
    ("8536490000", "реле"):                     "继电器",
    ("8536490000", "промежуточн"):              "继电器",
    ("8536490000", "тепловое"):                 "热继电器",
    ("8536490000", "тепловой"):                 "热继电器",

    # ── 8536508000 Load Isolation Switches ────────────────────────────────
    ("8536508000", "рубильник"):                "开关负载隔离开关",
    ("8536508000", "переключател"):             "开关负载隔离开关",
    ("8536508000", "пакетный"):                 "开关负载隔离开关",
    ("8536508000", "выключател"):               "开关负载隔离开关",
    ("8536508000", "разъединител"):             "开关负载隔离开关",

    # ── 8536901000 Branch Clamps / Inverter Spare Parts ──────────────────
    # Spare-parts keywords (逆变器板备件) take priority over branch-clamp ones.
    ("8536901000", "запасные части для инвертора"): "逆变器板备件",
    ("8536901000", "запчасти для инвертора"):        "逆变器板备件",
    ("8536901000", "комплектующие для инвертора"):   "逆变器板备件",
    ("8536901000", "плата инвертора"):               "逆变器板备件",
    ("8536901000", "инверторная плата"):             "逆变器板备件",
    ("8536901000", "запасные части к инвертору"):    "逆变器板备件",
    ("8536901000", "запчасти к инвертору"):          "逆变器板备件",
    ("8536901000", "spare parts"):                   "逆变器板备件",
    ("8536901000", "запасные части"):                "配件",
    ("8536901000", "запчасти"):                      "配件",
    ("8536901000", "комплектующие"):                 "配件",
    ("8536901000", "плата"):                         "电路板",
    ("8536901000", "зажим"):                         "分支夹",
    ("8536901000", "клемм"):                         "分支夹",
    ("8536901000", "соединител"):                    "分支夹",

    # ── 8546901000 Insulating Parts ───────────────────────────────────────
    ("8546901000", "изолятор"):                 "绝缘件",
    ("8546901000", "изолирующ"):                "绝缘件",
    ("8546901000", "изоляц"):                   "绝缘件",

    # ── 8526920000 Radio Control Kits ─────────────────────────────────────
    ("8526920000", "радиоуправления"):          "无线电控制套件",
    ("8526920000", "радиопульт"):               "无线电控制套件",
    ("8526920000", "радио"):                    "无线电控制套件",

    # ── 9030339100 Digital Panel Voltmeters ───────────────────────────────
    ("9030339100", "вольтметр"):                "数字面板电压表",
    ("9030339100", "амперметр"):                "数字面板电流表",
    ("9030339100", "панельный"):                "数字面板表",
    ("9030339100", "измерител"):                "数字面板表",

    # ── 9107000000 Time Switches ──────────────────────────────────────────
    ("9107000000", "переключатели"):            "时间开关",
    ("9107000000", "временные"):                "时间开关",
    ("9107000000", "таймер"):                   "时间开关",
    ("9107000000", "реле времени"):             "时间继电器",
    ("9107000000", "реле"):                     "时间继电器",

    # ── 9403700008 Sealed Electrical Panels ───────────────────────────────
    ("9403700008", "герметичный"):              "电气密封面板",
    ("9403700008", "щит"):                      "电气密封面板",
    ("9403700008", "панель"):                   "电气密封面板",
    ("9403700008", "шкаф"):                     "电气柜",
    ("9403700008", "ящик"):                     "电气箱",

    # ── 3925902000 Perforated Cable Ducts ─────────────────────────────────
    ("3925902000", "перфорированный"):          "穿孔线槽",
    ("3925902000", "перфориров"):               "穿孔线槽",
    ("3925902000", "короб"):                    "穿孔线槽",
    ("3925902000", "лоток"):                    "穿孔线槽",
    ("3925902000", "кабельный"):                "穿孔线槽",

    # ── 8302100000 Hinges / Hardware ──────────────────────────────────────
    ("8302100000", "петли"):                    "铰链",
    ("8302100000", "петля"):                    "铰链",
    ("8302100000", "шарнир"):                   "铰链",
    ("8302100000", "замок"):                    "锁",
    ("8302100000", "ручка"):                    "把手",
    ("8302100000", "фурнитур"):                 "铰链",

    # ── 7308909809 Steel Structural Parts ─────────────────────────────────
    ("7308909809", "профиль"):                  "钢铁结构件",
    ("7308909809", "стойка"):                   "钢铁结构件",
    ("7308909809", "стальн"):                   "钢铁结构件",
    ("7308909809", "рама"):                     "钢铁结构件",
    ("7308909809", "каркас"):                   "钢铁结构件",

    # ── 7326909807 Other Steel Products ───────────────────────────────────
    ("7326909807", "метиз"):                    "钢铁制品",
    ("7326909807", "болт"):                     "螺栓",
    ("7326909807", "гайка"):                    "螺母",
    ("7326909807", "шайба"):                    "垫圈",
    ("7326909807", "саморез"):                  "自攻螺钉",
    ("7326909807", "скоба"):                    "钢铁制品",
    ("7326909807", "крепеж"):                   "钢铁制品",
    ("7326909807", "винт"):                     "螺钉",

    # ── 3924100000 Plastic Tableware / Kitchen Goods ──────────────────────
    # Silicone kitchen tools (longest match wins → placed before generic entries)
    ("3924100000", "силиконовая лопатка"):      "硅胶刮刀",
    ("3924100000", "лопатк"):                   "硅胶刮刀",
    ("3924100000", "силиконовая кисточка"):     "硅胶刷",
    ("3924100000", "кисточк"):                  "硅胶刷",
    # Baking/chocolate molds
    ("3924100000", "форма для выпечки для шоколод"): "巧克力烘焙模具",
    ("3924100000", "шоколод"):                  "巧克力烘焙模具",   # misspelling of шоколад
    ("3924100000", "шоколад"):                  "巧克力模具",
    ("3924100000", "форма для выпечки"):        "烤盘",
    ("3924100000", "форма"):                    "烤盘",
    # Generic plastic tableware
    ("3924100000", "посуд"):                    "塑料餐具",
    ("3924100000", "тарелк"):                   "塑料餐具",
    ("3924100000", "чашк"):                     "塑料餐具",
    ("3924100000", "кружк"):                    "塑料кружка",
    ("3924100000", "миск"):                     "塑料餐具",
    ("3924100000", "пластик"):                  "塑料餐具",
    ("3924100000", "хозяйств"):                 "塑料餐具",
    ("3924100000", "товар"):                    "塑料餐具",
    ("3924100000", "лоток"):                    "塑料托盘",
    ("3924100000", "контейнер"):                "塑料容器",

    # ── 3923100009 Plastic Boxes / Cases ──────────────────────────────────
    # Cake packaging (longest match first)
    ("3923100009", "тара для упаковка торт"):   "蛋糕包装容器",
    ("3923100009", "торт"):                     "蛋糕包装容器",
    ("3923100009", "тара"):                     "塑料容器",
    ("3923100009", "короб"):                    "塑料箱",
    ("3923100009", "ящик"):                     "塑料箱",
    ("3923100009", "контейнер"):                "塑料容器",
    ("3923100009", "пластик"):                  "塑料箱",
    ("3923100009", "коробк"):                   "塑料箱",

    # ── 3923301090 Plastic Bottles / Flasks ───────────────────────────────
    ("3923301090", "бутылк"):                   "塑料瓶",
    ("3923301090", "фляг"):                     "塑料水壶",
    ("3923301090", "канистр"):                  "塑料桶",
    ("3923301090", "пластик"):                  "塑料瓶",
    ("3923301090", "флакон"):                   "塑料瓶",

    # ── 7323930000 Iron/Steel Kitchen Utensils ────────────────────────────
    # Metal baking molds (longest match first)
    ("7323930000", "форма для выпечки"):        "金属烤盘",
    ("7323930000", "форма"):                    "铁制模具",
    ("7323930000", "кастрюл"):                  "铁制炊具",
    ("7323930000", "сковород"):                 "铁制平底锅",
    ("7323930000", "ведр"):                     "金属桶",
    ("7323930000", "таз"):                      "铁制盆",
    ("7323930000", "кухн"):                     "铁制厨房用具",
    ("7323930000", "хозяйств"):                 "铁制厨房用具",
    ("7323930000", "металл"):                   "铁制厨房用具",

    # ── 6702900000 Artificial Flowers ─────────────────────────────────────
    ("6702900000", "искусств"):                 "人造花卉",
    ("6702900000", "цвет"):                     "人造花卉",
    ("6702900000", "декор"):                    "装饰花卉",
    ("6702900000", "пластик"):                  "人造花卉",
    ("6702900000", "растени"):                  "人造植物",

    # ── 9505900000 Festive Articles ───────────────────────────────────────
    # Carnival crown (longest match first)
    ("9505900000", "карновальных праздников"):  "嘉年华庆典产品、儿童皇冠",  # misspelling
    ("9505900000", "карновальн"):               "嘉年华庆典产品、儿童皇冠",  # misspelling
    ("9505900000", "карнавальн"):               "嘉年华庆典产品",
    ("9505900000", "корон"):                    "儿童皇冠",
    ("9505900000", "праздн"):                   "节日用品",
    ("9505900000", "игрушк"):                   "节日玩具",
    ("9505900000", "новогодн"):                 "节日用品",
    ("9505900000", "декор"):                    "节日装饰品",
    ("9505900000", "сувенир"):                  "节日用品",

    # ── 3406000000 Candles ────────────────────────────────────────────────
    # Twisted birthday/cake candles (longest match first)
    ("3406000000", "витые свечи для торта"):    "蛋糕用节日扭曲蜡烛",
    ("3406000000", "витые свечи"):              "节日扭曲蜡烛",
    ("3406000000", "витые"):                    "扭曲蜡烛",
    ("3406000000", "свеч"):                     "蜡烛",
    ("3406000000", "восков"):                   "蜡烛",

    # ── 4823908597 Paper Articles ─────────────────────────────────────────
    # Gift packaging filler (longest match first)
    ("4823908597", "наполнитель бумажный"):     "礼品包装纸填料",
    ("4823908597", "наполнитель"):              "纸制填料",
    ("4823908597", "бумаг"):                    "纸制品",
    ("4823908597", "картон"):                   "纸板制品",
    ("4823908597", "упаков"):                   "纸质包装",
    ("4823908597", "пакет"):                    "纸袋",

    # ── 8421310009 Air Filters ────────────────────────────────────────────
    ("8421310009", "фильтр"):                   "空气过滤器",
    ("8421310009", "воздушн"):                  "空气过滤器",
    ("8421310009", "воздух"):                   "空气过滤器",
    ("8421310009", "двигател"):                 "空气过滤器",
    ("8421310009", "масл"):                     "机油滤清器",
    ("8421310009", "топлив"):                   "燃油滤清器",
    ("8421310009", "очистител"):                "空气过滤器",
}

# ── hs_code → default name (when no keyword matched) ──────────────────────
_HS_DEFAULT: dict[str, str] = {
    "8504312909": "电流互感器",
    "8504409100": "变频器",
    "8535210000": "真空断路器",
    "8536201007": "断路器",
    "8536209007": "断路器",
    "8536490000": "接触器",
    "8536508000": "开关负载隔离开关",
    "8536901000": "分支夹",
    "8546901000": "绝缘件",
    "8526920000": "无线电控制套件",
    "9030339100": "数字面板电压表",
    "9107000000": "时间开关",
    "9403700008": "电气密封面板",
    "3925902000": "穿孔线槽",
    "8302100000": "铰链",
    "7308909809": "钢铁结构件",
    "7326909807": "钢铁制品",
    # From declaration_ATAF1305.pdf
    "3924100000": "塑料餐具",
    "3923100009": "塑料箱",
    "3923301090": "塑料瓶",
    "7323930000": "铁制厨房用具",
    "6702900000": "人造花卉",
    "9505900000": "节日用品",
    "3406000000": "蜡烛",
    "4823908597": "纸制品",
    # From declaration_DBT3004.pdf
    "8421310009": "空气过滤器",
}


# ── Missing-translation helpers ───────────────────────────────────────────

def _record_missing_translation(hs_code: str, russian_name: str) -> None:
    """Record a value that reached Stage-3 (no dictionary match found)."""
    entry = {"hs_code": hs_code, "russian_name": russian_name}
    if entry not in _missing_translations:
        _missing_translations.append(entry)
    log.warning(
        "UNTRANSLATED: hs=%s  ru=%r — falling back to online translation. "
        "Add this HS code/keyword pair to _KW_MAP or _HS_DEFAULT.",
        hs_code, russian_name[:60],
    )


def write_missing_translations_report(path: str = "missing_translations.txt") -> None:
    """
    Write every value that could not be translated via dictionary to *path*.

    Called at the end of the PDF generation pipeline.  If there are no
    missing translations, the file is not written (or is cleared if it
    already exists from a previous run).
    """
    if not _missing_translations:
        log.info("All values were translated via dictionary — no missing_translations.txt written.")
        return

    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("Missing Translations Report\n")
            f.write("=" * 60 + "\n")
            f.write(
                "These values were NOT matched in _KW_MAP or _HS_DEFAULT\n"
                "and fell back to online translation or 商品.\n"
                "Add each HS code + keyword pair to translation.py.\n\n"
            )
            for idx, item in enumerate(_missing_translations, 1):
                f.write(
                    f"[{idx}]\n"
                    f"  HS code:     {item['hs_code']}\n"
                    f"  Russian name: {item['russian_name']}\n\n"
                )
        log.warning(
            "missing_translations.txt written: %d untranslated value(s) → %s",
            len(_missing_translations), path,
        )
    except OSError as exc:
        log.error("Could not write missing_translations.txt: %s", exc)


def clear_missing_translations() -> None:
    """Reset the missing-translation log (useful between multiple parse runs)."""
    _missing_translations.clear()


# ── Online fallback translation ────────────────────────────────────────────

def _translate_to_simplified_chinese(text: str) -> str:
    """
    Translate `text` to Simplified Chinese (zh-CN) using deep-translator.
    Falls back to "商品" if the library is unavailable or the call fails.
    """
    if not text or not text.strip():
        return "商品"
    try:
        from deep_translator import GoogleTranslator
        result = GoogleTranslator(source="auto", target="zh-CN").translate(text.strip())
        return result if result else "商品"
    except Exception as exc:
        log.warning("Online translation failed for %r: %s — using 商品", text[:40], exc)
        return "商品"


def _extract_product_type(russian_name: str) -> str:
    """
    Extract only the product TYPE from a Russian product description,
    removing model numbers, voltage/power specs, and other technical details.

    Examples:
      "Инвертор ЕМ-60 1,5 220В"  → "Инвертор"
      "Магнитный пускатель BRW 40-11" → "Магнитный пускатель"
      "Реле времени электронное (таймер задержки)" → "Реле времени электронное"
      "Клеммный блок (клеммник) UKK" → "Клеммный блок"
      "Изолятор шинный ступенчатый CT4" → "Изолятор шинный ступенчатый"
      "Запчасти Инвертора Плата" → "Запчасти Инвертора Плата"
    """
    if not russian_name:
        return russian_name

    name = russian_name.strip()

    # Remove content in parentheses (usually alternative names or clarifications)
    name = re.sub(r'\s*\([^)]*\)', '', name)

    # Remove trailing model/serial numbers: sequences starting with uppercase
    # Latin letter(s) followed by digits, or pure alphanumeric codes
    # e.g. "BRW 40-11", "CT4", "UKK", "CL225-3", "ЕМ-60 1,5 220В"
    # Strategy: find the last Cyrillic word and cut everything after that
    # consists of Latin letters, digits, voltage specs, etc.

    # Split into tokens
    tokens = name.split()
    result_tokens = []

    for token in tokens:
        # If token is purely Latin alphanumeric (model number like BRW, UKK, CT4)
        if re.match(r'^[A-Za-z0-9\-.,/]+$', token):
            break
        # If token starts with Cyrillic letter followed by Latin (like ЕМ-60) — model
        if re.match(r'^[А-Яа-яЁё]+[-][A-Za-z0-9]', token):
            break
        # If token is a number with voltage/power suffix (like "220В", "380В", "1,5")
        if re.match(r'^[\d,.\-]+[А-Яа-яA-Za-z]*$', token) and any(c.isdigit() for c in token):
            # Check if this is part of product name (like "времени") or a spec
            if re.match(r'^[\d,.\-]+', token):
                break
        # If token is purely numeric
        if re.match(r'^[\d,.\-/]+$', token):
            break
        result_tokens.append(token)

    if result_tokens:
        return ' '.join(result_tokens)

    # If nothing was kept, return original name (minus parentheses)
    return name


# ── Main product-name lookup ──────────────────────────────────────────────

def get_chinese_name(hs_code: str | int, russian_name: str) -> str:
    """
    Return the Chinese customs product name for a given HS code + Russian name.

    Tarjima to'g'ridan-to'g'ri Google Translate orqali amalga oshiriladi
    (soddalashtirilgan xitoy tiliga — zh-CN). HS kodga bog'liq emas.

    Agar online tarjima ishlamasa, lug'atdan qidiriladi (fallback sifatida).
    """
    hs = normalize_hs(hs_code)
    ru_norm = normalize_text(russian_name)

    # ── Stage 1: Online translation (primary method) ──────────────────────
    # Tarjimadan oldin mahsulot nomini tozalash — raqamlar, model va
    # texnik parametrlarni olib tashlash, faqat mahsulot turini tarjima qilish.
    clean_name = _extract_product_type(russian_name)
    translated = _translate_to_simplified_chinese(clean_name)

    if translated and translated != "商品":
        log.debug(
            "Online translation: ru=%r → clean=%r → zh=%s",
            russian_name[:40], clean_name[:40], translated,
        )
        return translated

    # ── Stage 2 (fallback): keyword match ─────────────────────────────────
    best_match: Optional[str] = None
    best_len = 0
    for (code, kw), zh in _KW_MAP.items():
        if code != hs:
            continue
        kw_norm = normalize_text(kw)
        if kw_norm in ru_norm and len(kw_norm) > best_len:
            best_match = zh
            best_len   = len(kw_norm)

    if best_match:
        log.debug(
            "Fallback KW match: hs=%s ru=%r → %s",
            hs, russian_name[:40], best_match,
        )
        return best_match

    # ── Stage 3 (fallback): HS code default ───────────────────────────────
    if hs in _HS_DEFAULT:
        zh = _HS_DEFAULT[hs]
        log.debug("Fallback HS default: hs=%s → %s", hs, zh)
        return zh

    # ── Stage 4: record missing and return generic ────────────────────────
    _record_missing_translation(hs, russian_name)
    return "商品"


# ── Company-name conversion ────────────────────────────────────────────────

def get_chinese_company_name(name: str) -> str:
    """
    Convert a Latin-script Chinese company name to simplified Chinese.

    Applies three replacement passes (longest match first):
      1. Legal-form suffixes  (CO., LTD → 有限公司, …)
      2. City/province names  (YIWU → 义乌, XIAN → 西安, …)
      2b. Known company brand / proper-noun parts (DELIXI → 德力西, …)
      3. Business-sector terms (INTERNATIONAL TRADE → 国际贸易, …)

    Parts that do not match any rule are kept as-is (Latin script).
    Leading/trailing whitespace is stripped before processing.

    If the input is already Chinese (contains CJK characters), it is returned
    as-is without modification.
    """
    if not name:
        return name

    # If already contains Chinese, return as-is
    if re.search(r"[\u4e00-\u9fff]", name):
        return name

    _CITY_MAP = [
        ("YIWU",       "义乌"),
        ("HANGZHOU",   "杭州"),
        ("SHANGHAI",   "上海"),
        ("BEIJING",    "北京"),
        ("SHENZHEN",   "深圳"),
        ("GUANGZHOU",  "广州"),
        ("XI'AN",      "西安"),
        ("XIAN",       "西安"),
        ("NINGBO",     "宁波"),
        ("QINGDAO",    "青岛"),
        ("WENZHOU",    "温州"),
        ("CHENGDU",    "成都"),
        ("WUHAN",      "武汉"),
        ("NANJING",    "南京"),
        ("TIANJIN",    "天津"),
        ("SUZHOU",     "苏州"),
        ("DONGGUAN",   "东莞"),
        ("FOSHAN",     "佛山"),
        ("ZHONGSHAN",  "中山"),
        ("XIAMEN",     "厦门"),
        ("CHONGQING",  "重庆"),
        ("HEFEI",      "合肥"),
        ("ZHENGZHOU",  "郑州"),
        ("ZHUHAI",     "珠海"),
        ("CHANGSHA",   "长沙"),
        ("WUXI",       "无锡"),
        ("SHENYANG",   "沈阳"),
        ("JINAN",      "济南"),
        ("HARBIN",     "哈尔滨"),
        ("ZHEJIANG",   "浙江"),
        ("FUJIAN",     "福建"),
        ("SHANDONG",   "山东"),
        ("JIANGSU",    "江苏"),
        ("GUANGDONG",  "广东"),
        ("YUNNAN",     "云南"),
        ("SICHUAN",    "四川"),
        ("HUNAN",      "湖南"),
        ("HUBEI",      "湖北"),
        ("LIAONING",   "辽宁"),
        ("HEBEI",      "河北"),
        ("HENAN",      "河南"),
        ("SHAANXI",    "陕西"),
        ("SHANXI",     "山西"),
        ("ANHUI",      "安徽"),
        ("JIANGXI",    "江西"),
        ("GUIZHOU",    "贵州"),
        ("GANSU",      "甘肃"),
        ("QINGHAI",    "青海"),
        ("CHINA",      "中国"),
    ]

    _SECTOR_MAP = [
        ("IMPORT & EXPORT",         "进出口"),
        ("IMPORT AND EXPORT",       "进出口"),
        ("IMPORT&EXPORT",           "进出口"),
        ("INTERNATIONAL TRADE",     "国际贸易"),
        ("INTERNATIONAL COMMERCE",  "国际商贸"),
        ("INTERNATIONAL",           "国际"),
        ("SUPPLY CHAIN",            "供应链"),
        ("AUTO PARTS",              "汽车配件"),
        ("AUTO ACCESSORIES",        "汽车配件"),
        ("AUTOMOTIVE",              "汽车"),
        ("INVERTER",                "逆变器"),
        ("ELECTRONICS",             "电子"),
        ("ELECTRICAL",              "电气"),
        ("ELECTRIC",                "电气"),
        ("TECHNOLOGY",              "科技"),
        ("TECH",                    "科技"),
        ("MACHINERY",               "机械设备"),
        ("EQUIPMENT",               "设备"),
        ("INDUSTRIAL",              "实业"),
        ("INDUSTRY",                "实业"),
        ("CHEMICAL",                "化工"),
        ("PHARMACEUTICAL",          "制药"),
        ("FOOD",                    "食品"),
        ("TEXTILE",                 "纺织"),
        ("GARMENT",                 "服装"),
        ("CLOTHING",                "服装"),
        ("FURNITURE",               "家具"),
        ("HARDWARE",                "五金"),
        ("PLASTIC",                 "塑料"),
        ("RUBBER",                  "橡胶"),
        ("PRINTING",                "印刷"),
        ("PACKAGING",               "包装"),
        ("LOGISTICS",               "物流"),
        ("TRADING",                 "贸易"),
        ("TRADE",                   "贸易"),
        ("COMMERCIAL",              "商贸"),
        ("COMMERCE",                "商贸"),
        ("IMPORT",                  "进口"),
        ("EXPORT",                  "出口"),
        ("GENERAL",                 "综合"),
        ("HOLDING",                 "控股"),
        ("GROUP",                   "集团"),
    ]

    _SUFFIX_MAP = [
        ("CO., LTD.",   "有限公司"),
        ("CO., LTD",    "有限公司"),
        ("CO.,LTD.",    "有限公司"),
        ("CO.,LTD",     "有限公司"),
        ("CO. LTD.",    "有限公司"),
        ("CO. LTD",     "有限公司"),
        ("CO.LTD",      "有限公司"),
        ("CO LTD",      "有限公司"),
        ("LIMITED",     "有限公司"),
        ("LTD.",        "有限公司"),
        ("LTD",         "有限公司"),
        ("LLC",         "有限责任公司"),
        ("INC.",        "股份有限公司"),
        ("INC",         "股份有限公司"),
        ("CORP.",       "集团有限公司"),
        ("CORP",        "集团有限公司"),
    ]

    _COMPANY_PARTS_MAP = [
        ("BANGHENG",  "邦恒"),
        ("DELIXI",    "德力西"),
        ("CHINT",     "正泰"),
        ("TAIYUE",    "泰岳"),
        ("HOLLEY",    "和而泰"),
        ("SUPOR",     "苏泊尔"),
        ("HAIER",     "海尔"),
        ("HISENSE",   "海信"),
        ("MIDEA",     "美的"),
        ("GREE",      "格力"),
        ("HUAWEI",    "华为"),
        ("XIAOMI",    "小米"),
        ("LENOVO",    "联想"),
    ]

    result = name.strip().upper()

    # Remove parentheses but KEEP the content inside them (e.g. (HANGZHOU) → HANGZHOU)
    result = re.sub(r'[()]', '', result)
    result = re.sub(r'\s+', ' ', result).strip()

    # Pass 1 — legal-form suffix (strip from end first so it doesn't block sector words)
    for eng, chn in _SUFFIX_MAP:
        if result.endswith(eng):
            result = result[: -len(eng)].rstrip(". ,") + chn
            break

    # Pass 2 — city/province names (word-boundary match, longest first)
    for eng, chn in sorted(_CITY_MAP, key=lambda x: -len(x[0])):
        result = re.sub(r"(?<![A-Z])" + re.escape(eng) + r"(?![A-Z])", chn, result)

    # Pass 2b — known company brand / proper-noun parts (longest first)
    for eng, chn in sorted(_COMPANY_PARTS_MAP, key=lambda x: -len(x[0])):
        result = re.sub(r"(?<![A-Z])" + re.escape(eng) + r"(?![A-Z])", chn, result)

    # Pass 3 — sector terms (longest match first to avoid partial replacements)
    for eng, chn in sorted(_SECTOR_MAP, key=lambda x: -len(x[0])):
        result = re.sub(r"(?<![A-Z])" + re.escape(eng) + r"(?![A-Z])", chn, result)

    # Convert ASCII parentheses enclosing Chinese text to Chinese brackets.
    # e.g. (杭州) → （杭州）,  (杭州逆变器) → （杭州逆变器）
    result = re.sub(
        r"\(([^)]*[\u4e00-\u9fff][^)]*)\)",
        lambda m: "（" + m.group(1) + "）",
        result,
    )

    # Cleanup: remove spaces adjacent to Chinese characters.
    result = re.sub(r"\s+([\u4e00-\u9fff])", r"\1", result)
    result = re.sub(r"([\u4e00-\u9fff])\s+", r"\1", result)
    result = result.strip()

    return result


def list_known_hs_codes() -> list[str]:
    """Return sorted list of all recognised HS codes."""
    return sorted(set(_HS_DEFAULT.keys()))
