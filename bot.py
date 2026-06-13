"""
bot.py — Telegram bot: Excel (.xlsx) → Customs declaration PDF

Ishga tushirish:
  python bot.py

Talablar:
  .env faylida BOT_TOKEN o'rnatilgan bo'lishi kerak.
"""

from __future__ import annotations

import io
import logging
import os
import tempfile
from dataclasses import dataclass, field

from dotenv import load_dotenv
from telegram import Update, Document
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

load_dotenv()

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)

# ── Foydalanuvchi holati ───────────────────────────────────────────────────
@dataclass
class UserState:
    entry_no: str = ""       # 18 raqamli bojxona raqami (ixtiyoriy)
    date: str     = ""       # YYYYMMDD formatida sana (ixtiyoriy)
    pending_files: list[str] = field(default_factory=list)

_user_states: dict[int, UserState] = {}

def _state(user_id: int) -> UserState:
    if user_id not in _user_states:
        _user_states[user_id] = UserState()
    return _user_states[user_id]


# ── Komandalar ─────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "👋 *Bojxona deklaratsiyasi boti*\n\n"
        "Ishlatish:\n"
        "1️⃣ `.xlsx` faylini yuboring → PDF qaytariladi\n\n"
        "Ixtiyoriy sozlamalar:\n"
        "• `/entry 941520260000090882` — bojxona raqamini o'rnating\n"
        "• `/date 20260317` — sanani o'rnating (YYYYMMDD)\n"
        "• `/settings` — joriy sozlamalarni ko'rish\n"
        "• `/reset` — sozlamalarni tozalash\n\n"
        "ℹ️ Sozlamalar barcha keyingi fayllar uchun saqlanadi.",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "📋 *Komandalar ro'yxati:*\n\n"
        "`/start` — botni ishga tushirish\n"
        "`/entry <raqam>` — 18 raqamli bojxona raqamini o'rnating\n"
        "`/date <YYYYMMDD>` — deklaratsiya sanasini o'rnating\n"
        "`/settings` — joriy sozlamalarni ko'rish\n"
        "`/reset` — barcha sozlamalarni tozalash\n\n"
        "📎 Shunchaki `.xlsx` faylini yuboring — PDF tayyor bo'ladi!",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_entry(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    if not ctx.args:
        await update.message.reply_text(
            "❌ Raqam kiriting.\nMisol: `/entry 941520260000090882`",
            parse_mode=ParseMode.MARKDOWN,
        )
        return
    no = ctx.args[0].strip()
    if not no.isdigit():
        await update.message.reply_text("❌ Faqat raqamlardan iborat bo'lishi kerak.")
        return
    _state(user_id).entry_no = no
    await update.message.reply_text(f"✅ Bojxona raqami saqlandi: `{no}`", parse_mode=ParseMode.MARKDOWN)


async def cmd_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    if not ctx.args:
        await update.message.reply_text(
            "❌ Sana kiriting.\nMisol: `/date 20260317`",
            parse_mode=ParseMode.MARKDOWN,
        )
        return
    d = ctx.args[0].strip()
    if len(d) != 8 or not d.isdigit():
        await update.message.reply_text("❌ Format noto'g'ri. YYYYMMDD formatida kiriting.\nMisol: `20260317`", parse_mode=ParseMode.MARKDOWN)
        return
    _state(user_id).date = d
    await update.message.reply_text(f"✅ Sana saqlandi: `{d}`", parse_mode=ParseMode.MARKDOWN)


async def cmd_settings(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    st = _state(user_id)
    entry_str = f"`{st.entry_no}`" if st.entry_no else "_o'rnatilmagan_"
    date_str  = f"`{st.date}`"     if st.date      else "_o'rnatilmagan_"
    await update.message.reply_text(
        f"⚙️ *Joriy sozlamalar:*\n\n"
        f"• Bojxona raqami: {entry_str}\n"
        f"• Sana: {date_str}",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_reset(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    _user_states[user_id] = UserState()
    await update.message.reply_text("🔄 Barcha sozlamalar tozalandi.")


# ── .xlsx fayl qabul qilish ────────────────────────────────────────────────
async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    doc: Document = update.message.document
    user_id = update.effective_user.id

    # Faqat .xlsx qabul qilinadi
    fname = doc.file_name or ""
    if not fname.lower().endswith(".xlsx"):
        await update.message.reply_text(
            "⚠️ Faqat `.xlsx` formatidagi fayllar qabul qilinadi.",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    status_msg = await update.message.reply_text("⏳ Fayl yuklanmoqda…")

    with tempfile.TemporaryDirectory() as tmp_dir:
        xlsx_path = os.path.join(tmp_dir, fname)
        tg_file = await ctx.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(xlsx_path)

        await status_msg.edit_text("⚙️ Excel fayl tahlil qilinmoqda…")

        # ── Excel tahlili ─────────────────────────────────────────────────
        try:
            from excel_parser import ExcelParser
            data = ExcelParser().parse(xlsx_path)
        except Exception as exc:
            log.error("Excel parse error: %s", exc, exc_info=True)
            await status_msg.edit_text(f"❌ Excel faylni o'qishda xato:\n`{exc}`", parse_mode=ParseMode.MARKDOWN)
            return

        if not data.items:
            await status_msg.edit_text(
                "❌ Faylda mahsulot topilmadi.\n"
                "Iltimos, to'g'ri invoice/packing list formatini tekshiring."
            )
            return

        await status_msg.edit_text(
            f"🔄 PDF yaratilmoqda…\n"
            f"📦 Mahsulotlar: {len(data.items)} ta\n"
            f"📋 Kontrakt: `{data.contract_no}`",
            parse_mode=ParseMode.MARKDOWN,
        )

        # ── PDF yaratish ──────────────────────────────────────────────────
        st = _state(user_id)
        try:
            from pdf_editor import PDFEditor
            from translation import write_missing_translations_report, clear_missing_translations
            clear_missing_translations()
            pdf_bytes = PDFEditor().edit(
                data,
                pre_entry_no=st.entry_no,
                declaration_date=st.date,
            )
        except Exception as exc:
            log.error("PDF generation error: %s", exc, exc_info=True)
            await status_msg.edit_text(f"❌ PDF yaratishda xato:\n`{exc}`", parse_mode=ParseMode.MARKDOWN)
            return

        # ── Missing translations hisoboti ─────────────────────────────────
        report_path = os.path.join(tmp_dir, "missing_translations.txt")
        write_missing_translations_report(report_path)

        # ── PDF yuborish ──────────────────────────────────────────────────
        safe_contract = data.contract_no.replace("/", "-").replace("\\", "-") or "declaration"
        pdf_name = f"declaration_{safe_contract}.pdf"

        await status_msg.delete()
        await update.message.reply_document(
            document=io.BytesIO(pdf_bytes),
            filename=pdf_name,
            caption=(
                f"✅ *PDF tayyor!*\n\n"
                f"📋 Kontrakt: `{data.contract_no}`\n"
                f"🏢 Oluvchi: {data.consignee or '—'}\n"
                f"📦 Mahsulotlar: {len(data.items)} ta\n"
                f"📦 Joylar: {data.packages} ta\n"
                f"⚖️ Brutto: {data.gross_weight:.1f} kg\n"
                f"💰 Jami: ${data.total_amount_usd:.2f}"
            ),
            parse_mode=ParseMode.MARKDOWN,
        )

        # Agar tarjima qilinmagan qiymatlar bo'lsa — hisobotni ham yuborish
        if os.path.exists(report_path) and os.path.getsize(report_path) > 0:
            with open(report_path, "rb") as rf:
                await update.message.reply_document(
                    document=rf,
                    filename="missing_translations.txt",
                    caption="⚠️ Quyidagi mahsulotlar tarjima qilinmadi (lug'atga qo'shish tavsiya etiladi).",
                )


# ── Noma'lum xabarlar ──────────────────────────────────────────────────────
async def handle_unknown(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "ℹ️ `.xlsx` faylini yuboring yoki `/help` ni bosing.",
        parse_mode=ParseMode.MARKDOWN,
    )


# ── Xatolarni ushlab qolish ────────────────────────────────────────────────
async def error_handler(update: object, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    log.error("Xato yuz berdi: %s", ctx.error, exc_info=ctx.error)
    if isinstance(update, Update) and update.message:
        await update.message.reply_text("❌ Ichki xato yuz berdi. Iltimos, qayta urinib ko'ring.")


# ── Asosiy qism ────────────────────────────────────────────────────────────
def main() -> None:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError(
            "BOT_TOKEN topilmadi!\n"
            ".env faylini yarating va BOT_TOKEN=<tokeningiz> qatorini qo'shing.\n"
            "Token olish: https://t.me/BotFather"
        )

    app = (
        Application.builder()
        .token(token)
        .build()
    )

    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("help",     cmd_help))
    app.add_handler(CommandHandler("entry",    cmd_entry))
    app.add_handler(CommandHandler("date",     cmd_date))
    app.add_handler(CommandHandler("settings", cmd_settings))
    app.add_handler(CommandHandler("reset",    cmd_reset))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_unknown))
    app.add_error_handler(error_handler)

    log.info("Bot ishga tushdi. To'xtatish uchun Ctrl+C bosing.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
