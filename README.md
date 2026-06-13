# 🤖 Bojxona deklaratsiyasi boti

Excel (.xlsx) fayldan Xitoy bojxona deklaratsiyasi (PDF) yaratuvchi Telegram bot.

---

## ⚡ Tez ishga tushirish

### 1. Python o'rnatish
Python 3.10 yoki undan yuqori versiya kerak.  
Yuklab olish: https://www.python.org/downloads/

Tekshirish:
```
python --version
```

---

### 2. Loyihani tayyorlash

ZIP faylni ochib, papkaga kiring:
```
cd customs_bot
```

Virtual muhit yarating (tavsiya etiladi):
```
# Windows:
python -m venv venv
venv\Scripts\activate

# Mac / Linux:
python -m venv venv
source venv/bin/activate
```

Paketlarni o'rnating:
```
pip install -r requirements.txt
```

---

### 3. Bot tokenini olish

1. Telegramda [@BotFather](https://t.me/BotFather) ga yozing
2. `/newbot` komandasini yuboring
3. Bot nomini kiriting (masalan: `MyCustomsBot`)
4. Username kiriting (masalan: `my_customs_bot`)
5. BotFather sizga token beradi:  
   `1234567890:ABCDEFGHijklmnopqrstuvwxyz123456789`

---

### 4. `.env` faylini sozlash

`.env.example` faylini `.env` deb ko'chiring:
```
# Windows:
copy .env.example .env

# Mac / Linux:
cp .env.example .env
```

`.env` faylini matn muharririda oching va tokenni kiriting:
```
BOT_TOKEN=1234567890:ABCDEFGHijklmnopqrstuvwxyz123456789
```

---

### 5. Windows uchun SimSun shrift

PDF da Xitoy harflari to'g'ri chiqishi uchun **SimSun** shrift kerak.  
Windows da bu shrift allaqachon mavjud (`C:\Windows\Fonts\simsun.ttc`).

**Mac / Linux** da `fonts/` papkasiga `simsun.ttc` ni qo'ying.  
Yoki bepul NotoSansCJKsc shrifti avtomatik yuklanadi.

---

### 6. Botni ishga tushirish

```
python bot.py
```

Muvaffaqiyatli ishga tushsa quyidagi xabar chiqadi:
```
INFO: Bot ishga tushdi. To'xtatish uchun Ctrl+C bosing.
```

**Botni to'xtatish:** `Ctrl + C`

---

## 📱 Botdan foydalanish

| Komanda | Tavsif |
|---------|--------|
| `/start` | Botni ishga tushirish |
| `/entry 941520260000090882` | Bojxona raqamini o'rnating (18 raqam) |
| `/date 20260317` | Deklaratsiya sanasini o'rnating (YYYYMMDD) |
| `/settings` | Joriy sozlamalarni ko'rish |
| `/reset` | Sozlamalarni tozalash |
| `.xlsx fayl yuborish` | PDF yaratish |

---

## 📁 Loyiha tuzilishi

```
customs_bot/
├── bot.py              ← Telegram bot (asosiy ishga tushirish fayli)
├── main.py             ← CLI orqali ishlatish (terminal)
├── excel_parser.py     ← Excel faylni tahlil qilish
├── pdf_editor.py       ← PDF yaratish
├── translation.py      ← Tarjima tizimi (RU → CN)
├── config.py           ← PDF koordinatalari va sozlamalar
├── requirements.txt    ← Python paketlari
├── .env.example        ← Token namunasi
├── .env                ← Sizning tokeningiz (git ga KIRITMANG!)
├── template/
│   └── namuna.pdf      ← PDF shablon
├── fonts/              ← SimSun shriftini bu yerga qo'ying (Mac/Linux)
└── output/             ← Yaratilgan PDF fayllar (avtomatik yaratiladi)
```

---

## 🔧 Muammolar va yechimlar

**"No module named 'pymupdf'"**  
→ `pip install -r requirements.txt` ni qayta ishga tushiring

**"BOT_TOKEN topilmadi"**  
→ `.env` fayli mavjudligini va `BOT_TOKEN=...` qatorini tekshiring

**PDF da harflar chiqmayapti**  
→ `fonts/` papkasiga `simsun.ttc` ni qo'ying (Mac/Linux uchun)

**"getaddrinfo failed"**  
→ Internet aloqasini tekshiring; bot Telegram serverlariga ulanishi kerak

---

## 🖥️ CLI orqali ishlatish (bot siz)

```
python main.py manba.xlsx
python main.py manba.xlsx --entry-no 941520260000090882 --date 20260317
python main.py manba.xlsx --output result.pdf
```
