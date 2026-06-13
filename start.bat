@echo off
chcp 65001 >nul
echo ========================================
echo   Bojxona deklaratsiyasi boti
echo ========================================
echo.

:: Virtual muhit mavjudligini tekshirish
if exist "venv\Scripts\activate.bat" (
    echo [*] Virtual muhit topildi, faollashtirilmoqda...
    call venv\Scripts\activate.bat
) else (
    echo [*] Virtual muhit yaratilmoqda...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo [*] Paketlar o'rnatilmoqda...
    pip install -r requirements.txt
)

echo.
echo [*] Bot ishga tushmoqda...
echo [*] To'xtatish uchun Ctrl+C bosing
echo.
python bot.py

pause
