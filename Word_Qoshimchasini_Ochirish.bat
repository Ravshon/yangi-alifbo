@echo off
chcp 65001 > nul
echo ========================================================
echo   Microsoft Word "Yangi Alifbo" Qo'shimchasini O'chirish
echo ========================================================
echo.
set "TARGET=%APPDATA%\Microsoft\Word\STARTUP\YangiAlifbo.dotm"

if exist "%TARGET%" (
    del /f /q "%TARGET%"
    echo [Bajarildi] Microsoft Word qo'shimchasi muvaffaqiyatli o'chirildi.
) else (
    echo [Ma'lumot] Qo'shimcha topilmadi yoki allaqachon o'chirilgan.
)
echo.
pause
